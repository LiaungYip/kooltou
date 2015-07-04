# -*- coding: utf-8 -*-

# Third-party libs from PyPI
import unicodedata      # for slugifying
import win32com.client  # part of pypiwin32
import pywintypes       # Part of pypiwin32
import pytz             # timezones info - for converting local time to UTC
import easygui
# Python standard libs
import re
import datetime
import os
import logging
import time


def set_up_logging ():
    # set up logging
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)
    # log all messages to file
    l1 = logging.FileHandler('log-%s.txt' % time.strftime("%Y-%m-%dT%H%M%S"))
    l1.setLevel(logging.DEBUG)
    # log warnings and errors to console
    l2 = logging.StreamHandler()
    l2.setLevel(logging.DEBUG)
    fm = logging.Formatter('%(levelname)s: %(message)s')
    for handler in [l1,l2]:
        handler.setFormatter(fm)
        root_logger.addHandler(handler)
    logging.info("Starting up....")

def get_subfolders( folder, f_list = list() ):
    # Return a flat list of Outlook `Folder` objects,
    # including all sub-folders of `folder` as well as `folder` itself.
    # Call as folder_list = get_subfolders( folder ).
    logging.debug("get_subfolders: %s (%s)" % (folder.Name, folder.FolderPath))
    f_list.append (folder)
    for f in folder.Folders:
        get_subfolders(f, f_list)
    return f_list

def slugify(value):
    """Normalise string to ASCII, remove strange characters, convert all
    whitespace characters (including tabs!!) to spaces."""
    # http://stackoverflow.com/questions/295135/turn-a-string-into-a-valid-filename-in-python
    value = unicodedata.normalize('NFKD', value).encode('ascii', 'ignore')
    value = unicode(re.sub('[^\w\s-]', '', value).strip())
    value = re.sub('[\s]+', ' ', value)  # Strip \t and other naughty things
    return value

def user_select_outlook_folder(ol_namespace):
    logging.info("Asking user to select Outlook folder...")
    while True:
        selected_folder = ol_namespace.PickFolder() # Brings up a GUI selection dialogue
        if selected_folder is not None:
            break
        if easygui.buttonbox("Failed to pick an Outook folder to back up. Try again, or quit?","Error",("Try again","Quit")) == "Quit":
            exit()
    logging.info("User selected the folder %s" % selected_folder.FolderPath)
    return selected_folder

def user_select_filesystem_dir():
    logging.info("Asking user to select filesystem directory...")
    while True:
        dir = easygui.diropenbox("Pick a folder to save the email .msg files to.","Choose backup destination")
        if dir is not None:
            break
        if easygui.buttonbox("Failed to pick a directory to save the emails. Try again, or quit?","Error",("Try again","Quit")) == "Quit":
            exit()
    logging.info("User selected the directory %s" % dir)
    return dir

def get_local_timezone():
    TZ = "Australia/Perth" #TODO: Make this configurable
    return pytz.timezone(TZ)
    logging.info("Using timezone %s" % TZ)

def make_directories(directory_path):
    if os.path.isdir (directory_path):
        logging.info("Folder %s already exists." % directory_path)
    else:
        logging.info("Folder %s does not exist. Attempting to create it." % directory_path)
        os.makedirs(directory_path)
        assert os.path.isdir (directory_path)
        logging.info("Created folder %s." % directory_path)

def get_mailitem_utc_time_string ( mailitem, local_timezone ):
    # mailitem is an Outlook.MailItem.
    # Times are returned as `pywintypes.Time` objects.
    # See http://timgolden.me.uk/python/win32_how_do_i/use-a-pytime-value.html.
    try:
        timestamp = int(mailitem.ReceivedTime) # Seconds since epoch
    except ValueError:
        # ReceivedTime is 01/01/1901 00:00:00 for certain program-generated
        # emails, i.e. from CDEGS license manager
        # Try using CreationTime instead
        timestamp = int(mailitem.CreationTime)
        logging.warning("MailItem didn't have a ReceivedTime. Using CreationTime instead. (Mail item: %s, CreationTime: %s)" % (subject, timestamp))
    tz_aware_time = datetime.datetime.fromtimestamp(timestamp,local_timezone)
    utc_time = tz_aware_time.astimezone(pytz.utc)
    # Like ISO format, but without the :'s (not OK for filenames) and
    # without the time-zone ID at the end.
    utc_time_string = utc_time.__format__("%Y-%m-%dT%H%M%SZ")
    return utc_time_string


set_up_logging()
logging.info("Opening Outlook Application...")
ol_application = win32com.client.Dispatch("Outlook.Application")
ol_namespace = ol_application.GetNamespace("MAPI") # Equivalent to ol_application.Session

selected_ol_folder = user_select_outlook_folder(ol_namespace)
ol_folder_list = get_subfolders(selected_ol_folder)
save_to_folder = user_select_filesystem_dir()
local_timezone = get_local_timezone()
num_processed = 0
num_saved = 0

for folder in ol_folder_list:
    ol_folder_path_str = folder.FolderPath # Example: \\Outlook Data File\J9 Administrivia\Expenses
    logging.info("Entering folder %s" % ol_folder_path_str)
    ol_folder_path_parts = ol_folder_path_str.split("\\")[1:] # [1:] to skip empty part due to "\\" at start
    ol_folder_path = os.sep.join( [slugify(p) for p in ol_folder_path_parts] ) # clean naughty characters

    dest_folder_path = os.path.join (save_to_folder, ol_folder_path)
    make_directories(dest_folder_path)

    for mi in folder.Items: #mi = MailItem
        if num_processed % 100 == 0:
            logging.info("... %i emails processed ..." % num_processed)
        num_processed += 1

        # "IPM.Note" is a normal email message.
        # "IPM.Note.EnterpriseVault.Shortcut" is an email message archived by Enterprise Vault.
        # other types include:
        # IPM.Appointment, IPM.Contact, IPM.Schedule.Meeting.Resp.Pos, and so on.
        if not mi.MessageClass.startswith("IPM.Note"):
            logging.debug("Ignoring item `%s` in folder `%s` of type `%s`" %
                          (mi.Subject, ol_folder_path_str, mi.MessageClass))
            continue

        raw_subject = mi.Subject
        subject = slugify(raw_subject[:100])
        if raw_subject != subject:
            logging.debug("Raw name `%s` was cleaned to `%s`" %
                          (raw_subject, subject))

        utc_time_string = get_mailitem_utc_time_string (mi, local_timezone)

        file_name = utc_time_string + ' ' + subject + ".MSG"
        file_path = os.path.join ( dest_folder_path , file_name )

        logging.debug("Trying to save %s" % file_path)
        try:
            mi.SaveAs ( file_path, 9 ) # Magic number 9 = olUnicodeMsg.
            num_saved += 1
            if num_saved % 50 == 0:
                logging.info("... %i emails saved to .msg ..." % num_saved)
        except pywintypes.com_error:
            logging.exception("Failure in MailItem.SaveAs().")
            logging.error("Details... MessageClass: `%s`, FolderPath: `%s`, Subject `%s`, ReceivedTime `%s`, CreationTime `%s`" %
                          (mi.MessageClass, ol_folder_path_str, raw_subject, mi.ReceivedTime, mi.CreationTime))


logging.info("DONE. Processed: %i, Saved: %i" % (num_processed, num_saved))