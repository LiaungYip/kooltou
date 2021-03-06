# -*- coding: utf-8 -*-
__version__ = "v0.0.6 - 2015-10-02"


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

# ---------------------------------------------------------------------------- #
# Function definitions
# ---------------------------------------------------------------------------- #

def halt_catch_fire():
    raw_input("Fatal error. Can't continue.\nConsult 'log-...-summary.txt' for details of error(s).\nPress enter to exit...")
    exit()

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
        if easygui.buttonbox("Failed to pick an Outlook folder to back up. Try again, or quit?","Error",("Try again","Quit")) == "Quit":
            exit()
    logging.info("User selected the folder %s" % selected_folder.FolderPath)
    return selected_folder

def make_directories(directory_path):
    if os.path.isdir (directory_path):
        logging.info("Folder %s already exists." % directory_path)
    else:
        logging.info("Folder %s does not exist. Attempting to create it." % directory_path)
        try:
            os.makedirs(directory_path)
        except WindowsError:
            logging.exception("Do you have permission to write to this location?")
            halt_catch_fire()

        try:
            assert os.path.isdir (directory_path)
        except AssertionError:
            logging.exception("Failed to create directory.")
            halt_catch_fire()

        logging.info("Created folder %s." % directory_path)

def get_mailitem_utc_time ( mailitem, local_timezone ):
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
        subject = mailitem.Subject
        logging.warning("MailItem didn't have a ReceivedTime. Using CreationTime instead. (Mail item: %s, CreationTime: %s)" % (subject, timestamp))
    tz_aware_time = datetime.datetime.fromtimestamp(timestamp,local_timezone)
    utc_time = tz_aware_time.astimezone(pytz.utc)
    return utc_time

def get_mailitem_utc_time_string ( utc_time ):
    # Like ISO format, but without the :'s (not OK for filenames) and
    # without the time-zone ID at the end.
    return utc_time.__format__("%Y-%m-%dT%H%M%SZ")

def ol_category_exists ( ol_namespace, category_name ):
    return not (ol_namespace.Categories[category_name] is None)

def create_ol_category ( ol_namespace, category_name ):
    # Adds a new category. The category colour is the next un-used colour.
    # The category hot-key is not assigned.
    if not ol_category_exists( ol_namespace, category_name ):
        ol_namespace.Categories.Add (category_name)

def email_in_ol_category (mailitem, category_name):
    return (category_name in mailitem.Categories.split(", "))

def set_ol_category (mailitem, category_name):
    if True: #not email_in_ol_category(mailitem, category_name):
        mailitem.Categories = ", ".join((mailitem.Categories, category_name))
        mailitem.Save()

def unset_ol_category (mailitem, category_name):
    cats = str(mailitem.Categories).split(", ")
    mailitem.Categories = ", ".join( (c for c in cats if c != category_name) )

# ---------------------------------------------------------------------------- #
# User settings load/save
# ---------------------------------------------------------------------------- #

from settings_interface import SETTINGS_FILE_NAME, Settings

# ---------------------------------------------------------------------------- #
# Main script
# ---------------------------------------------------------------------------- #

#TODO: Pywin32 packaging.

# set up logging
root_logger = logging.getLogger()
root_logger.setLevel(logging.DEBUG)
# log all messages to file. 10 second delay to avoid nasty disk seeking.
l1 = logging.FileHandler('log-%s-detailed.txt' % time.strftime("%Y-%m-%dT%H%M%S"), encoding="utf-8", delay=10)
l1.setLevel(logging.DEBUG)
# log only info and above to a second file.
l2 = logging.FileHandler('log-%s-summary.txt' % time.strftime("%Y-%m-%dT%H%M%S"), encoding="utf-8", delay=10)
l2.setLevel(logging.INFO)
# log to console - if in debug mode, log everything. Otherwise, only INFO and above.
l3 = logging.StreamHandler()
l3.setLevel(logging.INFO)
fm = logging.Formatter('%(levelname)s: %(message)s')
for handler in [l1,l2,l3]:
    handler.setFormatter(fm)
    root_logger.addHandler(handler)
logging.info("Starting up....")
logging.info("kooltou version %s" % __version__)
logging.info("https://github.com/LiaungYip/kooltou")

# noinspection PyBroadException
try:
    settings = Settings(SETTINGS_FILE_NAME)

    if settings.debug_mode:
        l3.setLevel(logging.DEBUG)
        logging.debug("Setting logger `l3` to level DEBUG")

    logging.info("Opening Outlook Application...")
    ol_application = win32com.client.Dispatch("Outlook.Application")
    ol_namespace = ol_application.GetNamespace("MAPI") # Equivalent to ol_application.Session

    selected_ol_folder = user_select_outlook_folder(ol_namespace)
    save_to_folder = settings.dest_dir

    ol_folder_list = get_subfolders(selected_ol_folder)

    local_timezone = pytz.timezone(settings.timezone_name)
    now = datetime.datetime.now(pytz.utc)

    CATEGORY_NAME = "Saved as MSG"
    create_ol_category(ol_namespace,CATEGORY_NAME)

    num_already_archived = 0
    num_processed = 0
    num_saved = 0

    for folder in ol_folder_list:
        ol_folder_path_str = folder.FolderPath # Example: \\Outlook Data File\J9 Administrivia\Expenses
        logging.info("Entering folder %s" % ol_folder_path_str)
        ol_folder_path_parts = ol_folder_path_str.split("\\")[2:] # [2:] to skip empty parts due to "\\" at start
        ol_folder_path = os.sep.join( [slugify(p) for p in ol_folder_path_parts] ) # clean naughty characters

        dest_folder_path = os.path.join (save_to_folder, ol_folder_path)
        make_directories(dest_folder_path)

        # 255 characters max path length on windows. Less one character for trailing "\".
        max_filename_length = 255 - len(dest_folder_path) - 1

        for mi in folder.Items: #mi = MailItem
            num_processed += 1
            raw_subject = mi.Subject
            # "IPM.Note" is a normal email message.
            # "IPM.Note.EnterpriseVault.Shortcut" is an email message archived by Enterprise Vault.
            # other types include:
            # IPM.Appointment, IPM.Contact, IPM.Schedule.Meeting.Resp.Pos, and so on.
            if not mi.MessageClass.startswith("IPM.Note"):
                logging.debug("Ignoring item `%s` in folder `%s` of type `%s`" %
                              (raw_subject, ol_folder_path_str, mi.MessageClass))
                continue

            if email_in_ol_category(mi, CATEGORY_NAME):
                logging.debug("Item `%s` already archived (`%s`). Skipping." %
                              (raw_subject, CATEGORY_NAME))
                num_already_archived += 1
                continue

            utc_time = get_mailitem_utc_time(mi, local_timezone)
            delta = now - utc_time
            if delta.days < settings.days_old:
                logging.debug("Item `%s` of date `%s` is less than %i days old. Skipping." %
                              (raw_subject, utc_time, settings.days_old))
                continue
            else:
                logging.debug("Item `%s` of date `%s` is %i days old, archiving..." %
                              (raw_subject, utc_time, delta.days))

            utc_time_string = get_mailitem_utc_time_string (utc_time)

            # Calculate maximum permitted number of characters for subject, to stay within 255 chars for entire path.
            # Capped at 100 characters.
            max_subject_length = max_filename_length - len(utc_time_string) - len (" ")  - len(".MSG")
            subject_length = min(100, max_subject_length)

            subject = slugify(raw_subject[:subject_length])
            if raw_subject != subject:
                logging.debug("Raw name `%s` was cleaned to `%s`" %
                              (raw_subject, subject))

            file_name = utc_time_string + ' ' + subject + ".MSG"
            file_path = os.path.join ( dest_folder_path , file_name )

            logging.debug("Trying to save %s" % file_path)
            try:
                if len(file_path) > 255:
                    raise ValueError("File path exceeds 255 chars: %s" % file_path)
                mi.SaveAs ( file_path, 9 ) # Magic number 9 = olUnicodeMsg.
                if settings.mark_as_saved:
                    set_ol_category(mi, CATEGORY_NAME)
                num_saved += 1
            except pywintypes.com_error, ValueError:
                logging.exception("Failure in MailItem.SaveAs().")
                logging.error("Details... MessageClass: `%s`, FolderPath: `%s`, Subject `%s`, ReceivedTime `%s`, CreationTime `%s`" %
                              (mi.MessageClass, ol_folder_path_str, raw_subject, mi.ReceivedTime, mi.CreationTime))

            if num_saved % 50 == 0:
                logging.info("... %i emails saved to .msg ..." % num_saved)

            if num_processed % 100 == 0:
                logging.info("... %i emails processed ..." % num_processed)

    logging.info("DONE. Processed: %i, Saved: %i, Already archived (skipped): %i" %
                 (num_processed, num_saved, num_already_archived))

    raw_input("Press any key to exit...")

except Exception:
    logging.exception("Fatal error in main program.")
    halt_catch_fire()