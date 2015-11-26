__author__ = 'layip'

import easygui
import logging
import os
import pytz

SETTINGS_FILE_NAME = "easygui_settings.dat"
class Settings(easygui.EgStore):
    def __init__(self, filename):
        logging.info("Initialising settings...")
        self.filename = filename
        self.timezone_name = ""
        self.debug_mode = ""
        self.days_old = -1
        self.dest_dir = ""
        self.mark_as_saved = True

        if os.path.isfile(self.filename):
            logging.info("Config file %s found. Attempting to load..."
                         % SETTINGS_FILE_NAME)
            self.restore()
        else:
            logging.info("No config file %s found. Asking user for settings..."
                         % SETTINGS_FILE_NAME)
            self.enter_settings()

        while True:
            logging.info("Validating settings...")
            if self.settings_valid() and self.user_confirm_settings():
                break
            logging.info("Settings invalid. Re-asking user for settings...")
            self.enter_settings()

        logging.info("Storing settings to file %s..." % SETTINGS_FILE_NAME)
        self.store()

    def settings_valid(self):
        try:
            assert type (self.timezone_name) is str
            assert self.timezone_name != ""
            assert type(self.debug_mode) is bool
            assert type (self.days_old) is int
            assert self.days_old >= 0
            assert self.dest_dir != ""
            assert os.path.isdir(self.dest_dir)
            assert type (self.mark_as_saved) is bool
            return True
        except (AttributeError, AssertionError):
            logging.exception("Setting not set yet, or invalid setting.")
            return False

    def enter_settings(self):
        s1 = "What is your local time zone? (i.e. Australia/Perth.) Used to determine the UTC/GMT time for file naming purposes."
        self.timezone_name = easygui.choicebox(s1,
                                               "Choose local time zone",
                                               pytz.common_timezones)

        s2 = "Enable debug mode? (Displays very detailed program execution information on screen.)"
        a = easygui.buttonbox(s2,"Debug mode?",("Normal mode (recommended)","Debug mode (experts only)"))
        self.debug_mode = True if a.lower().startswith("debug") else False

        s3 = "Save emails older than X days: (default 60 days; 0 for all emails regardless of age.)"
        self.days_old = easygui.integerbox(s3,"Email age?",default=60,
                                           lowerbound=0, upperbound=99999999)

        self.dest_dir = easygui.diropenbox("Pick a folder to save the email .msg files to.",
                                           "Choose backup destination",
                                           "%HOMEPATH%")

        b = easygui.ynbox("Apply 'Saved as MSG' tag to emails that are saved?",
                          title="Mark emails as saved?",
                          choices=("Yes (recommended)","No (experts only)"))
        self.mark_as_saved = bool(b)

    def user_confirm_settings(self):
        s1 = """Your settings are:
Time zone: `%s`
Debug mode: `%s`
Save emails older than `%i` days
Destination directory: `%s`
Mark emails as saved: `%s`

Use these settings, or select new settings?
"""
        s1 = s1 % ( self.timezone_name,
                    self.debug_mode,
                    self.days_old,
                    self.dest_dir,
                    self.mark_as_saved,
        )
        logging.info(s1)
        d = easygui.buttonbox(s1,"Use these settings?",("Use these settings","New settings"))
        logging.info("User input: %s" % d)
        return (d == "Use these settings")