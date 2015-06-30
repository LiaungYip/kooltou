__author__ = 'layip'

import win32com.client

mailbox = 'Mailbox - Yip, Li-Aung'
folderindex = 3 # Magic?

ol_application = win32com.client.Dispatch("Outlook.Application")
ol_namespace = ol_application.GetNamespace("MAPI") # Equivalent to ol_application.Session

# Folders object: https://msdn.microsoft.com/en-us/library/office/ff860950(v=office.14).aspx
# Supports methods Add, GetFirst, GetLast, GetNext, GetPrevious, Item, Remove.
# Item() can be used to get a folder by name, i.e. Folders.Item("GroupDiscussion").
# ol_folders = ol_namespace.Folders

selected_folder = ol_namespace.PickFolder()

if selected_folder is None:
    print ("Nothing!")
else:
    print selected_folder.Name

def get_subfolders( folder ):
    # Returns a list of folders under the given folder.
    for f in folder.Folders:
        print f.Name
        get_subfolders(f)