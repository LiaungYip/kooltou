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

def get_subfolders( folder, f_list ):
    # `folder` is an Outlook `Folder` object, and `f_list` is an empty python list().
    # At finish, f_list contains a flat list of Outlook `Folder` objects,
    # including all sub-folders of `folder` as well as `folder` itself.
    f_list.append (folder)
    for f in folder.Folders:
        get_subfolders(f, f_list)

def print_subjects (folder):
    for i in folder.Items:
        print i.subject

folder_list = list()
get_subfolders(selected_folder, folder_list)

for n in xrange(3):
    print_subjects(folder_list[n])