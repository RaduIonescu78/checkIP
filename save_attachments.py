from tkinter import messagebox

import win32com
from win32com.client.gencache import EnsureDispatch as Dispatch
import os


class SaveAtt:

    def __init__(self, folder_name):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6).Folders.Item(folder_name)
        self.folder_name = folder_name
        self.all_emails = inbox.Items
        self.all_emails.Sort("[ReceivedTime]", False)
        self.last_email = self.all_emails.GetLast()
        self.file_wr = str(self.last_email) + '.txt'
        self.localpath = os.path.abspath(__file__)
        self.attach_folder = (self.localpath[0:len(self.localpath) - 19])

    def save_attachments(self):
        try:
            attach = self.last_email.Attachments
            #print(self.last_email.Attachments)
            att = 0
            for i in range(0, len(attach)):
                attach[i].SaveAsFile(os.path.join(self.attach_folder + str(attach[i])))
                messagebox.showinfo(message=str(attach[i]) + " was saved to " + self.attach_folder)
                print(str(attach[i]) + " was saved to ", self.attach_folder)
                # print(attach(i))
                att += 1
            if att == 0:
                messagebox.showerror(message="No attachments found in the last email !!!")
                print("No attachments found in the last email !!!")
        except:
            print("Attachments error")

    def save_to_file(self, file, string):
        saved_file = os.path.join(self.attach_folder, file + ".txt")
        f = open(saved_file, 'a')
        f.write(string)
        f.close()