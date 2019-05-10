# saveMails.py Save numbers of sent emails to specific folder

import win32com.client
import os.path
import tkinter
from tkinter import messagebox

class saveEmailGUI:
    # GUI for the program
    def __init__(self):

        # Create main window
        self.main_window = tkinter.Tk()

        # Create top and bot frames
        self.top_frame = tkinter.Frame()
        self.mid_frame = tkinter.Frame()
        self.bot_frame = tkinter.Frame()

        # Create widgets for top frame
        self.num_msg = tkinter.Label(self.top_frame, text = 'Enter number of messages:')
        self.msg_entry = tkinter.Entry(self.top_frame, width = 10)

        # Pack top frame widgets
        self.num_msg.pack(side='left')
        self.msg_entry.pack(side='left')

        # Create widgets for mid frame
        self.save_path = tkinter.Label(self.mid_frame, text = 'Enter the folder path:')
        self.path_entry = tkinter.Entry(self.mid_frame, width = 50)

        # Pack mid frame widgets
        self.save_path.pack(side='left')
        self.path_entry.pack(side='left')   

        # Create widgets for bot frame
        self.process_btn = tkinter.Button(self.bot_frame, text='Save', command=self.convert)
        self.quit_btn = tkinter.Button(self.bot_frame,text='Quit', command=self.main_window.destroy)

        # Pack bot frame widgets
        self.process_btn.pack(side='left')
        self.quit_btn.pack(side='left')

        # Pack the frames
        self.top_frame.pack()
        self.mid_frame.pack()
        self.bot_frame.pack()

        # Enter main loop
        tkinter.mainloop()

    def convert(self):
    # Connect to outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Get default folder: sent items is number 5
        sent = outlook.GetDefaultFolder(5)

        # Save type for save email
        OlSaveAsType = {
            "olTXT": 0,
            "olRTF": 1,
            "olTemplate": 2,
            "olMSG": 3,
            "olDoc": 4,
            "olHTML": 5,
            "olVCard": 6,
            "olVCal": 7,
            "olICal": 8
        }

        # Get all the messages in the sent folder
        messages = sent.Items

        # Sort messages based on sent time
        messages.Sort("[ReceivedTime]", True)

        # Path to save the messages
        path = self.path_entry.get()
    
        # Get user input for how many messages to save
        num_mails = int(self.msg_entry.get())

        # Save the messages
        for i in range(num_mails):
            messages[i].SaveAs(path + '\\' + ''.join(messages[i].Subject.split(':')) + ".msg", OlSaveAsType['olMSG'])
        
        
        
        # TODO: Fix to check files end with .msg
        # Check if all the emails are in the folder
        if len(os.listdir(path)) == num_mails:
            messagebox.showinfo('Done!', 'Saved.')
        else:
            messagebox.showerror('Error!', 'Can only save ' + str(len(os.listdir(path))) + ' emails.' )
        

save_mail = saveEmailGUI()