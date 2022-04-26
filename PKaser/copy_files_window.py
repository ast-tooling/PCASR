import tkinter as tk
import os
from shutil import copyfile
import threading
import queue


'''Defines the 'CopyFiles' Window and handles copying logic
    
    I'm sorry if you're trying to debug this, I know it's messy
'''
class copyFilesWindow:

    the_queue = queue.Queue()

    def __init__(self,parent,subFolder_path,subDir,files):

        # save parent for further use
        self.the_parent = parent

        self.subFolder_path = subFolder_path
        self.subDir = subDir
        self.files = files
        self.already_copied = False

        print("subFolder_path: "+str(subFolder_path))
        print("subDir: "+str(subDir))
        print("files: "+str(files))

        self.copy_window = tk.Tk()
        self.copy_window.title("Copying Files")
        self.copy_window.geometry('+%d+%d'%(parent.toplevel.winfo_x(),parent.toplevel.winfo_y()))

        # Tell the window manager to give focus back after the X button is hit
        self.copy_window.protocol("WM_DELETE_WINDOW", self.kill_window)

        # Set this window to appear on top of other windows
        self.copy_window.attributes("-topmost",True)


        self.textbox = tk.Text(master=self.copy_window)
        self.textbox.pack(side=tk.TOP)
        self.start = tk.Button(master=self.copy_window,text="Copy Files",command=self.copy_status)
        self.start.pack(side=tk.BOTTOM)
        

    '''wrapper method for textbox.insert to handle state disabling
    '''
    def text_insert(self,text):
        self.textbox.configure(state='normal')
        self.textbox.insert(tk.END, text)
        self.textbox.configure(state='disabled')
        self.textbox.see(tk.END)  

    '''Opens threads and hands off files to copy over to copy_main method
    '''
    def copy_status(self):
        self.threadlist = []
        for file in self.files:
            print(str(file))
            if isinstance(file,list):
                print("I'm here in the copy files window!")
                self.text_insert("Copying " + file[1] + ' Directory\n')

                # If This is a Folder
                if file[2] != 0:
                    relative_path = file[1]
                    full_dir = self.subFolder_path+'\\'+file[0]+'\\'+relative_path
                    mid_path = file[0]+'\\'+relative_path
                    for sub_file in os.listdir(full_dir):
                        if os.path.isfile(full_dir+'\\'+sub_file): 
                            print(full_dir)
                            print(sub_file)
                            self.copy_main(sub_file,mid_path)
                # Not a Folder
                else:
                    self.copy_main(file[0],file[1])

            # Single File Copy Shorthand from push files elif clause
            else:
                self.copy_main(file,"")


            thread = threading.Thread(target=self.done_check, args=[self.threadlist])
            thread.start()
            self.textbox.after(100,self.thread_check)


    '''Handles destination specific copying logic before hand-off to copyfile method
    '''
    def copy_main(self,file,dir):
        if "\\SAMPLE_DATA" in self.subDir:
            self.text_insert("       Copying to FTP ROOT"+ '\n')
            server = "\\\\ssnj-netapp01\\imtest\\imstage01\\ftproot\\"+self.the_parent.getCustString()
            thread = threading.Thread(target=copyfile, args=[self.subFolder_path+"\\"+dir+file,server+"\\"+file])
            self.threadlist.append(thread)
            thread.start()
        else:
            self.text_insert(file+' \n')
            for server in self.the_parent.server_list:
                if dir:
                    try:
                        os.mkdir("\\\\"+server+self.subDir+"\\"+dir)
                        self.text_insert('      '+dir + ' Folder Created\n')
                    except:
                        self.text_insert('      '+dir + ' Folder Already Exists\n')
                self.text_insert("       Copying to \\\\"+server+ '\n')
                thread = threading.Thread(target=copyfile, args=[self.subFolder_path+'\\'+dir+file,"\\\\"+server+self.subDir+"\\"+dir+file])
                self.threadlist.append(thread)
                thread.start()


    '''works in conjunction with thread_check to provide updates to progress
    '''
    def done_check(self,threadlist):
        print(threadlist)
        while True:
            if all(not thread.isAlive() for thread in threadlist):
                self.the_queue.put("Done")
                break

    '''works in conjunction with done_check to provide updates to progress
    '''
    def thread_check(self):
        try:
            message = self.the_queue.get(block=False)
        except queue.Empty:
            self.textbox.after(100, self.thread_check)
            self.text_insert(".")
            return

        if message == "Done":
            self.text_insert("\nDone!")
        else:
            self.textbox.after(100,self.thread_check)

    def kill_window(self):
        self.the_parent.toplevel.wm_attributes("-disabled",False)
        self.copy_window.destroy()