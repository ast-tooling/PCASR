import tkinter as tk
from tkinter import *
import os
from pk_wrappers.buttonsWrapper import TkinterCustomButton

def QuickCommment():

    root = tk.Tk()
    root.geometry("649x270")
    root.title("Quick Comment")
    root.resizable(False, False)
    
    user = os.getlogin()
    data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser" %user
    appIcon_file = data_folder+"\\images\\billTrustIcon.png"

    iconPhotoImage = tk.PhotoImage(file = appIcon_file)

    def getTextInput():
        result=commentBox.get("1.0","end")
        print(result)
        
        return result
        
    scrollbar = Scrollbar(root)
    scrollbar.pack(side=RIGHT, fill=Y)
    commentBox=tk.Text(root, width=90 ,height=13,font=('Calibri',11))
    commentBox.pack()
    commentBox.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=commentBox.yview)

    btnRead = TkinterCustomButton(master=root,bg_color="#ffffcc",fg_color="#003035",corner_radius=10,text= "Submit",text_color="white",hover_color="#aadb1e",width=80,height=24,command=getTextInput)
                    
    btnRead.pack()

    root.iconphoto(False,iconPhotoImage)
    root.mainloop()

    
QuickCommment()