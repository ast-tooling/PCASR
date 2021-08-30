import tkinter as tk
import os
from tkinter_custom_button import TkinterCustomButton


root = tk.Tk()
root.geometry("400x190")
root.title("The PKaser - Quick Comment")

 
user = os.getlogin()
data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser" %user
appIcon_file = data_folder+"\\images\\billTrustIcon.png"

iconPhotoImage = tk.PhotoImage(file = appIcon_file)

def getTextInput():
    result=textExample.get("1.0","end")
    print(result)

    return result

textExample=tk.Text(root, height=10)
textExample.pack()
btnRead = TkinterCustomButton(master=root,bg_color="#ffffcc",fg_color="#003035",corner_radius=10,text= "Submit",text_color="white",hover_color="#aadb1e",width=80,height=24,command=getTextInput)
                
btnRead.pack()

root.iconphoto(False,iconPhotoImage)
root.mainloop()

        