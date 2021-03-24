import tkinter as tk
from tkinter import ttk
import os

class filePickerWindow:

    def __init__(self,parent,subFolder_path):
        #save parent for further use
        self.the_parent = parent

        self.picker_window = tk.Tk()
        self.picker_window.title("File Selector")

        self.picker_frame = tk.Frame(self.picker_window,relief='raised', borderwidth=1,width=400,height=200)
        self.picker_frame.pack_propagate(0)


        self.toplevel =parent.toplevel
        self.subFolder_path = subFolder_path
        self.picker_window.geometry('+%d+%d'%(parent.toplevel.winfo_x(),parent.toplevel.winfo_y()))

        # Set this window to appear on top of all other windows
        self.picker_window.attributes("-topmost",True)

        # Tell the window manager to give focus back after the X button is hit
        self.picker_window.protocol("WM_DELETE_WINDOW", self.kill_window)

        self.pcase_tree()
        self.files = []

        ok_button = tk.Button(self.picker_window,text="Select",command=self.ok_button)
        cancel_button = tk.Button(self.picker_window,text="Cancel", command=self.kill_window)

        self.picker_frame.pack()
        ok_button.pack(side='right', padx=5, pady=5)  
        cancel_button.pack(side='right')

    def ok_button(self):
        selected =  self.tree.selection()
        for i in selected:
            print("ancestry.com results: "+self.lineage_hanlder(i,''))
            self.files.append([self.tree.item(i)['text'],self.lineage_hanlder(i,""),len(self.tree.get_children(i))])
        self.kill_window()

    def show(self):
        self.picker_window.wait_window()
        return self.files

    def kill_window(self):
        self.the_parent.toplevel.wm_attributes("-disabled",False)
        self.picker_window.destroy()

    def lineage_hanlder(self,child,childString):
        sub_dir_path = self.get_lineage(child,childString)

        if "\\" not in sub_dir_path:
            return ""
        else:
            return "\\".join(sub_dir_path.split("\\")[:-1])+'\\'

    def get_lineage(self,child,childString):
        if child:
            childString = self.tree.item(child)['text'] + '\\' + childString
            child = self.tree.parent(child)
            if "Z:" in self.tree.item(child)['text']:
                child = ""
            return self.get_lineage(child,childString)

        return childString[:-1]

    def pcase_tree(self):

        dir = self.subFolder_path

        self.nodes = dict()
        frame = self.picker_frame

        self.tree = ttk.Treeview(frame,height=10,selectmode='none')
        self.tree.bind("<Double-1>", self.select)

        self.ysb = ttk.Scrollbar(frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscroll=self.ysb.set)
        self.tree.heading('#0', text='PCase SubFolder', anchor='w')

        self.ysb.pack(side='right',fill='y')
        self.tree.pack(fill='both')
        

        abspath = dir
        self.insert_node('', abspath, abspath)
        self.tree.bind('<<TreeviewOpen>>', self.open_node)

        self.tree.focus(self.tree.get_children()[0])
        self.tree.item(self.tree.focus(),open=True)
        self.open_node('')

    def select(self,event=None):
        self.tree.selection_toggle(self.tree.focus())
        return 'break'


    def insert_node(self, parent, text, abspath):
        node = self.tree.insert(parent, 'end', text=text, open=False)
        if os.path.isdir(abspath):
            self.nodes[node] = abspath
            self.tree.insert(node, 'end')

    def open_node(self, event):
        node = self.tree.focus()
        abspath = self.nodes.pop(node, None)
        if abspath:
            self.tree.delete(self.tree.get_children(node))
            for p in os.listdir(abspath):
                self.insert_node(node, p, os.path.join(abspath, p))