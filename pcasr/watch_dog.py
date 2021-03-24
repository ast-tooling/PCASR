import tkinter as tk
import subprocess
from tkinter import messagebox, filedialog, RIGHT, RAISED, Listbox, END, MULTIPLE, TOP, BOTTOM, ttk, Frame, Label, Text, Scrollbar, Y,X, Message, Button, Menu, Entry, DISABLED,ACTIVE, BOTH
import webbrowser
import validators
import os
from shutil import copyfile
import time
import threading
import queue
import json
from watchdog.observers import Observer
from watchdog.events import PatternMatchingEventHandler
from simple_salesforce import Salesforce
import configparser

class watchDog:

    def __init__(self,parent,path):
        patterns = "*"
        ignore_patterns = ""
        ignore_directories = False
        case_sensitive = True
        self.parent = parent
        my_event_handler = PatternMatchingEventHandler(patterns, ignore_patterns, ignore_directories, case_sensitive)

        def on_any_event(event):
            parent.ftpListUpdate()

        my_event_handler.on_any_event = on_any_event

        go_recursively = False
        my_observer = Observer()
        my_observer.schedule(my_event_handler, path, recursive=go_recursively)

        my_observer.start()
        while self.parent.keepWatching():
            time.sleep(1)
        my_observer.stop()
        my_observer.join()

    def stopWatching(self):
        self.watching = False