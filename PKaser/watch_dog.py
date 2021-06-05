import time
from watchdog.observers import Observer
from watchdog.events import PatternMatchingEventHandler

'''
Handles watching and notifying of changes on a given path
'''
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