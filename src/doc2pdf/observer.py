import os
import threading
from doc2pdf import util

import win32file
import win32con

class Observer(threading.Thread):
    ACTION_CREATED = 1
    ACTION_DELETED = 2
    ACTION_UPDATED = 3
    ACTION_RENAMED_FROM = 4
    ACTION_RENAMED_TO = 5
    
    FILE_LIST_DIRECTORY = 0x0001
    
    DEFAULT_BUFFER_SIZE = 4098
    DEFAULT_RANDOM_LENGTH = 10
    
    EVENTS = ["created", "updated", "deleted", "renamed"]
    
    def __init__(self, path,
                 buffer_size=DEFAULT_BUFFER_SIZE,
                 random_length=DEFAULT_RANDOM_LENGTH):
        threading.Thread.__init__(self)
        self.__path = path
        self.__buffer_size = buffer_size
        self.__random_length = random_length
        self.__stop_event = threading.Event()
        self.__hdir = win32file.CreateFile(self.__path,
                                           Observer.FILE_LIST_DIRECTORY,
                                           win32con.FILE_SHARE_READ |
                                           win32con.FILE_SHARE_WRITE |
                                           win32con.FILE_SHARE_DELETE,
                                           None,
                                           win32con.OPEN_EXISTING,
                                           win32con.FILE_FLAG_BACKUP_SEMANTICS,
                                           None)
        self.__callbacks = {k:None for k in Observer.EVENTS}
        self.__from_path = None
    
    def subscribe(self, event, callback):
        if event not in Observer.EVENTS: raise ValueError("illegal event")
        self.__callbacks[event] = callback
    
    def run(self):
        while True:
            results = win32file.ReadDirectoryChangesW(self.__hdir,
                                                      Observer.DEFAULT_BUFFER_SIZE,
                                                      True,
                                                      win32con.FILE_NOTIFY_CHANGE_FILE_NAME |
                                                      win32con.FILE_NOTIFY_CHANGE_DIR_NAME |
                                                      win32con.FILE_NOTIFY_CHANGE_ATTRIBUTES |
                                                      win32con.FILE_NOTIFY_CHANGE_SIZE |
                                                      win32con.FILE_NOTIFY_CHANGE_LAST_WRITE |
                                                      win32con.FILE_NOTIFY_CHANGE_SECURITY,
                                                      None,
                                                      None)
            if self.__stop_event.is_set(): break
            for action, subpath in results:
                path = os.path.join(self.__path, subpath)
                self.__handle_action(action, path);
    
    def __handle_action(self, action, path):
        if action == Observer.ACTION_CREATED and self.__callbacks["created"]:
            self.__callbacks["created"](path)
        elif action == Observer.ACTION_UPDATED and self.__callbacks["updated"]:
            self.__callbacks["updated"](path)
        elif action == Observer.ACTION_DELETED and self.__callbacks["deleted"]:
            self.__callbacks["deleted"](path)
        elif action == Observer.ACTION_RENAMED_FROM and self.__callbacks["renamed"]:
            self.__from_path = path
        elif action == Observer.ACTION_RENAMED_TO and self.__callbacks["renamed"]:
            self.__callbacks["renamed"](self.__from_path, path)
    
    def interrupt(self):
        dummy = util.mkrandomdirname(self.__path, self.__random_length)
        self.__stop_event.set()
        os.mkdir(dummy)
        os.rmdir(dummy)
