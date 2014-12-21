import os
import threading
import util

import win32file
import win32con

ACTION_CREATED = 1
ACTION_DELETED = 2
ACTION_UPDATED = 3
ACTION_RENAMED_FROM = 4
ACTION_RENAMED_TO = 5

FILE_LIST_DIRECTORY = 0x0001

BUFFER_SIZE = 4098
RANDOM_LENGTH = 10

class Observer(threading.Thread):
    def __init__(self, path, callback, buffer_size=BUFFER_SIZE, random_length=RANDOM_LENGTH):
        threading.Thread.__init__(self)
        self.__path = path
        self.__buffer_size = buffer_size
        self.__random_length = random_length
        self.__stop_event = threading.Event()
        self.__callback = callback
        self.__hdir = win32file.CreateFile(self.__path,
                                           FILE_LIST_DIRECTORY,
                                           win32con.FILE_SHARE_READ |
                                           win32con.FILE_SHARE_WRITE |
                                           win32con.FILE_SHARE_DELETE,
                                           None,
                                           win32con.OPEN_EXISTING,
                                           win32con.FILE_FLAG_BACKUP_SEMANTICS,
                                           None)
    def run(self):
        while True:
            results = win32file.ReadDirectoryChangesW(self.__hdir,
                                                      BUFFER_SIZE,
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
                filepath = os.path.join(self.__path, subpath)
                self.__callback(action, filepath)
    def interrupt(self):
        dummy = util.mkrandomdirname(self.__path, self.__random_length)
        self.__stop_event.set()
        os.mkdir(dummy)
        os.rmdir(dummy)
