import os
import threading
import logging

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
    
    EVENTS = [ "created", "updated", "deleted", "renamed" ]
    
    def __init__(self, path, buffer_size=DEFAULT_BUFFER_SIZE, restart=True):
        logging.info("observer initialization...")
        threading.Thread.__init__(self)
        self.__path = path
        self.__buffer_size = buffer_size
        self.__restart = restart
        self.__stop_event = threading.Event()
        # https://msdn.microsoft.com/en-us/library/windows/desktop/aa363858(v=vs.85).aspx
        self.__callbacks = {k:None for k in Observer.EVENTS}
        self.__from_path = None
    def subscribe(self, event, callback):
        if event not in Observer.EVENTS:
            logging.warning("subscribe failed.")
            raise ValueError("illegal event")
        self.__callbacks[event] = callback
        logging.info("subscribe done.")
    def __init(self):
        self.__hdir = win32file.CreateFile(
            self.__path,
            Observer.FILE_LIST_DIRECTORY,
            win32con.FILE_SHARE_READ |
                win32con.FILE_SHARE_WRITE |
                win32con.FILE_SHARE_DELETE,
            None,
            win32con.OPEN_EXISTING,
            win32con.FILE_FLAG_BACKUP_SEMANTICS,
            None
        )
    def __deinit(self):
        win32file.CloseHandle(self.__hdir)
    def run(self):
        logging.info("observer started.")
        self.__init()
        logging.info("observer initialized.")
        while True:
            try:
                # https://msdn.microsoft.com/en-us/library/windows/desktop/aa365465(v=vs.85).aspx
                results = win32file.ReadDirectoryChangesW(
                    self.__hdir,
                    self.__buffer_size,
                    True,
                    win32con.FILE_NOTIFY_CHANGE_FILE_NAME |
                        win32con.FILE_NOTIFY_CHANGE_DIR_NAME |
                        win32con.FILE_NOTIFY_CHANGE_ATTRIBUTES |
                        win32con.FILE_NOTIFY_CHANGE_SIZE |
                        win32con.FILE_NOTIFY_CHANGE_LAST_WRITE |
                        win32con.FILE_NOTIFY_CHANGE_SECURITY,
                    None,
                    None
                )
                if self.__stop_event.is_set(): break
                for action, subpath in results:
                    path = os.path.join(self.__path, subpath)
                    self.__handle_action(action, path);
            except Exception as e:
                logging.info("observer error: %s" % e)
                if not isinstance(e, tuple): raise e
                if len(e) != 3: raise e
                if e[1] != "ReadDirectoryChangesW": raise e
                logging.info("observer restarting...")
                self.__deinit()
                self.__init()
                logging.info("observer restarted.")
        self.__deinit()
        logging.info("observer ended.")
    def __handle_action(self, action, path):
        logging.info("observed action: action %d, path %s" % (action, path))
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
        logging.info("observer interrupt...")
        self.__stop_event.set()
        # https://msdn.microsoft.com/en-us/library/windows/desktop/aa363792.aspx
        win32file.CancelIoEx(self.__hdir)
        logging.info("observer interrupted.")
