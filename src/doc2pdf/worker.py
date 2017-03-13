from __future__ import print_function

import os
import time
import logging
import tempfile
import collections

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

from doc2pdf import util
from doc2pdf import converter
from doc2pdf import queue

class EventHandler(FileSystemEventHandler):
    def __init__(self, on_created=None, on_deleted=None, on_modified=None, on_moved=None):
        FileSystemEventHandler.__init__(self)
        self.__on_created = on_created
        self.__on_deleted = on_deleted
        self.__on_modified = on_modified
        self.__on_moved = on_moved
    def on_created(self, event):
        if self.__on_created is not None: self.__on_created(event.src_path)
    def on_deleted(self, event):
        if self.__on_deleted is not None: self.__on_deleted(event.src_path)
    def on_modified(self, event):
        if self.__on_modified is not None: self.__on_modified(event.src_path)
    def on_moved(self, event):
        if self.__on_moved is not None: self.__on_moved(event.src_path, event.dest_path)

# TODO: synchronize handlers or queue actions
class Worker:
    EXTENSIONS = [ "doc", "docx", "xls", "xlsx" ]
    
    def __init__(self, config):
        self.__time = time.time
        self.__config = config
        logging.info("check config...")
        
        if not self.check_config():
            logging.error("config malformed!")
            raise Exception()
        logging.info("config okay.")
        
        self.__queue = queue.SyncedQueue(self.__time)
        logging.info("queue created.")
        
        self.__converter = converter.Converter(config["converter_timeout"], config["converter_retries"], self.__queue)
        logging.info("converter created.")
        
        if not os.path.exists(config["temporary_directory"]):
            os.mkdir(config["temporary_directory"])
        else:
            if not os.path.isdir(config["temporary_directory"]):
                logging.error("tmp path doesnt point to a dir!")
                raise Exception()
        util.cleardir(config["temporary_directory"])
        tempfile.tempdir = config["temporary_directory"]
        logging.info("tmp space created.")
        
        self.__observers = []
        include_paths = self.__config["include_paths"]
        for p in include_paths:
            o = self.__create_observer(p)
            self.__observers.append(o)
        logging.info("observers created.")
    def __create_observer(self, path):
        event_handler = EventHandler(
            self.__handle_created_updated,
            self.__handle_deleted,
            self.__handle_created_updated,
            self.__handle_moved
        )
        
        o = Observer()
        o.schedule(event_handler, path, recursive=True)
        
        return o
    def check_config(self):
        # TODO: check config
        return True
    def start(self):
        self.__converter.start()
        for o in self.__observers:
            o.start()
    def stop(self):
        self.__converter.interrupt()
        for o in self.__observers:
            o.stop()
            o.join()
    def __use_path(self, path):
        exclude_paths = self.__config["exclude_paths"]
        for p in exclude_paths:
            if path.startswith(p): return False
        name = os.path.basename(path)
        if name.startswith("~"): return False
        if name.endswith(".tmp"): return False
        split = path.rsplit(".", 1)
        if len(split) < 2: return False
        ext = split[1]
        if ext not in Worker.EXTENSIONS: return False
        return True
    def __pdfpath(self, path):
        return util.replaceextension(path, "pdf")
    def __handle_created_updated(self, path, delay=None):
        if not self.__use_path(path): return
        if len(self.__queue) >= self.__config["queue_capacity"]:
            logging.warning("queue overflow, aborting.")
            return
        logging.info("queue " + path)
        paths = (path, self.__pdfpath(path))
        if delay == None: delay = self.__config["converter_delay"]
        if self.__queue.contains(paths): self.__queue.removeFirst(paths);
        r = self.__queue.put(paths, self.__time() + delay)
        if not r: logging.warning("queue insert failed")
    def __handle_deleted(self, path):
        if not self.__use_path(path): return
        if not self.__config["autodelete"]: return
        logging.info("delete " + path)
        try:
            os.remove(self.__pdfpath(path))
        except:
            pass
    def __handle_moved(self, from_path, to_path):
        if not self.__use_path(from_path): return
        if not self.__use_path(to_path): return
        logging.info("rename " + from_path + " to " + to_path)
        util.silentremove(self.__pdfpath(from_path))
        self.__handle_created_updated(to_path, 0)
