from __future__ import print_function

import os
import traceback
import argparse
import json
import logging

from doc2pdf import util
from doc2pdf import converter
from doc2pdf.observer import Observer

EXAMPLE_CONFIG = """
{
    "autodelete": false,
    "timeout": 60,
    "include_paths": ["C:\\\\incpath1", "C:\\\\incpath2"],
    "exclude_paths": ["C:\\\\incpath1\\\\excpath1"]
}
""".strip().encode("utf-8")

# TODO: switch from print to log function, add timestamp, fix newline
class watcher:
    EXTENSIONS = ["doc", "docx", "xls", "xlsx"]
    
    def __init__(self, config_path):
        logging.info("starting doc2pdf...")
        
        config_file = open(config_path)
        self.__settings = json.load(config_file)
        
        logging.info("config loaded.")
        logging.info("check config...")
        
        if not self.check_config():
            logging.error("config malformed.")
            raise Exception()
        logging.info("config okay.")
        
        self.__converter = converter.pdf_converter()
        logging.info("converter created.")
        
        self.__observers = []
        include_paths = self.__settings["include_paths"]
        for p in include_paths:
            o = self.__create_observer(p)
            self.__observers.append(o);
        logging.info("observers created.")
    def __create_observer(self, path):
        o = Observer(path)
        o.subscribe("created", self.__handle_created_updated);
        o.subscribe("updated", self.__handle_created_updated);
        o.subscribe("deleted", self.__handle_deleted);
        o.subscribe("renamed", self.__handle_renamed);
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
            o.interrupt()
    def __use_path(self, path):
        exclude_paths = self.__settings["exclude_paths"]
        for p in exclude_paths:
            if path.startswith(p): return False
        name = os.path.basename(path)
        if name.startswith("~"): return False
        if name.endswith(".tmp"): return False
        split = path.rsplit(".", 1)
        if len(split) < 2: return False
        ext = split[1]
        if ext not in watcher.EXTENSIONS: return False
        return True;
    def __pdfpath(self, path):
        return util.replaceextension(path, "pdf")
    def __handle_created_updated(self, path):
        if not self.__use_path(path): return
        logging.info("convert " + path)
        self.__converter.add(path, self.__pdfpath(path))
    def __handle_deleted(self, path):
        if not self.__use_path(path): return
        if not self.__settings["autodelete"]: return
        logging.info("delete " + path)
        try:
            os.remove(self.__pdfpath(path))
        except Exception:
            self.__log(traceback.format_exc());
    def __handle_renamed(self, from_path, to_path):
        if not self.__use_path(from_path): return
        if not self.__use_path(to_path): return
        logging.info("rename from " + from_path + " to " + to_path)
        try:
            os.rename(self.__pdfpath(from_path), self.__pdfpath(to_path))
        except Exception:
            logging.warning(traceback.format_exc());
            self.__handle_created_updated(to_path)

def main():
    logging.basicConfig(format="%(asctime)s %(message)s", level=logging.DEBUG)
    
    parser = argparse.ArgumentParser(description="automatic ms office to pdf converter")
    parser.add_argument("config", help="path to the config file")
    parser.add_argument("-c", dest="create", action="store_const", const=True, help="create sample config")
    args = parser.parse_args()
    
    if args.create:
        config_file = open(args.config, "wb")
        config_file.write(EXAMPLE_CONFIG)
        config_file.close()
        return
    
    w = watcher(args.config)
    w.start()

if __name__ == "__main__":
    main()
