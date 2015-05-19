from __future__ import print_function

import os
import traceback
import argparse
import time
import datetime
import json

from doc2pdf import util
from doc2pdf import converter
from doc2pdf.observer import Observer

EXAMPLE_CONFIG = """
{
    "include_paths": ["C:\\\\incpath1", "C:\\\\incpath2"],
    "exclude_paths": ["C:\\\\incpath1\\\\excpath1"],
    "log_file": "C:\\\\pathtolog.txt",
    "autodelete": False
}
""".strip().encode("utf-8")

# TODO: switch from print to log function, add timestamp, fix newline
class doc2pdf:
    EXTENSIONS = ["doc", "docx", "xls", "xlsx"]
    
    def __init__(self, config_path):
        self.__log_file = None
        self.__log("starting doc2pdf...")
        
        config_file = open(config_path)
        self.__settings = json.load(config_file)
        
        log_path = self.__settings.get("log_file", None)
        if log_path: self.__log_file = open(log_path, "a")
        
        self.__log("config loaded.")
        self.__log("check config...")
        
        if not self.check_config():
            self.__log("config malformed.")
        self.__log("config okay.")
        
        include_paths = self.__settings["include_paths"]
        
        self.__observers = []
        for p in include_paths:
            o = self.__create_observer(p)
            self.__observers.append(o);
        self.__log("observers created.")
    def __log(self, message):
        t = time.time()
        ts = datetime.datetime.fromtimestamp(t).strftime('%Y-%m-%d %H:%M:%S')
        line = ts + "  " + message + os.linesep
        
        print(line, end="")
        if self.__log_file:
            self.__log_file.write(line)
            self.__log_file.flush()
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
        for o in self.__observers:
            o.start()
    def stop(self):
        for o in self.__observers:
            o.interrupt()
    def __use_path(self, path):
        exclude_paths = self.__settings["exclude_paths"]
        for p in exclude_paths:
            if path.startswith(p): return False
        name = os.path.basename(path)
        if name.startswith("~"): return False
        split = path.rsplit(".", 1)
        if len(split) < 2: return False
        ext = split[1]
        if ext not in doc2pdf.EXTENSIONS: return False
        return True;
    def __pdfpath(self, path):
        return util.replaceextension(path, "pdf")
    def __handle_created_updated(self, path):
        if not self.__use_path(path): return
        self.__log("convert " + path)
        converter.convert_to_pdf(path, self.__pdfpath(path))
    def __handle_deleted(self, path):
        if not self.__use_path(path): return
        if not self.__settings["autodelete"]: return
        self.__log("delete " + path)
        try:
            os.remove(self.__pdfpath(path))
        except Exception:
            self.__log(traceback.format_exc());
    def __handle_renamed(self, from_path, to_path):
        if not self.__use_path(from_path): return
        self.__log("rename from " + from_path + " to " + to_path)
        try:
            os.rename(self.__pdfpath(from_path), self.__pdfpath(to_path))
        except Exception:
            self.__log(traceback.format_exc());
            self.__handle_created_updated(to_path)

def main():
    parser = argparse.ArgumentParser(description="automatic ms office to pdf converter")
    parser.add_argument("config", help="path to the config file")
    parser.add_argument("-c", dest="create", action="store_const", const=True, help="create sample config")
    args = parser.parse_args()
    
    if args.create:
        config_file = open(args.config, "wb")
        config_file.write(EXAMPLE_CONFIG)
        config_file.close()
        return
    
    d2p = doc2pdf(args.config)
    d2p.start()

if __name__ == "__main__":
    main()
