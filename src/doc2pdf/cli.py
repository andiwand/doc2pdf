import logging
import argparse
import json
import sys
import traceback

from doc2pdf import watcher


EXAMPLE_CONFIG = """
{
    "log_file": "C:\\\\pathtolog.txt",
    "autodelete": false,
    "converter_timeout": 60,
    "converter_delay": 5,
    "queue_capacity": 100,
    "temporary_directory": "C:\\\\pathtotmp",
    "include_paths": ["C:\\\\incpath1", "C:\\\\incpath2"],
    "exclude_paths": ["C:\\\\incpath1\\\\excpath1"]
}
""".strip().encode("utf-8")


def catchexcept(etype, value, tb):
    if excepthook_old: excepthook_old()
    logging.error("uncatched exception...")
    logging.error("type: %s, value: %s, traceback: %s" % (etype.__name__, value, "".join(traceback.format_tb(tb))))

#def hookexcept():
#    sys.excepthook = catchexcept

def hookexcept():
    old_print_exception = traceback.print_exception
    def custom_print_exception(etype, value, tb, limit=None, file=None):
        catchexcept(etype, value, tb)
        old_print_exception(etype, value, tb, limit=limit, file=file)
    traceback.print_exception = custom_print_exception

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
    
    rootLogger = logging.getLogger()
    rootLogger.setLevel(logging.DEBUG)
    logFormatter = logging.Formatter("%(asctime)s %(message)s")
    
    consoleHandler = logging.StreamHandler(sys.stdout)
    consoleHandler.setFormatter(logFormatter)
    rootLogger.addHandler(consoleHandler)
    
    logging.info("starting doc2pdf...")
    
    logging.info("hook exceptions...")
    hookexcept()
    
    config_file = open(args.config)
    config = json.load(config_file)
    logging.info("config loaded.")
    
    fileHandler = logging.FileHandler(config["log_file"])
    fileHandler.setFormatter(logFormatter)
    rootLogger.addHandler(fileHandler)
    
    w = watcher.Watcher(config)
    w.start()

if __name__ == "__main__":
    main()
