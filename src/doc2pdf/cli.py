import logging
import argparse
import json
import sys
import traceback
import threading

from doc2pdf import worker


EXAMPLE_CONFIG = """
{
    "log_file": "C:\\\\pathtolog.txt",
    "autodelete": false,
    "converter_timeout": 60,
    "converter_retries": 3,
    "converter_delay": 5,
    "queue_capacity": 100,
    "temporary_directory": "C:\\\\pathtotmp",
    "include_paths": ["C:\\\\incpath1", "C:\\\\incpath2"],
    "exclude_paths": ["C:\\\\incpath1\\\\excpath1"]
}
""".strip().encode("utf-8")


def catchexcept(etype, value, tb):
    logging.error("uncatched exception...")
    logging.error("type: %s, value: %s, traceback: %s" % (etype.__name__, value, "".join(traceback.format_tb(tb))))

def hookexcept():
    """
    Workaround for `sys.excepthook` thread bug from:
    http://bugs.python.org/issue1230540

    Call once from the main thread before creating any threads.
    """

    sys.excepthook = catchexcept
    init_original = threading.Thread.__init__

    def init(self, *args, **kwargs):
        init_original(self, *args, **kwargs)
        run_original = self.run

        def run_with_except_hook(*args2, **kwargs2):
            try:
                run_original(*args2, **kwargs2)
            except Exception:
                sys.excepthook(*sys.exc_info())

        self.run = run_with_except_hook

    threading.Thread.__init__ = init

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
    
    w = worker.Worker(config)
    w.start()

if __name__ == "__main__":
    main()
