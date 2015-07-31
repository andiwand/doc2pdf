import logging
import argparse
import os
import shutil
import time
import util
import errno

def is_int(s):
    try:
        int(s)
        return True
    except ValueError:
        return False

def touch(path, times=None):
    with open(path, "a"):
        os.utime(path, times)

def silentremove(filename):
    try:
        os.remove(filename)
    except OSError as e:
        if e.errno != errno.ENOENT:
            raise

def main():
    parser = argparse.ArgumentParser(description="test script for doc2pdf")
    parser.add_argument("-t", "--timeout", default="10", help="seconds to wait for the pdf")
    parser.add_argument("-s", "--src", help="path to the file to copy")
    parser.add_argument("path", help="path to the created file")
    args = parser.parse_args(["-t", "100", "asdf"])
    
    logging.basicConfig(level=logging.DEBUG, format="%(asctime)s %(message)s") #TODO: outsource
    
    logging.info("starting doc2pdf tester...")
    logging.debug("validate arguments")
    if args.src and not os.path.isfile(args.src):
        logging.error("src file does not exist")
        return 1
    if not is_int(args.timeout):
        logging.error("timeout is not a number")
        return 2
    
    logging.info("create test file %s" % args.path);
    if args.src:
        logging.info("copying from %s" % args.src)
        shutil.copyfile(args.src, args.path)
    else:
        touch(args.path)
    
    time.sleep(int(args.timeout))
    
    logging.info("checking for pdf file");
    pdf_file = util.replaceextension(args.path, "pdf")
    
    result = os.path.isfile(pdf_file)
    if result: logging.info("pdf found")
    else: logging.error("pdf not found")
    
    logging.info("clear files");
    silentremove(pdf_file)
    silentremove(args.path)
    
    logging.info("exit");
    if not result: return 3
    return 0

if __name__ == "__main__":
    main()
