import os.path
import threading
import logging
import traceback
import multiprocessing as mp
import tempfile
import shutil

from doc2pdf import util

import pythoncom
from win32com import client

#http://msdn.microsoft.com/en-us/library/bb238158.aspx
FORMAT_DICT = {
    "default": 16,
    "doc": 0,
    "pdf": 17,
    "xps": 18
}

def convert_word(src, dst, fmt):
    pythoncom.CoInitialize()
    c = None
    try:
        c = client.DispatchEx("Word.Application")
        c.DisplayAlerts = False
        doc = c.Documents.Open(src, ConfirmConversions=False, ReadOnly=True, Revert=True, Visible=False, NoEncodingDialog=True)
        fmt_code = FORMAT_DICT[fmt]
        doc.SaveAs(dst, FileFormat=fmt_code)
        doc.Close(SaveChanges=0)
    finally:
        if c: c.Quit()
        pythoncom.CoUninitialize()

def convert_excel_to_pdf(src, dst):
    pythoncom.CoInitialize()
    c = None
    try:
        c = client.DispatchEx("Excel.Application")
        c.DisplayAlerts = False
        book = c.Workbooks.Open(src, ReadOnly=True, IgnoreReadOnlyRecommended=True, Notify=False)
        book.ExportAsFixedFormat(0, dst, OpenAfterPublish=False)
        book.Close()
    finally:
        if c: c.Quit()
        pythoncom.CoUninitialize()

def convert_to_pdf(src, dst):
    if src.endswith(".doc") or src.endswith(".docx"):
        convert_word(src, dst, "pdf")
    elif src.endswith(".xls") or src.endswith(".xlsx"):
        convert_excel_to_pdf(src, dst)
    else:
        raise ValueError("unsupported extension")

class pdf_converter(threading.Thread):
    def __init__(self, timeout, queue):
        threading.Thread.__init__(self)
        self.__stop_event = threading.Event()
        self.__timeout = timeout
        self.__queue = queue
    def __convert_tmp(self, src, dst):
        suffix = "." + util.getextension(src)
        src_tmp = util.tmpfile(suffix=suffix)
        dst_tmp = util.tmpfile(suffix=".pdf")
        
        shutil.copyfile(src, src_tmp)
        self.__convert(src_tmp, dst_tmp)
        try:
            shutil.copyfile(dst_tmp, dst)
        except Exception:
            logging.warning("cannot write pdf file, aborting.")
        
        os.remove(src_tmp)
        os.remove(dst_tmp)
    def __convert(self, src, dst):
        logging.info("convert %s ..." % src)
        if not os.path.isfile(src):
            logging.error("file not found: %s" % src)
            return
        
        try:
            p = mp.Process(target=convert_to_pdf, args=(src, dst))
            p.start()
            p.join(self.__timeout)
            if p.is_alive():
                logging.error("timeout: %s" % src)
                p.terminate()
                if p.is_alive():
                    logging.error("cannot terminate: %s" % src)
                p.join(self.__timeout) #TODO: other timeout
        except Exception:
            logging.warning(traceback.format_exc())
    def run(self):
        while not self.__stop_event.is_set():
            paths, _ = self.__queue.get()
            if not paths:
                self.__stop_event.set()
                break
            self.__convert_tmp(paths[0], paths[1])
    def interrupt(self):
        self.__stop_event.set()
