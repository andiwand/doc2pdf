import os.path
import threading
import logging
import shutil
import traceback
import pythoncom
from win32com import client

from doc2pdf import util

#http://msdn.microsoft.com/en-us/library/bb238158.aspx
FORMAT_DICT = {
    "default": 16,
    "doc": 0,
    "pdf": 17,
    "xps": 18
}

class PdfConverter:
    def __init__(self):
        pass
    def __del__(self):
        pass
    def convert(self, src, dst):
        pass

class Word2PdfConverter(PdfConverter):
    def __init__(self):
        PdfConverter.__init__(self)
        pythoncom.CoInitialize()
        self.__client = client.DispatchEx("Word.Application")
        self.__client.DisplayAlerts = False
    def __del__(self):
        self.__client.Quit()
        pythoncom.CoUninitialize()
        PdfConverter.__del__(self)
    def convert(self, src, dst):
        if not os.path.isfile(src): return False
        if not os.path.isfile(dst): return False
        try:
            doc = self.__client.Documents.Open(src, ConfirmConversions=False, ReadOnly=True, Revert=True, Visible=False, NoEncodingDialog=True)
            doc.SaveAs(dst, FileFormat=17)
            doc.Close(SaveChanges=0)
        except Exception:
            logging.error("convert word failed...")
            logging.error(traceback.format_exc())
            return False
        return True

class Excel2PdfConverter(PdfConverter):
    def __init__(self):
        PdfConverter.__init__(self)
        pythoncom.CoInitialize()
        self.__client = client.DispatchEx("Excel.Application")
        self.__client.DisplayAlerts = False
    def __del__(self):
        self.__client.Quit()
        pythoncom.CoUninitialize()
        PdfConverter.__del__(self)
    def convert(self, src, dst):
        if not os.path.isfile(src): return False
        if not os.path.isfile(dst): return False
        try:
            book = self.__client.Workbooks.Open(src, ReadOnly=True, IgnoreReadOnlyRecommended=True, Notify=False)
            book.ExportAsFixedFormat(0, dst, OpenAfterPublish=False)
            book.Close()
        except Exception:
            logging.error("convert excel failed...")
            logging.error(traceback.format_exc())
            return False
        return True

class Office2PdfConverter(PdfConverter):
    def __init__(self):
        PdfConverter.__init__(self)
        self.__word = Word2PdfConverter()
        self.__excel = Excel2PdfConverter()
    def __del__(self):
        PdfConverter.__del__(self)
    def convert(self, src, dst):
        if src.endswith(".doc") or src.endswith(".docx"):
            return self.__word.convert(src, dst)
        elif src.endswith(".xls") or src.endswith(".xlsx"):
            return self.__excel.convert(src, dst)
        else:
            raise ValueError("unsupported extension")
            return False

class Converter(threading.Thread):
    def __init__(self, timeout, queue):
        threading.Thread.__init__(self)
        self.__stop_event = threading.Event()
        self.__timeout = timeout
        self.__queue = queue
        self.__converter = None
    def __convert_retry(self, src, dst):
        #TODO: quickfix, outsource
        successful = False
        for _ in range(3):
            if self.__convert_tmp(src, dst):
                successful = True
                break
            else:
                logging.warning("retry convert..." % (src, dst))
        if not successful: logging.warning("retry exceeded")
        return successful
    def __convert_tmp(self, src, dst):
        suffix = "." + util.getextension(src)
        src_tmp = util.tmpfile(suffix=suffix)
        dst_tmp = util.tmpfile(suffix=".pdf")
        
        successful = True
        
        try:
            logging.info("copy %s to %s ..." % (src, src_tmp))
            shutil.copyfile(src, src_tmp)
        except Exception:
            successful = False
            logging.warning("cannot copy office file, aborting.")
        
        if successful:
            self.__convert(src_tmp, dst_tmp)
            try:
                logging.info("copy %s to %s ..." % (dst_tmp, dst))
                shutil.copyfile(dst_tmp, dst)
            except Exception:
                successful = False
                logging.warning("cannot copy pdf file, aborting.")
        
        if successful: logging.info("convert successful.")
        else: logging.warning("convert failed. %s" % src)
        
        os.remove(src_tmp)
        os.remove(dst_tmp)
        
        return successful
    def __convert(self, src, dst):
        logging.info("convert %s to %s ..." % (src, dst))
        if not os.path.isfile(src):
            logging.error("file not found: %s" % src)
            return False
        
        return self.__converter.convert(src, dst)
    def run(self):
        self.__converter = Office2PdfConverter()
        
        while not self.__stop_event.is_set():
            paths, _ = self.__queue.get()
            if not paths:
                self.__stop_event.set()
                break
            self.__convert_retry(paths[0], paths[1])
    def interrupt(self):
        self.__stop_event.set()
