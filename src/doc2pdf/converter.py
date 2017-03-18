import os.path
import threading
import logging
import shutil
import traceback
import tempfile
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
        self.__init()
    def __del__(self):
        self.__dest()
        PdfConverter.__del__(self)
    def __init(self):
        pythoncom.CoInitialize()
        self.__client = client.DispatchEx("Word.Application")
        self.__client.DisplayAlerts = False
    def __dest(self):
        try:
            if self.__client: self.__client.Quit()
        except:
            logging.error("quit word failed.")
            logging.error(traceback.format_exc())
        pythoncom.CoUninitialize()
    def __recover(self):
        logging.error("recover word process...")
        self.__dest()
        self.__init()
    def convert(self, src, dst):
        try:
            doc = self.__client.Documents.Open(src, ConfirmConversions=False, ReadOnly=True, Revert=True, Visible=False, NoEncodingDialog=True)
            doc.SaveAs(dst, FileFormat=17)
            doc.Close(SaveChanges=0)
        except:
            logging.error("convert word failed %s ..." % src)
            logging.error(traceback.format_exc())
            self.__recover()
            return False
        return True

class Excel2PdfConverter(PdfConverter):
    def __init__(self):
        PdfConverter.__init__(self)
        self.__init()
    def __del__(self):
        self.__dest()
        PdfConverter.__del__(self)
    def __init(self):
        pythoncom.CoInitialize()
        self.__client = client.DispatchEx("Excel.Application")
        self.__client.DisplayAlerts = False
    def __dest(self):
        try:
            if self.__client: self.__client.Quit()
        except:
            logging.error("quit excel failed.")
            logging.error(traceback.format_exc())
        pythoncom.CoUninitialize()
    def __recover(self):
        logging.error("recover excel process...")
        self.__dest()
        self.__init()
    def convert(self, src, dst):
        try:
            book = self.__client.Workbooks.Open(src, ReadOnly=True, IgnoreReadOnlyRecommended=True, Notify=False)
            book.ExportAsFixedFormat(0, dst, OpenAfterPublish=False)
            book.Close()
        except:
            logging.error("convert excel failed %s ..." % src)
            logging.error(traceback.format_exc())
            self.__recover()
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
    def __init__(self, timeout, retries, queue):
        threading.Thread.__init__(self)
        self.__stop_event = threading.Event()
        self.__timeout = timeout
        self.__retries = retries
        self.__queue = queue
        self.__converter = None
    def __convert_retry(self, src, dst):
        successful = False
        for i in range(self.__retries):
            if self.__convert_tmp(src, dst):
                successful = True
                break
            else:
                logging.warning("retry %d convert %s ..." % (i, src))
        if not successful:
            logging.error("retry exceeded %s" % src)
        return successful
    def __convert_tmp(self, src, dst):
        tmpdir = tempfile.mkdtemp()
        src_tmp = os.path.join(tmpdir, os.path.basename(src))
        dst_tmp = os.path.join(tmpdir, os.path.basename(dst))
        # TODO: use try finally
        successful = True
        
        try:
            logging.info("copy %s to %s ..." % (src, src_tmp))
            shutil.copyfile(src, src_tmp)
        except:
            successful = False
            logging.error("cannot copy office file, aborting")
        if successful:
            successful = self.__convert(src_tmp, dst_tmp)
        
        if successful:
            logging.info("convert successful")
            try:
                logging.info("move %s to %s ..." % (dst_tmp, dst))
                shutil.move(dst_tmp, dst)
            except:
                successful = False
                logging.error("cannot move pdf file, aborting")
        else:
            logging.error("convert failed. %s" % src)
        
        if not util.silentremove(src_tmp): logging.warning("cannot remove %s" % src_tmp)
        os.rmdir(tmpdir)
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
            paths, _ = self.__queue.pop()
            if not paths:
                self.__stop_event.set()
                continue
            try:
                self.__convert_retry(paths[0], paths[1])
            except:
                logging.error("uncaught converter exception...")
                logging.error(traceback.format_exc())
    def interrupt(self):
        self.__stop_event.set()
