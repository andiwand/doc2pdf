import time
import os.path

import util
import converter
import observer

import pythoncom

from_docpath = None
from_pdfpath = None

def observe(action, path):
    global from_docpath, from_pdfpath
    
    name = os.path.basename(path)
    if name.startswith("~$"): return
    ext = os.path.splitext(path)[1]
    if ext != ".doc" and ext != ".docx": return
    
    pdfpath = util.replacelast(path, ext, ".pdf", 1)
    
    if action == observer.ACTION_DELETED:
        handle_delete(path, pdfpath)
    elif action in (observer.ACTION_CREATED, observer.ACTION_UPDATED):
        handle_update(path, pdfpath)
    elif action == observer.ACTION_RENAMED_FROM:
        from_docpath = path
        from_pdfpath = pdfpath
    elif action == observer.ACTION_RENAMED_TO:
        handle_rename(from_docpath, from_pdfpath, path, pdfpath)

def handle_delete(docpath, pdfpath):
    try:
        os.remove(pdfpath)
    except Exception, e:
        print e

def handle_rename(from_docpath, from_pdfpath, to_docpath, to_pdfpath):
    try:
        os.rename(from_pdfpath, to_pdfpath)
    except Exception, e:
        print e

def handle_update(docpath, pdfpath):
    pythoncom.CoInitialize()
    converter.convert(docpath, pdfpath, "pdf")
    pythoncom.CoUninitialize()

path = r"C:\Users\Andreas\Desktop\test"
o = observer.Observer(path, observe)
o.start()
time.sleep(1000)
o.interrupt()
