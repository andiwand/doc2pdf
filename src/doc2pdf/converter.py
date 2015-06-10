from win32com import client
import pythoncom

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
    except Exception, e:
        print(e)
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
    except Exception, e:
        print(e)
    finally:
        if c: c.Quit()
        pythoncom.CoUninitialize()

def convert_to_pdf(src, dst):
    if src.endswith(".doc") or src.endswith(".docx"):
        convert_word(src, dst, "pdf")
    elif src.endswith(".xls") or src.endswith(".xlsx"):
        convert_excel_to_pdf(src, dst)
    else:
        raise ValueError("not able to convert")
