from win32com import client

#http://msdn.microsoft.com/en-us/library/bb238158.aspx
FORMAT_DICT = {
    "default": 16,
    "doc": 0,
    "pdf": 17,
    "xps": 18
}

def convert(src, dst, fmt):
    try:
        word = client.DispatchEx("Word.Application")
        doc = word.Documents.Open(src)
        fmt_code = FORMAT_DICT[fmt]
        doc.SaveAs(dst, FileFormat = fmt_code)
        doc.Close()
    except Exception, e:
        print e
    finally:
        word.Quit()
