import os
import string
import random
import tempfile
import shutil

def mkrandomdirname(path, length):
    while True:
        name = ''.join(random.choice(string.ascii_letters + string.digits) for _ in range(length))
        dirpath = os.path.join(path, name)
        if not os.path.exists(dirpath): return dirpath

def replacelast(s, old, new, occurrence):
    return new.join(s.rsplit(old, occurrence))

def replaceextension(path, new_extension):
    return path.rsplit(".", 1)[0] + "." + new_extension

def getextension(path):
    return path.rsplit(".", 1)[1]

def tmpfile(suffix="", prefix="tmp"):
    tmp = tempfile.NamedTemporaryFile(mode="wb", delete=False,
                                      suffix=suffix, prefix=prefix)
    result = tmp.name
    tmp.close()
    return result

def cleardir(path):
    for f in os.listdir(path):
        path = os.path.join(path, f)
        try:
            if os.path.isfile(path):
                os.unlink(path)
            elif os.path.isdir(path):
                shutil.rmtree(path)
        except Exception, e:
            print e
