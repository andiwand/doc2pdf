import os
import string
import random

class Tee:
    def __init__(self, f1, f2, autoflush=False):
        self.__f1 = f1
        self.__f2 = f2
        self.__autoflush = autoflush
    def write(self, data):
        self.__f1.write(data)
        self.__f2.write(data)
        if self.__autoflush: self.flush()
    def flush(self):
        self.__f1.flush()
        self.__f2.flush()

def mkrandomdirname(path, length):
    while True:
        name = ''.join(random.choice(string.ascii_letters + string.digits) for _ in range(length))
        dirpath = os.path.join(path, name)
        if not os.path.exists(dirpath): return dirpath

def replacelast(s, old, new, occurrence):
    return new.join(s.rsplit(old, occurrence))

def replaceextension(path, new_extension):
    return path.rsplit(".", 1)[0] + "." + new_extension
