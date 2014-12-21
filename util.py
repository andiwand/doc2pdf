import os
import string
import random

def mkrandomdirname(path, length):
    while True:
        name = ''.join(random.choice(string.ascii_letters + string.digits) for _ in range(length))
        dirpath = os.path.join(path, name)
        if not os.path.exists(dirpath): return dirpath

def replacelast(s, old, new, occurrence):
    return new.join(s.rsplit(old, occurrence))
