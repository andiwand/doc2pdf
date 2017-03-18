import threading
import bisect
import time
import util

# TODO: use deque and insertion sort
class TimedQueue:
    def __init__(self):
        self.__queue = []
        self.__element_to_times = {}
        self.__time_to_elements = {}
    def __len__(self):
        return len(self.__queue)
    def put(self, element, t):
        util.dict_get_set(self.__element_to_times, element, []).append(t)
        util.dict_get_set(self.__time_to_elements, t, []).append(element)
        bisect.insort(self.__queue, t)
        return True
    def pop(self, t):
        if len(self.__queue) <= 0: return None
        if t < self.__queue[0]: return None
        t = self.__queue[0]
        element = self.__removeTime(t)
        del self.__queue[0]
        return (element, t)
    def contains(self, element):
        return element in self.__element_to_times
    def next(self):
        if not self.__queue: return None
        return self.__queue[0]
    def __removeTime(self, t):
        elements = self.__time_to_elements.get(t)
        if len(elements) <= 0: return None
        element = elements[0]
        self.__removeElement(element)
        return element
    def __removeElement(self, element):
        times = self.__element_to_times.get(element)
        if times == None or len(times) <= 0: return None
        t = times.pop(0)
        if len(times) > 0: del self.__element_to_times[element]
        elements = self.__time_to_elements[t]
        elements.remove(element)
        if len(elements) > 0: del self.__time_to_elements[t]
        return t
    def removeFirst(self, element):
        t = self.__removeElement(element)
        if t == None: return False
        index = bisect.bisect(self.__queue, t)
        del self.__queue[index - 1]
        return True

class SyncedQueue(TimedQueue):
    def __init__(self, timefunc=time.time):
        TimedQueue.__init__(self)
        self.__condition = threading.Condition()
        self.__timefunc = timefunc
    def __len__(self):
        self.__condition.acquire()
        result = TimedQueue.__len__(self)
        self.__condition.release()
        return result
    def put(self, element, t):
        self.__condition.acquire()
        result = TimedQueue.put(self, element, t)
        if result: self.__condition.notifyAll()
        self.__condition.release()
        return result
    def pop(self):
        self.__condition.acquire()
        result = None
        while True:
            if len(self) <= 0:
                self.__condition.wait()
                continue
            t = self.__timefunc()
            result = TimedQueue.pop(self, t)
            if result is not None: break
            self.__condition.wait(self.next() - t)
        self.__condition.release()
        return result
    def next(self):
        self.__condition.acquire()
        result = TimedQueue.next(self)
        self.__condition.release()
        return result
    def contains(self, element):
        self.__condition.acquire()
        result = TimedQueue.contains(self, element)
        self.__condition.release()
        return result
    def removeFirst(self, element):
        self.__condition.acquire()
        result = TimedQueue.removeFirst(self, element)
        self.__condition.release()
        return result
