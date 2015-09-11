import threading
import bisect
import time

def dict_get_set(d, k, v):
    tmp = d.get(k, v);
    if tmp == v: d[k] = v
    return tmp

class TimedQueue:
    def __init__(self):
        self.__queue = []
        self.__element_to_times = {}
        self.__time_to_elements = {}
    def __len__(self):
        return len(self.__queue)
    def put(self, element, time):
        dict_get_set(self.__element_to_times, element, []).append(time)
        dict_get_set(self.__time_to_elements, time, []).append(element)
        
        bisect.insort(self.__queue, time)
        return True
    def get(self, time):
        if not self.__queue: return None
        if time < self.__queue[0]: return None
        time = self.__queue[0]
        element = self.__removeTime(time)
        del self.__queue[0]
        return (element, time)
    def contains(self, element):
        return element in self.__element_to_times
    def next(self):
        if not self.__queue: return None
        return self.__queue[0]
    def __removeTime(self, time):
        elements = self.__time_to_elements.get(time)
        if not elements: return None
        element = elements[0]
        self.__removeElement(element)
        return element
    def __removeElement(self, element):
        times = self.__element_to_times.get(element)
        if not times: return None
        time = times.pop(0)
        if not times: del self.__element_to_times[element]
        
        elements = self.__time_to_elements[time]
        elements.remove(element)
        if not elements: del self.__time_to_elements[time]
        
        return time
    def removeFirst(self, element):
        time = self.__removeElement(element)
        if time == None: return False
        
        index = bisect.bisect(self.__queue, time)
        del self.__queue[index - 1]
        
        return True

class SyncedQueue(TimedQueue):
    def __init__(self, time=time.time):
        TimedQueue.__init__(self)
        self.__condition = threading.Condition()
        self.__time = time
    def __len__(self):
        self.__condition.acquire()
        result = TimedQueue.__len__(self)
        self.__condition.release()
        return result
    def put(self, element, time):
        self.__condition.acquire()
        result = TimedQueue.put(self, element, time)
        if result: self.__condition.notify()
        self.__condition.release()
        return result
    def get(self):
        self.__condition.acquire()
        result = None
        while True:
            if len(self) <= 0: self.__condition.wait()
            time = self.__time();
            result = TimedQueue.get(self, time)
            if result != None: break
            self.__condition.wait(self.next() - time)
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
