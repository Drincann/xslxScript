
import os


def extname(path):
    return os.path.splitext(path)[1]


def basename(path):
    return os.path.splitext(path)[0]


def sn(n): return 5 * n * (n + 1)
def snSub(begin, end): return sn(end) - sn(begin)


def calHealth(failureTimes, notUploadedTimes, processTimeTimes):
    return 100 - (20 * notUploadedTimes) - (30 * processTimeTimes) - sn(failureTimes)
