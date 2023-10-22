import time, os
from multiprocessing import current_process
from threading import Thread, current_thread
from datetime import datetime
import speech2text

class noWorkThread(Thread):

    # override the constructor
    def __init__(self, value):
        # first execute the base constructor
        Thread.__init__(self)
        # now store the additional value you need in this class
        self.sec = value
        self.valretd= ""
    # override the run function
    def run(self):
        pid = os.getpid()
        threadName = current_thread().name
        processName = current_process().name
        print("Instantating session of noworkthread")
        print(f"{pid} * {processName} * {threadName} ---> No work thread..."+ "\n")
        print((datetime.now().strftime("%m/%d/%Y, %H:%M:%S.%f"))+(" is the start time inside no work thread\n"))

        #messagebox.showinfo("Recording Stopped","OK")
        #print((datetime.now().strftime("%m/%d/%Y, %H:%M:%S.%f"))+(" is the time before recording\n"))
        self.valretd=speech2text.from_mic()
        #print((datetime.now().strftime("%m/%d/%Y, %H:%M:%S.%f"))+(" is the time after recording\n"))
        #time.sleep(self.sec)
        print(f"{pid} * {processName} * {threadName} #	---> Exiting thread..." + "\n")
        print((datetime.now().strftime("%m/%d/%Y, %H:%M:%S.%f"))+(" is the end time inside no work thread\n"))

    def io_bound(sec):
        pid = os.getpid()
        threadName = current_thread().name
        processName = current_process().name
        print(f"{pid} * {processName} * {threadName} ---> No work thread..."+ "\n")
        print((datetime.now().strftime("%m/%d/%Y, %H:%M:%S.%f"))+(" is the start time inside no work thread\n"))
        time.sleep(sec)
        #i=2000000
        #while (i!=0):
        #i=i-1
        #if ((i%20000)==0):
        #print(str(i))
        print(f"{pid} * {processName} * {threadName} #	---> Finished sleeping..." + "\n")
        print((datetime.now().strftime("%m/%d/%Y, %H:%M:%S.%f"))+(" is the end time inside no work thread\n"))
# block for a moment