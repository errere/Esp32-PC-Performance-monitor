import mmap
import contextlib
import time

import json
import win32api
import win32con
import serial
import threading
import xmltodict
import socket
import sys
import os



class pySerial:
    hander = serial.Serial()
    IsSerialOpen = False

    def begin(self,portx,bps):
        print("open:")
        ret = False
        try:
           self.hander = serial.Serial(portx, bps, timeout=None)
           if(self.hander.is_open):
               ret = True
        except Exception as e:
           print("[open]---异常---：", e)
           win32api.MessageBox(0, "[open]---异常---：%s" % str(e), "error",win32con.MB_OK)
        self.IsSerialOpen = ret
        return ret

    def end(self):
        print("close:")
        try:
            self.hander.close()
            self.IsSerialOpen = False
        except Exception as e:
           print("[close]---异常---：", e)

        pass

    def read(self,len):
        if(self.IsSerialOpen):
            try:
                return self.hander.read(len).decode("gbk")
            except Exception as e:
                print("[read]---异常---：", e)

            pass

        return ""

    def print(self,buf):
        if(self.IsSerialOpen):
            try:
                result = self.hander.write(buf.encode("gbk"))#写数据
            except Exception as e:
                print("[print]---异常---：", e)
                return -1

            return result

        return -1

    def write(self,buf):
        if(self.IsSerialOpen):
            try:
                result = self.hander.write(bytes.fromhex(buf))#写数据
            except Exception as e:
                print("[write]---异常---：", e)
                return -1

            return result

        return -1


    def stute(self):
        return self.IsSerialOpen
        pass

    pass


def GetMemory():
    try:
        with contextlib.closing(mmap.mmap(-1, 1024, tagname='AIDA64_SensorValues', access=mmap.ACCESS_READ)) as m:
            s = m.read(40960)
            Sdata = s.decode('UTF-8', 'ignore').strip().strip(b'\x00'.decode())
            Stu = 1
            if(Sdata == ''):
                Stu = 0
                Sdata = 'NULL'
    except Exception as e:
        Stu = 0
        Sdata = "ERROR : %s"

    return Stu,Sdata

TheSerial = pySerial()

#task
def TaskSerial():
    while True:

        Res,Sdata = GetMemory()
        if (Res == 1):
            temp = f'<root>{Sdata}</root>'
            xs = xmltodict.parse(temp)
            data = {'Msg':'OK','Result': xs}
        else:
            data = {'Msg':'Fail','Result': Sdata}
            pass
        TheSerial.write("a0")
        TheSerial.print(json.dumps(data))
        TheSerial.print("\n")
        time.sleep(2)

        pass
    pass

def main():
    argc = len(sys.argv)
    
    if(argc != 3):
        win32api.MessageBox(0, "use :\" python <fileName> <PORTX> <BPS> \"", "error",win32con.MB_OK)
        #print("use :\" python <fileName> <PORTX> <BPS> <HTTPPORT>\"")
        #print("like pythonw file.pyw COM3 115200 8888")
        return

    serOpen = TheSerial.begin(sys.argv[1],int(sys.argv[2]))
    if(serOpen == False):
        win32api.MessageBox(0, "Serial Open Error", "error",win32con.MB_OK)
        win32api.MessageBox(0, "use :\" python <fileName> <PORTX> <BPS>\"", "error",win32con.MB_OK)
        #print("use :\" python <fileName> <PORTX> <BPS> <HTTPPORT>\"")
        #print("like pythonw file.pyw COM3 115200 8888")
        return
    
    win32api.MessageBox(0, "Serial Open OK" , "Info",win32con.MB_OK)

    try:
        threadUPD = threading.Thread(target=TaskSerial)
        threadUPD.setDaemon(True)  # thread1,它做为程序主线程的守护线程,当主线程退出时,thread1线程也会退出,由thread1启动的其它子线程会同时退出,不管是否执行完任务
        threadUPD.start()
    except:
        print("---异常---: 无法启动线程TaskSerial")
    pass



 

if __name__ == '__main__':
    main()
    while(True):
        inS = input()
        if(inS == 'exit'):
            TheSerial.end()
            break
        pass

    #pid = os.getpid()
    #os.popen('taskkill.exe /pid:'+str(pid))     LIl1i   O0o
