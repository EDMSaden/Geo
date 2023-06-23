import cv2 as cv
import numpy as np
from time import time
import win32gui, win32ui, win32con

class WindowCapture:
    def __init__(self, windowname):
        self.hwnd = win32gui.FindWindow(None, windowname)
        #Get the window size
        self.left, self.top, self.right, self.bot = win32gui.GetWindowRect(self.hwnd)
        self.w = self.right - self.left
        self.h = self.bot - self.top
        self.hwnd = None

    def window_capture(self):
        wDC = win32gui.GetWindowDC(self.hwnd)
        dcObj = win32ui.CreateDCFromHandle(wDC)
        cDC = dcObj.CreateCompatibleDC()
        dataBitMap = win32ui.CreateBitmap()
        dataBitMap.CreateCompatibleBitmap(dcObj, self.w, self.h)
        cDC.SelectObject(dataBitMap)
        cDC.BitBlt((0,0,),(self.w,self.h), dcObj, (self.left, self.top), win32con.SRCCOPY)

        signedIntsArray = dataBitMap.GetBitmapBits(True)
        img = np.fromstring(signedIntsArray, dtype='uint8')
        img.shape = (self.h, self.w, 4)

        #Free Resources
        dcObj.DeleteDC()
        cDC.DeleteDC()
        win32gui.ReleaseDC(self.hwnd,wDC)
        win32gui.DeleteObject(dataBitMap.GetHandle())

        #drop the alpha channel
        img = img[...,:3]
        img = np.ascontiguousarray(img)

        return img

    #Поиск окн
    @staticmethod
    def list_window_names():
        def winEnumHandler(hwnd, ctx):
            if win32gui.IsWindowVisible(hwnd):
                print(hex(hwnd), win32gui.GetWindowText(hwnd))
        win32gui.EnumWindows(winEnumHandler, None)


if __name__ == '__main__':
    wincap = WindowCapture('10.31.6.59/inst/ — Mozilla Firefox')
    loop_time = time()
    while True:
        screenshot = wincap.window_capture()      
        cv.imshow('Computer Vision', screenshot)
        print(f'FPS {int(1 / (time() - loop_time))}')
        loop_time = time()
        if cv.waitKey(1) == ord('q'):
            cv.destroyAllWindows
            break

    print('Done')
    WindowCapture.list_window_names()