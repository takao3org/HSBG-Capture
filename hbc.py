# -*- coding: utf-8 -*-
#  No License: This is free and unencumbered software released into
#  the public domain.

from PIL import Image
import win32con
import win32api
import win32gui
import win32ui
import sys
import os
import datetime

import threading
import queue
import time
import tkinter
from functools import partial

class HsBgCap():
    def __init__(self):
        self.hwnd = 0
        self.size = [0, 0]
        self.offs = [0, 0]

    # public methods
    def Update(self):
        self.hwnd = win32gui.FindWindow(None, 'Hearthstone')
        if self.hwnd == 0:
            self.size = [0, 0]
            self.offs = [0, 0]
            return False

        # window location
        rect = win32gui.GetWindowRect(self.hwnd)
        (wx, wy)  = (rect[0], rect[1])

        # client size
        rect = win32gui.GetClientRect(self.hwnd)
        self.size = (rect[2], rect[3])

        # client location
        (cx, cy)  = win32gui.ClientToScreen(self.hwnd, (0, 0))
        self.offs = (cx - wx, cy - wy)
        return True

    def Capture(self, x, y, w, h):
        if self.hwnd == 0:
            return None
        try:
            # create bitmap in memory
            wdc = win32gui.GetWindowDC(self.hwnd)
            hdc = win32ui.CreateDCFromHandle(wdc)
            bmp = win32ui.CreateBitmap()
            bmp.CreateCompatibleBitmap(hdc, w, h)

            # bitblt to bitmap
            cdc = hdc.CreateCompatibleDC()
            cdc.SelectObject(bmp)
            x += self.offs[0]
            y += self.offs[1]
            cdc.BitBlt((0, 0), (w, h), hdc, (x, y), win32con.SRCCOPY)

            # create PIL instance from bitmap
            img = Image.frombuffer('RGB', (w, h), bmp.GetBitmapBits(True), 
                                   'raw', 'BGRX', 0, 1)
            cdc.DeleteDC()
            win32gui.DeleteObject(bmp.GetHandle())
            hdc.DeleteDC()
            win32gui.ReleaseDC(self.hwnd, wdc)
            return img
        except:
            SetError('Cannot capture (%r)' %win32api.GetLastError())
            return None
    
    def IsHome(self):
        # capture '1st Place' part at lobby
        (x, y) = self.__CalcXY(1451, 194)
        img = self.Capture(x - 24, y, 48, 2)
        if img == None:
            return False

        # search purple pixel
        hsv = img.convert('HSV').split()
        c = self.__CntPixel(hsv, 48, 2, (202, 154, 211), (203, 156, 255))
        if c >= 12:
            return True
        else:
            return False

    def IsRank(self):
        # capture 'flag' part near new rating
        (x, y) = self.__CalcXY(965, 748)
        img = self.Capture(x - 24, y, 48, 2)
        if img == None:
            return False

        # search dark purple pixel
        hsv = img.convert('HSV').split()
        c = self.__CntPixel(hsv, 48, 2, (209, 141,  96), (212, 151,  99))
        if c >= 48:
            return True
        else:
            return False

    def GetRate(self):
        # capture numbers below 'rating' at lobby
        (x, y) = self.__CalcXY(1380, 502)
        (w, h) = self.__CalcWH(148, 44)
        img = self.Capture(x, y, w, h)
        if img == None:
            return -1
        hsv = img.convert('HSV').split()

        # check if top edge of captured image is dark
        c = self.__CntPixel(hsv, w, 1, (208, 0, 0), (216, 40, 64))
        if c <= w / 2:
            return -1

        # recognize numbers from captured image
        numb = 0
        for r in self.__GetRect(hsv, w, h, 5):
            res = -1
            # check the size of character roughly
            if r[2] - r[0] >= h / 6 and r[2] - r[0] <= h and \
               r[3] - r[1] >= h / 2:
                res = self.__GetNumb(hsv, w, h, r)
            if res == -1:
                SetError('Cannot recognize Rate')
                img.save('Rate_%d_%d_%d_%d.bmp' % (x, y, w, h))
                return -1
            numb = numb * 10 + res
        img.save(path+'\\rate.bmp')
        return numb

    def GetRank(self):
        # capture rank number in result display
        (x, y) = self.__CalcXY(858, 632)
        (w, h) = self.__CalcWH(148, 50)
        img = self.Capture(x, y, w, h)
        if img == None:
            return -1
        hsv = img.convert('HSV').split()

        # check if top edge of captured image is dark
        c = self.__CntPixel(hsv, w, 1, (188, 76, 0), (200, 116, 64))
        if c <= w / 2:
            return -1

        # recognize a number from captured image
        numb = 0
        for r in self.__GetRect(hsv, w, h, 1):
            res = -1
            # check the size of character roughly
            if r[2] - r[0] >= h / 6 and r[2] - r[0] <= h and \
               r[3] - r[1] >= h / 2:
                res = self.__GetNumb(hsv, w, h, r)
            if res == -1:
                SetError('Cannot recognize Rank')
                img.save('Rank_%d_%d_%d_%d.bmp' % (x, y, w, h))
                return -1
            numb = res
        img.save(path+'\\rank.bmp')
        return numb

    # private methods
    def __CalcXY(self, x, y):
        # scale (x, y) for current client from 1920/1080
        s = self.size[1] / 1080
        return (int((x - 960) * s + (self.size[0] / 2) + .5),
                int((y - 540) * s + (self.size[1] / 2) + .5))
    
    def __CalcWH(self, w, h):
        # scale (w, h) for current client from 1920/1080
        s = self.size[1] / 1080
        return (int(w * s + .5), int(h * s + .5))

    def __GetRect(self, hsv, w, h, n):
        # measure number's letter size
        data_v = hsv[2].getdata()
        rect = []
        (lx, rx) = (-1, -2)
        (lu, ru) = (0, 0)
        for x in range(w + 1):
            flag = 0
            for y in range(h + 1):
                if x < w and y < h:
                    p = data_v[y * w + x]
                else:
                    p = 0
                if flag < 2 and p >= 200:
                    flag = 2
                if flag < 1 and p >= 100:
                    flag = 1
            # search bright pixel
            if flag == 2 and lx != -1 and rx == -2:
                rx = -1
            # search black or grey pixel for right edge
            if flag <= 1 and lx != -1 and rx == -1:
                rx = x
                if flag == 1:
                    ru = 0.5
            # search white or grey pixel for left edge
            if flag >= 1 and lx == -1:
                lx = x
                if flag == 1:
                    lu = 0.5
            if rx <= -1:
                continue

            (ty, by) = (-1, -2)
            (tv, bv) = (0, 0)
            for v in range(h + 1):
                flag = 0
                for u in range(lx, rx + 1):
                    if u < w and v < h:
                        p = data_v[v * w + u]
                    else:
                        p = 0
                    if flag < 2 and p >= 200:
                        flag = 2
                    if flag < 1 and p >= 100:
                        flag = 1
                # search bright pixel
                if flag == 2 and ty != -1 and by == -2:
                    by = -1
                # search black or grey pixel for bottom edge
                if flag <= 1 and ty != -1 and by == -1:
                    by = v
                    if flag == 1:
                        bv = 0.5
                # search white or grey pixel for top edge
                if flag >= 1 and ty == -1:
                    ty = v
                    if flag == 1:
                        tv = 0.5
            if by >= 0:
                rect.append((lx, ty, rx, by, lu, tv, ru, bv))
            if len(rect) >= n:
                break
            (lx, rx) = (-1, -2)
            (lu, ru) = (0, 0)
        return rect

    def __GetNumb(self, hsv, w, h, r):
        # sample location
        sample = (
            ( 1, 11), ( 2,  5), ( 2, 14), ( 2, 17), ( 3,  8), ( 5, 15),
            ( 8, 16), ( 9,  6), ( 9,  9), ( 9, 13), (11,  6), (13, 15),
            (15, 10), (15, 19))

        # reference data (0: black, 255: white, -1: not sure)
        result = (
            (255, -1,255,255,255,255, -1, -1,  0,  0,255,255,255, -1),
            ( -1, -1, -1, -1,255,255,  0,  0,  0,  0,  0,  0,  0,  0),
            ( -1, -1,  0,  0,255,  0,255,  0,  0, -1,255,  0, -1,255),
            (  0, -1,  0, -1,  0,  0,  0,255,255, -1, -1,255,  0,  0),
            ( -1,  0,255,  0, -1,255, -1,255,255,255,255,255,  0, -1),
            ( -1, -1,  0, -1,255,  0,  0,  0,255, -1,  0,255,  0,  0),
            (255, -1,255, -1,255,255, -1,  0,255,  0,  0,255, -1, -1),
            (  0, -1,  0, -1,  0,255,255, -1,255, -1, -1,  0,  0,  0),
            (  0,255,255,255,255, -1,  0, -1, -1, -1,255,255,  0,  0),
            ( -1,255, -1,  0,255, -1,  0, -1, -1,255,255,255,255,  0))

        # reference letter size for each number
        size = (22.5, 21.0, 23.0, 22.5, 21.0, 22.5, 22.0, 22.5, 22.5, 22.5)

        # reference letter top/left offset
        offs = ((0.5, 0.5), (  0, 0.5), (0.5, 0.5), (0.5, 0.5), (0.5, 0.5),
                (0.5, 0.5), (0.5,   0), (  0,   0), (  0, 0.5), (  0, 0.5))

        (lx, ty, rx, by, lu, tv, ru, bv) = r
        data = hsv[2].getdata()
        rate = [0] * 10
        for n in range(10):
            s = (by + bv - ty - tv) / size[n]
            for i, (x, y) in enumerate(sample):
                x = lx + int((x - offs[n][0]) * s + lu + .5)
                y = ty + int((y - offs[n][1]) * s + tv + .5)
                if x < w and y < h:
                    p = data[y * w + x]
                else:
                    p = 0
                r = result[n][i]
                if (r == 0 and p >= 150) or (r == 255 and p < 150):
                    rate[n] += 1
        # most certainly number would be minimum rate
        ((n0, r0), (n1, r1)) = \
            sorted(zip(range(10), rate), key=lambda t: t[1])[0:2]
        if r0 == 0 or (r0 == 1 and r1 >= 2):
            return n0
        return -1

    def __CntPixel(self, hsv, w, h, cl, ch):
        data_h = hsv[0].getdata()
        data_s = hsv[1].getdata()
        data_v = hsv[2].getdata()
        cntr = 0
        for y in range(h):
            rptr = y * w
            for x in range(w):
                if data_h[rptr+x] < cl[0] or data_h[rptr+x] > ch[0]:
                    continue
                if data_s[rptr+x] < cl[1] or data_s[rptr+x] > ch[1]:
                    continue
                if data_v[rptr+x] < cl[2] or data_v[rptr+x] > ch[2]:
                    continue
                cntr += 1
        return cntr

def SetError(s):
    global error
    error.write(str(datetime.datetime.now()) + ': %s\n' % s)
    error.flush()

def MainLoop():
    global hbcap
    state = 'null'
    while True:
        print('state: %s' % state, end='')
        # check if hearthstone client exists or not
        if hbcap.Update() == False or hbcap.size == (0, 0):
            state = 'null'
        else:
            sprev = state
            state = 'idle'
            # check client display
            if hbcap.IsRank():
                if sprev != 'rank':
                    # if new to rank display, get number
                    numb = hbcap.GetRank()
                    if numb != -1:
                        print('get rank -> %d' % numb, end='')
                        try:
                            with open(path_rank, mode='a') as file:
                                file.write('%d' % numb)
                            with open(path_rank, mode='r') as file:
                                vars_rank.set(file.read())
                        except:
                            pass
                        state = 'rank'
                    else:
                        state = 'idle'
                else:
                    state = 'rank'
            if hbcap.IsHome():
                if sprev != 'home':
                    # if new to home display, get number
                    numb = hbcap.GetRate()
                    if numb != -1:
                        print('get rate -> %d' % numb, end='')
                        with open(path_rate, mode='w') as file:
                            file.write('%d' % numb)
                        with open(path_rate, mode='r') as file:
                            vars_rate.set(file.read())
                        state = 'home'
                    else:
                        state = 'idle'
                else:
                    state = 'home'
        print('')

        # command from GUI
        while trans.empty() == False:
            try:
                t = trans.get(block=False)
            except queue.Empty:
                break           
            print('trans: comm = %s, data = %s' % (t['comm'], t['data']))
            if t['comm'] == 'init':
                try:
                    if os.path.exists(path_rank):
                        with open(path_rank, mode='r') as file:
                            vars_rank.set(file.read())
                    if os.path.exists(path_rate):
                        with open(path_rate, mode='r') as file:
                            vars_rate.set(file.read())
                except:
                    pass
            elif t['comm'] == 'rank':
                print(path_rank)
                try:
                    with open(path_rank, mode='w') as file:
                        file.write(t['data'])
                except:
                    pass
                vars_rank.set(t['data'])
            elif t['comm'] == 'rate':
                try:
                    with open(path_rate, mode='w') as file:
                        file.write(t['data'])
                except:
                    pass
                vars_rate.set(t['data'])
            elif t['comm'] == 'exit':
                return

        # sleep till next
        if state == 'null':
            time.sleep(1)
        else:
            time.sleep(.2)

def ExitLoop():
    trans.put({'comm': 'exit', 'data': ''})
    root.destroy()

def SetupGUI(root):
    # frames for layout
    frame0 = tkinter.Frame(root)
    frame1 = tkinter.Frame(frame0)
    frame2 = tkinter.Frame(frame0)
    
    # setup entry box for showing rank history
    font10 = (u'Meiryo', 10)
    label0 = tkinter.Label(frame1, font=font10, width=6, text='Rank: ')
    entry0 = tkinter.Entry(frame1, font=font10, width=30, relief='solid',
                           textvariable=vars_rank)
    button0 = tkinter.Label(frame1, font=font10, width=8, relief='solid',
                            bd=1, bg='#e1e1e1', text=u'クリア')
    button1 = tkinter.Label(frame1, font=font10, width=8, relief='solid',
                            bd=1, bg='#e1e1e1', text=u'更新')
    label0.pack(side='left')
    entry0.pack(side='left', padx=4, expand=1, fill=tkinter.X)
    button1.pack(side='right', padx=4)
    button0.pack(side='right', padx=4)
    
    # setup entry box for showing rate
    label1 = tkinter.Label(frame2, font=font10, width=6, text='Rate: ')
    entry1 = tkinter.Entry(frame2, font=font10, width=8, relief='solid',
                           textvariable=vars_rate)
    button2 = tkinter.Label(frame2, font=font10, width=8, relief='solid',
                            bd=1, bg='#e1e1e1', text=u'更新')
    label1.pack(side='left')
    entry1.pack(side='left', padx=4)
    button2.pack(side='right', padx=4)

    frame1.pack(side='top', padx=4, pady=2, ipadx=4, ipady=0, expand=1,
                fill=tkinter.X, anchor=tkinter.W)
    frame2.pack(side='top', padx=4, pady=2, ipadx=4, ipady=0, expand=1,
                fill=tkinter.X, anchor=tkinter.W)
    frame0.pack(pady=4, expand=1, fill=tkinter.X)

    # setup events of button
    button0.bind("<Enter>", partial(SetColor, button0, '#e5f1fb'))
    button1.bind("<Enter>", partial(SetColor, button1, '#e5f1fb'))
    button2.bind("<Enter>", partial(SetColor, button2, '#e5f1fb'))
    button0.bind("<Leave>", partial(SetColor, button0, '#e1e1e1'))
    button1.bind("<Leave>", partial(SetColor, button1, '#e1e1e1'))
    button2.bind("<Leave>", partial(SetColor, button2, '#e1e1e1'))
    button0.bind("<ButtonPress>",   partial(SetColor, button0, '#cce4f7'))
    button1.bind("<ButtonPress>",   partial(SetColor, button1, '#cce4f7'))
    button2.bind("<ButtonPress>",   partial(SetColor, button2, '#cce4f7'))
    button0.bind("<ButtonRelease>", partial(SetColor, button0, '#e5f1fb'))
    button1.bind("<ButtonRelease>", partial(SetColor, button1, '#e5f1fb'))
    button2.bind("<ButtonRelease>", partial(SetColor, button2, '#e5f1fb'))
    button0.bind("<ButtonPress>", partial(ClrRank), add='+')
    button1.bind("<ButtonPress>", partial(SetRank), add='+')
    button2.bind("<ButtonPress>", partial(SetRate), add='+')

# event functions
def SetColor(wgt, c, ev):
    wgt.configure(background=c)

def ClrRank(ev):
    trans.put({'comm': 'rank', 'data': ''})

def SetRank(ev):
    trans.put({'comm': 'rank', 'data': vars_rank.get()})

def SetRate(ev):
    trans.put({'comm': 'rate', 'data': vars_rate.get()})

if __name__=='__main__':
    # setup path
    path = os.path.join(os.path.dirname(sys.argv[0]))
    path_rank = path + '\\rank.txt'
    path_rate = path + '\\rate.txt'
    path_error = path + '\\error.log'
    error = open(path_error, mode='a')
    hbcap = HsBgCap()

    # setup queue
    trans = queue.Queue()
    trans.put({'comm': 'init', 'data': ''})

    # setup root window
    root = tkinter.Tk()
    root.protocol('WM_DELETE_WINDOW', ExitLoop)
    root.title('Hearthstone BG Capture')
    root.resizable(1, 0)
    vars_rank = tkinter.StringVar()
    vars_rate = tkinter.StringVar()
    SetupGUI(root)

    # setup thread
    mloop = threading.Thread(target=MainLoop)
    mloop.start()
    root.mainloop()