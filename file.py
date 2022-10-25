from tkinter import *
from PIL import Image, ImageDraw
import pyautogui
import os
from os import path

width, height= pyautogui.size()

tk = Tk()
tk.attributes('-fullscreen', True)
cvs = Canvas(tk, width=500,height=500)
cvs.configure(bg='white')
cvs.pack(fill=BOTH, expand=True)


img = Image.new('RGB',(width,height),(255,255,255))
draw = ImageDraw.Draw(img)

mousePressed = False
last = None

def press(evt):
    global mousePressed
    mousePressed = True

def release(evt):
    global mousePressed
    mousePressed = False

cvs.bind_all('<ButtonPress-1>', press)
cvs.bind_all('<ButtonRelease-1>', release)

def finish():
    if path.exists('img.png')== True:
        os.remove('img.png')
    img.save("C:\\Users\\user\\Desktop\\img.png")
    tk.destroy()

Button(tk,text='Επόμενο',command=finish).pack()

def move(evt):
    global mousePressed, last
    x,y = evt.x,evt.y
    if mousePressed:
        if last is None:
            last = (x,y)
            return
        draw.line(((x,y),last), (0,0,0))
        cvs.create_line(x,y,last[0],last[1])
        last = (x,y)
    else:
        last = (x,y)

cvs.bind_all('<Motion>', move)

tk.mainloop()
