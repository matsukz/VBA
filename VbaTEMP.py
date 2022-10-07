from asyncore import write
from cgitb import text
from multiprocessing import AuthenticationError
import tkinter, tkinter.messagebox
import pyautogui
import pyperclip

root = tkinter.Tk()
root.title("VbaTEMPLATES")
root.geometry('250x300')
root.attributes("-topmost", True)
root.resizable(0,0)

mtxtbox = tkinter.Text(font=("", 16))
mtxtbox.place(x=60, y=10, width=150, height=30)
mlbl = tkinter.Label(text='メイン')
mlbl.place(x=25, y=12)

otxtbox = tkinter.Text(font=("", 16))
otxtbox.place(x=60, y=50, width=150, height=30)
olbl = tkinter.Label(text='オプション')
olbl.place(x=10, y=52)

mckb = tkinter.BooleanVar()
mckb.set(True)

def range():
    if not mtxtbox.get("1.0", "end-1c") == "" :
    
        try:
            main = ""
            msell = mtxtbox.get("1.0", "end-1c")
            osell = otxtbox.get("1.0", "end-1c")
            if not osell == "" :
                Lmain = "Range(\""
                Rmain =  "\")"      

                pyautogui.click(48, 0)
                pyautogui.write(Lmain + msell + Rmain)

                if not osell=="" :
                    pyperclip.copy(osell)
                    pyautogui.hotkey("ctrl", "v")
                else :
                    pass

                pyautogui.press("Return")

        except Exception as e :
            tkinter.messagebox.showerror("ERROR", "例外が発生したため実行できませんでした。")
    
    else :
        tkinter.messagebox.showerror("ERROR", "文字を入力してください。")

def sub():
    count = 0
    if not mtxtbox.get("1.0", "end-1c") == "" :

        try:
            Copy = ""
            pyperclip.copy(mtxtbox.get("1.0", "end-1c"))
            print(Copy)

            pyautogui.click(48, 0)
            pyautogui.write("Sub ")
            pyautogui.hotkey("ctrl", "v")

            pyautogui.write(" ()")
            pyautogui.press("Return")
            pyautogui.press("Tab")

            if mckb.get()==True :
                pyautogui.write("Cells.delet")
            
            pyautogui.press("Return")
        
        except Exception as e :
            tkinter.messagebox.showerror("ERROR", "例外が発生したため続行できません。")
        
    else :
        tkinter.messagebox.showerror("ERROR", "文字を入力してください。")

def fill():
    if not mtxtbox.get("1.0", "end-1c") == "" :

        msell = mtxtbox.get("1.0", "end-1c")
        osell = otxtbox.get("1.0", "end-1c")
        Lmain = "Range(\""
        Nmain = ".Autofill Destination:="
        Rmain = "\")"

        try:
            Copy = ""
            pyperclip.copy(mtxtbox.get("1.0", "end-1c"))

            pyautogui.click(48, 0)
            pyautogui.write(Lmain)

            pyautogui.hotkey("ctrl", "v")
            pyautogui.write(Rmain)

            pyautogui.write(Nmain)
            pyperclip.copy(osell)
            pyautogui.write(Lmain)
            pyautogui.hotkey("ctrl", "v")
            
            pyautogui.write(Rmain)
            pyautogui.press("Return")
        
        except Exception as e :
            tkinter.messagebox.showerror("ERROR", "例外が発生したため続行できません。")
        
    else :
        tkinter.messagebox.showerror("ERROR", "文字を入力してください。")

def mdelete():
    mtxtbox.delete("1.0", "end-1c")
    mtxtbox.focus()

def odelete():
    otxtbox.delete("1.0", "end-1c")
    otxtbox.focus()

Range_button = tkinter.Button(text='Range',command=range,width=16,height=3)
Range_button.place(x=60,y=100)

AutoFill_button = tkinter.Button(text='AutoFill',command=fill,width=16,height=3)
AutoFill_button.place(x=60,y=160)

Sub_button = tkinter.Button(text='Sub',command=sub,width=16,height=3)
Sub_button.place(x=60,y=220)

mdelete_button = tkinter.Button(text="削除",command=mdelete,width=3,height=1)
mdelete_button.place(x=215,y=12)

odelete_button = tkinter.Button(text="削除",command=odelete,width=3,height=1)
odelete_button.place(x=215,y=52)

ckbox = tkinter.Checkbutton(root, variable=mckb, text="Cells.deleteの入力")
ckbox.place(x=60, y=275)

root.mainloop()