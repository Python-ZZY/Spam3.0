try:
    import pyi_splash
    pyi_splash.close()
except:
    pass

from tkinter import *
from tkinter.scrolledtext import *
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from functools import partial
import os
import pickle
import sys
import random
import time
import threading
import pyautogui
import win32gui

from scripts.configs import *
from scripts.style import *

__version__ = "3.0.1 Alpha"

pyautogui.PAUSE = 0

def path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")

    return os.path.normpath(os.path.join(base_path, relative_path))

class Thread(threading.Thread):
    def __init__(self, target=None):
        super().__init__()
        self.flag = threading.Event()
        self.flag.clear()
        self.target = target

    def run(self):
        while True:
            self.flag.wait()
            self.target()
            
class StatusBar(Frame):
    def __init__(self, master, **kw):
        setattr(master, "statusbar", self)
        
        super().__init__(master, **kw)
        self.labels = {}
        self.set_label("", "")

    def set_label(self, name, text="", side="left", width=0):
        if name not in self.labels:
            label = Label(self, borderwidth=0, anchor="w")
            label.pack(side=side, pady=0, padx=4)
            self.labels[name] = label
        else:
            label = self.labels[name]
        if width != 0:
            label.config(width=width)
        label.config(text=text)
        
class Button(ttk.Button):
    def __init__(self, master, tip=None, statusbar=None, **kw):
        super().__init__(master, cursor="hand2", **kw)

        if tip and statusbar:
            bind_status(self, tip, statusbar)

class NumBox(Spinbox):
    def __init__(self, master, allowneg1=False, value=None, **kw):
        self.allowneg1 = allowneg1
        self.value = value

        super().__init__(master, value=value, **kw)
        self.bind("<FocusOut>", self._vcmd)

    def get(self):
        res = int(super().get())
        self.insert(res)
        return res
    
    def insert(self, text):
        self.delete(0, "end")
        super().insert(0, text)
        
    def _vcmd(self, x):
        try:
            value = self.get()
            assert value >= 0
        except ValueError:
            self.insert(self.value)
        except AssertionError:
            if self.allowneg1 and self.get() == "-1":
                return
            self.insert(self.value)

class SelectKeyToplevel(Toplevel):
    def __init__(self, master, keybox):
        super().__init__(master)
        self.title("选择按键...")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()
        self.focus_force()
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.bind("<Key>", self.add_key)

        self.hotkeys = []

        self.keylabel = Entry(self, relief="groove", state="disabled", cursor="arrow")
        self.keylabel.grid(row=1, column=0, columnspan=3, padx=40, pady=20)

        self.okbutton = ttk.Button(self, text="确定", command=self.ok)
        self.okbutton.grid(row=2, column=0, padx=5, pady=20)

        self.cancelbutton = ttk.Button(self, text="取消", command=self.cancel)
        self.cancelbutton.grid(row=2, column=1, padx=5, pady=20)

        self.rbutton = ttk.Button(self, text="重置", command=self.reset)
        self.rbutton.grid(row=2, column=2, padx=5, pady=20)

        self.wait_window(self)

    def add_key(self, event):
        if len(self.hotkeys) < 3:
            if event.char not in ("", " ", "\b", "\r"):
                if event.char not in self.hotkeys and event.char in chars:
                    self.hotkeys.append(event.char)
            else:
                keysym = event.keysym.replace("Control", "Ctrl")\
                         .replace("Return", "Enter")\
                         .replace("space", "Space")
                if keysym.endswith("_L") or keysym.endswith("_R"):
                    keysym = keysym.replace("_L", "").replace("_R", "")
                if keysym in hotkeylist:
                    if keysym not in self.hotkeys:
                        self.hotkeys.append(keysym)
                    else:
                        if event.keysym == "Return":
                            self.ok()
                            return
        else:
            if event.keysym == "Return":
                self.ok()
                return
        self.update_entry()

    def update_entry(self):
        self.keylabel.config(state="normal")
        self.keylabel.delete(0, "end")
        for key in self.hotkeys:
            self.keylabel.insert("end", key)
            if key != self.hotkeys[-1]:
                self.keylabel.insert("end", "+")
        self.keylabel.config(state="disabled")
        
    def ok(self):
        self.grab_release()
        self.destroy()

    def cancel(self):
        self.hotkeys = False
        self.ok()

    def reset(self):
        self.hotkeys = []
        self.update_entry()
        
class KeyBox(Spinbox):
    def __init__(self, master, key="<Enter>"):
        super().__init__(master, cursor="hand2", wrap=True, justify="center",
                         state="readonly", values=("<Enter>", "<Ctrl-Enter>"))
        self.insert(key)
        self.bind("<Button-1>", self.select_key)
        
    def get(self):
        return super().get()[1:-1].split("-")
    
    def insert(self, text):
        self.config(state="normal")
        self.delete(0, "end")
        super().insert(0, text)
        self.config(state="readonly")

    def select_key(self, event):
        if self.identify(event.x, event.y) == "entry":
            key = SelectKeyToplevel(self.master, self)
            if key.hotkeys != False:
                self.insert("<" + "-".join(key.hotkeys) + ">")

class Info(Toplevel):
    def __init__(self):
        super().__init__(root)
        self.title(APPNAME)
        self.iconbitmap(ICONNAME)
        self.transient(root)
        self.resizable(False, False)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.config(bg="#fff")

        self.image = PhotoImage(file=path("assets/about.ppm"))
        Label(self, image=self.image, bg="#fff").pack()

        Label(self, text=INTRO, bg="#fff").pack(pady=10)
        Label(self, text=ABOUT, justify="left", bg="#fff").pack(pady=10)

    def cancel(self):
        self.grab_release()
        self.destroy()
        
APPNAME = "刷屏神器"
ICONNAME = path("assets/icon.ico")
INTRO = "一款简易的、功能多样的自动化GUI键盘输入控制工具"
ABOUT = f"""版本: {__version__}
开发者: PythonZZY
邮箱: pythonzzy@foxmail.com
Github: https://github.com/Python-ZZY/"""

hotkeylist = ("Enter", "Ctrl", "Alt", "Shift", "F1", "F2", "F3", "F4",
              "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12")
settings = {"循环倒计时(ms)":(NumBox, {"value":2500, "from_":0, "to":1000000}, "设置开始循环前等待的时间"),
            "循环次数":(NumBox, {"allowneg1":True, "value":-1, "from_":-1, "to":1000000}, "循环次数，-1表示无限"),
            "单次循环延迟(ms)":(NumBox, {"value":600, "from_":0, "to":1000000}, "循环执行一次后等待的时间"),
            "单次循环按键":(KeyBox, {}, "在分隔符位置或单次循环后按下的按键")}
variables = {
    "amount":("amount", "已循环次数"),
    "surplus":("surplus", "剩余循环次数"),
    "total":("total", "循环总次数"),
    "random":("round(random.random(), 3)", "随机返回0-1之间的一个三位小数"),
    "randint":("random.randint(0, 9)", "随机返回0-9之间的一个整数"),
    "date":("datetime.now().strftime('%Y-%m-%d')", "当前日期"),
    "time":("datetime.now().strftime('%H:%M:%S')", "当前时间"),
    " & ":(None, "分组运算符：将分割的片段进行分组，作为多条消息输出"),
    " / ":(None, "随机运算符：在当前消息组中，将多个消息片段分割并随机输出"),
    " # ":(None, "跳过运算符：在当前消息组的消息片段中，跳过输出后面的消息"),
    }
chars = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

def status_enter(text, x):
    if not keyboardthread.flag.isSet():
        statusbar.set_label("", text=text)

def status_leave(x):
    if not keyboardthread.flag.isSet():
        statusbar.set_label("", text="")
    
def bind_status(widget, tip, statusbar):
    widget.bind("<Enter>", partial(status_enter, tip))
    widget.bind("<Leave>", status_leave)
    
def insert_var(name):
    if variables[name][0]:
        name = "%" + name + "%"
    stext.window_create("insert", window=Label(stext, text=name,
                                               fg="#0000ff", bg="#ffffff"))
    stext.focus_set()
    stext.see("insert")

def get_text(amount, surplus, total):
    res = [[""]]
    for t in export_file_obj():
        if t[0] == "window":
            var_text = t[1].strip()
            if var_text.startswith("%"):
                if type(res[-1][-1]) == str:
                    res[-1][-1] += str(eval(variables[var_text[1:-1]][0]))
            else:
                if var_text == "&":
                    res.append([""])
                elif var_text == "/":
                    res[-1].append("")
                elif var_text == "#":
                    if res[-1][-1] == "":
                        res[-1][-1] = None
        else:
            res[-1][-1] += t[1]

    for text_list in res:
        yield random.choice(text_list)
                
def send():
    run_button.config(text="终止", command=keyboardthread.flag.clear)
    root.title(APPNAME + "(运行中)")
    
    pretime, total, delay, hotkeys = [w.get() for w in setting_list]

    for i in range(pretime // 10):
        statusbar.set_label("", f"运行倒计时{round((pretime / 10 - i)/100, 2)}s")
        time.sleep(0.01)
        if not keyboardthread.flag.isSet():
            break
    else:
        status_label = "循环总次数:%d 已循环次数:{} 剩余循环次数:{}"%total
        amount = 1
        surplus = total
        text = get_text(amount, surplus, total)
        while total == -1 or amount <= total:
            if not keyboardthread.flag.isSet():
                break

            try:
                txt = next(text)
                if txt == None:
                    continue
            except StopIteration:
                amount += 1
                surplus = -1 if total == -1 else (total - amount)
                text = get_text(amount, surplus, total)
            else:
                root.clipboard_clear()
                root.clipboard_append(txt)
                pyautogui.hotkey("ctrl", "v")
                pyautogui.hotkey(*hotkeys)

                statusbar.set_label("", status_label.format(amount, surplus))
                
                for i in range(delay // 10):
                    statusbar.set_label("", status_label.format(amount, surplus) + \
                                        f" 下一次输出倒计时:{round((delay / 10 - i)/100, 2)}s")
                    time.sleep(0.01)
                    if not keyboardthread.flag.isSet():
                        break

    statusbar.set_label("", "")
    keyboardthread.flag.clear()
    run_button.config(text="运行", command=keyboardthread.flag.set)
    root.title(APPNAME + "(空闲)")

def export_file_obj():
    text_dump = stext.dump("1.0", "end", window=True, text=True)
    if text_dump[-1][1] == "\n":
        text_dump = text_dump[:-1]

    for t in text_dump:
        if t[0] == "window":
            yield t[0], root.nametowidget(t[1]).cget("text")
        else:
            yield t
        
def load_file_obj(text_dump):
    for t in text_dump:
        if t[0] == "window":
            if t[1].startswith("%"):
                insert_var(t[1][1:-1])
            else:
                insert_var(t[1])
        else:
            stext.insert("end", t[1])
            
def file_open(x=None):
    name = filedialog.askopenfilename(
        defaultextension=".sm",
        filetypes=[("刷屏消息模板文件",".sm")]
        )
    if name:
        if not messagebox.askokcancel("提示", "确认用载入文件替代当前文本？"):
            return
        stext.delete("1.0", "end")
        try:
            load_file_obj(pickle.load(open(name, "rb")))
        except:
            messagebox.showerror("错误", "文件载入失败！")

def file_export(x=None):
    name = filedialog.asksaveasfilename(
        defaultextension=".sm",
        filetypes=[("刷屏消息模板文件",".sm")]
        )
    if name:
        file = open(name, "wb")
        pickle.dump(tuple(export_file_obj()), file)

def win_topmost(*x):
    hwnd = int(root.frame(), base=16)
    geometry = root.geometry()

    if win_topmost_var.get():
        win32gui.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 594)
    else:
        win32gui.SetWindowPos(hwnd, -2, 0, 0, 0, 0, 594)
    root.geometry(geometry)

def help_about():
    Info()

def help_docs(x=None):
    pass

root = Tk(className="")
root.title(APPNAME + "(空闲)")
root.iconbitmap(default=ICONNAME)
root.update()
configure_style()

root.bind("<Control-O>", file_open)
root.bind("<Control-o>", file_open)
root.bind("<Control-E>", file_export)
root.bind("<Control-e>", file_export)
root.bind("<Control-Q>", lambda x: sys.exit())
root.bind("<Control-q>", lambda x: sys.exit())
root.bind("<F1>", help_docs)
root.bind("<F2>", lambda x:win_topmost_var.set(not win_topmost_var.get()))
root.bind("<F3>", lambda x: root.iconify())

menubar = Menu(root)
root.config(menu=menubar)

file_menu = Menu(menubar, tearoff=False)
menubar.add_cascade(label="文件", menu=file_menu)

file_menu.add_command(label="导入", command=file_open, accelerator="Ctrl+O")
file_menu.add_command(label="导出", command=file_export, accelerator="Ctrl+E")
file_menu.add_separator()
file_menu.add_command(label="退出", command=sys.exit, accelerator="Ctrl+Q")

win_menu = Menu(menubar, tearoff=False)
menubar.add_cascade(label="窗口", menu=win_menu)

win_topmost_var = BooleanVar()
win_topmost_var.trace_add("write", win_topmost)
win_menu.add_checkbutton(label="置顶", variable=win_topmost_var, accelerator="F2")
win_menu.add_command(label="最小化", command=root.iconify, accelerator="F3")

help_menu = Menu(menubar, tearoff=False)
menubar.add_cascade(label="帮助", menu=help_menu)

help_menu.add_command(label="关于...", command=help_about)
help_menu.add_separator()
help_menu.add_command(label="帮助文档", command=help_docs, accelerator="F1")

settings_f = ttk.LabelFrame(root, text="设置", padding=10)
settings_f.pack(fill="x", padx=10, pady=10)

stext = ScrolledText(root, width=50, height=15, undo=True, autoseparators=False)
stext.pack(fill="x", padx=10, pady=5)
stext.bind("<Key>", lambda x: stext.edit_separator())
stext.focus_force()

variables_f = ttk.LabelFrame(root, text="变量&运算符", padding=10)
variables_f.pack(fill="x", padx=10, pady=5)
variables_f_text = Text(variables_f, bg="SystemButtonFace", relief="flat", height=6,
                        selectbackground="SystemButtonFace", cursor="arrow")
variables_f_text.pack(fill="both")

statusbar = StatusBar(root)
statusbar.pack(side="bottom", anchor="w")

setting_list = []
for i, (name, (Widg, cnf, tip)) in enumerate(settings.items()):
    label = Label(settings_f, text=name+":")
    label.grid(row=i, column=0, sticky="w", pady=2)
    bind_status(label, tip, statusbar)
    
    w = Widg(settings_f, **cnf)
    w.grid(row=i, column=1, sticky="e", pady=2)
    setting_list.append(w)

has_split = False
for i, (name, (code, helper)) in enumerate(variables.items()):
    b = Button(variables_f_text, tip=helper, statusbar=statusbar, text=name,
               command=partial(insert_var, name))

    if not (code or has_split):
        variables_f_text.insert("end", "\n")
        has_split = True
        
    variables_f_text.window_create("end", window=b)

variables_f_text.config(state="disabled")

keyboardthread = Thread(target=send)
keyboardthread.setDaemon(True)
keyboardthread.start()

run_button = Button(root, text="运行", command=keyboardthread.flag.set)
run_button.pack(padx=10, pady=5)

mainloop()
