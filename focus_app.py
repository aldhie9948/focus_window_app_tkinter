# %%
import os
import time
import ctypes
import win32gui
import win32con
import win32api
import win32com.client
import threading
import pythoncom
import tkinter 
from tkinter import ttk as tk
from tkinter import font
from tkinter import StringVar
import sv_ttk

selected_window = None  # (hwnd, title)
running = False         # apakah thread sedang jalan


# --- Konfigurasi ---
IDLE_THRESHOLD_SECONDS = 10
CHECK_INTERVAL_SECONDS = 2

def enum_window_titles():
  windows = []

  def callback(hwnd, _):
    # Ambil judul window
    title = win32gui.GetWindowText(hwnd)
    # Cek apakah window visible dan punya judul
    if win32gui.IsWindowVisible(hwnd) and title:
      windows.append((hwnd, title))
  
  win32gui.EnumWindows(callback, None)

  return windows


def get_idle_duration():
    class LASTINPUTINFO(ctypes.Structure):
        _fields_ = [('cbSize', ctypes.c_uint), ('dwTime', ctypes.c_uint)]

    lii = LASTINPUTINFO()
    lii.cbSize = ctypes.sizeof(lii)
    if ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii)):
        millis = win32api.GetTickCount() - lii.dwTime
        return millis / 1000.0
    return 0


def is_window_focused(hwnd: int):
    return hwnd == win32gui.GetForegroundWindow()


def focus_window(hwnd: int):
    pythoncom.CoInitialize()  # Inisialisasi COM di thread baru
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys("%")
    win32gui.SetForegroundWindow(hwnd)
    pythoncom.CoUninitialize()



def set_focused_window(hwnd: int, title: str):
    global running, selected_window

    running = True
    while running:
        if selected_window != (hwnd, title):
            # Jika user memilih window lain, keluar dari loop
            break

        if not is_window_focused(hwnd):
            idle = get_idle_duration()
            print(f"[{title}] not focused. Idle: {idle:.2f}s")
            if idle >= IDLE_THRESHOLD_SECONDS:
                print(f"→ Focusing: {title}")
                focus_window(hwnd)
            else:
                print("→ User is active. Skip.")
        else:
            print(f"[{title}] already focused.")
        
        time.sleep(CHECK_INTERVAL_SECONDS)

    print(f"Thread untuk '{title}' dihentikan.")



def make_focus_command(hwnd: int, title: str):
    def func():
        global selected_window_title, selected_window
        selected_window = (hwnd, title)  # ubah target, ini akan hentikan thread lama
        selected_window_title.set(title)
        thread = threading.Thread(target=set_focused_window, args=(hwnd, title), daemon=True)
        thread.start()
    return func


# %%
root = tkinter.Tk()
root.title("Focus Window App")

root.maxsize(width=800, height=600)
root.minsize(width=800, height=600)

selected_window_title = StringVar()
selected_window_title.set("Belom ada window dipilih.")

def on_close():
    global running
    running = False  # Hentikan loop
    root.destroy()   # Tutup jendela Tkinter

root.protocol("WM_DELETE_WINDOW", on_close)

sv_ttk.set_theme(root=root, theme="dark")

# Definisikan font sekali
h1_font = font.Font(family="Helvetica", size=24, weight="bold")
h2_font = font.Font(family="Helvetica", size=18, weight="bold")
h3_font = font.Font(family="Helvetica", size=14, weight="bold")
small_font = font.Font(family="Helvetica", size=12, weight="normal")

main_frame = tk.Frame(root, padding=20)
main_frame.pack(fill="both", expand=True)



# %%
from utils import enum_window_titles

windows = enum_window_titles()

tk.Label(main_frame, text="Daftar Aplikasi Berjalan", font=h3_font, justify="center", anchor="center").pack(fill="x", pady=10)

tk.Label(main_frame, text="Pilih aplikasi yang akan fokuskan secara berkala.", justify="center", anchor="center", font=small_font).pack(fill="x", pady=10)

tk.Label(main_frame, textvariable=selected_window_title, font=h2_font, anchor="center", justify="center", wraplength=800).pack(fill="both", expand=True, ipady=20)

button_frame = tk.Frame(main_frame, padding=(0, 50, 0, 0)).pack(fill="both")

for hwnd, title in windows:
  button = tk.Button(main_frame,text=title,command=make_focus_command(hwnd, title))
  button.pack(fill="x", pady=10)

# %%
root.mainloop()


