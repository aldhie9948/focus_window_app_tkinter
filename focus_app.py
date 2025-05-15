# %%
import time
import ctypes
import win32gui
import win32con
import win32api
import win32com.client
import threading
import pythoncom
import tkinter 
from tkinter import StringVar, ttk as tk, font , messagebox
import sv_ttk
import pystray
from PIL import Image
import os
import sys
import msvcrt

selected_window = None  # (hwnd, title)
running = True         # apakah thread sedang jalan

# --- Konfigurasi ---
IDLE_THRESHOLD_SECONDS = 10
CHECK_INTERVAL_SECONDS = 2

lock_file = "focus_app.lock"
lock = None

try:
  lock = open(lock_file, "w")
  msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
except Exception as e:
  messagebox.showerror("Error", "Aplikasi sudah berjalan.")
  print("Aplikasi sudah berjalan.")
  sys.exit()


def resource_path(relative_path):
    """Dapatkan path absolut, baik saat dijalankan biasa atau via PyInstaller"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS,"static", relative_path)
    return os.path.join(os.path.abspath("."),"static", relative_path)

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

# %%
class MyApp(tkinter.Tk):
  def __init__(self):
    super().__init__()

    self.icon_path = resource_path("icon.ico")

    self.iconbitmap(self.icon_path)

    self.selected_window_title = StringVar()
    self.selected_window_title.set("Belum ada window yang dipilih.")

    self.h1_font = font.Font(family="Helvetica", size=24, weight="bold")
    self.h2_font = font.Font(family="Helvetica", size=18, weight="bold")
    self.h3_font = font.Font(family="Helvetica", size=14, weight="bold")
    self.small_font = font.Font(family="Helvetica", size=12, weight="normal")

    self.title("Focus Window App")
    self.minsize(width=800, height=600)
    self.maxsize(width=800, height=1080)

    self.windows = enum_window_titles()

    sv_ttk.set_theme(root=self, theme="dark")

    main_frame = tk.Frame(self, padding=50)
    main_frame.pack(fill="both", expand=True)

    tk.Label(main_frame, text="Daftar Aplikasi Berjalan", font=self.h3_font, justify="center", anchor="center").pack(fill="x", pady=10)

    tk.Label(main_frame, text="Pilih aplikasi yang akan difokuskan secara berkala.", font=self.small_font, justify="center", anchor="center").pack(fill="x", pady=10)

    tk.Label(main_frame, textvariable=self.selected_window_title, font=self.h2_font, justify="center", anchor="center", wraplength=800).pack(fill="both", ipady=20, expand=True)

    app_list_frame = tk.Frame(main_frame)
    app_list_frame.pack(fill="both", expand=True)

    for hwnd, title in self.windows:
      tk.Button(app_list_frame, text=title, command=self.make_focus_command(hwnd, title)).pack(fill="x", pady=10)
    
    self.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)

    # init icon
    self.icon = pystray.Icon("focus_window_app_icon")
    self.icon.icon = Image.open(self.icon_path)
    self.icon.menu = pystray.Menu(
      pystray.MenuItem("Show", self.show_window, default=True),
      pystray.MenuItem("Quit", self.quit_window)
    )

    threading.Thread(target=self.icon.run, daemon=True).start()

    self.update_windows_list()

  
  def minimize_to_tray(self):
    self.withdraw()
  
  def show_window(self, _=None):
    self.after(0, self.deiconify())

  def quit_window(self, _=None):
    global running
    running = False
    self.icon.stop()

    try:
      msvcrt.locking(lock.fileno(), msvcrt.LK_UNLCK, 1)
      lock.close()
      os.remove(lock_file)
    except Exception as e:
      print(f"Gagal menghapus file lock: {e}")

    self.destroy()


  def make_focus_command(self, hwnd: int, title: str):
    def func():
        global selected_window
        selected_window = (hwnd, title)  # ubah target, ini akan hentikan thread lama
        self.selected_window_title.set(title)
        thread = threading.Thread(target=set_focused_window, args=(hwnd, title), daemon=True)
        thread.start()
    return func

  def update_windows_list(self):
    self.windows = enum_window_titles()
    self.after(2000, self.update_windows_list)

# %%
if __name__ == "__main__":
  app = MyApp()
  app.mainloop()


