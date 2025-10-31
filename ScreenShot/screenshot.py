import os
import time
import ctypes
import threading
import tkinter as tk
from tkinter import messagebox
from PIL import ImageGrab, ImageTk
from ctypes import wintypes

# ========= Config =========
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SAVE_DIR = os.path.join(SCRIPT_DIR, "images")
STIPPLE_PATTERN = "gray25"  # gray12 (lighter) / gray50 (darker)

# Hotkeys (global) — we’ll try multiple variants for “/”
MOD_NOREPEAT = 0x4000
MOD_SHIFT = 0x0004

VK_OEM_2   = 0xBF   # main keyboard '/?' on many layouts
VK_DIVIDE  = 0x6F   # numpad '/'
VK_F7      = 0x76
VK_F8      = 0x77
VK_F9      = 0x78   # fallback capture

# Combos we’ll attempt for capture:
CAPTURE_HOTKEY_CANDIDATES = [
    (VK_OEM_2, 0),                     # '/'
    (VK_OEM_2, MOD_SHIFT),             # Shift + '/'
    (VK_DIVIDE, 0),                    # Numpad '/'
    (VK_F9, 0),                        # F9 fallback
]

# ========= DPI Awareness (Windows) =========
def set_dpi_aware():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-monitor (Win 8.1+)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()     # Legacy
        except Exception:
            pass

# ========= Helpers =========
def ensure_save_dir():
    os.makedirs(SAVE_DIR, exist_ok=True)
    return SAVE_DIR

def timestamp_name():
    return time.strftime("%Y%m%d_%H%M%S")

def grab_fullscreen():
    try:
        return ImageGrab.grab(all_screens=True)
    except TypeError:
        return ImageGrab.grab()

def combo_to_text(vk, mod):
    names = []
    if mod & MOD_SHIFT: names.append("Shift")
    key = {VK_OEM_2: " / ", VK_DIVIDE: "Numpad /", VK_F9: "F9"}.get(vk, f"VK_{vk:02X}")
    names.append(key)
    return "+".join(names).replace("  ", " ").strip()

# ========= Global Hotkey Thread =========
class HotkeyThread(threading.Thread):
    """
    Registers global hotkeys (multiple fallbacks).
    Dispatches to Tk via root.after(0, ...).
    """
    def __init__(self, on_capture, on_show, on_hide, tk_root):
        super().__init__(daemon=True)
        self.on_capture = on_capture
        self.on_show = on_show
        self.on_hide = on_hide
        self.tk_root = tk_root
        self.registered_ids = []  # list of (id, vk, mod)

    def run(self):
        user32 = ctypes.windll.user32

        # Register overlay show/hide first
        user32.RegisterHotKey(None, 1002, MOD_NOREPEAT, VK_F7)  # show
        user32.RegisterHotKey(None, 1003, MOD_NOREPEAT, VK_F8)  # hide

        # Try capture variants
        next_id = 2000
        for vk, mod in CAPTURE_HOTKEY_CANDIDATES:
            if user32.RegisterHotKey(None, next_id, mod | MOD_NOREPEAT, vk):
                self.registered_ids.append((next_id, vk, mod))
                print(f"[Hotkey] Registered capture: {combo_to_text(vk, mod)}")
                next_id += 1
            else:
                print(f"[Hotkey] Could not register: {combo_to_text(vk, mod)}")

        if not self.registered_ids:
            print("[Hotkey] No capture hotkeys registered. F9 should have worked—"
                  "if not, another program is using it. Try closing conflicts or running this as admin.")

        msg = wintypes.MSG()
        while True:
            ret = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
            if ret == 0:  # WM_QUIT
                break

            if msg.message == 0x0312:  # WM_HOTKEY
                hotkey_id = msg.wParam
                if hotkey_id == 1002:       # F7
                    self.tk_root.after(0, self.on_show)
                elif hotkey_id == 1003:     # F8
                    self.tk_root.after(0, self.on_hide)
                else:
                    # Any registered capture ID
                    for (rid, vk, mod) in self.registered_ids:
                        if hotkey_id == rid:
                            self.tk_root.after(0, self.on_capture)
                            break

            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

        # Unregister on exit
        user32.UnregisterHotKey(None, 1002)
        user32.UnregisterHotKey(None, 1003)
        for rid, vk, mod in self.registered_ids:
            user32.UnregisterHotKey(None, rid)

# ========= Main Tool =========
class SnippingTool:
    HANDLE_SIZE = 8
    EDGE_TOL = 6
    MIN_W = 20
    MIN_H = 20

    def __init__(self):
        set_dpi_aware()
        ensure_save_dir()

        self.root = tk.Tk()
        self.root.withdraw()  # start hidden so you can keep working
        self.root.title("Snipping Overlay")
        self.root.attributes('-topmost', True)
        self.root.overrideredirect(True)

        # Prepare background and size
        self.bg_img = grab_fullscreen()
        self.W, self.H = self.bg_img.size
        self.root.geometry(f"{self.W}x{self.H}+0+0")

        self.canvas = tk.Canvas(self.root, width=self.W, height=self.H, highlightthickness=0, bd=0)
        self.canvas.pack(fill="both", expand=True)
        self.root.configure(cursor="tcross")

        self.tk_bg = ImageTk.PhotoImage(self.bg_img)
        self.bg_item = self.canvas.create_image(0, 0, image=self.tk_bg, anchor="nw")

        # Initial rectangle (centered)
        margin = min(self.W, self.H) // 6
        self.x1, self.y1 = margin, margin
        self.x2, self.y2 = self.W - margin, self.H - margin

        self.rect = self.canvas.create_rectangle(self.x1, self.y1, self.x2, self.y2,
                                                 outline="#00e0ff", width=2)
        self.handles = {}
        self._create_handles()
        self.update_overlay()

        # State
        self.mode = None
        self.start_x = 0
        self.start_y = 0
        self.orig = None

        # Bindings (when overlay visible)
        self.canvas.bind("<Motion>", self.on_motion)
        self.canvas.bind("<Button-1>", self.on_button1)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.root.bind("<Escape>", lambda e: self.quit_app())

        # Global hotkeys
        hk = HotkeyThread(
            on_capture=self.capture_now,
            on_show=self.show_overlay,
            on_hide=self.hide_overlay,
            tk_root=self.root
        )
        hk.start()

        print("Hotkeys:")
        print("  F7  -> show overlay")
        print("  F8  -> hide overlay")
        print("  '/' or Shift+'/' or Numpad '/' or F9 -> capture")

        # Start Tk
        self.root.mainloop()

    # ---- Overlay control ----
    def show_overlay(self):
        # Refresh background and geometry to current desktop
        self.bg_img = grab_fullscreen()
        self.W, self.H = self.bg_img.size
        self.root.geometry(f"{self.W}x{self.H}+0+0")
        self.tk_bg = ImageTk.PhotoImage(self.bg_img)
        self.canvas.itemconfigure(self.bg_item, image=self.tk_bg)

        self.update_handles()
        self.update_overlay()
        self.root.deiconify()
        self.root.lift()
        self.root.attributes('-topmost', True)

    def hide_overlay(self):
        self.root.withdraw()

    def quit_app(self):
        try:
            ctypes.windll.user32.PostQuitMessage(0)  # stop hotkey loop
        except Exception:
            pass
        self.root.destroy()

    # ---- Rectangle + Handles ----
    def _create_handles(self):
        for key in ["nw", "n", "ne", "e", "se", "s", "sw", "w"]:
            self.handles[key] = self.canvas.create_rectangle(0,0,0,0, fill="#00e0ff", outline="#00e0ff")
        self.update_handles()

    def update_handles(self):
        x1, y1, x2, y2 = self._norm_rect(self.x1, self.y1, self.x2, self.y2)
        cx = (x1 + x2) // 2
        cy = (y1 + y2) // 2
        hs = self.HANDLE_SIZE
        def box(x, y): return (x - hs, y - hs, x + hs, y + hs)
        coords = {
            "nw": box(x1, y1),
            "n" : box(cx, y1),
            "ne": box(x2, y1),
            "e" : box(x2, cy),
            "se": box(x2, y2),
            "s" : box(cx, y2),
            "sw": box(x1, y2),
            "w" : box(x1, cy),
        }
        for k, r in coords.items():
            self.canvas.coords(self.handles[k], *r)
        self.canvas.coords(self.rect, x1, y1, x2, y2)

    def update_overlay(self):
        x1, y1, x2, y2 = self._norm_rect(self.x1, self.y1, self.x2, self.y2)
        if hasattr(self, 'mask_items'):
            for it in self.mask_items:
                self.canvas.delete(it)
        self.mask_items = []
        fill = "black"
        stip = STIPPLE_PATTERN
        self.mask_items.append(self.canvas.create_rectangle(0, 0, self.W, y1, fill=fill, stipple=stip, outline=""))
        self.mask_items.append(self.canvas.create_rectangle(0, y2, self.W, self.H, fill=fill, stipple=stip, outline=""))
        self.mask_items.append(self.canvas.create_rectangle(0, y1, x1, y2, fill=fill, stipple=stip, outline=""))
        self.mask_items.append(self.canvas.create_rectangle(x2, y1, self.W, y2, fill=fill, stipple=stip, outline=""))

    # ---- Hit testing ----
    def _norm_rect(self, x1, y1, x2, y2):
        return (min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))

    def _point_in_rect(self, x, y):
        x1, y1, x2, y2 = self._norm_rect(self.x1, self.y1, self.x2, self.y2)
        return x1 <= x <= x2 and y1 <= y <= y2

    def _edge_hit(self, x, y):
        x1, y1, x2, y2 = self._norm_rect(self.x1, self.y1, self.x2, self.y2)
        tol = self.EDGE_TOL
        near_left   = abs(x - x1) <= tol
        near_right  = abs(x - x2) <= tol
        near_top    = abs(y - y1) <= tol
        near_bottom = abs(y - y2) <= tol
        if near_left and near_top: return "nw"
        if near_right and near_top: return "ne"
        if near_right and near_bottom: return "se"
        if near_left and near_bottom: return "sw"
        if near_top: return "n"
        if near_bottom: return "s"
        if near_left: return "w"
        if near_right: return "e"
        return None

    def _cursor_for(self, where):
        return {
            "n": "size_ns",
            "s": "size_ns",
            "e": "size_we",
            "w": "size_we",
            "ne": "size_ne_sw",
            "sw": "size_ne_sw",
            "nw": "size_nw_se",
            "se": "size_nw_se",
            "move": "fleur",
            None: "tcross"
        }.get(where, "tcross")

    # ---- Events (when overlay visible) ----
    def on_motion(self, event):
        where = self._edge_hit(event.x, event.y)
        if where:
            self.canvas.configure(cursor=self._cursor_for(where))
        elif self._point_in_rect(event.x, event.y):
            self.canvas.configure(cursor=self._cursor_for("move"))
        else:
            self.canvas.configure(cursor=self._cursor_for(None))

    def on_button1(self, event):
        where = self._edge_hit(event.x, event.y)
        self.start_x, self.start_y = event.x, event.y
        self.orig = (self.x1, self.y1, self.x2, self.y2)
        if where:
            self.mode = where
        elif self._point_in_rect(event.x, event.y):
            self.mode = "move"
        else:
            self.mode = "new"
            self.x1 = self.x2 = event.x
            self.y1 = self.y2 = event.y
            self.update_handles()
            self.update_overlay()

    def on_drag(self, event):
        x = max(0, min(event.x, self.W))
        y = max(0, min(event.y, self.H))
        ox1, oy1, ox2, oy2 = self.orig

        if self.mode == "move":
            dx, dy = x - self.start_x, y - self.start_y
            nx1, ny1, nx2, ny2 = ox1 + dx, oy1 + dy, ox2 + dx, oy2 + dy
            w = nx2 - nx1
            h = ny2 - ny1
            nx1 = max(0, min(nx1, self.W - w))
            ny1 = max(0, min(ny1, self.H - h))
            nx2, ny2 = nx1 + w, ny1 + h
            self.x1, self.y1, self.x2, self.y2 = nx1, ny1, nx2, ny2

        elif self.mode in ("n","s","e","w","ne","nw","se","sw"):
            x1, y1, x2, y2 = ox1, oy1, ox2, oy2
            if "n" in self.mode: y1 = min(y, y2 - self.MIN_H)
            if "s" in self.mode: y2 = max(y, y1 + self.MIN_H)
            if "w" in self.mode: x1 = min(x, x2 - self.MIN_W)
            if "e" in self.mode: x2 = max(x, x1 + self.MIN_W)
            x1, x2 = max(0, x1), min(self.W, x2)
            y1, y2 = max(0, y1), min(self.H, y2)
            self.x1, self.y1, self.x2, self.y2 = x1, y1, x2, y2

        elif self.mode == "new":
            self.x2, self.y2 = x, y

        self.update_handles()
        self.update_overlay()

    def on_release(self, event):
        self.mode = None
        self.orig = None

    # ---- Capture (callable from global hotkey) ----
    def capture_now(self):
        x1, y1, x2, y2 = map(int, self._norm_rect(self.x1, self.y1, self.x2, self.y2))
        if x2 - x1 < self.MIN_W or y2 - y1 < self.MIN_H:
            try:
                if self.root.state() == 'normal':
                    messagebox.showwarning("Too small", "Selection is too small to capture.")
            except Exception:
                pass
            return

        # Hide overlay to avoid borders in the image
        was_visible = (self.root.state() == 'normal')
        try:
            self.root.withdraw()
            self.root.update_idletasks()
            time.sleep(0.12)  # let the OS repaint
        except Exception:
            pass

        # Capture the bbox directly
        try:
            img = ImageGrab.grab(bbox=(x1, y1, x2, y2), all_screens=True)  # Pillow ≥ 9.2
        except TypeError:
            full = grab_fullscreen()
            img = full.crop((x1, y1, x2, y2))

        save_dir = ensure_save_dir()
        base = f"capture_{timestamp_name()}.png"
        path = os.path.join(save_dir, base)
        counter = 1
        while os.path.exists(path):
            path = os.path.join(save_dir, f"capture_{timestamp_name()}_{counter:03d}.png")
            counter += 1
        img.save(path)

        # Restore overlay if it was visible before
        if was_visible:
            try:
                self.root.deiconify()
                self.root.lift()
                self.root.attributes('-topmost', True)
                messagebox.showinfo("Saved", f"Saved to:\n{path}")
            except Exception:
                pass

if __name__ == "__main__":
    SnippingTool()
