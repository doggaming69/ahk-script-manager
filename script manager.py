"""
AHK Manager
pip install customtkinter keyboard pystray Pillow
"""
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, colorchooser, messagebox
import subprocess, json, os, sys, threading, time, re
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageTk, ImageSequence

try:
    import keyboard as kb
    HAS_KEYBOARD = True
except ImportError:
    HAS_KEYBOARD = False

try:
    import pystray
    HAS_TRAY = True
except ImportError:
    HAS_TRAY = False

# ── Paths ──────────────────────────────────────────────────────────────────────
DATA_DIR = Path(os.getenv("APPDATA", "~")) / "AhkManager"
CFG_FILE = DATA_DIR / "config.json"
THEME_FILE = DATA_DIR / "theme.json"
DATA_DIR.mkdir(parents=True, exist_ok=True)

# ── AHK detection ──────────────────────────────────────────────────────────────
def _find_all_ahk() -> dict:
    found, seen = {}, set()
    roots = [
        r"C:\Program Files\AutoHotkey",
        r"C:\Program Files (x86)\AutoHotkey",
        r"C:\AutoHotkey",
        os.path.expanduser(r"~\AutoHotkey"),
    ]
    def _add(path: str):
        path = os.path.normpath(path)
        if path in seen or not os.path.isfile(path): return
        seen.add(path)
        name = Path(path).name
        lower = path.lower()
        ver = "v2" if "\\v2\\" in lower or "/v2/" in lower else "v1"
        bits = "64" if re.search(r"64|u64", name, re.I) else "32"
        label = f"AHK {ver} ({bits}-bit) — {name}"
        base, n = label, 1
        while label in found:
            label = f"{base} [{n}]"; n += 1
        found[label] = path
    for root in roots:
        if not os.path.isdir(root): continue
        for dp, _, files in os.walk(root):
            for f in files:
                if f.lower().startswith("autohotkey") and f.lower().endswith(".exe"):
                    _add(os.path.join(dp, f))
    for directory in os.environ.get("PATH", "").split(os.pathsep):
        for name in ("AutoHotkey.exe","AutoHotkey64.exe","AutoHotkeyU64.exe",
                     "AutoHotkeyU32.exe","AutoHotkey32.exe"):
            _add(os.path.join(directory, name))
    return found

AHK_EXES: dict = _find_all_ahk()

def _detect_ver(script_path: str) -> str:
    try:
        with open(script_path, encoding="utf-8", errors="ignore") as f:
            head = "".join(f.readline() for _ in range(40))
        hl = head.lower()
        if "#requires autohotkey v2" in hl: return "v2"
        if "#requires autohotkey v1" in hl: return "v1"
        v2 = (hl.count("msgbox(") + hl.count("inputbox(") +
              len(re.findall(r'\b\w+\s*\(.*\)\s*=>', hl)) +
              len(re.findall(r'\bclass\s+\w+', hl)) +
              len(re.findall(r'\bMap\s*\(', hl)))
        v1 = (hl.count("msgbox,") + hl.count("inputbox,") +
              len(re.findall(r'#noenv|#persistent|#singleinstance|gosub,|goto,', hl)))
        return "v2" if v2 > v1 else "v1"
    except:
        return "v1"

def _best_exe(ver: str) -> str:
    for prefer64 in (True, False):
        for label, path in AHK_EXES.items():
            if ver not in label.lower(): continue
            if prefer64 == ("64" in label) and os.path.isfile(path): return path
    for path in AHK_EXES.values():
        if os.path.isfile(path): return path
    return ""

# ── Theme ──────────────────────────────────────────────────────────────────────
DEFAULT_THEME = {
    "accent":"#DC3232","bg":"#0F0F12","surface":"#16161C","surface2":"#1E1E28",
    "border":"#2A2A36","text":"#E1E1EB","text_dim":"#6E6E8A","green":"#3CC864",
    "bg_image":"","bg_opacity":35,
}
COLOR_PRESETS = [
    ("Red","#DC3232","#0F0A0A"), ("Blue","#2882F0","#0A0C12"),
    ("Green","#32C85A","#0A0F0C"),("Purple","#A050F0","#0D0A12"),
    ("Orange","#F08228","#120E0A"),("Pink","#E650A0","#120A0E"),
    ("Cyan","#28D4D4","#0A0F0F"), ("White","#D2D2D8","#0E0E10"),
]

def load_theme():
    try:
        if THEME_FILE.exists():
            t = dict(DEFAULT_THEME)
            t.update(json.loads(THEME_FILE.read_text()))
            t.pop("logo_path", None)
            return t
    except:
        pass
    return dict(DEFAULT_THEME)

def save_theme(t):
    THEME_FILE.write_text(json.dumps(
        {k: v for k, v in t.items() if k != "logo_path"},
        indent=2
    ))

def hex_to_rgb(h):
    h = h.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

def shift_color(h, amt):
    r, g, b = hex_to_rgb(h)
    if (r + g + b) // 3 < 128:
        return f"#{min(r+amt,255):02x}{min(g+amt,255):02x}{min(b+amt,255):02x}"
    return f"#{max(r-amt,0):02x}{max(g-amt,0):02x}{max(b-amt,0):02x}"

def blend_color(a, b, t):
    ra, ga, ba = hex_to_rgb(a)
    rb, gb, bb = hex_to_rgb(b)
    return f"#{int(ra + (rb - ra) * t):02x}{int(ga + (gb - ga) * t):02x}{int(ba + (bb - ba) * t):02x}"

def _cover_scale(img: Image.Image, w: int, h: int) -> Image.Image:
    """Stretch img to exactly (w, h) so it always fills the window completely."""
    if img.width == 0 or img.height == 0: return img
    return img.resize((w, h), Image.LANCZOS)

# ── Script model ───────────────────────────────────────────────────────────────
class AhkScript:
    def __init__(self, path, hotkey="", mods="", ahk_exe=""):
        self.path = path
        self.name = Path(path).stem
        self.hotkey = hotkey
        self.mods = mods
        self.ahk_exe = ahk_exe
        self.process = None
        self.last_toggled = None
        self._hotkey_combo = None  # normalized combo string for this script

    @property
    def is_running(self):
        return self.process is not None and self.process.poll() is None

    @property
    def hotkey_display(self):
        if not self.hotkey: return "—"
        return "+".join([m.capitalize() for m in self.mods.split("+") if m]
                        + [self.hotkey.upper()])

    def to_dict(self):
        return {"path": self.path, "hotkey": self.hotkey, "mods": self.mods, "ahk_exe": self.ahk_exe}

    @classmethod
    def from_dict(cls, d):
        return cls(d.get("path", ""), d.get("hotkey", ""), d.get("mods", ""), d.get("ahk_exe", ""))

# ── Script manager ─────────────────────────────────────────────────────────────
class ScriptManager:
    def __init__(self):
        self.scripts: list[AhkScript] = []
        self._hotkey_callbacks = {}  # combo → actual callback function

    def resolve_exe(self, s):
        if s.ahk_exe and os.path.isfile(s.ahk_exe): return s.ahk_exe
        return _best_exe(_detect_ver(s.path))

    def remove(self, s):
        self.stop(s)
        self.unregister_hotkey(s)
        self.scripts.remove(s)

    def toggle(self, s):
        if s.is_running:
            self.stop(s)
        else:
            self.start(s)

    def start(self, s):
        if s.is_running: return
        exe = self.resolve_exe(s)
        if not exe:
            raise RuntimeError("No AutoHotkey executable found.\nInstall from https://www.autohotkey.com")
        if not os.path.exists(s.path):
            raise FileNotFoundError(f"Script not found:\n{s.path}")
        s.process = subprocess.Popen([exe, s.path], creationflags=subprocess.CREATE_NO_WINDOW)
        s.last_toggled = time.strftime("%H:%M:%S")

    def stop(self, s):
        if not s.is_running: return
        try:
            s.process.kill()
        except:
            pass
        s.process = None
        s.last_toggled = time.strftime("%H:%M:%S")

    def stop_all(self):
        for s in self.scripts:
            self.stop(s)

    def register_hotkey(self, s, cb):
        if not HAS_KEYBOARD or not s.hotkey:
            return
        self.unregister_hotkey(s)  # Remove any previous registration for this script

        mods = s.mods.lower() if s.mods else ""
        key = s.hotkey.lower()
        combo = f"{mods}+{key}" if mods else key
        s._hotkey_combo = combo

        def wrapped():
            cb()

        try:
            kb.add_hotkey(combo, wrapped, suppress=False)
            self._hotkey_callbacks[combo] = wrapped
        except Exception as e:
            print(f"Failed to register hotkey '{combo}': {e}")
            s._hotkey_combo = None

    def unregister_hotkey(self, s):
        if not HAS_KEYBOARD or not s._hotkey_combo:
            return
        combo = s._hotkey_combo
        if combo in self._hotkey_callbacks:
            try:
                kb.remove_hotkey(combo)
            except:
                pass
            del self._hotkey_callbacks[combo]
        s._hotkey_combo = None

    def save(self):
        CFG_FILE.write_text(json.dumps([s.to_dict() for s in self.scripts], indent=2))

    def load(self):
        if not CFG_FILE.exists(): return
        try:
            for d in json.loads(CFG_FILE.read_text()):
                if os.path.exists(d.get("path", "")):
                    self.scripts.append(AhkScript.from_dict(d))
        except:
            pass

# ── Hotkey dialog ──────────────────────────────────────────────────────────────
class HotkeyDialog(ctk.CTkToplevel):
    def __init__(self, parent, theme, hk="", mods=""):
        super().__init__(parent)
        self.theme = theme
        self.result_key = hk
        self.result_mods = mods
        self.confirmed = False
        self.listening = False
        self.title("Bind Key")
        self.geometry("400x250")
        self.resizable(False, False)
        self.configure(fg_color=theme["bg"])
        self.grab_set()
        self.focus_force()

        hdr = ctk.CTkFrame(self, fg_color=theme["surface"], corner_radius=0, height=44)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="BIND KEY", font=("Segoe UI", 11, "bold"),
                     text_color=theme["accent"]).pack(side="left", padx=16)

        ctk.CTkLabel(self,
            text="Click the box, then press any key combination.\n"
                 "Ctrl / Alt / Shift supported. Esc to cancel.",
            font=("Segoe UI", 9), text_color=theme["text_dim"], justify="left"
        ).pack(anchor="w", padx=16, pady=(12, 0))

        self.cap_var = tk.StringVar(value=self._fmt())
        self.cap_btn = ctk.CTkButton(self, textvariable=self.cap_var,
            font=("Consolas", 13, "bold"), fg_color=theme["surface2"],
            hover_color=theme["border"], border_color=theme["border"], border_width=1,
            text_color=theme["text"], height=52, corner_radius=4, command=self._start)
        self.cap_btn.pack(fill="x", padx=16, pady=12)

        row = ctk.CTkFrame(self, fg_color="transparent")
        row.pack(fill="x", padx=16)
        ctk.CTkButton(row, text="Apply", width=110, height=36, fg_color=theme["accent"],
                      font=("Segoe UI", 10, "bold"), command=self._apply).pack(side="left")
        ctk.CTkButton(row, text="Clear", width=110, height=36, fg_color=theme["surface2"],
                      command=self._clear).pack(side="left", padx=8)
        ctk.CTkButton(row, text="Cancel", width=110, height=36, fg_color=theme["surface2"],
                      command=self.destroy).pack(side="left")

        self.bind("<KeyPress>", self._key)

    def _start(self):
        self.listening = True
        self.cap_var.set("Press keys…")
        self.cap_btn.configure(border_color=self.theme["accent"],
                               text_color=self.theme["text_dim"])
        self.focus_force()

    def _key(self, e):
        if not self.listening: return
        k = e.keysym
        if k in ("Control_L", "Control_R", "Shift_L", "Shift_R",
                 "Alt_L", "Alt_R", "Super_L", "Super_R"): return
        if k == "Escape":
            self.listening = False
            self.cap_var.set(self._fmt())
            return
        mods = []
        if e.state & 0x4: mods.append("ctrl")
        if e.state & 0x1: mods.append("shift")
        if e.state & 0x20000: mods.append("alt")
        k = k.lower()
        if not (len(k) == 1 or (k.startswith("f") and k[1:].isdigit())):
            k = k.capitalize()
        self.result_key = k
        self.result_mods = "+".join(mods)
        self.listening = False
        self.cap_btn.configure(border_color=self.theme["border"], text_color=self.theme["text"])
        self.cap_var.set(self._fmt())

    def _fmt(self):
        if not self.result_key: return "Click here to capture"
        return " + ".join([m.capitalize() for m in self.result_mods.split("+") if m]
                          + [self.result_key.upper()])

    def _apply(self):
        self.confirmed = True
        self.destroy()

    def _clear(self):
        self.result_key = ""
        self.result_mods = ""
        self.cap_var.set("Click here to capture")

# =============================================================================
# MAIN APP
# =============================================================================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.theme = load_theme()
        self.manager = ScriptManager()
        self.manager.load()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        self.title("AHK Manager")
        self.geometry("980x600")
        self.minsize(780, 480)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._raw_path = ""
        self._raw_frames = []
        self._raw_delays = []
        self._is_gif = False
        self._scaled_frames = []
        self._scaled_delays = []
        self._scale_gen = 0
        self._gif_idx = 0
        self._gif_job = None
        self._gif_active = False
        self._app_focused = True
        self._bg_label = None
        self._current_img = None
        self._crop_entries = []
        self._crop_photos = {}
        self._crop_items = {}
        self._resize_job = None
        self._last_wh = (0, 0)

        self._build()
        self._reregister_hotkeys()
        self._poll()

        if HAS_TRAY:
            threading.Thread(target=self._build_tray, daemon=True).start()

    def _build(self):
        self._stop_playback()
        for w in self.winfo_children(): w.destroy()
        self._bg_label = None
        self._current_img = None
        self._crop_entries = []
        self._crop_photos = {}
        self._crop_items = {}
        self._resize_job = None
        self._last_wh = (0, 0)

        t = self.theme
        self.configure(fg_color=t["bg"])

        # Background label (lowest layer)
        self._bg_label = ctk.CTkLabel(self, text="", fg_color=t["bg"])
        self._bg_label.place(x=0, y=0, relwidth=1, relheight=1)
        self._bg_label.lower()

        # Tab bar
        TAB_H = 38
        tab_bar = tk.Frame(self, bg=t["bg"], height=TAB_H, bd=0, highlightthickness=0)
        tab_bar.pack(fill="x", side="top")
        tab_bar.pack_propagate(False)

        self._tab_widgets = {}
        for key, label in [("scripts", "Scripts"), ("settings", "Customization")]:
            lbl = tk.Label(tab_bar, text=label, font=("Segoe UI", 9),
                           bg=t["bg"], fg=t["text_dim"], cursor="hand2",
                           padx=20, activebackground=t["bg"], activeforeground=t["text"])
            lbl.pack(side="left", fill="y")
            lbl.bind("<Button-1>", lambda e, k=key: self._show(k))
            self._tab_widgets[key] = lbl

        tk.Frame(self, bg=t["border"], height=1, bd=0, highlightthickness=0).pack(fill="x", side="top")

        # Body
        self._body = tk.Frame(self, bg=t["bg"], bd=0, highlightthickness=0)
        self._body.pack(fill="both", expand=True, side="top")

        self._pg_scripts = self._build_scripts_page()
        self._pg_settings = self._build_settings_page()
        self._show("scripts")

        self.bind("<FocusIn>", self._on_focus_in)
        self.bind("<FocusOut>", self._on_focus_out)
        self.bind("<Configure>", self._on_configure)
        self.after(100, self._on_resize_settle)

    def _show(self, which):
        t = self.theme
        for key, lbl in self._tab_widgets.items():
            if key == which:
                lbl.configure(font=("Segoe UI", 9, "bold"), fg=t["text"])
            else:
                lbl.configure(font=("Segoe UI", 9), fg=t["text_dim"])

        self._pg_scripts.pack_forget()
        self._pg_settings.pack_forget()
        if which == "scripts":
            self._pg_scripts.pack(fill="both", expand=True)
        else:
            self._pg_settings.pack(fill="both", expand=True)

    def _on_focus_in(self, event):
        if not self._app_focused:
            self._app_focused = True
            if self._is_gif and self._scaled_frames and not self._gif_active:
                self._gif_active = True
                self._gif_tick()

    def _on_focus_out(self, event):
        if str(event.widget) == str(self):
            self._app_focused = False
            self._gif_active = False

    def _on_configure(self, event):
        if event.widget is not self: return
        wh = (event.width, event.height)
        if wh == self._last_wh: return
        self._last_wh = wh
        if self._resize_job:
            self.after_cancel(self._resize_job)
        self._resize_job = self.after(200, self._on_resize_settle)

    def _on_resize_settle(self):
        self._resize_job = None
        w = self.winfo_width()
        h = self.winfo_height()
        if w < 10 or h < 10: return
        bg_path = self.theme.get("bg_image", "")
        if bg_path != self._raw_path:
            self._load_raw_frames(bg_path)
        if not self._raw_frames:
            self._clear_bg()
            return
        self._start_rescale(w, h)

    def _load_raw_frames(self, bg_path):
        self._stop_playback()
        self._raw_path = bg_path
        self._raw_frames = []
        self._raw_delays = []
        self._scaled_frames = []
        self._scaled_delays = []
        self._is_gif = False
        if not bg_path or not os.path.isfile(bg_path): return
        try:
            raw = Image.open(bg_path)
            frames_raw = []
            for frame in ImageSequence.Iterator(raw):
                delay = max(20, int(frame.info.get("duration", 80)))
                frames_raw.append((frame.convert("RGB"), delay))
            if not frames_raw: return
            self._is_gif = len(frames_raw) > 1
            for img, delay in frames_raw:
                self._raw_frames.append(img)
                self._raw_delays.append(delay)
        except Exception as e:
            print(f"[RAW LOAD] {e}")

    def _start_rescale(self, w, h):
        self._scale_gen += 1
        gen = self._scale_gen
        opacity = max(5, min(100, int(self.theme.get("bg_opacity", 35))))
        bg_rgb = hex_to_rgb(self.theme.get("bg", "#0F0F12"))
        raw_frames = list(self._raw_frames)
        raw_delays = list(self._raw_delays)

        def worker():
            result = []
            solid = Image.new("RGB", (w, h), bg_rgb)
            for raw, delay in zip(raw_frames, raw_delays):
                if gen != self._scale_gen: return
                try:
                    frame = raw.resize((w, h), Image.LANCZOS)
                    if opacity < 100:
                        frame = Image.blend(solid, frame, opacity / 100.0)
                    result.append((frame, delay))
                except:
                    pass
            if gen == self._scale_gen:
                self.after(0, lambda: self._apply_scaled(result, gen))

        threading.Thread(target=worker, daemon=True).start()

    def _apply_scaled(self, result, gen):
        if gen != self._scale_gen or not result: return
        self._scaled_frames = [f for f, _ in result]
        self._scaled_delays = [d for _, d in result]
        self._gif_idx = min(self._gif_idx, len(self._scaled_frames) - 1)
        if self._is_gif:
            self._gif_active = self._app_focused
            self._show_frame(self._gif_idx)
            if self._gif_active and not self._gif_job:
                self._gif_tick()
        else:
            self._gif_active = False
            self._show_frame(0)

    def _gif_tick(self):
        if not self._gif_active or not self._scaled_frames:
            self._gif_job = None
            return
        self._show_frame(self._gif_idx)
        delay = self._scaled_delays[self._gif_idx]
        self._gif_idx = (self._gif_idx + 1) % len(self._scaled_frames)
        self._gif_job = self.after(delay, self._gif_tick)

    def _stop_playback(self):
        self._gif_active = False
        if self._gif_job:
            try:
                self.after_cancel(self._gif_job)
            except:
                pass
            self._gif_job = None

    def _show_frame(self, idx):
        if not self._scaled_frames: return
        img = self._scaled_frames[idx % len(self._scaled_frames)]
        self._current_img = img
        ctk_img = ctk.CTkImage(light_image=img, dark_image=img,
                               size=(img.width, img.height))
        if self._bg_label and self._bg_label.winfo_exists():
            self._bg_label.configure(image=ctk_img)
            self._bg_label._bg_img_ref = ctk_img  # GC guard
        self._paint_scroll_crops(img)

    def _clear_bg(self):
        t = self.theme
        if self._bg_label and self._bg_label.winfo_exists():
            self._bg_label.configure(image="", fg_color=t["bg"])
        self._current_img = None

    def _register_scroll_canvas(self, canvas):
        self._crop_entries.append({"canvas": canvas})

    def _paint_scroll_crops(self, full_img):
        if not full_img: return
        self._crop_entries = [e for e in self._crop_entries if e["canvas"].winfo_exists()]
        if not self._crop_entries: return
        try:
            ref_rx = self.winfo_rootx()
            ref_ry = self.winfo_rooty()
        except:
            return
        W, H = full_img.size
        for entry in list(self._crop_entries):
            canvas = entry["canvas"]
            try:
                if not canvas.winfo_exists(): continue
                cw = canvas.winfo_width()
                ch = canvas.winfo_height()
                if cw < 2 or ch < 2: continue
                cx = canvas.winfo_rootx() - ref_rx
                cy = canvas.winfo_rooty() - ref_ry
                x0 = max(0, cx)
                y0 = max(0, cy)
                x1 = min(cx + cw, W)
                y1 = min(cy + ch, H)
                if x1 <= x0 or y1 <= y0: continue
                crop = full_img.crop((x0, y0, x1, y1))
                if crop.size != (cw, ch):
                    crop = crop.resize((cw, ch), Image.NEAREST)
                photo = ImageTk.PhotoImage(crop)
                self._crop_photos[id(canvas)] = photo
                item = self._crop_items.get(id(canvas))
                if item:
                    try:
                        canvas.itemconfigure(item, image=photo)
                        canvas.coords(item, 0, 0)
                        continue
                    except tk.TclError:
                        pass
                item = canvas.create_image(0, 0, anchor="nw", image=photo)
                canvas.tag_lower(item)
                self._crop_items[id(canvas)] = item
            except Exception as e:
                print(f"[CROP] {e}")

    # ── Scripts page ──────────────────────────────────────────────────────────
    def _build_scripts_page(self):
        t = self.theme
        pg = tk.Frame(self._body, bg=t["bg"], bd=0, highlightthickness=0)

        act = tk.Frame(pg, bg=t["bg"], height=50, bd=0, highlightthickness=0)
        act.pack(fill="x", side="top")
        act.pack_propagate(False)

        ctk.CTkButton(act, text="+ Add Script", width=120, height=32,
                      font=("Segoe UI", 9, "bold"), fg_color=t["surface2"],
                      hover_color=t["border"], border_width=1, border_color=t["border"],
                      command=self._add_script).pack(side="left", padx=(12, 4), pady=9)

        ctk.CTkButton(act, text="Stop All", width=100, height=32,
                      font=("Segoe UI", 9, "bold"), fg_color="#3C1212",
                      hover_color="#501818", border_width=1, border_color="#641A1A",
                      text_color="#E05050", command=self._stop_all
                      ).pack(side="left", padx=4, pady=9)

        ahk_n = sum(1 for v in AHK_EXES.values() if v and os.path.isfile(v))
        tk.Label(act, text=f"● {ahk_n} AHK version(s) found" if ahk_n else "✗ AHK not found",
                 font=("Segoe UI", 8), bg=t["bg"],
                 fg=t["green"] if ahk_n else t["accent"]).pack(side="right", padx=16)

        tk.Frame(pg, bg=t["border"], height=1, bd=0, highlightthickness=0).pack(fill="x", side="top")

        hdr = tk.Frame(pg, bg=t["bg"], height=26, bd=0, highlightthickness=0)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)

        for txt, x in [("ON/OFF", 14), ("NAME", 76), ("DETECTED", 320),
                       ("AHK EXE", 455), ("HOTKEY", 640), ("LAST", 760)]:
            tk.Label(hdr, text=txt, font=("Segoe UI", 7, "bold"), bg=t["bg"],
                     fg=t["text_dim"]).place(x=x, rely=0.5, anchor="w")

        tk.Frame(pg, bg=t["border"], height=1, bd=0, highlightthickness=0).pack(fill="x", side="top")

        self._scroll_frame = ctk.CTkScrollableFrame(pg, fg_color="transparent",
                                                     scrollbar_button_color=t["surface2"],
                                                     corner_radius=0)
        self._scroll_frame.pack(fill="both", expand=True, side="top")

        sc = getattr(self._scroll_frame, "_parent_canvas", None)
        if sc:
            sc.configure(bg=t["bg"])
            self._scroll_inner_canvas = sc
            self._register_scroll_canvas(sc)
            self.after(30, self._fix_scroll_interior)

        tk.Frame(pg, bg=t["border"], height=1, bd=0, highlightthickness=0).pack(fill="x", side="top")

        sbar = tk.Frame(pg, bg=t["bg"], height=26, bd=0, highlightthickness=0)
        sbar.pack(fill="x", side="top")
        sbar.pack_propagate(False)

        self._status_lbl = tk.Label(sbar, text="No scripts", font=("Segoe UI", 8),
                                    bg=t["bg"], fg=t["text_dim"])
        self._status_lbl.pack(side="left", padx=14)

        self._rebuild_rows()
        return pg

    def _fix_scroll_interior(self):
        sc = getattr(self, "_scroll_inner_canvas", None)
        if sc is None or not sc.winfo_exists(): return
        for child in sc.winfo_children():
            try:
                child.configure(bg=self.theme["bg"])
            except:
                pass

    def _rebuild_rows(self):
        for w in self._scroll_frame.winfo_children():
            w.destroy()
        for s in self.manager.scripts:
            self._build_row(s)
        n = len(self.manager.scripts)
        r = sum(1 for s in self.manager.scripts if s.is_running)
        self._status_lbl.configure(text=f" {n} scripts · {r} running")
        self.after(60, self._fix_scroll_interior)
        self.after(80, lambda: self._paint_scroll_crops(self._current_img) if self._current_img else None)

    def _build_row(self, script):
        t = self.theme
        row = tk.Frame(self._scroll_frame, bg=t["bg"], height=52, bd=0, highlightthickness=0)
        row.pack(fill="x", pady=(0, 2))
        row.pack_propagate(False)

        if script.is_running:
            tk.Frame(row, bg=t["accent"], width=3, bd=0, highlightthickness=0).place(x=0, y=6, height=40)

        var = tk.BooleanVar(value=script.is_running)
        ctk.CTkSwitch(row, text="", variable=var, width=44, progress_color=t["accent"],
                      fg_color=t["surface2"], button_color="#FFF", button_hover_color="#EEE",
                      command=lambda s=script: self._toggle(s)).place(x=10, rely=0.5, anchor="w")

        tk.Label(row, text="●", font=("Segoe UI", 10), bg=t["bg"],
                 fg=t["green"] if script.is_running else t["border"]).place(x=62, rely=0.5, anchor="w")

        tk.Label(row, text=script.name, font=("Segoe UI", 10, "bold"),
                 bg=t["bg"], fg=t["text"]).place(x=78, rely=0.5, anchor="w")

        ver = _detect_ver(script.path)
        vc = t["green"] if ver == "v2" else "#E08030"
        ctk.CTkLabel(row, text=f"AHK {ver}", font=("Consolas", 8, "bold"),
                     text_color=vc, fg_color=t["surface2"], corner_radius=4,
                     width=62, height=22).place(x=322, rely=0.5, anchor="w")

        options = ["Auto (smart detect)"] + list(AHK_EXES.keys())
        cur = "Auto (smart detect)"
        if script.ahk_exe:
            for lbl, p in AHK_EXES.items():
                if p == script.ahk_exe:
                    cur = lbl
                    break
        vv = tk.StringVar(value=cur)
        ctk.CTkOptionMenu(row, variable=vv, values=options, width=175, height=26,
                          fg_color=t["surface2"], button_color=t["border"],
                          button_hover_color=t["surface2"], dropdown_fg_color=t["surface"],
                          dropdown_hover_color=t["surface2"], text_color=t["text_dim"],
                          font=("Segoe UI", 8),
                          command=lambda val, s=script: self._set_ver(s, val)
                          ).place(x=392, rely=0.5, anchor="w")

        hk = script.hotkey_display
        tk.Label(row, text=hk,
                 font=("Consolas", 9, "bold") if hk != "—" else ("Segoe UI", 9),
                 bg=t["bg"], fg=t["accent"] if hk != "—" else t["text_dim"]
                 ).place(x=578, rely=0.5, anchor="w")

        tk.Label(row, text=script.last_toggled or "—", font=("Segoe UI", 8),
                 bg=t["bg"], fg=t["text_dim"]).place(x=682, rely=0.5, anchor="w")

        ctk.CTkButton(row, text="Bind", width=62, height=26, font=("Segoe UI", 8),
                      fg_color=t["surface2"], hover_color=t["border"],
                      border_width=1, border_color=t["border"], text_color=t["text_dim"],
                      command=lambda s=script: self._assign_hotkey(s)
                      ).place(relx=1.0, x=-138, rely=0.5, anchor="w")

        ctk.CTkButton(row, text="Remove", width=70, height=26, font=("Segoe UI", 8),
                      fg_color="#2A0E0E", hover_color="#3C1212",
                      border_width=1, border_color="#501414", text_color="#E05050",
                      command=lambda s=script: self._remove(s)
                      ).place(relx=1.0, x=-60, rely=0.5, anchor="w")

    def _set_ver(self, s, label):
        s.ahk_exe = "" if label.startswith("Auto") else AHK_EXES.get(label, "")
        self.manager.save()

    def _toggle(self, s):
        try:
            self.manager.toggle(s)
        except Exception as e:
            messagebox.showerror("Error", str(e))
        self._rebuild_rows()

    def _stop_all(self):
        self.manager.stop_all()
        self._rebuild_rows()

    def _add_script(self):
        paths = filedialog.askopenfilenames(title="Select AutoHotkey Script(s)",
                                            filetypes=[("AHK Scripts", "*.ahk"), ("All", "*.*")])
        for p in paths:
            if not any(s.path == p for s in self.manager.scripts):
                self.manager.scripts.append(AhkScript(p))
        self._rebuild_rows()
        self.manager.save()

    def _remove(self, s):
        if not messagebox.askyesno("Remove", f"Remove '{s.name}'?"):
            return
        self.manager.remove(s)
        self._rebuild_rows()
        self.manager.save()

    def _assign_hotkey(self, script):
        dlg = HotkeyDialog(self, self.theme, script.hotkey, script.mods)
        self.wait_window(dlg)
        if not dlg.confirmed:
            return
        self.manager.unregister_hotkey(script)
        script.hotkey = dlg.result_key
        script.mods = dlg.result_mods
        if script.hotkey:
            self.manager.register_hotkey(
                script,
                lambda s=script: self.after(0, lambda sc=s: self._toggle(sc))
            )
            messagebox.showinfo("Hotkey Set", f"'{script.hotkey_display}' → '{script.name}'")
        self._rebuild_rows()
        self.manager.save()

    def _reregister_hotkeys(self):
        for s in self.manager.scripts:
            if s.hotkey:
                self.manager.register_hotkey(
                    s,
                    lambda sc=s: self.after(0, lambda s=sc: self._toggle(s))
                )

    def _poll(self):
        changed = False
        for s in self.manager.scripts:
            if s.process is not None and s.process.poll() is not None:
                s.process = None
                s.last_toggled = time.strftime("%H:%M:%S")
                changed = True
        if changed:
            self._rebuild_rows()
        self.after(1000, self._poll)

    # ── Settings page ─────────────────────────────────────────────────────────
    def _build_settings_page(self):
        t = self.theme
        pg = tk.Frame(self._body, bg=t["bg"], bd=0, highlightthickness=0)
        sc = ctk.CTkScrollableFrame(pg, fg_color="transparent",
                                    scrollbar_button_color=t["surface2"], corner_radius=0)
        sc.pack(fill="both", expand=True)

        left = ctk.CTkFrame(sc, fg_color="transparent")
        right = ctk.CTkFrame(sc, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nw", padx=(20, 10), pady=10)
        right.grid(row=0, column=1, sticky="nw", padx=(10, 20), pady=10)
        sc.grid_columnconfigure(0, weight=1)
        sc.grid_columnconfigure(1, weight=1)

        self._sect(left, "Color Presets", 0)
        pf = ctk.CTkFrame(left, fg_color="transparent")
        pf.grid(row=1, column=0, sticky="w", pady=(0, 14))
        for i, (name, accent, _) in enumerate(COLOR_PRESETS):
            ctk.CTkButton(pf, text=name, width=90, height=34, fg_color=t["surface2"],
                          hover_color=accent, font=("Segoe UI", 8), text_color=t["text"],
                          border_width=2, border_color=accent,
                          command=lambda idx=i: self._preset(idx)
                          ).grid(row=i//4, column=i%4, padx=3, pady=3)

        self._sect(left, "Custom Colors", 2)
        cf = ctk.CTkFrame(left, fg_color="transparent")
        cf.grid(row=3, column=0, sticky="w", pady=(0, 14))
        for i, (lbl, cmd) in enumerate([("Accent", self._pick_accent),
                                        ("Background", self._pick_bg),
                                        ("Text", self._pick_text)]):
            ctk.CTkButton(cf, text=lbl, width=120, height=32, fg_color=t["surface2"],
                          hover_color=t["border"], command=cmd).grid(row=0, column=i, padx=(0, 6))

        self._sect(left, "AHK Versions Detected", 4)
        af = ctk.CTkFrame(left, fg_color=t["surface"], corner_radius=4)
        af.grid(row=5, column=0, sticky="w", pady=(0, 14), ipadx=4, ipady=4)
        if AHK_EXES:
            for i, (lbl, p) in enumerate(AHK_EXES.items()):
                ok = p and os.path.isfile(p)
                ctk.CTkLabel(af, text=f"{'●' if ok else '✗'} {lbl}", font=("Segoe UI", 8),
                             text_color=t["green"] if ok else t["accent"], anchor="w"
                             ).grid(row=i, column=0, sticky="w", padx=10, pady=2)
        else:
            ctk.CTkLabel(af, text="No AHK installations found.", font=("Segoe UI", 8),
                         text_color=t["accent"]).grid(row=0, column=0, sticky="w", padx=10, pady=4)

        ctk.CTkButton(af, text="+ Add Custom AHK Exe…", width=240, height=28,
                      font=("Segoe UI", 8), fg_color=t["surface2"], hover_color=t["border"],
                      command=self._add_custom_ahk
                      ).grid(row=len(AHK_EXES)+1, column=0, sticky="w", padx=10, pady=(4, 8))

        self._sect(left, "Reset", 6)
        ctk.CTkButton(left, text="Reset to Defaults", width=160, height=32,
                      fg_color="#3C1212", hover_color="#501818", text_color="#E05050",
                      command=self._reset_theme).grid(row=7, column=0, sticky="w")

        self._sect(right, "Background Image / GIF", 0)
        prev_wrap = tk.Frame(right, width=220, height=110, bg=t["surface2"], bd=0)
        prev_wrap.grid(row=1, column=0, sticky="w", pady=(0, 8))
        prev_wrap.grid_propagate(False)
        self._bg_prev_lbl = tk.Label(prev_wrap, bg=t["surface2"], text="No image selected",
                                     fg=t["text_dim"], font=("Segoe UI", 8))
        self._bg_prev_lbl.place(relx=0.5, rely=0.5, anchor="center")
        self._bg_prev_photo = None
        self._refresh_bg_prev()

        bf = ctk.CTkFrame(right, fg_color="transparent")
        bf.grid(row=2, column=0, sticky="w", pady=(0, 4))
        ctk.CTkButton(bf, text="Choose Image / GIF…", width=170, height=32,
                      fg_color=t["surface2"], hover_color=t["border"],
                      command=self._pick_bg_img).pack(side="left")
        ctk.CTkButton(bf, text="Clear", width=80, height=32,
                      fg_color=t["surface2"], hover_color=t["border"],
                      command=self._clear_bg_img).pack(side="left", padx=6)

        self._sect(right, "Image Opacity", 3)
        op = ctk.CTkFrame(right, fg_color="transparent")
        op.grid(row=4, column=0, sticky="w", pady=(0, 14))
        self._bg_op_lbl = ctk.CTkLabel(op, text=f"{self.theme.get('bg_opacity', 35)}%",
                                       font=("Segoe UI", 8), text_color=t["text_dim"], width=40)
        ctk.CTkSlider(op, from_=5, to=100, number_of_steps=95, width=200,
                      progress_color=t["accent"], button_color=t["accent"],
                      command=self._set_bg_op).pack(side="left")
        self._bg_op_lbl.pack(side="left", padx=8)

        self._sect(right, "Window Opacity", 5)
        wo = ctk.CTkFrame(right, fg_color="transparent")
        wo.grid(row=6, column=0, sticky="w")
        self._win_op_lbl = ctk.CTkLabel(wo, text="100%", font=("Segoe UI", 8),
                                        text_color=t["text_dim"], width=40)
        ctk.CTkSlider(wo, from_=40, to=100, number_of_steps=60, width=200,
                      progress_color=t["accent"], button_color=t["accent"],
                      command=self._set_win_op).pack(side="left")
        self._win_op_lbl.pack(side="left", padx=8)

        return pg

    def _sect(self, p, text, row):
        ctk.CTkLabel(p, text=text.upper(), font=("Segoe UI", 8, "bold"),
                     text_color=self.theme["text_dim"]).grid(row=row, column=0, sticky="w", pady=(14, 4))

    def _pick_bg_img(self):
        p = filedialog.askopenfilename(title="Select Background Image or GIF",
            filetypes=[("Images & GIFs", "*.png *.jpg *.jpeg *.bmp *.gif *.webp"), ("All", "*.*")])
        if not p: return
        self.theme["bg_image"] = p
        save_theme(self.theme)
        self._load_raw_frames(p)
        self._on_resize_settle()
        self._refresh_bg_prev()

    def _clear_bg_img(self):
        self.theme["bg_image"] = ""
        save_theme(self.theme)
        self._stop_playback()
        self._raw_path = ""
        self._raw_frames = []
        self._raw_delays = []
        self._scaled_frames = []
        self._scaled_delays = []
        self._scale_gen += 1
        self._clear_bg()
        self._refresh_bg_prev()

    def _set_bg_op(self, val):
        self.theme["bg_opacity"] = int(float(val))
        save_theme(self.theme)
        if self._raw_frames:
            self._on_resize_settle()
        try:
            self._bg_op_lbl.configure(text=f"{int(float(val))}%")
        except:
            pass

    def _refresh_bg_prev(self):
        try:
            p = self.theme.get("bg_image", "")
            if p and os.path.isfile(p):
                img = Image.open(p).convert("RGB").resize((220, 110), Image.LANCZOS)
                self._bg_prev_photo = ImageTk.PhotoImage(img)
                self._bg_prev_lbl.configure(image=self._bg_prev_photo, text="")
            else:
                self._bg_prev_photo = None
                self._bg_prev_lbl.configure(image="", text="No image selected")
        except:
            pass

    def _preset(self, idx):
        _, acc, bg = COLOR_PRESETS[idx]
        self.theme.update({
            "accent": acc,
            "bg": bg,
            "surface": shift_color(bg, 8),
            "surface2": shift_color(bg, 15),
            "border": shift_color(bg, 25)
        })
        save_theme(self.theme)
        self._rebuild_ui()

    def _pick_accent(self):
        c = colorchooser.askcolor(color=self.theme["accent"], title="Accent")
        if c[1]:
            self.theme["accent"] = c[1]
            save_theme(self.theme)
            self._rebuild_ui()

    def _pick_bg(self):
        c = colorchooser.askcolor(color=self.theme["bg"], title="Background")
        if c[1]:
            bg = c[1]
            self.theme.update({
                "bg": bg,
                "surface": shift_color(bg, 8),
                "surface2": shift_color(bg, 15),
                "border": shift_color(bg, 25)
            })
            save_theme(self.theme)
            self._rebuild_ui()

    def _pick_text(self):
        c = colorchooser.askcolor(color=self.theme["text"], title="Text")
        if c[1]:
            self.theme["text"] = c[1]
            self.theme["text_dim"] = blend_color(c[1], self.theme["bg"], 0.5)
            save_theme(self.theme)
            self._rebuild_ui()

    def _reset_theme(self):
        if not messagebox.askyesno("Reset", "Reset all theme settings?"):
            return
        keep = {k: self.theme[k] for k in ("bg_image", "bg_opacity") if k in self.theme}
        self.theme = dict(DEFAULT_THEME)
        self.theme.update(keep)
        save_theme(self.theme)
        self._rebuild_ui()

    def _add_custom_ahk(self):
        p = filedialog.askopenfilename(title="Select AHK Executable",
                                       filetypes=[("Executables", "*.exe"), ("All", "*.*")])
        if not p or not os.path.isfile(p): return
        lbl = f"Custom — {Path(p).name}"
        AHK_EXES[lbl] = p
        messagebox.showinfo("Added", f"Added: {lbl}")

    def _set_win_op(self, val):
        self.attributes("-alpha", int(float(val)) / 100)
        try:
            self._win_op_lbl.configure(text=f"{int(float(val))}%")
        except:
            pass

    def _rebuild_ui(self):
        self.configure(fg_color=self.theme["bg"])
        self._build()

    def _build_tray(self):
        if not HAS_TRAY: return
        sz = 64
        img = Image.new("RGBA", (sz, sz), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        r, g, b = hex_to_rgb(self.theme["accent"])
        d.ellipse([2, 2, sz-3, sz-3], fill=(r, g, b, 255))
        try:
            fnt = ImageFont.truetype("segoeuib.ttf", sz//2)
        except:
            fnt = ImageFont.load_default()
        d.text((sz//2, sz//2), "A", fill="white", font=fnt, anchor="mm")

        menu = pystray.Menu(
            pystray.MenuItem("Show", lambda: self.after(0, self.deiconify)),
            pystray.MenuItem("Exit", lambda: self.after(0, self._force_exit))
        )
        self._tray_icon = pystray.Icon("AhkManager", img, "AHK Manager", menu)
        self._tray_icon.run()

    def _on_close(self):
        self.withdraw()
        if not HAS_TRAY:
            self._force_exit()

    def _force_exit(self):
        self._stop_playback()
        self._scale_gen += 1
        self.manager.stop_all()
        self.manager.save()
        if HAS_TRAY:
            try:
                self._tray_icon.stop()
            except:
                pass
        self.destroy()
        sys.exit(0)

if __name__ == "__main__":
    app = App()
    app.mainloop()
