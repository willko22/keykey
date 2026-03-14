#!/usr/bin/env python3
"""
Key Heatmap – tracks keystrokes with heatmap visualization.
Auto-detects keyboard layout (QWERTY/QWERTZ/AZERTY/Dvorak/Colemak).
"""

import json, sys, threading, platform, math, os, time, ctypes
from collections import Counter
from pathlib import Path
from pynput import keyboard as pynput_kb
import tkinter as tk
from tkinter import ttk

HAS_TRAY = platform.system() == "Windows"
TRAY_IMPORT_ERROR = None if HAS_TRAY else "Native tray is only supported on Windows"

# ── storage ───────────────────────────────────────────────────────────────────
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
SAVE_FILE  = BASE_DIR / "stats.json"
PREFS_FILE = Path.home() / ".key_heatmap_prefs.json"
ICON_FILE  = BASE_DIR / "keykey.ico"
APP_ALL_ID = "__all__"
UNATTRIBUTED_APP_ID = "__unattributed__"
DOTA2_APP_ID = "__dota2__"
WINDOW_TITLE_PREFIX = "window::"

def _foreground_window_title_windows(hwnd):
    user32 = ctypes.windll.user32
    title_len = user32.GetWindowTextLengthW(hwnd)
    if title_len <= 0:
        return ""
    buf = ctypes.create_unicode_buffer(title_len + 1)
    user32.GetWindowTextW(hwnd, buf, title_len + 1)
    return buf.value or ""

def _infer_game_app_id_from_title(title):
    t = (title or "").lower()
    if "dota" in t and "2" in t:
        return DOTA2_APP_ID
    return None

def _window_title_app_id(title):
    clean = (title or "").strip()
    if not clean:
        return None
    return f"{WINDOW_TITLE_PREFIX}{clean}"

def load_stats():
    if not SAVE_FILE.exists():
        return {}
    with open(SAVE_FILE, encoding="utf-8") as f:
        raw = json.load(f)
    # Backward compatibility: if stats are in a flat key->count shape,
    # treat it as an all-app aggregate bucket.
    if raw and all(isinstance(v, int) for v in raw.values()):
        return {APP_ALL_ID: Counter(raw)}
    out = {}
    for app_id, app_counts in raw.items():
        if isinstance(app_counts, dict):
            out[app_id] = Counter({k: int(v) for k, v in app_counts.items()})
    return out

def save_stats(all_counts):
    payload = {app_id: dict(counter) for app_id, counter in all_counts.items()}
    with open(SAVE_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)

def load_prefs():
    if PREFS_FILE.exists():
        with open(PREFS_FILE) as f: return json.load(f)
    return {}
def save_prefs(p):
    with open(PREFS_FILE,"w") as f: json.dump(p,f,indent=2)

counts_by_app  = load_stats()
current_app_id = None
last_active_app_id = None
counts_lock    = threading.Lock()
held_keys      = set()
stats_dirty    = False

def _foreground_app_path_windows():
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    psapi = ctypes.windll.kernel32

    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return None

    title = _foreground_window_title_windows(hwnd)

    pid = ctypes.c_ulong()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    if not pid.value or pid.value == os.getpid():
        return None

    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    hproc = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid.value)
    if not hproc:
        # If process details are blocked, attribute by exact window title.
        inferred = _infer_game_app_id_from_title(title)
        if inferred:
            return inferred
        return _window_title_app_id(title)
    try:
        buf_len = ctypes.c_ulong(1024)
        buf = ctypes.create_unicode_buffer(buf_len.value)
        if not psapi.QueryFullProcessImageNameW(hproc, 0, buf, ctypes.byref(buf_len)):
            inferred = _infer_game_app_id_from_title(title)
            if inferred:
                return inferred
            return _window_title_app_id(title)
        return str(Path(buf.value))
    finally:
        kernel32.CloseHandle(hproc)

def _display_app_label(app_id):
    if not app_id:
        return None
    if app_id == DOTA2_APP_ID:
        return "Dota 2"
    if app_id.startswith(WINDOW_TITLE_PREFIX):
        return app_id[len(WINDOW_TITLE_PREFIX):]
    if app_id == UNATTRIBUTED_APP_ID:
        return "Unknown / Unattributed"
    if app_id.startswith("process_"):
        return app_id.replace("process_", "PID ") + " (restricted)"
    p = Path(app_id)
    if p.suffix.lower() == ".exe":
        return p.stem
    return p.name or app_id

def get_foreground_app_id():
    if platform.system() != "Windows":
        return None
    try:
        return _foreground_app_path_windows()
    except Exception:
        return None

def _focus_watcher():
    global current_app_id, last_active_app_id
    while True:
        app_id = get_foreground_app_id()
        with counts_lock:
            current_app_id = app_id
            if app_id:
                last_active_app_id = app_id
        time.sleep(0.2)

# ── key normalisation ─────────────────────────────────────────────────────────
# On Windows, pynput gives key.vk for every key.
# We map VK codes to stable key_ids so stats are layout-independent at the
# physical-key level, then the layout definition tells us what to display.
#
# For alpha keys (A-Z): vk 65-90 → "a"-"z"
# For digit keys: vk 48-57 → "0"-"9"  (but key.char gives "1"/"!" etc.)
# For special keys: Key.xxx
#
# Strategy: use key.char when it's a single printable character we recognise,
# fall back to vk-based id for everything else.

# VK → stable key_id (Windows VK codes for symbol/special keys)
_VK_TO_ID = {
    8:   "backspace",
    9:   "tab",
    13:  "return",
    20:  "caps_lock",
    32:  "space",
    160: "shift",     # VK_LSHIFT
    161: "shift_r",   # VK_RSHIFT
    162: "ctrl",      # VK_LCONTROL
    163: "ctrl_r",    # VK_RCONTROL
    164: "alt",       # VK_LMENU
    165: "alt_r",     # VK_RMENU
    91:  "cmd",       # VK_LWIN
    92:  "cmd_r",     # VK_RWIN
    93:  "menu",      # VK_APPS
    # number row symbols — VK codes for the physical key (unshifted label)
    192: "`",    # VK_OEM_3  (`~)
    189: "-",    # VK_OEM_MINUS
    187: "=",    # VK_OEM_PLUS  (= key, shifted gives +)
    219: "[",    # VK_OEM_4
    221: "]",    # VK_OEM_6
    220: "\\",   # VK_OEM_5
    186: ";",    # VK_OEM_1
    222: "'",    # VK_OEM_7
    188: ",",    # VK_OEM_COMMA
    190: ".",    # VK_OEM_PERIOD
    191: "/",    # VK_OEM_2
    # numpad — alias to same IDs as regular keys
    96:  "0",    # VK_NUMPAD0
    97:  "1",    # VK_NUMPAD1
    98:  "2",    # VK_NUMPAD2
    99:  "3",    # VK_NUMPAD3
    100: "4",    # VK_NUMPAD4
    101: "5",    # VK_NUMPAD5
    102: "6",    # VK_NUMPAD6
    103: "7",    # VK_NUMPAD7
    104: "8",    # VK_NUMPAD8
    105: "9",    # VK_NUMPAD9
    106: "8",    # VK_MULTIPLY  (* = shift+8)
    107: "=",    # VK_ADD       (numpad + → same as = key)
    109: "-",    # VK_SUBTRACT  (numpad -)
    110: ".",    # VK_DECIMAL   (numpad .)
    111: "/",    # VK_DIVIDE    (numpad /)
}

# shifted variants → same key_id as the base key (so "+" maps to "=", "!" to "1")
_SHIFTED_TO_BASE = {
    "~":"`", "!":"1", "@":"2", "#":"3", "$":"4", "%":"5",
    "^":"6", "&":"7", "*":"8", "(":"9", ")":"0",
    "_":"-", "+":"=",
    "{":"[", "}":"]", "|":"\\",
    ":":";", '"':"'",
    "<":",", ">":".", "?":"/",
}

_SPECIAL_TO_ID = {
    "shift_l":"shift", "shift":"shift",
    "ctrl_l":"ctrl",   "ctrl":"ctrl",
    "alt_l":"alt",     "alt":"alt",
    "alt_gr":"alt_r",
    "cmd_l":"cmd",     "cmd":"cmd",
    "enter":"return",
}

def normalise_key(key):
    # 1. Try VK code first (most reliable on Windows for symbol keys)
    vk = getattr(key, "vk", None)
    if vk is not None:
        if vk in _VK_TO_ID:
            return _VK_TO_ID[vk]
        # digit keys: vk 48-57
        if 48 <= vk <= 57:
            return str(vk - 48) if vk > 48 else "0"
        # alpha keys: vk 65-90
        if 65 <= vk <= 90:
            return chr(vk + 32)  # lowercase

    # 2. Try key.char for anything with a printable character
    try:
        ch = key.char
        if ch and len(ch) == 1:
            base = _SHIFTED_TO_BASE.get(ch)
            if base:
                return base
            return ch.lower()
    except AttributeError:
        pass

    # 3. Special keys by name
    try:
        raw = str(key).replace("Key.", "").lower()
        return _SPECIAL_TO_ID.get(raw, raw)
    except Exception:
        return None

def on_press(key):
    global stats_dirty
    k = normalise_key(key)
    if not k:
        return

    with counts_lock:
        if k in held_keys:
            return
        held_keys.add(k)

        # If foreground app cannot be identified by EXE or window title, ignore.
        app_id = current_app_id
        if not app_id:
            return
        app_counts = counts_by_app.setdefault(app_id, Counter())
        app_counts[k] += 1
        stats_dirty = True

def on_release(key):
    k = normalise_key(key)
    if not k:
        return
    with counts_lock:
        held_keys.discard(k)

def flush_stats_if_dirty():
    global stats_dirty
    with counts_lock:
        if not stats_dirty:
            return
        save_stats(counts_by_app)
        stats_dirty = False

listener = pynput_kb.Listener(on_press=on_press, on_release=on_release)
listener.daemon = True
listener.start()

if platform.system() == "Windows":
    focus_thread = threading.Thread(target=_focus_watcher, daemon=True)
    focus_thread.start()

# ── layout detection ──────────────────────────────────────────────────────────
_LANG_TO_LAYOUT = {
    0x0409:"qwerty", 0x0809:"qwerty", 0x0c09:"qwerty", 0x1009:"qwerty",
    0x0407:"qwertz", 0x0807:"qwertz", 0x0c07:"qwertz",
    0x041b:"qwertz", 0x0405:"qwertz", 0x040e:"qwertz",
    0x040c:"azerty", 0x080c:"azerty", 0x100c:"azerty",
}

def detect_layout():
    if platform.system() == "Windows":
        try:
            import ctypes
            u = ctypes.WinDLL("user32", use_last_error=True)
            hkl  = u.GetKeyboardLayout(u.GetWindowThreadProcessId(u.GetForegroundWindow(), 0))
            lang = hkl & 0xFFFF
            name = _LANG_TO_LAYOUT.get(lang)
            if name: return name
            p = lang & 0xFF
            if p == 0x07: return "qwertz"
            if p == 0x0c: return "azerty"
        except Exception: pass
    try:
        import subprocess
        out = subprocess.check_output(["localectl","status"],text=True,stderr=subprocess.DEVNULL)
        for line in out.splitlines():
            v = line.split(":")[-1].strip().lower()
            if "dvorak"  in v: return "dvorak"
            if "colemak" in v: return "colemak"
            if any(v.startswith(x) for x in ("de","sk","cz","hu")): return "qwertz"
            if any(v.startswith(x) for x in ("fr","be")):           return "azerty"
    except Exception: pass
    return "qwerty"

# ── layout definitions ────────────────────────────────────────────────────────
# Tuple: (normal_label, shifted_label, key_id, width)
# key_id must match what normalise_key() returns.
# For printable keys, key_id = the unshifted character pynput/VK gives us.
# Stats for a cell = counts[key_id]  (we no longer double-count shifted variants
# separately — everything normalises to the base key_id).

def pk(n, s, kid=None, w=1): return (n, s, kid or n, w)
def sk(lbl, kid, w=1):       return (lbl, None, kid, w)

LAYOUTS = {
"qwerty": [
    [pk("`","~"),pk("1","!"),pk("2","@"),pk("3","#"),pk("4","$"),pk("5","%"),
     pk("6","^"),pk("7","&"),pk("8","*"),pk("9","("),pk("0",")"),
     pk("-","_"),pk("=","+"),sk("Bksp","backspace",2)],
    [sk("Tab","tab",1.5),pk("q","Q"),pk("w","W"),pk("e","E"),pk("r","R"),
     pk("t","T"),pk("y","Y"),pk("u","U"),pk("i","I"),pk("o","O"),pk("p","P"),
     pk("[","{"),pk("]","}"),pk("\\","|",None,1.5)],
    [sk("Caps","caps_lock",1.75),pk("a","A"),pk("s","S"),pk("d","D"),pk("f","F"),
     pk("g","G"),pk("h","H"),pk("j","J"),pk("k","K"),pk("l","L"),
     pk(";",":"),pk("'",'"'),sk("Enter","return",2.25)],
    [sk("Shift","shift",2.25),pk("z","Z"),pk("x","X"),pk("c","C"),pk("v","V"),
     pk("b","B"),pk("n","N"),pk("m","M"),pk(",","<"),pk(".",">"),(  "/","?","/",1),
     sk("Shift","shift_r",2.75)],
    [sk("Ctrl","ctrl",1.25),sk("Win","cmd",1.25),sk("Alt","alt",1.25),
     sk("Space","space",6.25),
     sk("Alt","alt_r",1.25),sk("Win","cmd_r",1.25),sk("Menu","menu",1.25),sk("Ctrl","ctrl_r",1.25)],
],
"qwertz": [
    [pk("^","°","^"),pk("1","!"),pk("2",'"'),pk("3","§"),pk("4","$"),pk("5","%"),
     pk("6","&"),pk("7","/"),pk("8","("),pk("9",")"),pk("0","="),
     pk("ß","?","ß"),pk("´","`","´"),sk("Bksp","backspace",2)],
    [sk("Tab","tab",1.5),pk("q","Q"),pk("w","W"),pk("e","E"),pk("r","R"),
     pk("t","T"),pk("z","Z"),pk("u","U"),pk("i","I"),pk("o","O"),pk("p","P"),
     pk("ü","Ü","ü"),pk("+","*","="),pk("#","'","\\",1.5)],
    [sk("Caps","caps_lock",1.75),pk("a","A"),pk("s","S"),pk("d","D"),pk("f","F"),
     pk("g","G"),pk("h","H"),pk("j","J"),pk("k","K"),pk("l","L"),
     pk("ö","Ö","ö"),pk("ä","Ä","ä"),sk("Enter","return",2.25)],
    [sk("Shift","shift",2.25),pk("<",">","<"),pk("y","Y"),pk("x","X"),pk("c","C"),
     pk("v","V"),pk("b","B"),pk("n","N"),pk("m","M"),pk(",",";"),pk(".",":"),
     pk("-","_"),sk("Shift","shift_r",2.75)],
    [sk("Ctrl","ctrl",1.25),sk("Win","cmd",1.25),sk("Alt","alt",1.25),
     sk("Space","space",6.25),
     sk("AltGr","alt_r",1.25),sk("Win","cmd_r",1.25),sk("Menu","menu",1.25),sk("Ctrl","ctrl_r",1.25)],
],
"azerty": [
    [pk("²","","²"),pk("&","1","1"),pk("é","2","2"),pk('"',"3","3"),pk("'","4","4"),pk("(","5","5"),
     pk("-","6","6"),pk("è","7","7"),pk("_","8","8"),pk("ç","9","9"),pk("à","0","0"),
     pk(")","°",")"),pk("=","+","="),sk("Bksp","backspace",2)],
    [sk("Tab","tab",1.5),pk("a","A"),pk("z","Z"),pk("e","E"),pk("r","R"),
     pk("t","T"),pk("y","Y"),pk("u","U"),pk("i","I"),pk("o","O"),pk("p","P"),
     pk("^","¨","["),pk("$","£","]"),pk("*","μ","\\",1.5)],
    [sk("Caps","caps_lock",1.75),pk("q","Q"),pk("s","S"),pk("d","D"),pk("f","F"),
     pk("g","G"),pk("h","H"),pk("j","J"),pk("k","K"),pk("l","L"),
     pk("m","M",";"),pk("ù","%","'"),sk("Enter","return",2.25)],
    [sk("Shift","shift",2.25),pk("<",">","<"),pk("w","W"),pk("x","X"),pk("c","C"),
     pk("v","V"),pk("b","B"),pk("n","N"),pk(",","?",","),pk(";",".","."),
     pk(":","/","/"),(  "!","§","/",1),sk("Shift","shift_r",2.75)],
    [sk("Ctrl","ctrl",1.25),sk("Win","cmd",1.25),sk("Alt","alt",1.25),
     sk("Space","space",6.25),
     sk("AltGr","alt_r",1.25),sk("Win","cmd_r",1.25),sk("Menu","menu",1.25),sk("Ctrl","ctrl_r",1.25)],
],
"dvorak": [
    [pk("`","~"),pk("1","!"),pk("2","@"),pk("3","#"),pk("4","$"),pk("5","%"),
     pk("6","^"),pk("7","&"),pk("8","*"),pk("9","("),pk("0",")"),
     pk("[","{","["),pk("]","}","]"),sk("Bksp","backspace",2)],
    [sk("Tab","tab",1.5),pk("'",'"'),pk(",","<"),pk(".",">"),(  "p","P","p",1),
     pk("y","Y"),pk("f","F"),pk("g","G"),pk("c","C"),pk("r","R"),pk("l","L"),
     pk("/","?"),pk("=","+"),pk("\\","|",None,1.5)],
    [sk("Caps","caps_lock",1.75),pk("a","A"),pk("o","O"),pk("e","E"),pk("u","U"),
     pk("i","I"),pk("d","D"),pk("h","H"),pk("t","T"),pk("n","N"),pk("s","S"),
     pk("-","_"),sk("Enter","return",2.25)],
    [sk("Shift","shift",2.25),pk(";",":"),pk("q","Q"),pk("j","J"),pk("k","K"),
     pk("x","X"),pk("b","B"),pk("m","M"),pk("w","W"),pk("v","V"),pk("z","Z"),
     sk("Shift","shift_r",2.75)],
    [sk("Ctrl","ctrl",1.25),sk("Win","cmd",1.25),sk("Alt","alt",1.25),
     sk("Space","space",6.25),
     sk("Alt","alt_r",1.25),sk("Win","cmd_r",1.25),sk("Menu","menu",1.25),sk("Ctrl","ctrl_r",1.25)],
],
"colemak": [
    [pk("`","~"),pk("1","!"),pk("2","@"),pk("3","#"),pk("4","$"),pk("5","%"),
     pk("6","^"),pk("7","&"),pk("8","*"),pk("9","("),pk("0",")"),
     pk("-","_"),pk("=","+"),sk("Bksp","backspace",2)],
    [sk("Tab","tab",1.5),pk("q","Q"),pk("w","W"),pk("f","F"),pk("p","P"),
     pk("g","G"),pk("j","J"),pk("l","L"),pk("u","U"),pk("y","Y"),pk(";",":"),
     pk("[","{"),pk("]","}"),pk("\\","|",None,1.5)],
    [sk("Caps","caps_lock",1.75),pk("a","A"),pk("r","R"),pk("s","S"),pk("t","T"),
     pk("d","D"),pk("h","H"),pk("n","N"),pk("e","E"),pk("i","I"),pk("o","O"),
     pk("'",'"'),sk("Enter","return",2.25)],
    [sk("Shift","shift",2.25),pk("z","Z"),pk("x","X"),pk("c","C"),pk("v","V"),
     pk("b","B"),pk("k","K"),pk("m","M"),pk(",","<"),pk(".",">"),(  "/","?","/",1),
     sk("Shift","shift_r",2.75)],
    [sk("Ctrl","ctrl",1.25),sk("Win","cmd",1.25),sk("Alt","alt",1.25),
     sk("Space","space",6.25),
     sk("Alt","alt_r",1.25),sk("Win","cmd_r",1.25),sk("Menu","menu",1.25),sk("Ctrl","ctrl_r",1.25)],
],
}

TOTAL_UNITS = max(sum(t[3] for t in row) for row in LAYOUTS["qwerty"])
NUM_ROWS    = 5
PAD_RATIO   = 0.008

# ── colors ────────────────────────────────────────────────────────────────────
BG      = "#0d1117"
OUTLINE = "#2d3748"

HEAT_STOPS = [
    (0.00,(22,28,42)),  (0.04,(24,48,80)),  (0.08,(18,72,110)),
    (0.13,(16,96,130)), (0.18,(14,118,130)),(0.24,(16,135,118)),
    (0.30,(18,148,100)),(0.37,(30,158,70)), (0.44,(55,168,45)),
    (0.51,(90,175,25)), (0.58,(145,182,18)),(0.65,(195,190,18)),
    (0.72,(225,185,18)),(0.79,(238,155,16)),(0.86,(242,110,16)),
    (0.92,(242,65,18)), (0.96,(238,38,18)), (1.00,(230,20,10)),
]

def heat_color(ratio):
    if ratio <= 0: return "#16192a"
    ratio = math.pow(max(0.0, min(1.0, ratio)), 0.55)
    for i in range(len(HEAT_STOPS)-1):
        r0,c0 = HEAT_STOPS[i]; r1,c1 = HEAT_STOPS[i+1]
        if r0 <= ratio <= r1:
            t = (ratio-r0)/(r1-r0)
            return "#{:02x}{:02x}{:02x}".format(
                int(c0[0]+t*(c1[0]-c0[0])),
                int(c0[1]+t*(c1[1]-c0[1])),
                int(c0[2]+t*(c1[2]-c0[2])))
    return "#e6140a"

def text_color(bg):
    r,g,b = int(bg[1:3],16),int(bg[3:5],16),int(bg[5:7],16)
    return "#fff" if (0.299*r+0.587*g+0.114*b) < 145 else "#111"

def merge_all_counts(all_counts):
    merged = Counter()
    for app_counts in all_counts.values():
        merged.update(app_counts)
    return merged

def make_app_labels(app_ids):
    base_names = []
    for app_id in app_ids:
        base_names.append(_display_app_label(app_id))

    name_freq = Counter(base_names)
    out = {}
    used = set()
    for app_id, base in zip(app_ids, base_names):
        if name_freq[base] <= 1:
            label = base
        else:
            p = Path(app_id)
            folder = p.parent.name or str(p.parent)
            label = f"{base} << {folder}"

        if label in used:
            i = 2
            while f"{label} #{i}" in used:
                i += 1
            label = f"{label} #{i}"

        used.add(label)
        out[label] = app_id
    return out

# ── window ────────────────────────────────────────────────────────────────────
class HeatmapWindow:
    BARS_H = 175
    BARS_H_EXPANDED_MIN = 280
    MIN_W, MIN_H = 700, 420
    DEFAULT_W, DEFAULT_H = 1050, 640

    def __init__(self, start_hidden=False):
        self._start_hidden = start_hidden
        self._tray_mode = HAS_TRAY
        self.prefs       = load_prefs()
        self.layout_name = detect_layout()
        self.rows        = LAYOUTS[self.layout_name]

        self.root = tk.Tk()
        if self._start_hidden:
            # Keep the window hidden from first paint; user opens it via tray.
            self.root.withdraw()
        self.root.title(f"Key Heatmap  [{self.layout_name.upper()}]")
        self.root.configure(bg=BG)
        self.root.minsize(self.MIN_W, self.MIN_H)
        if ICON_FILE.exists():
            try:
                self.root.iconbitmap(default=str(ICON_FILE))
            except Exception:
                pass

        w = self.prefs.get("w", self.DEFAULT_W)
        h = self.prefs.get("h", self.DEFAULT_H)
        x, y = self.prefs.get("x"), self.prefs.get("y")
        self.root.geometry(f"{w}x{h}+{x}+{y}" if x is not None else f"{w}x{h}")

        self.key_cells   = []
        self.view_var    = tk.StringVar(value="All Apps")
        self.view_combo  = None
        self.view_map    = {"All Apps": APP_ALL_ID}
        self._view_items = ["All Apps"]
        self._default_view_set = False
        self._hover_idx  = None
        self._resize_job = None
        self._stats_expanded = False
        self._stats_frame = None
        self._stats_title = None
        self._bars_indicator_tag = "__bars_indicator__"
        self._bars_snap  = {}
        self._bars_total = 1

        self._build_ui()
        self._refresh()
        self.root.after(2000, self._auto_refresh)
        self.canvas.bind("<Configure>", self._on_canvas_resize)
        self.root.bind("<Configure>", self._on_root_configure)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close_btn)

    def _init_styles(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "Keykey.TCombobox",
            fieldbackground="#111827",
            background="#1f2937",
            foreground="#e5e7eb",
            bordercolor="#334155",
            lightcolor="#334155",
            darkcolor="#334155",
            arrowcolor="#93c5fd",
            padding=6,
        )
        style.map(
            "Keykey.TCombobox",
            fieldbackground=[("readonly", "#111827")],
            background=[("readonly", "#1f2937")],
            foreground=[("readonly", "#e5e7eb")],
            bordercolor=[("focus", "#38bdf8")],
            lightcolor=[("focus", "#38bdf8")],
            darkcolor=[("focus", "#38bdf8")],
        )

    def _build_ui(self):
        M = 14
        self._init_styles()
        hdr = tk.Frame(self.root, bg=BG)
        hdr.pack(fill="x", padx=M, pady=(M,4))
        tk.Label(hdr, text="KEY HEATMAP", font=("Courier New",14,"bold"),
                 fg="#e8f4fd", bg=BG).pack(side="left")
        tk.Label(hdr, text=f"[{self.layout_name.upper()}  auto-detected]",
                 font=("Courier New",9), fg="#374151", bg=BG).pack(side="left", padx=10)
        self.total_lbl = tk.Label(hdr, text="", font=("Courier New",10), fg="#6b7280", bg=BG)
        self.total_lbl.pack(side="right")

        bf = tk.Frame(self.root, bg=BG)
        bf.pack(fill="x", padx=M, pady=(0,6))

        def mkbtn(p,t,cmd,fg,abg):
            return tk.Button(p,text=t,command=cmd,bg="#1a2030",fg=fg,relief="flat",
                             font=("Courier New",9),padx=10,pady=3,cursor="hand2",
                             activebackground=abg,activeforeground=fg,bd=0)
        tk.Label(bf, text="View", font=("Courier New",9,"bold"), fg="#9ca3af", bg=BG).pack(side="left", padx=(0,8))
        self.view_combo = ttk.Combobox(
            bf,
            textvariable=self.view_var,
            values=self._view_items,
            width=34,
            state="readonly",
            style="Keykey.TCombobox",
            font=("Courier New",10),
        )
        self.view_combo.pack(side="left", padx=(0,10), ipady=4)
        self.view_combo.bind("<<ComboboxSelected>>", self._on_view_changed)

        mkbtn(bf,"Refresh",    self._refresh, "#7dd3fc","#1e3a50").pack(side="left",padx=(0,6))
        mkbtn(bf,"Reset Stats",self._reset,   "#f87171","#3d1f1f").pack(side="left")

        self.canvas = tk.Canvas(self.root, bg=BG, highlightthickness=0)
        self.canvas.pack(fill="both", expand=True, padx=M, pady=(0,4))

        self._stats_frame = tk.Frame(self.root, bg=BG, height=self.BARS_H)
        self._stats_frame.pack(fill="x", padx=M, pady=(0,M))
        self._stats_frame.pack_propagate(False)
        self._stats_title = tk.Label(
            self._stats_frame,
            text="TOP KEYS  [click chart to expand/collapse]",
            font=("Courier New",8,"bold"),
            fg="#6b7280",
            bg=BG,
        )
        self._stats_title.pack(anchor="w")

        bars_wrap = tk.Frame(self._stats_frame, bg=BG)
        bars_wrap.pack(fill="both", expand=True)

        self.bars_canvas = tk.Canvas(bars_wrap, bg=BG, highlightthickness=0, height=self.BARS_H-22)
        self.bars_canvas.pack(fill="both", expand=True)
        self.bars_canvas.configure(yscrollcommand=self._on_bars_yview)
        self.bars_canvas.bind("<Configure>", lambda e: self._redraw_bars())
        self.bars_canvas.bind("<Button-1>", self._toggle_stats_panel)
        self.bars_canvas.bind("<MouseWheel>", self._on_bars_mousewheel)
        self.bars_canvas.bind("<Button-4>", self._on_bars_mousewheel)
        self.bars_canvas.bind("<Button-5>", self._on_bars_mousewheel)

    def _selected_snapshot(self):
        selected_id = self.view_map.get(self.view_var.get(), APP_ALL_ID)
        with counts_lock:
            snap_by_app = {app_id: Counter(c) for app_id, c in counts_by_app.items()}
            active = current_app_id

        if selected_id == APP_ALL_ID:
            snap = merge_all_counts(snap_by_app)
        else:
            snap = Counter(snap_by_app.get(selected_id, Counter()))

        total = sum(snap.values()) or 1
        return snap, total, selected_id, active

    def _sync_app_selector(self):
        with counts_lock:
            app_ids = set(counts_by_app.keys())
            if last_active_app_id:
                app_ids.add(last_active_app_id)
            app_ids = sorted(app_ids, key=lambda s: s.lower())

        labels = make_app_labels(app_ids)
        new_map = {"All Apps": APP_ALL_ID}
        for label in sorted(labels.keys(), key=lambda s: s.lower()):
            new_map[label] = labels[label]
        new_items = list(new_map.keys())

        if new_items == self._view_items:
            return

        current = self.view_var.get()
        self.view_map = new_map
        self._view_items = new_items
        self.view_combo.configure(values=new_items)

        if not self._default_view_set:
            preferred = None
            with counts_lock:
                if last_active_app_id:
                    preferred = last_active_app_id
                elif current_app_id:
                    preferred = current_app_id

            if preferred and preferred in new_map.values():
                for label, app_id in new_map.items():
                    if app_id == preferred:
                        self.view_var.set(label)
                        break
            elif current in new_map:
                self.view_var.set(current)
            else:
                self.view_var.set("All Apps")
            self._default_view_set = True
        elif current in new_map:
            self.view_var.set(current)
        else:
            self.view_var.set("All Apps")

    def _on_view_changed(self, _=None):
        self._hover_idx = None
        self._refresh()

    def _draw_keys(self):
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        if cw < 20 or ch < 20: return

        pad   = max(2, int(cw * PAD_RATIO))
        unit  = (cw - pad*(TOTAL_UNITS+1)) / TOTAL_UNITS
        key_h = (ch - pad*(NUM_ROWS+1)) / NUM_ROWS
        fs    = max(6, min(13, int(unit*0.28)))
        fs_s  = max(5, fs-2)

        self.canvas.delete("all")
        self.key_cells.clear()

        snap, _, _, _ = self._selected_snapshot()
        max_v = max(snap.values(), default=1)

        def on_enter(idx):
            self._hover_idx = idx
            self._show_pct(idx)

        def on_leave(idx):
            self._hide_pct(idx)

        for row_i, row in enumerate(self.rows):
            x = pad
            y = pad + row_i*(key_h+pad)
            for tup in row:
                n, s, kid, width = tup
                w     = width*unit + (width-1)*pad
                v     = snap.get(kid, 0)
                ratio = v / max_v
                bg    = heat_color(ratio)
                fg    = text_color(bg)
                ol    = BG if ratio > 0.05 else OUTLINE

                rect = self.canvas.create_rectangle(x,y,x+w,y+key_h,
                                                    fill=bg,outline=ol,width=1)
                if s is None:
                    txt = self.canvas.create_text(x+w/2,y+key_h/2,
                          text=n,font=("Courier New",fs),fill=fg)
                    texts = [txt]
                else:
                    t1 = self.canvas.create_text(x+5,y+key_h*0.65,
                         text=n,font=("Courier New",fs,"bold"),fill=fg,anchor="w")
                    t2 = self.canvas.create_text(x+w-4,y+key_h*0.22,
                         text=s,font=("Courier New",fs_s),fill=fg,anchor="e")
                    texts = [t1,t2]

                idx = len(self.key_cells)
                self.key_cells.append({
                    "kid": kid,
                    "rect": rect,
                    "texts": texts,
                    "normal": n,
                    "shift": s,
                })

                for tag in [rect] + texts:
                    self.canvas.tag_bind(tag, "<Enter>", lambda e, i=idx: on_enter(i))
                    self.canvas.tag_bind(tag, "<Leave>", lambda e, i=idx: on_leave(i))
                x += w+pad

    def _on_canvas_resize(self, _=None):
        if self._resize_job: self.root.after_cancel(self._resize_job)
        self._resize_job = self.root.after(80, self._draw_keys)

    def _on_root_configure(self, event):
        if event.widget is not self.root:
            return
        if self._stats_expanded:
            self._layout_stats_panel()

    def _toggle_stats_panel(self, _=None):
        self._stats_expanded = not self._stats_expanded
        self._layout_stats_panel()
        self._redraw_bars()

    def _layout_stats_panel(self):
        if self._stats_expanded:
            # Expand chart over the keyboard area: hide keyboard canvas and let
            # stats fill the available center area.
            self.canvas.pack_forget()
            self._stats_frame.pack_forget()
            self._stats_frame.configure(height=max(self.BARS_H_EXPANDED_MIN, self.root.winfo_height() - 90))
            self._stats_frame.pack(fill="both", expand=True, padx=14, pady=(0,14))
            self._stats_title.configure(text="TOP KEYS  [click chart to collapse]")
        else:
            self._stats_frame.pack_forget()
            self.canvas.pack(fill="both", expand=True, padx=14, pady=(0,4))
            self._stats_frame.configure(height=self.BARS_H)
            self._stats_frame.pack(fill="x", padx=14, pady=(0,14))
            self._stats_title.configure(text="TOP KEYS  [click chart to expand]")

    def _on_bars_mousewheel(self, event):
        if not self._stats_expanded:
            return

        if hasattr(event, "delta") and event.delta:
            step = -1 if event.delta > 0 else 1
        elif getattr(event, "num", None) == 4:
            step = -1
        elif getattr(event, "num", None) == 5:
            step = 1
        else:
            return

        self.bars_canvas.yview_scroll(step, "units")
        self._draw_bars_indicators()

    def _on_bars_yview(self, first, last):
        # Called whenever yview changes; keep subtle overflow indicators in sync.
        self._draw_bars_indicators(float(first), float(last))

    def _draw_bars_indicators(self, first=None, last=None):
        bc = self.bars_canvas
        bc.delete(self._bars_indicator_tag)
        if not self._stats_expanded:
            return

        if first is None or last is None:
            first, last = bc.yview()

        cw = bc.winfo_width()
        ch = bc.winfo_height()
        x = cw - 10
        if first > 0.001:
            bc.create_polygon(
                x-4, 10, x+4, 10, x, 4,
                fill="#93c5fd", outline="", tags=self._bars_indicator_tag
            )
        if last < 0.999:
            bc.create_polygon(
                x-4, ch-10, x+4, ch-10, x, ch-4,
                fill="#93c5fd", outline="", tags=self._bars_indicator_tag
            )

    def _refresh(self):
        self._sync_app_selector()
        snap, total, selected_id, active = self._selected_snapshot()
        max_v = max(snap.values(), default=1)

        for cell in self.key_cells:
            v     = snap.get(cell["kid"],0)
            ratio = v/max_v
            bg    = heat_color(ratio)
            fg    = text_color(bg)
            ol    = BG if ratio>0.05 else OUTLINE
            self.canvas.itemconfig(cell["rect"],fill=bg,outline=ol)
            for txt in cell["texts"]:
                self.canvas.itemconfig(txt,fill=fg)

        if self._hover_idx is not None and 0 <= self._hover_idx < len(self.key_cells):
            self._show_pct(self._hover_idx)

        if selected_id == APP_ALL_ID:
            scope = "all apps"
        else:
            scope = self.view_var.get()
        active_label = ""
        if active:
            active_name = _display_app_label(active)
            active_label = f"  |  active: {active_name}"
        self.total_lbl.config(text=f"{total:,} keystrokes  |  {scope}{active_label}")
        self._update_bars(snap,total)

    def _update_bars(self,snap,total):
        self._bars_snap  = snap
        self._bars_total = total
        self._redraw_bars()

    def _redraw_bars(self):
        snap  = self._bars_snap
        total = self._bars_total or 1
        bc    = self.bars_canvas
        bc.delete("all")
        cw = bc.winfo_width()
        if cw < 20: return
        rows = sorted(snap.items(), key=lambda x:x[1], reverse=True)
        if not rows:
            bc.configure(scrollregion=(0, 0, cw, 0))
            return
        if not self._stats_expanded:
            rows = rows[:10]

        max_v  = rows[0][1]
        row_h  = 14
        lbl_w  = 90
        num_w  = 120
        bar_x0 = lbl_w+6
        bar_w  = max(10, cw-lbl_w-num_w-18)
        for i,(k,v) in enumerate(rows):
            y=i*(row_h+3)+2; cy=y+row_h/2
            ratio=v/max_v; pct=100*v/total
            bc.create_text(lbl_w-4,cy,text=repr(k) if len(k)==1 else k,
                           anchor="e",font=("Courier New",8),fill="#94a3b8")
            bc.create_rectangle(bar_x0,y,bar_x0+bar_w,y+row_h,fill="#1a2030",outline="")
            bc.create_rectangle(bar_x0,y,bar_x0+max(2,int(bar_w*ratio)),y+row_h,
                                fill=heat_color(ratio),outline="")
            bc.create_text(bar_x0+bar_w+6,cy,text=f"{v:,}  ({pct:.1f}%)",
                           anchor="w",font=("Courier New",8),fill="#64748b")
        content_h = len(rows) * (row_h + 3) + 6
        bc.configure(scrollregion=(0, 0, cw, content_h))
        self._draw_bars_indicators()

    def _auto_refresh(self):
        flush_stats_if_dirty()
        self._refresh()
        self.root.after(2000,self._auto_refresh)

    def _show_pct(self, idx):
        if not (0 <= idx < len(self.key_cells)):
            return
        cell = self.key_cells[idx]
        snap, total, _, _ = self._selected_snapshot()
        pct = (100.0 * snap.get(cell["kid"], 0)) / total
        txt = f"{pct:.1f}%"
        if len(cell["texts"]) == 1:
            self.canvas.itemconfig(cell["texts"][0], text=txt)
        else:
            self.canvas.itemconfig(cell["texts"][0], text=txt)
            self.canvas.itemconfig(cell["texts"][1], text="")

    def _hide_pct(self, idx):
        if not (0 <= idx < len(self.key_cells)):
            return
        cell = self.key_cells[idx]
        if self._hover_idx == idx:
            self._hover_idx = None
        if len(cell["texts"]) == 1:
            self.canvas.itemconfig(cell["texts"][0], text=cell["normal"])
        else:
            self.canvas.itemconfig(cell["texts"][0], text=cell["normal"])
            self.canvas.itemconfig(cell["texts"][1], text=cell["shift"])

    def _reset(self):
        import tkinter.messagebox
        selected_label = self.view_var.get()
        selected_id = self.view_map.get(selected_label, APP_ALL_ID)

        if selected_id == APP_ALL_ID:
            ok = tkinter.messagebox.askyesno("Reset", "Clear ALL keystroke data for all apps?")
            if not ok:
                return
            with counts_lock:
                counts_by_app.clear()
                global stats_dirty
                stats_dirty = True
        else:
            ok = tkinter.messagebox.askyesno("Reset", f"Clear keystroke data for '{selected_label}'?")
            if not ok:
                return
            with counts_lock:
                counts_by_app.pop(selected_id, None)
                global stats_dirty
                stats_dirty = True

        flush_stats_if_dirty()
        self._sync_app_selector()
        self._draw_keys()
        self._refresh()

    def _save_geometry(self):
        self.prefs.update({"w":self.root.winfo_width(),"h":self.root.winfo_height(),
                           "x":self.root.winfo_x(),"y":self.root.winfo_y()})
        save_prefs(self.prefs)

    def _on_close_btn(self):
        self._save_geometry()
        if self._tray_mode:
            self.hide()
        else:
            flush_stats_if_dirty()
            self.root.destroy()

    def show(self):
        self._refresh(); self.root.deiconify(); self.root.lift(); self.root.focus_force()
    def hide(self):
        self._save_geometry(); self.root.withdraw()
    def run(self):
        self.root.mainloop(); flush_stats_if_dirty()

# ── tray (native Windows) ─────────────────────────────────────────────────────
if platform.system() == "Windows":
    from ctypes import wintypes

class NativeTrayIcon:
    WM_USER = 0x0400
    WM_COMMAND = 0x0111
    WM_DESTROY = 0x0002
    WM_CLOSE = 0x0010

    WM_LBUTTONUP = 0x0202
    WM_LBUTTONDBLCLK = 0x0203
    WM_RBUTTONUP = 0x0205

    NIM_ADD = 0x00000000
    NIM_MODIFY = 0x00000001
    NIM_DELETE = 0x00000002

    NIF_MESSAGE = 0x00000001
    NIF_ICON = 0x00000002
    NIF_TIP = 0x00000004

    IMAGE_ICON = 1
    LR_LOADFROMFILE = 0x00000010
    LR_DEFAULTSIZE = 0x00000040

    MF_STRING = 0x00000000
    TPM_RETURNCMD = 0x0100
    TPM_NONOTIFY = 0x0080

    ID_SHOW = 1001
    ID_QUIT = 1002

    def __init__(self, on_show, on_quit, tooltip="Key Heatmap"):
        self._on_show = on_show
        self._on_quit = on_quit
        self._tooltip = tooltip

        self._ready = threading.Event()
        self._failed = threading.Event()
        self._thread = None

        self._hwnd = None
        self._menu = None
        self._hicon = None
        self._nid = None
        self._class_name = "KeyHeatmapTrayWindow"
        self._msg_tray = self.WM_USER + 1
        self._wnd_proc_ref = None

    class POINT(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

    class NOTIFYICONDATAW(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD),
            ("hWnd", wintypes.HWND),
            ("uID", wintypes.UINT),
            ("uFlags", wintypes.UINT),
            ("uCallbackMessage", wintypes.UINT),
            ("hIcon", wintypes.HICON),
            ("szTip", wintypes.WCHAR * 128),
            ("dwState", wintypes.DWORD),
            ("dwStateMask", wintypes.DWORD),
            ("szInfo", wintypes.WCHAR * 256),
            ("uTimeoutOrVersion", wintypes.UINT),
            ("szInfoTitle", wintypes.WCHAR * 64),
            ("dwInfoFlags", wintypes.DWORD),
            ("guidItem", ctypes.c_byte * 16),
            ("hBalloonIcon", wintypes.HICON),
        ]

    class WNDCLASSW(ctypes.Structure):
        _fields_ = [
            ("style", wintypes.UINT),
            ("lpfnWndProc", ctypes.c_void_p),
            ("cbClsExtra", ctypes.c_int),
            ("cbWndExtra", ctypes.c_int),
            ("hInstance", wintypes.HINSTANCE),
            ("hIcon", wintypes.HICON),
            ("hCursor", wintypes.HANDLE),
            ("hbrBackground", wintypes.HANDLE),
            ("lpszMenuName", wintypes.LPCWSTR),
            ("lpszClassName", wintypes.LPCWSTR),
        ]

    def start(self, timeout=2.0):
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        return self._ready.wait(timeout) and not self._failed.is_set()

    def stop(self, timeout=1.0):
        if self._hwnd:
            ctypes.windll.user32.PostMessageW(self._hwnd, self.WM_CLOSE, 0, 0)
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout)

    def _load_icon(self):
        user32 = ctypes.windll.user32
        hicon = None
        if ICON_FILE.exists():
            hicon = user32.LoadImageW(
                None,
                str(ICON_FILE),
                self.IMAGE_ICON,
                0,
                0,
                self.LR_LOADFROMFILE | self.LR_DEFAULTSIZE,
            )
        if not hicon:
            hicon = user32.LoadIconW(None, ctypes.c_wchar_p(32512))
        return hicon

    def _build_notify_data(self):
        nid = self.NOTIFYICONDATAW()
        nid.cbSize = ctypes.sizeof(self.NOTIFYICONDATAW)
        nid.hWnd = self._hwnd
        nid.uID = 1
        nid.uFlags = self.NIF_MESSAGE | self.NIF_ICON | self.NIF_TIP
        nid.uCallbackMessage = self._msg_tray
        nid.hIcon = self._hicon
        nid.szTip = self._tooltip
        return nid

    def _show_menu(self):
        user32 = ctypes.windll.user32
        pt = self.POINT()
        user32.GetCursorPos(ctypes.byref(pt))
        user32.SetForegroundWindow(self._hwnd)
        cmd = user32.TrackPopupMenu(
            self._menu,
            self.TPM_RETURNCMD | self.TPM_NONOTIFY,
            pt.x,
            pt.y,
            0,
            self._hwnd,
            None,
        )
        if cmd:
            user32.PostMessageW(self._hwnd, self.WM_COMMAND, int(cmd), 0)

    def _run(self):
        user32 = ctypes.windll.user32
        shell32 = ctypes.windll.shell32
        kernel32 = ctypes.windll.kernel32

        user32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT, ctypes.c_size_t, ctypes.c_size_t]
        user32.DefWindowProcW.restype = ctypes.c_ssize_t
        user32.TrackPopupMenu.argtypes = [wintypes.HMENU, wintypes.UINT, ctypes.c_int, ctypes.c_int, ctypes.c_int, wintypes.HWND, ctypes.c_void_p]
        user32.TrackPopupMenu.restype = wintypes.UINT
        user32.AppendMenuW.argtypes = [wintypes.HMENU, wintypes.UINT, ctypes.c_size_t, wintypes.LPCWSTR]
        user32.AppendMenuW.restype = wintypes.BOOL

        WNDPROC = ctypes.WINFUNCTYPE(
            ctypes.c_ssize_t,
            wintypes.HWND,
            wintypes.UINT,
            ctypes.c_size_t,
            ctypes.c_size_t,
        )

        def _wnd_proc(hwnd, msg, wparam, lparam):
            if msg == self._msg_tray:
                event_code = int(lparam) & 0xFFFF
                if event_code in (self.WM_LBUTTONUP, self.WM_LBUTTONDBLCLK):
                    self._on_show()
                elif event_code == self.WM_RBUTTONUP:
                    self._show_menu()
                return 0
            if msg == self.WM_COMMAND:
                cmd_id = int(wparam) & 0xFFFF
                if cmd_id == self.ID_SHOW:
                    self._on_show()
                    return 0
                if cmd_id == self.ID_QUIT:
                    self._on_quit()
                    return 0
            if msg == self.WM_DESTROY:
                if self._nid is not None:
                    shell32.Shell_NotifyIconW(self.NIM_DELETE, ctypes.byref(self._nid))
                user32.PostQuitMessage(0)
                return 0
            return int(user32.DefWindowProcW(hwnd, msg, wparam, lparam))

        try:
            self._wnd_proc_ref = WNDPROC(_wnd_proc)
            hinstance = kernel32.GetModuleHandleW(None)

            wc = self.WNDCLASSW()
            wc.lpfnWndProc = ctypes.cast(self._wnd_proc_ref, ctypes.c_void_p).value
            wc.hInstance = hinstance
            wc.lpszClassName = self._class_name

            user32.RegisterClassW(ctypes.byref(wc))

            self._hwnd = user32.CreateWindowExW(
                0,
                self._class_name,
                self._class_name,
                0,
                0,
                0,
                0,
                0,
                None,
                None,
                hinstance,
                None,
            )
            if not self._hwnd:
                raise RuntimeError("CreateWindowExW failed")

            self._menu = user32.CreatePopupMenu()
            user32.AppendMenuW(self._menu, self.MF_STRING, self.ID_SHOW, ctypes.c_wchar_p("Show Heatmap"))
            user32.AppendMenuW(self._menu, self.MF_STRING, self.ID_QUIT, ctypes.c_wchar_p("Quit"))

            self._hicon = self._load_icon()
            self._nid = self._build_notify_data()
            if not shell32.Shell_NotifyIconW(self.NIM_ADD, ctypes.byref(self._nid)):
                raise RuntimeError("Shell_NotifyIconW NIM_ADD failed")

            self._ready.set()

            msg = wintypes.MSG()
            while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
        except Exception:
            self._failed.set()
            self._ready.set()
        finally:
            if self._menu:
                user32.DestroyMenu(self._menu)
                self._menu = None
            if self._hwnd:
                user32.DestroyWindow(self._hwnd)
                self._hwnd = None

def run_with_tray(win):
    tray = NativeTrayIcon(
        on_show=lambda: win.root.after(0, win.show),
        on_quit=lambda: win.root.after(0, win.root.destroy),
    )

    if not tray.start(timeout=2.0):
        win._tray_mode = False
        win.show()
        win.run()
        return

    win.hide()
    win.run()
    tray.stop()
if __name__ == "__main__":
    import tkinter.messagebox

    if not HAS_TRAY and TRAY_IMPORT_ERROR:
        print(f"Tray support disabled: {TRAY_IMPORT_ERROR}")

    win = HeatmapWindow(start_hidden=HAS_TRAY)
    if HAS_TRAY:
        run_with_tray(win)
    else:
        win.run()







