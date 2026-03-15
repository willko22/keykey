MAIN_VERSION = 2
FEATURE_VERSION = 2
FIXES_VERSION = 0

"""
Key Heatmap – tracks keystrokes with heatmap visualization.
Auto-detects keyboard layout (QWERTY/QWERTZ/AZERTY/Dvorak/Colemak).
"""

import json, sys, threading, platform, math, os, time, ctypes, subprocess
from collections import Counter
from pathlib import Path
from pynput import keyboard as pynput_kb
import tkinter as tk
from tkinter import ttk

try:
    import tomllib
except Exception:
    tomllib = None

HAS_TRAY = platform.system() == "Windows"
TRAY_IMPORT_ERROR = None if HAS_TRAY else "Native tray is only supported on Windows"

# ── storage ───────────────────────────────────────────────────────────────────
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
SAVE_FILE  = BASE_DIR / "stats.json"
PREFS_FILE = BASE_DIR / "prefs.json"
LEGACY_PREFS_FILE = Path.home() / ".key_heatmap_prefs.json"
CONFIG_FILE = BASE_DIR / "keykey.toml"
ICON_FILE  = BASE_DIR / "keykey.ico"
APP_ALL_ID = "__all__"
UNATTRIBUTED_APP_ID = "__unattributed__"
DOTA2_APP_ID = "__dota2__"
WINDOW_TITLE_PREFIX = "window::"
APP_VERSION = f"{MAIN_VERSION}.{FEATURE_VERSION}.{FIXES_VERSION}"

GROUP_ORDER = ("game", "coding", "office", "creative")
OTHERS_GROUP_ID = "__others__"

DEFAULT_CONFIG = {
    "color": {
        "scale_mode": "adaptive",   # adaptive | smooth | stepped
        "percent_step": None,
        "color_count": 20,
        "gamma": 0.55,
        "outlier_mode": "log",      # log | none
        "log_base": 10.0,
        "cap_percentile": 95.0,
    },
    "apps": {
        "groups": {k: [] for k in GROUP_ORDER},
        "title_contains": {k: [] for k in GROUP_ORDER},
        "exclude": [],
    },
}

def _normalize_group_id(value):
    if not isinstance(value, str):
        return None
    gid = value.strip().lower().replace(" ", "_")
    if not gid or gid == OTHERS_GROUP_ID:
        return None
    return gid

def _configured_group_ids(apps_cfg):
    groups = apps_cfg.get("groups", {}) if isinstance(apps_cfg, dict) else {}
    titles = apps_cfg.get("title_contains", {}) if isinstance(apps_cfg, dict) else {}

    ids = set()
    for k in groups.keys():
        gid = _normalize_group_id(k)
        if gid:
            ids.add(gid)
    for k in titles.keys():
        gid = _normalize_group_id(k)
        if gid:
            ids.add(gid)

    if not ids:
        ids.update(GROUP_ORDER)

    return sorted(ids, key=lambda s: s.lower())

def _group_display_name(group_id):
    if group_id == OTHERS_GROUP_ID:
        return "Others"
    return str(group_id).replace("_", " ").title()

def _clean_string_list(v):
    if not isinstance(v, list):
        return []
    out = []
    for item in v:
        if isinstance(item, str):
            s = item.strip()
            if s:
                out.append(s)
    return out

def _to_float(v, default):
    try:
        return float(v)
    except Exception:
        return default

def _to_int(v, default):
    try:
        return int(v)
    except Exception:
        return default

def _load_config():
    cfg = {
        "color": dict(DEFAULT_CONFIG["color"]),
        "apps": {
            "groups": {k: [] for k in GROUP_ORDER},
            "title_contains": {k: [] for k in GROUP_ORDER},
            "exclude": [],
        },
    }

    if not CONFIG_FILE.exists() or tomllib is None:
        return cfg

    try:
        with open(CONFIG_FILE, "rb") as f:
            raw = tomllib.load(f)
    except Exception:
        return cfg

    color = raw.get("color")
    if isinstance(color, dict):
        mode = str(color.get("scale_mode", cfg["color"]["scale_mode"]))
        mode = mode.strip().lower()
        if mode in ("adaptive", "smooth", "stepped"):
            cfg["color"]["scale_mode"] = mode

        ps = color.get("percent_step", None)
        if ps is None:
            cfg["color"]["percent_step"] = None
        else:
            psv = _to_float(ps, None)
            if psv is not None and psv > 0:
                cfg["color"]["percent_step"] = psv

        cfg["color"]["color_count"] = max(2, _to_int(color.get("color_count", cfg["color"]["color_count"]), cfg["color"]["color_count"]))
        cfg["color"]["gamma"] = max(0.01, _to_float(color.get("gamma", cfg["color"]["gamma"]), cfg["color"]["gamma"]))

        outlier_mode = str(color.get("outlier_mode", cfg["color"]["outlier_mode"]))
        outlier_mode = outlier_mode.strip().lower()
        if outlier_mode in ("log", "none"):
            cfg["color"]["outlier_mode"] = outlier_mode

        cfg["color"]["log_base"] = max(1.0001, _to_float(color.get("log_base", cfg["color"]["log_base"]), cfg["color"]["log_base"]))

        cap_pct = _to_float(color.get("cap_percentile", cfg["color"]["cap_percentile"]), cfg["color"]["cap_percentile"])
        cfg["color"]["cap_percentile"] = min(100.0, max(50.0, cap_pct))

    apps = raw.get("apps")
    if isinstance(apps, dict):
        cfg["apps"]["exclude"] = _clean_string_list(apps.get("exclude", []))

        parsed_groups = None
        groups = apps.get("groups")
        if isinstance(groups, dict):
            parsed_groups = {}
            for key, vals in groups.items():
                gid = _normalize_group_id(key)
                if gid:
                    parsed_groups[gid] = _clean_string_list(vals)

        parsed_titles = None
        title_contains = apps.get("title_contains")
        if isinstance(title_contains, dict):
            parsed_titles = {}
            for key, vals in title_contains.items():
                gid = _normalize_group_id(key)
                if gid:
                    parsed_titles[gid] = _clean_string_list(vals)

        if parsed_groups is not None:
            cfg["apps"]["groups"] = parsed_groups
        if parsed_titles is not None:
            cfg["apps"]["title_contains"] = parsed_titles

        all_group_ids = set(cfg["apps"]["groups"].keys()) | set(cfg["apps"]["title_contains"].keys())
        if not all_group_ids:
            all_group_ids = set(GROUP_ORDER)
        cfg["apps"]["groups"] = {
            gid: _clean_string_list(cfg["apps"]["groups"].get(gid, [])) for gid in all_group_ids
        }
        cfg["apps"]["title_contains"] = {
            gid: _clean_string_list(cfg["apps"]["title_contains"].get(gid, [])) for gid in all_group_ids
        }

    return cfg

def _matches_group(app_id, group_id, apps_cfg):
    groups = apps_cfg.get("groups", {})
    titles = apps_cfg.get("title_contains", {})
    rules = [r.lower() for r in groups.get(group_id, [])]
    title_rules = [r.lower() for r in titles.get(group_id, [])]

    app_low = (app_id or "").lower()
    disp = (_display_app_label(app_id) or "").lower()
    stem = ""
    name = ""
    title = ""
    if app_id and app_id.startswith(WINDOW_TITLE_PREFIX):
        title = app_id[len(WINDOW_TITLE_PREFIX):].lower()
    else:
        p = Path(app_id or "")
        stem = p.stem.lower()
        name = p.name.lower()

    for rule in rules:
        if not rule:
            continue
        if rule in (stem, name, disp):
            return True
        if rule in app_low or (disp and rule in disp):
            return True

    hay_title = title or disp
    for rule in title_rules:
        if rule and rule in hay_title:
            return True

    return False

def _infer_app_group(app_id, apps_cfg):
    for g in _configured_group_ids(apps_cfg):
        if _matches_group(app_id, g, apps_cfg):
            return g
    return None

def _is_excluded_app(app_id, apps_cfg):
    if not app_id or not isinstance(apps_cfg, dict):
        return False

    excludes = [s.lower() for s in _clean_string_list(apps_cfg.get("exclude", []))]
    if not excludes:
        return False

    app_low = str(app_id).lower()
    disp = (_display_app_label(app_id) or "").lower()
    stem = ""
    name = ""
    title = ""

    if app_id.startswith(WINDOW_TITLE_PREFIX):
        title = app_id[len(WINDOW_TITLE_PREFIX):].lower()
    else:
        p = Path(app_id)
        stem = p.stem.lower()
        name = p.name.lower()

    for token in excludes:
        if not token:
            continue
        if token in app_low or (disp and token in disp) or (title and token in title):
            return True
        if token in (stem, name):
            return True
    return False

def _percentile(values, pct):
    if not values:
        return 0.0
    s = sorted(values)
    if len(s) == 1:
        return float(s[0])
    p = max(0.0, min(100.0, float(pct))) / 100.0
    i = p * (len(s) - 1)
    lo = int(math.floor(i))
    hi = int(math.ceil(i))
    if lo == hi:
        return float(s[lo])
    t = i - lo
    return float(s[lo] + (s[hi] - s[lo]) * t)

def _effective_color_bins(color_cfg):
    ps = color_cfg.get("percent_step")
    if ps is not None:
        ps = _to_float(ps, None)
        if ps is not None and ps > 0:
            return max(2, int(round(100.0 / ps)))
    return max(2, _to_int(color_cfg.get("color_count", 20), 20))

def _build_color_ratio_fn(snap, color_cfg):
    vals = [v for v in snap.values() if v > 0]
    if not vals:
        return lambda v: 0.0

    mode = color_cfg.get("scale_mode", "adaptive")
    outlier_mode = color_cfg.get("outlier_mode", "log")
    cap_pct = color_cfg.get("cap_percentile", 95.0)
    cap = max(vals)

    if mode in ("adaptive", "stepped"):
        cap = max(1.0, _percentile(vals, cap_pct))
    else:
        cap = max(1.0, float(cap))

    use_log = mode in ("adaptive", "stepped") and outlier_mode == "log"
    if use_log:
        base = max(1.0001, _to_float(color_cfg.get("log_base", 10.0), 10.0))
        denom = math.log(1.0 + cap, base)
        if denom <= 0:
            denom = 1.0

        def ratio_fn(v):
            vv = min(max(0.0, float(v)), cap)
            return max(0.0, min(1.0, math.log(1.0 + vv, base) / denom))

        return ratio_fn

    def ratio_fn(v):
        vv = min(max(0.0, float(v)), cap)
        return max(0.0, min(1.0, vv / cap))

    return ratio_fn

def _parse_semver(version_str):
    s = str(version_str or "").strip()
    if s.startswith("v"):
        s = s[1:]
    parts = s.split(".")
    nums = []
    for p in parts[:3]:
        token = ""
        for ch in p:
            if ch.isdigit():
                token += ch
            else:
                break
        nums.append(int(token) if token else 0)
    while len(nums) < 3:
        nums.append(0)
    return tuple(nums)

def _is_newer_version(current, other):
    return _parse_semver(current) > _parse_semver(other)

def _windows_process_exe_path(pid):
    kernel32 = ctypes.windll.kernel32
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
    if not handle:
        return None
    try:
        buf_len = ctypes.c_ulong(32768)
        buf = ctypes.create_unicode_buffer(buf_len.value)
        if not kernel32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(buf_len)):
            return None
        return str(Path(buf.value))
    finally:
        kernel32.CloseHandle(handle)

def _list_running_keykey_exe_paths(exe_name):
    if platform.system() != "Windows":
        return []

    kernel32 = ctypes.windll.kernel32
    TH32CS_SNAPPROCESS = 0x00000002

    class PROCESSENTRY32W(ctypes.Structure):
        _fields_ = [
            ("dwSize", ctypes.c_ulong),
            ("cntUsage", ctypes.c_ulong),
            ("th32ProcessID", ctypes.c_ulong),
            ("th32DefaultHeapID", ctypes.c_void_p),
            ("th32ModuleID", ctypes.c_ulong),
            ("cntThreads", ctypes.c_ulong),
            ("th32ParentProcessID", ctypes.c_ulong),
            ("pcPriClassBase", ctypes.c_long),
            ("dwFlags", ctypes.c_ulong),
            ("szExeFile", ctypes.c_wchar * 260),
        ]

    snap = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if snap == ctypes.c_void_p(-1).value:
        return []

    current_pid = os.getpid()
    exe_low = str(exe_name or "").lower()
    matches = []
    try:
        entry = PROCESSENTRY32W()
        entry.dwSize = ctypes.sizeof(PROCESSENTRY32W)
        ok = kernel32.Process32FirstW(snap, ctypes.byref(entry))
        while ok:
            pid = int(entry.th32ProcessID)
            pname = str(entry.szExeFile or "").lower()
            if pid != current_pid and pname == exe_low:
                pth = _windows_process_exe_path(pid)
                if pth:
                    matches.append((pid, pth))
            ok = kernel32.Process32NextW(snap, ctypes.byref(entry))
    finally:
        kernel32.CloseHandle(snap)

    return matches

def _query_exe_version(exe_path):
    if not exe_path:
        return None
    try:
        flags = 0
        if platform.system() == "Windows":
            flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        out = subprocess.check_output(
            [exe_path, "--version"],
            stderr=subprocess.STDOUT,
            timeout=4,
            creationflags=flags,
            text=True,
        )
        v = (out or "").strip().splitlines()
        return v[0].strip() if v else None
    except Exception:
        return None

def _terminate_pid_windows(pid):
    kernel32 = ctypes.windll.kernel32
    PROCESS_TERMINATE = 0x0001
    handle = kernel32.OpenProcess(PROCESS_TERMINATE, False, int(pid))
    if not handle:
        return False
    try:
        return bool(kernel32.TerminateProcess(handle, 0))
    finally:
        kernel32.CloseHandle(handle)

def _handle_existing_instance_upgrade():
    # This behavior targets packaged EXE workflow; avoid matching generic python.exe.
    if platform.system() != "Windows" or not getattr(sys, "frozen", False):
        return True

    exe_name = Path(sys.executable).name
    others = _list_running_keykey_exe_paths(exe_name)
    if not others:
        return True

    allow_start = True
    for pid, exe_path in others:
        running_ver = _query_exe_version(exe_path)
        if running_ver is None:
            allow_start = False
            continue

        if _is_newer_version(APP_VERSION, running_ver):
            _terminate_pid_windows(pid)
            continue

        allow_start = False

    return allow_start

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
        with open(PREFS_FILE, encoding="utf-8") as f:
            return json.load(f)
    if LEGACY_PREFS_FILE.exists():
        try:
            with open(LEGACY_PREFS_FILE, encoding="utf-8") as f:
                data = json.load(f)
            save_prefs(data)
            return data
        except Exception:
            pass
    return {}
def save_prefs(p):
    with open(PREFS_FILE, "w", encoding="utf-8") as f:
        json.dump(p, f, indent=2)

_RUNTIME_APPS_CFG = dict(DEFAULT_CONFIG["apps"])

def _set_runtime_apps_cfg(apps_cfg):
    global _RUNTIME_APPS_CFG
    if isinstance(apps_cfg, dict):
        _RUNTIME_APPS_CFG = apps_cfg

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
        if _is_excluded_app(app_id, _RUNTIME_APPS_CFG):
            app_id = None
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

def _key_category(kid):
    """Return 'alpha', 'numeral', 'symbols', 'ws', or 'special' for a key_id."""
    if kid in ("space", "tab"):  return "ws"
    if len(kid) == 1:
        if kid.isalpha():  return "alpha"
        if kid.isdigit():  return "numeral"
        return "symbols"   # single printable non-alpha non-digit (`, -, =, [, etc.)
    return "special"       # multi-char = named key (shift, return, ctrl, …)

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
        if not app_id or _is_excluded_app(app_id, _RUNTIME_APPS_CFG):
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

_set_runtime_apps_cfg(_load_config()["apps"])

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

def heat_color(ratio, gamma=0.55, bins=None):
    if ratio <= 0: return "#16192a"
    ratio = math.pow(max(0.0, min(1.0, ratio)), max(0.01, float(gamma)))
    if bins and bins > 1:
        ratio = round(ratio * (bins - 1)) / (bins - 1)
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
        self.config      = _load_config()
        self.color_cfg   = self.config["color"]
        self.apps_cfg    = self.config["apps"]
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
        self.filters_btn = None
        self._filters_panel_open = False
        self._filters_panel = None
        self._filters_checks_frame = None
        self._filter_group_vars = {}
        self._filter_app_vars = {}
        self._filters_structure_sig = None
        self._main_body = None
        self._main_left = None
        self._configured_groups = []
        self._group_ids = []
        self._group_to_apps = {}
        self._app_to_group = {}
        self._app_labels = {}
        self._app_ids = []
        self._group_state = {}
        self._app_state = {}
        self._filters_loaded_from_prefs = False
        self._hover_idx  = None
        self._resize_job = None
        self._geom_save_job = None
        self._stats_expanded = False
        self._stats_frame = None
        self._stats_title = None
        self._bars_indicator_tag = "__bars_indicator__"
        self._bars_snap  = {}
        self._bars_total = 1
        self._scale_mode_lbl = None
        self.show_alpha_var   = tk.BooleanVar(value=self.prefs.get("show_alpha",   True))
        self.show_numeral_var = tk.BooleanVar(value=self.prefs.get("show_numeral", True))
        self.show_symbols_var = tk.BooleanVar(value=self.prefs.get("show_symbols", True))
        self.show_special_var = tk.BooleanVar(value=self.prefs.get("show_special", True))
        self.show_ws_var      = tk.BooleanVar(value=self.prefs.get("show_ws",      True))

        self._build_ui()
        self._sync_filter_options()
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
        hdr = tk.Frame(self.root, bg=BG)
        hdr.pack(fill="x", padx=M, pady=(M,4))
        tk.Label(hdr, text="KEY HEATMAP", font=("Courier New",14,"bold"),
                 fg="#e8f4fd", bg=BG).pack(side="left")
        tk.Label(hdr, text=f"[{self.layout_name.upper()}  auto-detected]",
                 font=("Courier New",9), fg="#374151", bg=BG).pack(side="left", padx=10)
        self._scale_mode_lbl = tk.Label(hdr, text="", font=("Courier New",9), fg="#4b5563", bg=BG)
        self._scale_mode_lbl.pack(side="left", padx=8)
        self.total_lbl = tk.Label(hdr, text="", font=("Courier New",10), fg="#6b7280", bg=BG)
        self.total_lbl.pack(side="right")

        bf = tk.Frame(self.root, bg=BG)
        bf.pack(fill="x", padx=M, pady=(0,6))

        def mkbtn(p,t,cmd,fg,abg):
            return tk.Button(p,text=t,command=cmd,bg="#1a2030",fg=fg,relief="flat",
                             font=("Courier New",9),padx=10,pady=3,cursor="hand2",
                             activebackground=abg,activeforeground=fg,bd=0)
        self.filters_btn = mkbtn(bf, "Filters", self._toggle_filters_panel, "#facc15", "#3f2f00")
        self.filters_btn.pack(side="left", padx=(0, 10))

        mkbtn(bf,"Refresh",    self._on_refresh_clicked, "#7dd3fc","#1e3a50").pack(side="left",padx=(0,6))
        mkbtn(bf,"Reset Stats",self._reset,   "#f87171","#3d1f1f").pack(side="left")

        key_filters_wrap = tk.Frame(bf, bg=BG)
        key_filters_wrap.pack(side="right")
        tk.Label(key_filters_wrap, text="|", font=("Courier New",9), fg="#374151", bg=BG).pack(side="left", padx=(0,6))
        for _lbl, _var in (
            ("Alpha",   self.show_alpha_var),
            ("Numeral", self.show_numeral_var),
            ("Symbols", self.show_symbols_var),
            ("Special", self.show_special_var),
            ("WS",      self.show_ws_var),
        ):
            tk.Checkbutton(
                key_filters_wrap, text=_lbl,
                variable=_var,
                command=self._on_filter_changed,
                bg=BG, fg="#9ca3af", selectcolor="#1a2030",
                activebackground=BG, activeforeground="#e5e7eb",
                font=("Courier New", 9),
            ).pack(side="left", padx=(0, 4))

        self._main_body = tk.Frame(self.root, bg=BG)
        self._main_body.pack(fill="both", expand=True, padx=M, pady=(0, M))

        self._main_left = tk.Frame(self._main_body, bg=BG)
        self._main_left.pack(side="right", fill="both", expand=True)

        self.canvas = tk.Canvas(self._main_left, bg=BG, highlightthickness=0)
        self.canvas.pack(fill="both", expand=True, pady=(0,4))

        self._stats_frame = tk.Frame(self._main_left, bg=BG, height=self.BARS_H)
        self._stats_frame.pack(fill="x")
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

        self._filters_panel = tk.Frame(self._main_body, bg="#111827", width=290, highlightthickness=1, highlightbackground="#223046")
        self._filters_panel.pack_propagate(False)

        panel_title = tk.Label(
            self._filters_panel,
            text="APP FILTERS",
            font=("Courier New", 10, "bold"),
            fg="#facc15",
            bg="#111827",
        )
        panel_title.pack(anchor="w", padx=10, pady=(10, 6))

        panel_hint = tk.Label(
            self._filters_panel,
            text="Toggle groups or individual apps",
            font=("Courier New", 8),
            fg="#6b7280",
            bg="#111827",
        )
        panel_hint.pack(anchor="w", padx=10, pady=(0, 8))

        self._filters_checks_frame = tk.Frame(self._filters_panel, bg="#111827")
        self._filters_checks_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def _persist_prefs(self):
        self.prefs.update({
            "show_alpha":   self.show_alpha_var.get(),
            "show_numeral": self.show_numeral_var.get(),
            "show_symbols": self.show_symbols_var.get(),
            "show_special": self.show_special_var.get(),
            "show_ws":      self.show_ws_var.get(),
            "checked_groups": sorted([gid for gid in self._group_ids if self._group_state.get(gid, True)]),
            "checked_apps": sorted([app_id for app_id in self._app_ids if self._app_state.get(app_id, True)]),
        })
        save_prefs(self.prefs)

    def _reload_config(self):
        self.config = _load_config()
        self.color_cfg = self.config["color"]
        self.apps_cfg = self.config["apps"]
        _set_runtime_apps_cfg(self.apps_cfg)

    def _collect_app_ids(self):
        with counts_lock:
            app_ids = set(counts_by_app.keys())
            if last_active_app_id:
                app_ids.add(last_active_app_id)
            if current_app_id:
                app_ids.add(current_app_id)
        app_ids = [app_id for app_id in app_ids if not _is_excluded_app(app_id, self.apps_cfg)]
        return sorted(app_ids, key=lambda s: s.lower())

    def _sync_filter_options(self, reset=False):
        app_ids = self._collect_app_ids()
        labels_by_name = make_app_labels(app_ids)
        self._app_labels = {app_id: label for label, app_id in labels_by_name.items()}
        self._app_ids = app_ids

        self._configured_groups = _configured_group_ids(self.apps_cfg)
        self._group_ids = sorted(self._configured_groups + [OTHERS_GROUP_ID], key=lambda g: _group_display_name(g).lower())

        group_to_apps = {gid: [] for gid in self._group_ids}
        app_to_group = {}
        for app_id in app_ids:
            gid = _infer_app_group(app_id, self.apps_cfg) or OTHERS_GROUP_ID
            if gid not in group_to_apps:
                group_to_apps[gid] = []
            group_to_apps[gid].append(app_id)
            app_to_group[app_id] = gid

        for gid in group_to_apps:
            group_to_apps[gid].sort(key=lambda app_id: self._app_labels.get(app_id, app_id).lower())

        self._group_to_apps = group_to_apps
        self._app_to_group = app_to_group

        if not self._filters_loaded_from_prefs:
            checked_apps = self.prefs.get("checked_apps")
            checked_groups = self.prefs.get("checked_groups")

            if isinstance(checked_apps, list):
                checked_app_set = {str(v) for v in checked_apps}
                for app_id in app_ids:
                    self._app_state[app_id] = app_id in checked_app_set
            elif isinstance(checked_groups, list):
                checked_group_set = set()
                for value in checked_groups:
                    raw = str(value).strip().lower()
                    gid = _normalize_group_id(raw)
                    if gid:
                        checked_group_set.add(gid)
                    elif raw in ("others", "other", OTHERS_GROUP_ID):
                        checked_group_set.add(OTHERS_GROUP_ID)
                for app_id in app_ids:
                    gid = app_to_group.get(app_id, OTHERS_GROUP_ID)
                    self._app_state[app_id] = gid in checked_group_set
            else:
                for app_id in app_ids:
                    self._app_state[app_id] = True

            self._filters_loaded_from_prefs = True
        else:
            existing_state = dict(self._app_state)
            self._app_state = {}
            for app_id in app_ids:
                self._app_state[app_id] = existing_state.get(app_id, True)

        if reset:
            for app_id in app_ids:
                self._app_state[app_id] = True

        self._group_state = {}
        for gid in self._group_ids:
            apps = self._group_to_apps.get(gid, [])
            if not apps:
                self._group_state[gid] = False
            else:
                self._group_state[gid] = all(self._app_state.get(app_id, True) for app_id in apps)

        if self._filters_panel_open:
            sig = self._current_filters_structure_sig()
            if sig != self._filters_structure_sig:
                self._rebuild_filters_panel()
                self._filters_structure_sig = sig
            else:
                self._sync_filters_panel_values()

    def _selected_app_ids(self):
        return [app_id for app_id in self._app_ids if self._app_state.get(app_id, True)]

    def _rebuild_filters_panel(self):
        if self._filters_checks_frame is None:
            return

        for child in self._filters_checks_frame.winfo_children():
            child.destroy()

        self._filter_group_vars = {}
        self._filter_app_vars = {}

        for i, gid in enumerate(self._group_ids):
            apps = self._group_to_apps.get(gid, [])
            has_apps = bool(apps)
            gvar = tk.BooleanVar(value=self._group_state.get(gid, True))
            self._filter_group_vars[gid] = gvar
            gchk = tk.Checkbutton(
                self._filters_checks_frame,
                text=_group_display_name(gid),
                variable=gvar,
                command=lambda g=gid, v=gvar: self._on_group_panel_toggle(g, v),
                bg="#111827",
                fg="#9ca3af",
                selectcolor="#1a2030",
                activebackground="#111827",
                activeforeground="#e5e7eb",
                font=("Courier New", 9, "bold"),
                anchor="w",
                padx=4,
                bd=0,
                highlightthickness=0,
                state=("normal" if has_apps else "disabled"),
                disabledforeground="#4b5563",
            )
            gchk.pack(fill="x", pady=(1, 1))

            for app_id in apps:
                avar = tk.BooleanVar(value=self._app_state.get(app_id, True))
                self._filter_app_vars[app_id] = avar
                tk.Checkbutton(
                    self._filters_checks_frame,
                    text=f"   {self._app_labels.get(app_id, _display_app_label(app_id))}",
                    variable=avar,
                    command=lambda a=app_id, v=avar: self._on_app_panel_toggle(a, v),
                    bg="#111827",
                    fg="#9ca3af",
                    selectcolor="#1a2030",
                    activebackground="#111827",
                    activeforeground="#e5e7eb",
                    font=("Courier New", 9),
                    anchor="w",
                    padx=6,
                    bd=0,
                    highlightthickness=0,
                ).pack(fill="x", pady=(0, 0))

            if i < len(self._group_ids) - 1:
                tk.Frame(self._filters_checks_frame, bg="#223046", height=1).pack(fill="x", pady=(4, 4))

    def _sync_filters_panel_values(self):
        for gid, gvar in self._filter_group_vars.items():
            gvar.set(self._group_state.get(gid, True))
        for app_id, avar in self._filter_app_vars.items():
            avar.set(self._app_state.get(app_id, True))

    def _toggle_filters_panel(self):
        if self._filters_panel_open:
            self._filters_panel.pack_forget()
            self._filters_panel_open = False
            return

        self._sync_filter_options()
        sig = self._current_filters_structure_sig()
        needs_rebuild = (sig != self._filters_structure_sig) or (not self._filter_group_vars and not self._filter_app_vars)
        if needs_rebuild:
            self._rebuild_filters_panel()
            self._filters_structure_sig = sig
        else:
            self._sync_filters_panel_values()
        self._filters_panel.pack(side="left", fill="y", padx=(0, 10))
        self._filters_panel_open = True

    def _current_filters_structure_sig(self):
        return (
            tuple(self._group_ids),
            tuple((gid, tuple(self._group_to_apps.get(gid, []))) for gid in self._group_ids),
            tuple(self._app_labels.get(app_id, "") for app_id in self._app_ids),
        )

    def _on_group_panel_toggle(self, group_id, var):
        if not self._group_to_apps.get(group_id):
            self._group_state[group_id] = False
            var.set(False)
            return
        new_state = bool(var.get())
        self._group_state[group_id] = new_state
        for app_id in self._group_to_apps.get(group_id, []):
            self._app_state[app_id] = new_state
            app_var = self._filter_app_vars.get(app_id)
            if app_var is not None:
                app_var.set(new_state)
        self._persist_prefs()
        self._hover_idx = None
        self._refresh()

    def _on_app_panel_toggle(self, app_id, var):
        new_state = bool(var.get())
        self._app_state[app_id] = new_state
        gid = self._app_to_group.get(app_id, OTHERS_GROUP_ID)
        apps = self._group_to_apps.get(gid, [])
        group_checked = all(self._app_state.get(a, True) for a in apps) if apps else False
        self._group_state[gid] = group_checked
        gvar = self._filter_group_vars.get(gid)
        if gvar is not None:
            gvar.set(group_checked)
        self._persist_prefs()
        self._hover_idx = None
        self._refresh()

    def _on_refresh_clicked(self):
        self._reload_config()
        self.show_alpha_var.set(True)
        self.show_numeral_var.set(True)
        self.show_symbols_var.set(True)
        self.show_special_var.set(True)
        self.show_ws_var.set(True)
        self._sync_filter_options(reset=True)
        self._persist_prefs()
        self._hover_idx = None
        self._draw_keys()
        self._refresh()

    def _on_filter_changed(self):
        self._persist_prefs()
        self._refresh()

    def _scale_mode_text(self):
        mode = self.color_cfg.get("scale_mode", "adaptive")
        if mode == "stepped":
            bins = _effective_color_bins(self.color_cfg)
            ps = self.color_cfg.get("percent_step")
            if ps is not None:
                return f"[stepped {ps:g}% -> {bins} bins]"
            return f"[stepped {bins} bins]"
        if mode == "adaptive":
            return "[adaptive scale]"
        return "[smooth scale]"

    def _filtered_snap(self, snap):
        show = {
            "alpha":   self.show_alpha_var.get(),
            "numeral": self.show_numeral_var.get(),
            "symbols": self.show_symbols_var.get(),
            "special": self.show_special_var.get(),
            "ws":      self.show_ws_var.get(),
        }
        if all(show.values()):
            return snap
        return Counter({k: v for k, v in snap.items() if show[_key_category(k)]})

    def _selected_snapshot(self):
        selected_ids = self._selected_app_ids()
        with counts_lock:
            snap_by_app = {app_id: Counter(c) for app_id, c in counts_by_app.items()}
            active = current_app_id

        snap = Counter()
        for app_id in selected_ids:
            snap.update(snap_by_app.get(app_id, Counter()))

        total = sum(snap.values()) or 1
        return snap, total, selected_ids, active

    def _draw_keys(self):
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        if cw < 20 or ch < 20: return

        pad   = max(2, int(cw * PAD_RATIO))
        key_ratio = 1.0  # Keep key aspect consistent while scaling up/down.
        unit_w = (cw - pad * (TOTAL_UNITS + 1)) / TOTAL_UNITS
        unit_h = (ch - pad * (NUM_ROWS + 1)) / (NUM_ROWS * key_ratio)
        unit = max(1.0, min(unit_w, unit_h))
        key_h = unit * key_ratio

        kb_w = pad * (TOTAL_UNITS + 1) + TOTAL_UNITS * unit
        kb_h = pad * (NUM_ROWS + 1) + NUM_ROWS * key_h
        left = max(0, (cw - kb_w) / 2)
        top = max(0, (ch - kb_h) / 2)

        fs    = max(6, min(13, int(unit*0.28)))
        fs_s  = max(5, fs-2)

        self.canvas.delete("all")
        self.key_cells.clear()

        snap, _, _, _ = self._selected_snapshot()
        snap = self._filtered_snap(snap)
        ratio_fn = _build_color_ratio_fn(snap, self.color_cfg)
        gamma = self.color_cfg.get("gamma", 0.55)
        bins = _effective_color_bins(self.color_cfg) if self.color_cfg.get("scale_mode") == "stepped" else None

        def on_enter(idx):
            self._hover_idx = idx
            self._show_pct(idx)

        def on_leave(idx):
            self._hide_pct(idx)

        for row_i, row in enumerate(self.rows):
            x = left + pad
            y = top + pad + row_i*(key_h+pad)
            for tup in row:
                n, s, kid, width = tup
                w     = width*unit + (width-1)*pad
                v     = snap.get(kid, 0)
                ratio = ratio_fn(v)
                bg    = heat_color(ratio, gamma=gamma, bins=bins)
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
        if self._geom_save_job:
            self.root.after_cancel(self._geom_save_job)
        self._geom_save_job = self.root.after(250, self._save_geometry)
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
        self._sync_filter_options()
        snap, total, selected_ids, active = self._selected_snapshot()
        snap = self._filtered_snap(snap)
        total = sum(snap.values()) or 1
        ratio_fn = _build_color_ratio_fn(snap, self.color_cfg)
        gamma = self.color_cfg.get("gamma", 0.55)
        bins = _effective_color_bins(self.color_cfg) if self.color_cfg.get("scale_mode") == "stepped" else None

        for cell in self.key_cells:
            v     = snap.get(cell["kid"],0)
            ratio = ratio_fn(v)
            bg    = heat_color(ratio, gamma=gamma, bins=bins)
            fg    = text_color(bg)
            ol    = BG if ratio>0.05 else OUTLINE
            self.canvas.itemconfig(cell["rect"],fill=bg,outline=ol)
            for txt in cell["texts"]:
                self.canvas.itemconfig(txt,fill=fg)

        if self._hover_idx is not None and 0 <= self._hover_idx < len(self.key_cells):
            self._show_pct(self._hover_idx)

        if not self._app_ids or len(selected_ids) >= len(self._app_ids):
            scope = "all apps"
        else:
            scope = f"{len(selected_ids)} apps"
        active_label = ""
        if active:
            active_name = _display_app_label(active)
            active_label = f"  |  active: {active_name}"
        self.total_lbl.config(text=f"{total:,} keystrokes  |  {scope}{active_label}")
        if self._scale_mode_lbl is not None:
            self._scale_mode_lbl.config(text=self._scale_mode_text())
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
        ratio_fn = _build_color_ratio_fn(dict(rows), self.color_cfg)
        gamma = self.color_cfg.get("gamma", 0.55)
        bins = _effective_color_bins(self.color_cfg) if self.color_cfg.get("scale_mode") == "stepped" else None
        row_h  = 14
        lbl_w  = 90
        num_w  = 120
        bar_x0 = lbl_w+6
        bar_w  = max(10, cw-lbl_w-num_w-18)
        for i,(k,v) in enumerate(rows):
            y=i*(row_h+3)+2; cy=y+row_h/2
            ratio_raw=v/max_v
            ratio_col=ratio_fn(v)
            pct=100*v/total
            bc.create_text(lbl_w-4,cy,text=repr(k) if len(k)==1 else k,
                           anchor="e",font=("Courier New",8),fill="#94a3b8")
            bc.create_rectangle(bar_x0,y,bar_x0+bar_w,y+row_h,fill="#1a2030",outline="")
            bc.create_rectangle(bar_x0,y,bar_x0+max(2,int(bar_w*ratio_raw)),y+row_h,
                                fill=heat_color(ratio_col, gamma=gamma, bins=bins),outline="")
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
        snap, _, _, _ = self._selected_snapshot()
        snap = self._filtered_snap(snap)
        total = sum(snap.values()) or 1
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
        global stats_dirty
        import tkinter.messagebox

        selected_ids = self._selected_app_ids()
        if self._app_ids and not selected_ids:
            tkinter.messagebox.showinfo("Reset", "No apps are currently selected in Filters.")
            return
        if not self._app_ids or len(selected_ids) >= len(self._app_ids):
            ok = tkinter.messagebox.askyesno("Reset", "Clear ALL keystroke data for all apps?")
            if not ok:
                return
            with counts_lock:
                counts_by_app.clear()
                stats_dirty = True
        else:
            ok = tkinter.messagebox.askyesno("Reset", f"Clear keystroke data for {len(selected_ids)} selected app(s)?")
            if not ok:
                return
            with counts_lock:
                for app_id in selected_ids:
                    counts_by_app.pop(app_id, None)
                stats_dirty = True

        flush_stats_if_dirty()
        self._sync_filter_options()
        self._draw_keys()
        self._refresh()

    def _save_geometry(self):
        self._geom_save_job = None
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
        self._reload_config()
        self._sync_filter_options()
        self._refresh()
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
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

    argv = [str(a).strip().lower() for a in sys.argv[1:]]
    if "--version" in argv or "-v" in argv:
        print(APP_VERSION)
        sys.exit(0)

    if not _handle_existing_instance_upgrade():
        # Existing instance is same/newer, or its version could not be queried safely.
        sys.exit(0)

    if not HAS_TRAY and TRAY_IMPORT_ERROR:
        print(f"Tray support disabled: {TRAY_IMPORT_ERROR}")

    win = HeatmapWindow(start_hidden=HAS_TRAY)
    try:
        if HAS_TRAY:
            run_with_tray(win)
        else:
            win.run()
    except KeyboardInterrupt:
        flush_stats_if_dirty()
        try:
            win.root.destroy()
        except Exception:
            pass







