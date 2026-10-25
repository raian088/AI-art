# -*- coding: utf-8 -*-
"""
ブラウザ（Google Chrome）キオスクモード起動と、画面ロック用のキー制限。

- Chrome を専用プロファイルで --kiosk 起動する（常に全画面・アドレスバーなし）
- Chrome が落ちたり閉じられたりしたら自動で再起動する（管理画面から終了したときを除く）
- Chrome が前面にある間だけ Alt+Tab / Alt+F4 / Ctrl+W などのショートカットを無効化する
  （Windows キーはそのまま使える。Ctrl+Alt+Del は OS の仕様上ブロックできない）
"""
import json
import os
import subprocess
import sys
import threading
import time

IS_WINDOWS = sys.platform.startswith("win")


# ==========================================
# Chrome の場所を探す
# ==========================================
def find_chrome(custom_path: str = "") -> str:
    """chrome.exe のフルパスを返す。見つからなければ空文字。"""
    cands = []
    if custom_path:
        cands.append(custom_path)
    for env in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
        base = os.environ.get(env)
        if base:
            cands.append(os.path.join(base, "Google", "Chrome", "Application", "chrome.exe"))
    # レジストリ（App Paths）
    if IS_WINDOWS:
        try:
            import winreg
            for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                try:
                    with winreg.OpenKey(root, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe") as k:
                        cands.append(winreg.QueryValue(k, None))
                except OSError:
                    pass
        except Exception:
            pass
    for c in cands:
        if c and os.path.isfile(c):
            return c
    return ""


def find_edge() -> str:
    """Chrome が無い PC 用の予備。msedge.exe のフルパスを返す。"""
    for env in ("ProgramFiles(x86)", "ProgramFiles"):
        base = os.environ.get(env)
        if base:
            p = os.path.join(base, "Microsoft", "Edge", "Application", "msedge.exe")
            if os.path.isfile(p):
                return p
    return ""


# ==========================================
# 専用プロファイルの準備
# ==========================================
def _mark_profile_clean(profile_dir: str):
    """前回強制終了していても「ページを復元しますか？」が出ないようにする。"""
    pref = os.path.join(profile_dir, "Default", "Preferences")
    if not os.path.isfile(pref):
        return
    try:
        with open(pref, "r", encoding="utf-8") as f:
            data = json.load(f)
        prof = data.setdefault("profile", {})
        prof["exit_type"] = "Normal"
        prof["exited_cleanly"] = True
        with open(pref, "w", encoding="utf-8") as f:
            json.dump(data, f)
    except Exception as e:
        print(f"[kiosk] Preferences の修正に失敗（無視して続行）: {e}")


def build_browser_args(browser_path: str, url: str, profile_dir: str) -> list:
    return [
        browser_path,
        "--kiosk", url,
        f"--user-data-dir={profile_dir}",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-session-crashed-bubble",
        "--disable-infobars",
        "--noerrdialogs",
        "--disable-translate",
        "--disable-features=Translate,TranslateUI,OverscrollHistoryNavigation",
        "--overscroll-history-navigation=0",  # 2本指スワイプで「戻る」させない
        "--disable-pinch",                    # ピンチで拡大させない
        "--use-fake-ui-for-media-stream",     # マイク許可ダイアログを出さない（常に許可）
        "--autoplay-policy=no-user-gesture-required",
        "--password-store=basic",
        "--check-for-update-interval=31536000",
    ]


# ==========================================
# キオスクブラウザ（自動再起動つき）
# ==========================================
class KioskBrowser:
    def __init__(self, url: str, profile_dir: str, chrome_path: str = "", on_give_up=None):
        self.url = url
        self.profile_dir = profile_dir
        self.chrome_path = chrome_path
        self.on_give_up = on_give_up  # 再起動を諦めたときに呼ぶ
        self.proc = None
        self._closing = False
        self._lock = threading.Lock()
        self._restarts = []

    @property
    def pid(self):
        p = self.proc
        return p.pid if p is not None and p.poll() is None else None

    def _resolve_browser(self) -> str:
        path = find_chrome(self.chrome_path)
        if path:
            return path
        edge = find_edge()
        if edge:
            print("[kiosk] Chrome が見つからないため Edge で起動します")
        return edge

    def launch(self) -> bool:
        browser = self._resolve_browser()
        if not browser:
            print("[kiosk] Chrome / Edge が見つかりません")
            return False
        os.makedirs(self.profile_dir, exist_ok=True)
        _mark_profile_clean(self.profile_dir)
        args = build_browser_args(browser, self.url, self.profile_dir)
        print(f"[kiosk] 起動: {browser}")
        with self._lock:
            self.proc = subprocess.Popen(args, close_fds=True)
        return True

    def start(self) -> bool:
        if not self.launch():
            return False
        threading.Thread(target=self._watchdog, daemon=True).start()
        return True

    def _watchdog(self):
        """ブラウザが閉じられたら再起動する。1分間に5回以上落ちたら諦める。"""
        while not self._closing:
            p = self.proc
            if p is None:
                return
            p.wait()
            if self._closing:
                return
            now = time.time()
            self._restarts = [t for t in self._restarts if now - t < 60] + [now]
            if len(self._restarts) > 5:
                print("[kiosk] ブラウザが繰り返し終了したため再起動を停止します")
                if self.on_give_up:
                    self.on_give_up()
                return
            print("[kiosk] ブラウザが終了したので再起動します")
            time.sleep(1.5)
            if not self._closing:
                self.launch()

    def close(self):
        """管理画面から終了するとき。ブラウザを閉じて再起動もしない。"""
        self._closing = True
        p = self.proc
        if p is None or p.poll() is not None:
            return
        try:
            if IS_WINDOWS:
                # まず普通に閉じる → 3秒待ってダメなら強制終了
                subprocess.run(["taskkill", "/PID", str(p.pid), "/T"],
                               capture_output=True, creationflags=0x08000000)
                try:
                    p.wait(timeout=3)
                except subprocess.TimeoutExpired:
                    subprocess.run(["taskkill", "/PID", str(p.pid), "/T", "/F"],
                                   capture_output=True, creationflags=0x08000000)
            else:
                p.terminate()
        except Exception as e:
            print(f"[kiosk] ブラウザ終了に失敗: {e}")


# ==========================================
# キー制限（Windows 低レベルキーボードフック）
# ==========================================
# 仮想キーコード
VK_TAB, VK_ESCAPE, VK_SPACE = 0x09, 0x1B, 0x20
VK_LEFT, VK_RIGHT, VK_HOME = 0x25, 0x27, 0x24
VK_PRIOR, VK_NEXT, VK_DELETE = 0x21, 0x22, 0x2E
VK_CONTROL, VK_SHIFT = 0x11, 0x10
VK_APPS = 0x5D
VK_F1, VK_F3, VK_F4, VK_F5, VK_F6, VK_F7, VK_F10, VK_F11, VK_F12 = 0x70, 0x72, 0x73, 0x74, 0x75, 0x76, 0x79, 0x7A, 0x7B
VK_OEM_PLUS, VK_OEM_MINUS, VK_ADD, VK_SUBTRACT = 0xBB, 0xBD, 0x6B, 0x6D
BROWSER_KEYS = set(range(0xA6, 0xAD))  # 戻る/進む/更新/停止/検索/お気に入り/ホーム

# Alt と組み合わせたら止めるキー（Alt+` などの IME 切替は止めない）
ALT_BLOCK = {VK_TAB, VK_ESCAPE, VK_F4, VK_SPACE, VK_LEFT, VK_RIGHT, VK_HOME, ord('D'), ord('F'), ord('E')}
# Ctrl と組み合わせたら止めるキー（Ctrl+C/V/X/A/Z は入力欄で使うので残す）
CTRL_BLOCK = {ord(c) for c in "WNTPROSULHJFDEKGQIMB"} | {
    VK_TAB, VK_PRIOR, VK_NEXT, VK_F4, VK_F5, VK_OEM_PLUS, VK_OEM_MINUS, VK_ADD, VK_SUBTRACT, ord('0'),
}
# 単体で止めるキー（F6〜F10 は IME の変換で使うので F7 以外残す。F7 は Chrome のカーソルブラウズ）
SOLO_BLOCK = {VK_F1, VK_F3, VK_F5, VK_F7, VK_F11, VK_F12, VK_APPS} | BROWSER_KEYS


def should_block(vk: int, alt: bool, ctrl: bool, shift: bool) -> bool:
    """このキー入力を止めるか（テストしやすいよう純粋関数にしてある）。"""
    if ctrl and shift and vk == VK_ESCAPE:   # タスクマネージャー
        return True
    if alt and vk in ALT_BLOCK:
        return True
    if ctrl and vk in CTRL_BLOCK:
        return True
    if ctrl and shift and vk in (VK_DELETE,):
        return True
    if vk in SOLO_BLOCK:
        return True
    return False


class KeyboardBlocker:
    """指定プロセス（キオスクブラウザ）が前面のときだけショートカットを止める。"""

    def __init__(self, pid_getter):
        self.pid_getter = pid_getter
        self.enabled = True
        self._thread = None
        self._thread_id = None
        self._hook = None
        self._proc_ref = None

    def start(self):
        if not IS_WINDOWS:
            return
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self):
        self.enabled = False
        if IS_WINDOWS and self._thread_id:
            import ctypes
            ctypes.windll.user32.PostThreadMessageW(self._thread_id, 0x0012, 0, 0)  # WM_QUIT

    def _foreground_is_kiosk(self, user32) -> bool:
        import ctypes
        from ctypes import wintypes
        pid = self.pid_getter()
        if not pid:
            return False
        hwnd = user32.GetForegroundWindow()
        if not hwnd:
            return False
        fg_pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(fg_pid))
        return fg_pid.value == pid

    def _run(self):
        import ctypes
        from ctypes import wintypes

        user32 = ctypes.WinDLL("user32", use_last_error=True)
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

        WH_KEYBOARD_LL = 13
        WM_KEYDOWN, WM_SYSKEYDOWN = 0x0100, 0x0104
        LLKHF_ALTDOWN = 0x20
        LRESULT = ctypes.c_ssize_t

        class KBDLLHOOKSTRUCT(ctypes.Structure):
            _fields_ = [("vkCode", wintypes.DWORD), ("scanCode", wintypes.DWORD),
                        ("flags", wintypes.DWORD), ("time", wintypes.DWORD),
                        ("dwExtraInfo", ctypes.c_size_t)]

        HOOKPROC = ctypes.WINFUNCTYPE(LRESULT, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)
        user32.SetWindowsHookExW.argtypes = (ctypes.c_int, HOOKPROC, wintypes.HINSTANCE, wintypes.DWORD)
        user32.SetWindowsHookExW.restype = wintypes.HHOOK
        user32.CallNextHookEx.argtypes = (wintypes.HHOOK, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)
        user32.CallNextHookEx.restype = LRESULT
        user32.GetAsyncKeyState.restype = ctypes.c_short
        kernel32.GetModuleHandleW.restype = wintypes.HMODULE

        def proc(n_code, w_param, l_param):
            try:
                if n_code == 0 and self.enabled and w_param in (WM_KEYDOWN, WM_SYSKEYDOWN):
                    kb = ctypes.cast(l_param, ctypes.POINTER(KBDLLHOOKSTRUCT)).contents
                    alt = bool(kb.flags & LLKHF_ALTDOWN)
                    ctrl = bool(user32.GetAsyncKeyState(VK_CONTROL) & 0x8000)
                    shift = bool(user32.GetAsyncKeyState(VK_SHIFT) & 0x8000)
                    if should_block(kb.vkCode, alt, ctrl, shift) and self._foreground_is_kiosk(user32):
                        return 1
            except Exception:
                pass
            return user32.CallNextHookEx(None, n_code, w_param, l_param)

        self._proc_ref = HOOKPROC(proc)  # GC されないよう保持
        self._hook = user32.SetWindowsHookExW(WH_KEYBOARD_LL, self._proc_ref,
                                              kernel32.GetModuleHandleW(None), 0)
        if not self._hook:
            print(f"[kiosk] キーボードフックの設定に失敗: {ctypes.get_last_error()}")
            return
        self._thread_id = kernel32.GetCurrentThreadId()
        print("[kiosk] キー制限 ON（Windowsキー以外のショートカットを無効化）")
        msg = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
        user32.UnhookWindowsHookEx(self._hook)
