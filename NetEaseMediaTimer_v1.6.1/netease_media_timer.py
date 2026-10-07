# -*- coding: utf-8 -*-
"""
================================================================================
 网易云音乐定时播放助手（GSMTC 系统媒体会话版 + 无 SMTC 兼容模式） · 仅供技术学习演示  v1.6.1
================================================================================
【免责声明】
  1. 本程序仅用于 Windows GSMTC（Global System Media Transport Controls）
     全局媒体会话 API 的技术学习与演示；
  2. 自动化挂机存在平台风控风险，请勿用于任何违反平台服务条款的用途，
     因使用本程序产生的一切后果由使用者自行承担。

【使用前提（必须由用户手动完成，本程序绝不代劳）】
  1. 手动打开网易云音乐 PC 客户端；
  2. 手动选好歌单，开启列表循环；
  3. 手动点击播放，让网易云处于播放状态。
  —— 本程序不负责启动网易云 exe、不负责打开歌单、不负责初始播放。

【核心工作逻辑】
  1. 到达设定时间（或点击"立即开始计时"）后，程序开始倒计时；
  2. 期间仅通过 Windows 系统媒体会话 API（GSMTC）以 8~12 秒随机间隔低频
     轮询网易云播放状态（界面上的 1 秒倒计时为本地计时，不产生任何媒体查询）；
  3. 仅当检测到"已暂停"时，才发送一次系统媒体"播放"指令恢复；
     播放正常时零操作、零打扰，完全交给网易云原生连续循环播放；
  4. 倒计时走完用户设定时长，发送一次系统媒体"暂停"指令，任务结束；
  5. 若系统不存在网易云 SMTC 会话（网易云 Win32 原版客户端的官方限制），
     自动进入兼容模式：状态检测 = WASAPI 音频会话电平；
     静音挂机 = 会话静音开关（音乐继续走时长、不出声，可勾选开声音）；
     到点真暂停 = 冻结网易云进程（NtSuspendProcess，与系统睡眠同一机制，
     不模拟任何输入）——音乐立即停止、歌曲不再推进，次日定时或手动停止
     时自动解冻并从中断处继续。仅作用于网易云，不影响系统其它程序。

【硬性控制约束】
  - 禁止 pyautogui，不模拟鼠标坐标点击，不模拟键盘按键；
  - 不启动 / 不关闭 / 不查杀任何进程；
  - 所有播放/暂停控制均通过系统级 GSMTC 媒体会话接口完成，
    不干扰用户鼠标键盘的正常使用；
  - 兼容模式的静音控制为 Windows 音频混音器会话级操作，同样不模拟鼠标/键盘。

【运行环境】
  - Windows 10 (1903+) / Windows 11；
  - Python 3.11 ~ 3.13（建议 3.11 / 3.12）；
  - 依赖：pip install py-now-playing schedule psutil pywin32 pycaw comtypes
    （tkinter 随 Python 官方安装包自带；winsdk 由 py-now-playing 自动带入）
================================================================================
"""

# ============================== 标准库导入 ==============================
import asyncio              # 异步事件循环：驱动 py-now-playing / winsdk 异步接口
import ctypes               # 系统调用：防休眠 / 兼容模式进程映像判断
import json                 # 配置文件读写（记忆上次参数）
import os                   # 配置/图标/崩溃日志的路径处理
import queue                # 线程安全队列：asyncio 工作线程 → tkinter 主线程
import random               # 8~12 秒随机轮询间隔（避免固定周期特征）
import subprocess           # 仅用于读取 Get-StartApps 应用列表（只读、无窗口）
import sys                  # 平台检查 / 打包环境判断
import threading            # 后台 asyncio 事件循环线程
import time                 # 倒计时与日志时间戳
import tkinter as tk        # GUI 框架
try:                        # pywin32：兼容模式读取网易云主窗口标题（当前曲目）
    import win32gui
    import win32process
except Exception:           # 缺失时仅曲目标题显示不可用，核心功能不受影响
    win32gui = None
    win32process = None
from tkinter import messagebox, scrolledtext, ttk

# ========================= 第三方依赖（可选导入 + 优雅降级） =========================
# py-now-playing：GSMTC 的友好封装（0.1.x 类名 NowPlaying / 0.2.x 类名 PyNowPlaying），
#                 它会自动安装 winsdk（旧投影）或 winrt-*（新投影）作为底层依赖。
try:
    from py_now_playing import NowPlaying as PnpLegacy          # py-now-playing 0.1.x
except ImportError:
    PnpLegacy = None
try:
    from py_now_playing import PyNowPlaying as PnpModern        # py-now-playing 0.2.x(--pre)
except ImportError:
    PnpModern = None

# winsdk / winrt：Windows Runtime 投影库，GSMTC 媒体会话 API 的真正入口。
# 控制指令（播放/暂停）与状态读取最终都经由这两个命名空间完成。
try:
    from winsdk.windows.media.control import (
        GlobalSystemMediaTransportControlsSessionManager as MediaManager)
    GSMTC_BACKEND = "winsdk"
except ImportError:
    try:
        from winrt.windows.media.control import (
            GlobalSystemMediaTransportControlsSessionManager as MediaManager)
        GSMTC_BACKEND = "winrt"
    except ImportError:
        MediaManager = None
        GSMTC_BACKEND = None

try:
    import schedule     # 轻量级每日定时调度库
except ImportError:
    schedule = None

try:
    import psutil       # 只读检测网易云进程是否存在（不启动、不结束任何进程）
except ImportError:
    psutil = None

# ============================== 常量与外观配置 ==============================
APP_TITLE = "网易云定时播放助手 v1.6.1 · 静音挂机 + 真暂停"
APP_MUTEX = "DSH_NeteaseGsmtcMediaTimer_SingleInstanceMutex"
APP_VERSION = "v1.6.1"

POLL_MIN_SEC = 8.0        # 媒体状态轮询最小间隔（秒）—— 风控稳定重点：低频随机
POLL_MAX_SEC = 12.0       # 媒体状态轮询最大间隔（秒）

# GSMTC 播放状态枚举值（GlobalSystemMediaTransportControlsSessionPlaybackStatus）
GSMT_STATUS_ZH = {
    0: "已关闭(Closed)", 1: "已打开(Opened)", 2: "切换中(Changing)",
    3: "已停止(Stopped)", 4: "正在播放(Playing)", 5: "已暂停(Paused)",
}
GSMT_PLAYING = 4
GSMT_PAUSED = 5

# 匹配网易云媒体会话的关键字（同时用于 AUMID 与应用显示名称的小写匹配）
NETEASE_KEYWORDS = ("netease", "cloudmusic", "网易云", "云音乐")

MAX_MINUTES = 24 * 60     # 单次倒计时上限：24 小时
LOG_MAX_LINES = 2000      # 日志框自动瘦身阈值

# 网易云 Win32 原版客户端不注册 SMTC 的精确诊断文案（实测 + 社区文档确认）
SMTC_HINT = (
    "诊断结论：网易云 Win32 原版客户端（含 3.1.x）本身不向 Windows 注册 SMTC 系统媒体会话，"
    "这是客户端官方限制，不是本程序故障（Windows 自带媒体浮层同样看不到它）。\n"
    "解决方法（二选一）：\n"
    "  ① 给网易云安装 BetterNCM 扩展框架 + InfLink-rs 插件，让它注册标准媒体会话（步骤见 README「检测不到网易云」）；\n"
    "  ② 改用支持 SMTC 的播放器（QQ音乐 / Spotify / Edge、Chrome 网页播放等），本工具同样能识别与控制。"
)

# ---- 界面配色（浅色卡片风 + 网易云红点缀） ----
C_BG      = "#F3F4F8"     # 窗口背景
C_CARD    = "#FFFFFF"     # 卡片背景
C_ACCENT  = "#D33A31"     # 主题红（网易云风格）
C_ACCENT2 = "#B32E26"     # 主题红（按下/悬停加深）
C_TEXT    = "#26282E"     # 主文字
C_SUB     = "#7A7F8A"     # 次要文字
C_BORDER  = "#E4E7EF"     # 卡片描边
C_LOG_BG  = "#FAFBFD"     # 日志背景

# 状态指示灯配色（背景, 文字）
PILL_IDLE = ("#ECEEF3", "#5A5F6B")   # 灰：待机
PILL_OK   = ("#E3F5EC", "#157A47")   # 绿：正在播放 / 完成
PILL_WARN = ("#FDF0DC", "#B26A00")   # 橙：已暂停恢复中
PILL_BLUE = ("#E8F0FE", "#1A56B0")   # 蓝：每日定时等待中
PILL_ERR  = ("#FDECEA", "#B3261E")   # 红：异常 / 未找到


def app_dir() -> str:
    """程序所在目录（兼容 PyInstaller 打包后的单文件运行）。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


# ============================== 配置记忆（开箱即用体验优化） ==============================
CONFIG_FILE = os.path.join(app_dir(), "config.json")


def load_config() -> dict:
    """读取上次保存的参数（时长/定时时间），首次运行时给出合理默认值。"""
    data = {}
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f) or {}
    except Exception:
        data = {}
    return {
        "minutes": str(data.get("minutes", "60")),
        "hour": str(data.get("hour", "21")),
        "minute": str(data.get("minute", "30")),
        "play_sound": bool(data.get("play_sound", False)),
        "true_pause": bool(data.get("true_pause", True)),
    }


def save_config(minutes: str, hour: str, minute: str, play_sound: bool = False,
                true_pause: bool = True) -> None:
    """保存参数到程序目录（打包单文件 exe 后也写在其同目录）。"""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"minutes": minutes, "hour": hour, "minute": minute,
                       "play_sound": bool(play_sound),
                       "true_pause": bool(true_pause)},
                      f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ============================== pywin32 / ctypes 小工具 ==============================
_MUTEX_HOLD = None    # 全局持有互斥量引用，防止被垃圾回收后单实例保护失效


def acquire_single_instance() -> bool:
    """用 pywin32 互斥量保证程序单实例运行（防止两个实例同时控制媒体会话）。
    pywin32 不可用时跳过检查，不影响主功能。"""
    global _MUTEX_HOLD
    try:
        import win32api
        import win32event
        import winerror
        _MUTEX_HOLD = win32event.CreateMutex(None, False, APP_MUTEX)
        return win32api.GetLastError() != winerror.ERROR_ALREADY_EXISTS
    except Exception:
        return True


def prevent_sleep(enable: bool) -> None:
    """倒计时期间阻止系统休眠，结束后恢复系统默认电源策略。
    说明：pywin32 的 win32api 未封装 SetThreadExecutionState，
    故此处使用 ctypes 调用 kernel32（效果一致，同样不模拟任何键鼠输入）。"""
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    flags = ES_CONTINUOUS | (ES_SYSTEM_REQUIRED if enable else 0)
    try:
        import ctypes
        ctypes.windll.kernel32.SetThreadExecutionState(flags)
    except Exception:
        pass


def find_netease_processes():
    """只读枚举网易云音乐进程（cloudmusic.exe），绝不启动/结束任何进程。"""
    result = []
    if psutil is None:
        return result
    try:
        for proc in psutil.process_iter(attrs=("pid", "name")):
            try:
                name = (proc.info.get("name") or "").lower()
                if "cloudmusic" in name:
                    result.append((proc.info.get("pid"), name))
            except Exception:
                continue
    except Exception:
        pass
    return result


# ============================== 无 SMTC 兼容桥（方案 B） ==============================
# 实测结论（网易云 3.1.x / CEF）：
#   - 不注册 SMTC 媒体会话（GSMTC 看不到）；
#   - 不响应 WM_APPCOMMAND 窗口媒体指令（6 个候选窗口全部实测无响应）；
#   - 系统媒体键（SendInput）无效：媒体键只路由给前台焦点窗口，且网易云不处理；
#   - CEF 辅助树默认关闭，但用 WM_GETOBJECT 探测 3 次后 UIA 树完全可用：
#     可定位底部播放栏的播放/暂停按钮并 Invoke 调用（纯无障碍 API，
#     不模拟鼠标/键盘、不移动焦点）——这是“真暂停/恢复”的实测通道。
# 因此兼容模式采用：状态检测 = WASAPI 音频会话电平；
#   播放控制 = UIA 播放栏按钮（真暂停/恢复）+ 音频会话静音开关（无声挂机）。
PLAY_THRESHOLD = 0.015          # 峰值电平高于此值视为“正在播放”


_k32 = ctypes.WinDLL("kernel32", use_last_error=True)
_k32.OpenProcess.restype = ctypes.c_void_p
_k32.OpenProcess.argtypes = (ctypes.c_ulong, ctypes.c_long, ctypes.c_ulong)
_k32.QueryFullProcessImageNameW.restype = ctypes.c_long
_k32.QueryFullProcessImageNameW.argtypes = (
    ctypes.c_void_p, ctypes.c_ulong, ctypes.c_wchar_p,
    ctypes.POINTER(ctypes.c_ulong))
_k32.CloseHandle.restype = ctypes.c_long
_k32.CloseHandle.argtypes = (ctypes.c_void_p,)


def _pid_is_cloudmusic(pid):
    """用进程映像路径判断 pid 是否属于网易云（规范 ctypes 原型，句柄不截断；
    另有 psutil 兜底，双保险）。"""
    if not pid:
        return False
    try:
        h = _k32.OpenProcess(0x1000, False, int(pid) & 0xFFFFFFFF)
        if not h:
            return False
        try:
            buf = ctypes.create_unicode_buffer(512)
            size = ctypes.c_ulong(512)
            if _k32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
                p = buf.value.replace("\\", "/").lower()
                return "cloudmusic" in p or "netease" in p
        finally:
            _k32.CloseHandle(h)
    except Exception:
        pass
    try:
        import psutil
        n = (psutil.Process(int(pid)).name() or "").lower()
        return "cloudmusic" in n or "netease" in n
    except Exception:
        return False


class CompatBridge:
    """无 SMTC 兼容桥。
    状态：音频会话峰值电平（有声=播放，静默=暂停/停止）；
    控制：网易云音频会话 SetMute（到点停声 / 开始放声），不影响系统其它声音。"""

    def __init__(self, log):
        self._log = log
        self._com_tid = None
        self._main_hwnd = None       # 网易云主窗口缓存（读取标题 = 当前曲目）
        self._hwnd_checked = 0.0     # 上次窗口校验时间

    def _ensure_com(self):
        """pycaw/comtypes 需要在实际调用线程初始化 COM。"""
        tid = threading.get_ident()
        if tid != self._com_tid:
            try:
                import comtypes
                comtypes.CoInitialize()
                self._com_tid = tid
            except Exception as e:
                self._log(f"兼容模式 COM 初始化失败：{e!r}", "err")

    def _cloud_sessions(self):
        self._ensure_com()
        try:
            from pycaw.pycaw import AudioUtilities
            out = []
            for s in AudioUtilities.GetAllSessions():
                try:
                    name = (s.Process and s.Process.name() or "")
                except Exception:
                    name = ""
                if "cloudmusic" in str(name).lower() or \
                        _pid_is_cloudmusic(getattr(s, "ProcessId", 0)):
                    out.append(s)
            return out
        except Exception as e:
            self._log(f"兼容模式：枚举音频会话失败：{e!r}", "err")
            return []

    def set_mute(self, mute):
        """对网易云全部音频会话设置静音；返回成功设置的会话数。"""
        sess = self._cloud_sessions()
        n = 0
        for s in sess:
            try:
                s.SimpleAudioVolume.SetMute(bool(mute), None)
                n += 1
            except Exception as e:
                self._log(f"兼容模式：设置会话静音失败：{e!r}", "err")
        return n

    # ---- 主窗口定位与曲目标题（网易云把“歌名 - 歌手”写进主窗口标题，只读零打扰） ----
    def main_hwnd(self, refresh=False):
        if win32gui is None or win32process is None:
            return None
        now = time.time()
        if not refresh and self._main_hwnd \
                and win32gui.IsWindow(self._main_hwnd) \
                and now - self._hwnd_checked < 60:
            return self._main_hwnd
        found = None

        def cb(hwnd, _):
            nonlocal found
            if found:
                return True
            try:
                _, wpid = win32process.GetWindowThreadProcessId(hwnd)
                if not wpid or not _pid_is_cloudmusic(wpid):
                    return True
                if "orpheusbrowserhost" in win32gui.GetClassName(hwnd).lower() \
                        and win32gui.IsWindowVisible(hwnd):
                    found = hwnd
            except Exception:
                pass
            return True

        try:
            win32gui.EnumWindows(cb, None)
        except Exception:
            found = None
        self._main_hwnd = found
        self._hwnd_checked = now
        return found

    def track_title(self):
        """从主窗口标题读取当前曲目；不可用时返回空串。"""
        hwnd = self.main_hwnd()
        if not hwnd:
            return ""
        try:
            t = (win32gui.GetWindowText(hwnd) or "").strip()
        except Exception:
            return ""
        if not t or t == "MiniPlayer":
            return ""
        return t

    # ---- 进程级真暂停（NtSuspendProcess / NtResumeProcess，等同系统睡眠挂起） ----
    def suspend_resume_all(self, suspend: bool) -> int:
        """冻结/解冻全部网易云进程：真暂停 = 音频渲染立即停止、歌曲不再推进；
        解冻 = 从中断处继续播放。纯进程控制（与系统睡眠同一机制），
        不模拟鼠标/键盘、不操作窗口。返回成功操作的进程数。"""
        n = 0
        verb = "冻结" if suspend else "解冻"
        try:
            ntdll = ctypes.WinDLL("ntdll")
            ntdll.NtSuspendProcess.argtypes = [ctypes.c_void_p]
            ntdll.NtSuspendProcess.restype = ctypes.c_long
            ntdll.NtResumeProcess.argtypes = [ctypes.c_void_p]
            ntdll.NtResumeProcess.restype = ctypes.c_long
            fn = ntdll.NtSuspendProcess if suspend else ntdll.NtResumeProcess
            for pid, _name in (find_netease_processes() or []):
                h = _k32.OpenProcess(0x0800, False, pid)   # PROCESS_SUSPEND_RESUME
                if not h:
                    self._log(f"进程{verb}：OpenProcess(pid={pid}) 失败 "
                              f"err={ctypes.get_last_error()}", "warn")
                    continue
                try:
                    st = fn(h)
                    if st == 0:
                        n += 1
                    else:
                        self._log(f"进程{verb}：pid={pid} NTSTATUS=0x{st & 0xFFFFFFFF:08X}",
                                  "warn")
                finally:
                    _k32.CloseHandle(h)
        except Exception as e:
            self._log(f"网易云进程{verb}失败：{e!r}", "err")
        return n

    # ---- UIA 播放栏按钮（真暂停/恢复：纯无障碍 API，不模拟鼠标/键盘） ----
    def _uia_warm(self):
        """探测主窗口 WM_GETOBJECT 三次，促使 Chromium 开启辅助功能树（约 6 秒）。"""
        hwnd = self.main_hwnd()
        if not hwnd:
            return
        for _ in range(3):
            try:
                win32gui.SendMessage(hwnd, 0x003D, 0, 0xFFFFFFFC)   # WM_GETOBJECT/OBJID_CLIENT
            except Exception:
                pass
            time.sleep(2)

    def _uia_press_once(self):
        """定位底部播放栏的播放/暂停按钮并 Invoke；成功返回 True。"""
        import comtypes.client
        mod = comtypes.client.GetModule("UIAutomationCore.dll")
        uia = comtypes.client.CreateObject(
            "{ff48dba4-60ef-4201-aa87-54103eef594e}", interface=mod.IUIAutomation)
        hwnd = self.main_hwnd()
        if not hwnd:
            return False
        root = uia.ElementFromHandle(hwnd)
        arr = root.FindAll(4, uia.CreateTrueCondition())   # TreeScope_Descendants
        wr = root.CurrentBoundingRectangle
        best, best_d = None, None
        for i in range(arr.Length):
            el = arr.GetElement(i)
            try:
                if el.CurrentControlType != 50000:          # UIA_ButtonControlTypeId
                    continue
                name = (el.CurrentName or "").strip().lower()
                if name in ("play", "pause", "播放", "暂停"):
                    r = el.CurrentBoundingRectangle
                    # 播放栏固定在窗口底部：取最接近底部 93% 高度处的按钮
                    d = abs((r.top - wr.top) - (wr.bottom - wr.top) * 0.93)
                    if best_d is None or d < best_d:
                        best_d, best = d, el
            except Exception:
                continue
        if best is None:
            return False
        pat = best.GetCurrentPattern(10000)                 # UIA_InvokePatternId
        if pat is None:
            return False
        pat.QueryInterface(mod.IUIAutomationInvokePattern).Invoke()
        return True

    def press_play_pause(self):
        """点击网易云播放栏的播放/暂停按钮；首次使用自动唤醒辅助树（约 6 秒）。"""
        try:
            self._ensure_com()
            if self._uia_press_once():
                return True
            self._uia_warm()
            return self._uia_press_once()
        except Exception as e:
            self._log(f"兼容模式：播放栏按钮调用失败：{e!r}", "err")
            return False

    def snapshot(self, samples=8, interval=0.12):
        """一次枚举同时返回会话数/峰值电平/静音/会话音量/曲目标题。
        （减少 COM 枚举次数：状态检测 + 静音守卫 + 曲目显示共用一次采样）"""
        self._ensure_com()
        sess = self._cloud_sessions()
        if not sess:
            return {"n": 0, "peak": None, "mute": None,
                    "volume": None, "title": self.track_title()}
        from pycaw.pycaw import IAudioMeterInformation
        mx = 0.0
        for _ in range(samples):
            for s in sess:
                try:
                    v = s._ctl.QueryInterface(IAudioMeterInformation).GetPeakValue()
                    if v > mx:
                        mx = v
                except Exception:
                    pass
            time.sleep(interval)
        mute_vals, vol_vals = [], []
        for s in sess:
            try:
                mute_vals.append(bool(s.SimpleAudioVolume.GetMute()))
            except Exception:
                pass
            try:
                vol_vals.append(float(s.SimpleAudioVolume.GetMasterVolume()))
            except Exception:
                pass
        return {"n": len(sess), "peak": mx,
                "mute": all(mute_vals) if mute_vals else None,
                "volume": min(vol_vals) if vol_vals else None,
                "title": self.track_title()}


# ============================== 运行时开关：防休眠 / 开机自启 ==============================
class RuntimeFlags:
    """GUI 与后台线程共享的运行时开关（简单属性，跨线程读取安全）。"""
    def __init__(self):
        self.prevent_sleep = True   # 挂机期间阻止系统休眠（默认开启）
        self.play_sound = False     # 兼容模式：倒计时期间是否开声音（默认静音挂机刷时长）
        self.true_pause = True      # 到点后真暂停（冻结网易云进程，次日自动恢复）


AUTOSTART_NAME = "NetEaseMediaTimer"
RUN_KEY_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"


def get_autostart() -> bool:
    """查询是否已写入开机自启注册表项（HKCU Run，只读）。"""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY_PATH, 0,
                            winreg.KEY_QUERY_VALUE) as k:
            winreg.QueryValueEx(k, AUTOSTART_NAME)
            return True
    except Exception:
        return False


def _autostart_command() -> str:
    """生成本程序的开机启动命令（兼容打包 exe 与脚本两种形态）。"""
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"'
    return f'"{sys.executable}" "{os.path.abspath(__file__)}"'


def set_autostart(enable: bool) -> None:
    """写入/删除 HKCU Run 开机自启项（仅当前用户，无需管理员权限）。"""
    import winreg
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY_PATH, 0,
                        winreg.KEY_SET_VALUE) as k:
        if enable:
            winreg.SetValueEx(k, AUTOSTART_NAME, 0, winreg.REG_SZ, _autostart_command())
        else:
            try:
                winreg.DeleteValue(k, AUTOSTART_NAME)
            except FileNotFoundError:
                pass


# ============================== GSMTC 媒体桥 ==============================
class GsmtcBridge:
    """系统媒体会话桥：状态读取 + 播放/暂停控制。
    - 状态/曲目/控制优先走 py-now-playing（0.2.x PyNowPlaying 异步接口）；
    - py-now-playing 0.1.x 或缺失时，直接通过 winsdk/winrt 会话完成（等效）；
    - AUMID 自动发现：先按 AUMID 关键词匹配，再按应用显示名称兜底
      （py-now-playing 0.1.x 内置的名称发现在中文系统会因 GBK 编码报错，
       因此这里自行实现并做编码兼容）。"""

    def __init__(self, log_fn):
        self._log = log_fn
        self._manager = None            # GSMTC 会话管理器（缓存）
        self._pnp_legacy = None         # py-now-playing 0.1.x NowPlaying 实例
        self._pnp_modern = None         # py-now-playing 0.2.x PyNowPlaying 实例
        self._primary_aumid = None      # 已定位的网易云 AUMID
        self._start_apps_cache = None   # Get-StartApps {AppID: Name} 缓存

    # ---------- 初始化（幂等，可重复调用） ----------
    async def ensure_manager(self):
        if self._manager is None:
            if MediaManager is None:
                raise RuntimeError(
                    "未找到 winsdk / winrt 库，无法访问 GSMTC 媒体会话 API。"
                    "请先执行：pip install py-now-playing")
            self._manager = await MediaManager.request_async()
            self._log(f"已连接 Windows 系统媒体会话服务 GSMTC（底层：{GSMTC_BACKEND}）。", "ok")
        return self._manager

    async def ensure_py_now_playing(self):
        """初始化 py-now-playing 实例（两种版本 API 均兼容）。"""
        if PnpModern is not None and self._pnp_modern is None and self._primary_aumid:
            self._pnp_modern = await PnpModern.create(aumid=self._primary_aumid)
        if PnpLegacy is not None and self._pnp_legacy is None:
            self._pnp_legacy = PnpLegacy()
            await self._pnp_legacy.initalize_mediamanger()
        return self._pnp_modern or self._pnp_legacy

    # ---------- 网易云会话匹配 ----------
    def _netease_sessions(self, manager):
        """从当前所有 GSMTC 会话中筛出网易云的会话。"""
        sessions = []
        try:
            for s in manager.get_sessions():
                aumid = ""
                try:
                    aumid = s.source_app_user_model_id or ""
                except Exception:
                    pass
                if any(k in aumid.lower() for k in NETEASE_KEYWORDS):
                    sessions.append(s)
        except Exception:
            pass
        return sessions

    def _start_apps_map(self):
        """读取 Get-StartApps 应用列表 → {AppID: 显示名称}（带缓存）。
        用于 AUMID 关键词匹配失败时的"应用名称"兜底匹配；
        自行实现并兼容中文系统 GBK 输出编码（只读操作，不启动任何应用）。"""
        if self._start_apps_cache is not None:
            return self._start_apps_cache
        mapping = {}
        try:
            raw = subprocess.check_output(
                ["powershell.exe", "-NoProfile", "Get-StartApps | ConvertTo-Json"],
                creationflags=subprocess.CREATE_NO_WINDOW, timeout=20)
            text = None
            for enc in ("utf-8-sig", "gbk"):
                try:
                    text = raw.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            if text is None:
                text = raw.decode("utf-8", errors="replace")
            data = json.loads(text)
            if isinstance(data, dict):
                data = [data]
            for app in data:
                if isinstance(app, dict) and app.get("AppID"):
                    mapping[app["AppID"]] = str(app.get("Name", ""))
        except Exception:
            pass
        self._start_apps_cache = mapping
        return mapping

    async def discover_netease(self):
        """定位网易云媒体会话。返回 (aumid列表, 匹配方式描述)。"""
        manager = await self.ensure_manager()
        active = []
        try:
            active = [s.source_app_user_model_id or "" for s in manager.get_sessions()]
        except Exception:
            pass
        direct = [a for a in active if any(k in a.lower() for k in NETEASE_KEYWORDS)]
        if direct:
            return direct, "AUMID关键词"
        smap = self._start_apps_map()
        named = [a for a in active
                 if a in smap and any(k in smap[a].lower() for k in NETEASE_KEYWORDS)]
        if named:
            return named, "应用名称"
        return [], "未找到"

    def set_primary_aumid(self, aumid: str) -> None:
        self._primary_aumid = aumid
        self._pnp_modern = None   # 触发下次按新 AUMID 重建 0.2.x 实例

    # ---------- 状态读取 ----------
    async def _track_name(self, session) -> str:
        """读取会话当前曲目（标题 - 歌手），失败时返回占位文本。"""
        try:
            props = await session.try_get_media_properties_async()
            title = (getattr(props, "title", "") or "").strip()
            artist = (getattr(props, "artist", "") or "").strip()
            if title or artist:
                return f"{title} - {artist}" if artist else title
        except Exception:
            pass
        return "未知曲目"

    async def read_status(self):
        """读取网易云播放状态（带 GSMTC 管理器自愈：异常时重建后重试一次）。
        返回 dict：found=匹配到的会话数, status=GSMTC状态码, tracks=曲目列表, aumids=AUMID列表"""
        for attempt in (1, 2):
            try:
                manager = await self.ensure_manager()
                sessions = self._netease_sessions(manager)
                if not sessions:
                    try:
                        aumids = [s.source_app_user_model_id or ""
                                  for s in manager.get_sessions()]
                    except Exception:
                        aumids = []
                    return {"found": 0, "status": None, "tracks": [], "aumids": aumids}
                status = None
                try:
                    raw = sessions[0].get_playback_info().playback_status
                    status = int(getattr(raw, "value", raw))
                except Exception:
                    status = None
                tracks = []
                for s in sessions[:3]:
                    tracks.append(await self._track_name(s))
                return {"found": len(sessions), "status": status, "tracks": tracks,
                        "aumids": [getattr(s, "source_app_user_model_id", "") or ""
                                   for s in sessions]}
            except asyncio.CancelledError:
                raise
            except Exception:
                # 长时间挂机中 GSMTC 管理器可能失效：置空缓存，下次调用自动重建
                self._manager = None
                if attempt == 2:
                    raise
                await asyncio.sleep(0.5)

    async def describe_track(self) -> str:
        """用 py-now-playing 接口读取当前曲目描述（锦上添花，失败不影响主流程）。"""
        if self._primary_aumid is None:
            return ""
        if PnpModern is not None:
            try:
                pnp = await self.ensure_py_now_playing()
                info = await pnp.get_media_info()
                if info is not None:
                    text = f"{getattr(info, 'artist', '')} - {getattr(info, 'title', '')}".strip(" -")
                    if text:
                        return text
            except Exception:
                pass
        if PnpLegacy is not None:
            try:
                pnp = await self.ensure_py_now_playing()
                d = await pnp.get_now_playing(self._primary_aumid) or {}
                text = f"{d.get('artist') or ''} - {d.get('title') or ''}".strip(" -")
                if text.strip("- "):
                    return text
            except Exception:
                pass
        return ""

    # ---------- 控制：播放 / 暂停（系统媒体指令，非键鼠模拟） ----------
    async def send_play(self) -> None:
        await self._control(True)

    async def send_pause(self) -> None:
        await self._control(False)

    async def _control(self, play: bool) -> None:
        verb = "播放" if play else "暂停"
        handled = False
        # 路径一：py-now-playing 0.2.x 异步控制接口
        if PnpModern is not None and self._primary_aumid:
            try:
                pnp = await self.ensure_py_now_playing()
                ok = await (pnp.play() if play else pnp.pause())
                self._log(f"（py-now-playing）已发送系统“{verb}”指令：{'成功' if ok else '未生效'}。",
                          "ok" if ok else "warn")
                handled = True
            except Exception as e:
                self._log(f"py-now-playing 控制异常，改用直接会话控制：{e!r}", "warn")
        # 路径二：直接对匹配到的 GSMTC 会话发送系统媒体指令
        if not handled:
            try:
                manager = await self.ensure_manager()
                sessions = self._netease_sessions(manager)
                ok_n = 0
                for s in sessions:
                    try:
                        r = await (s.try_play_async() if play else s.try_pause_async())
                        ok_n += 1 if r else 0
                    except Exception as e:
                        self._log(f"发送{verb}指令异常：{e!r}", "err")
                self._log(f"已向 {ok_n}/{len(sessions)} 个网易云媒体会话发送系统“{verb}”指令。",
                          "ok" if ok_n else "warn")
            except Exception as e:
                self._manager = None   # 管理器可能失效（长时间挂机），下次自动重建
                self._log(f"发送{verb}指令失败：{e!r}（将重建 GSMTC 管理器后重试）", "err")


# ============================== asyncio 工作线程 ==============================
class AsyncWorker:
    """在独立线程中运行 asyncio 事件循环。
    tkinter 主线程绝不直接执行异步任务，双方通过线程安全接口通信，避免界面卡死。"""

    def __init__(self):
        self.loop = None
        self.thread = None
        self._ready = threading.Event()

    def start(self):
        self.thread = threading.Thread(target=self._run, name="GsmtcAsyncWorker", daemon=True)
        self.thread.start()
        self._ready.wait(timeout=5)

    def _run(self):
        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)
        self._ready.set()
        try:
            self.loop.run_forever()
        finally:
            try:
                pending = asyncio.all_tasks(self.loop)
                for t in pending:
                    t.cancel()
                if pending:
                    self.loop.run_until_complete(
                        asyncio.gather(*pending, return_exceptions=True))
            finally:
                self.loop.close()

    def submit(self, coro):
        """从 tkinter 线程把协程安全投递到事件循环。"""
        return asyncio.run_coroutine_threadsafe(coro, self.loop)

    def stop(self, timeout=3.0):
        try:
            if self.loop and self.loop.is_running():
                self.loop.call_soon_threadsafe(self.loop.stop)
            if self.thread:
                self.thread.join(timeout=timeout)
        except Exception:
            pass


# ============================== 核心业务控制器 ==============================
class TaskController:
    """三种模式：①立即开始计时 ②每日定时任务 ③停止全部任务。
    全部业务运行在 asyncio 工作线程，日志/状态经队列送回 GUI。
    队列消息格式：
      ("log", 时间戳, 文本, 等级)   —— 日志行
      ("timer", 文本)              —— 大字倒计时
      ("progress", 0~100 浮点)     —— 进度条
      ("pill", 文本, 背景色, 文字色) —— 状态指示灯
      ("eta", 文本)                —— 预计结束/下次执行说明行
    """

    def __init__(self, ui_queue: queue.Queue, flags: RuntimeFlags = None):
        self.q = ui_queue
        self.flags = flags or RuntimeFlags()
        self.worker = AsyncWorker()
        self.compat = CompatBridge(self._log)      # 无 SMTC 兼容桥（音频电平 + 会话静音）
        self._mode = "gsmtc"                      # 当前控制模式：gsmtc / compat
        self._muted_by_tool = False        # 兼容模式：静音是否由本程序执行
        self._silence_since = None                # 兼容模式：连续静默起点（一次性提示用）
        self._last_play_log = 0.0                 # 稳定播放日志节流（长挂机减噪）
        self._no_sess_streak = 0                  # 兼容模式：连续无会话计数（只提示一次）
        self.bridge = GsmtcBridge(self._log)
        self._loop = None
        self._sched = None            # schedule.Scheduler（每日定时）
        self._pump_task = None        # 驱动 schedule.run_pending 的常驻协程
        self._sched_gen = 0           # 调度代际号：重启/停止调度时让旧泵自动退出
        self._countdown_task = None   # 当前倒计时任务（真正的 asyncio.Task）
        self._busy = False            # 同步忙碌标志：消除“提交协程”与“任务创建”之间的竞态
        self._fire_count = 0          # 每日定时已自动执行的次数（挂机统计）

    # ---------- 各类 UI 消息推送（跨线程 → UI 队列） ----------
    def _log(self, msg, level="info"):
        self.q.put(("log", time.strftime("%H:%M:%S"), msg, level))

    def _timer(self, text):
        self.q.put(("timer", text))

    def _progress(self, percent: float):
        self.q.put(("progress", max(0.0, min(100.0, float(percent)))))

    def _pill(self, text, colors=PILL_IDLE):
        self.q.put(("pill", text, colors[0], colors[1]))

    def _eta(self, text):
        self.q.put(("eta", text))

    def _no_session_hint(self):
        """网易云进程在运行、但系统 0 媒体会话时的精确诊断输出。"""
        if find_netease_processes():
            self._log("诊断：网易云进程正在运行，但系统当前没有任何 SMTC 媒体会话。", "warn")
            self._log(SMTC_HINT, "warn")
        else:
            self._log("诊断：未检测到网易云进程，请先手动打开网易云并播放音乐。", "warn")

    def _announce_compat(self):
        """兼容模式就绪说明（系统无网易云 SMTC 会话时）。"""
        self._log("已就绪兼容模式：状态检测 = 网易云音频会话电平；到点停声 / 开始 = "
                  "会话静音开关（仅作用于网易云，不模拟鼠标键盘）。", "ok")
        self._log("如需完整的暂停/恢复控制，可安装 BetterNCM + InfLink-rs 插件（见 README）；"
                  "程序检测到 SMTC 会话后会自动升级控制方式。")

    def _idle_ui(self):
        """恢复到待机界面状态。"""
        self._timer("--:--:--")
        self._progress(0)
        self._pill("待机 · 请先手动打开网易云并开始播放", PILL_IDLE)
        self._eta("")

    # ---------- 防休眠守卫（长时间挂机核心保障） ----------
    def _awake_needed(self) -> bool:
        """是否需要阻止系统休眠：开关勾选 且 （有倒计时 或 每日定时驻留中）。"""
        if not self.flags.prevent_sleep:
            return False
        pump_alive = self._pump_task is not None and not self._pump_task.done()
        return self._busy or self._countdown_active() or pump_alive

    def _refresh_awake(self):
        """按当前状态刷新系统唤醒标记（仅在工作线程中调用）。"""
        prevent_sleep(self._awake_needed())

    def set_sleep_guard(self, enabled: bool):
        """GUI 勾选框回调：更新防休眠开关并立即生效。"""
        self.flags.prevent_sleep = bool(enabled)

        async def _apply():
            self._refresh_awake()

        self.worker.submit(_apply())

    # ---------- 启动 ----------
    def start(self):
        self.worker.start()
        self.worker.submit(self._boot())

    async def _boot(self):
        self._loop = asyncio.get_running_loop()
        try:
            # 依赖自检报告
            if PnpModern is not None:
                self._log("py-now-playing 0.2.x（PyNowPlaying 异步接口）就绪。")
            elif PnpLegacy is not None:
                self._log("py-now-playing 0.1.x（NowPlaying 接口）就绪。")
            else:
                self._log("未检测到 py-now-playing（不影响核心功能，将直接使用 winsdk/winrt）。",
                          "warn")
            if schedule is None:
                self._log("未安装 schedule 库，每日定时模式不可用（pip install schedule）。",
                          "err")
            if psutil is None:
                self._log("未安装 psutil，进程自检功能关闭（pip install psutil）。", "warn")

            # 网易云进程自检（只读）
            procs = find_netease_processes()
            if procs:
                self._log(f"检测到网易云音乐进程：{procs[0][1]} (pid={procs[0][0]})，"
                          f"共 {len(procs)} 个。", "ok")
            else:
                self._log("未检测到网易云进程 —— 请先手动打开网易云客户端并开始播放。",
                          "warn")

            # 媒体会话自检：列出当前系统全部会话，便于确认网易云是否已注册 SMTC
            try:
                manager = await self.bridge.ensure_manager()
                all_aumids = [s.source_app_user_model_id or ""
                              for s in manager.get_sessions()]
                if all_aumids:
                    self._log(f"当前系统媒体会话 {len(all_aumids)} 个：{'；'.join(all_aumids)}")
                    hit = await self.bridge.discover_netease()
                    if hit[0]:
                        self._log(f"已自动定位网易云媒体会话：{hit[0][0]}（匹配方式：{hit[1]}）。",
                                  "ok")
                        self.bridge.set_primary_aumid(hit[0][0])
                    else:
                        self._log("系统有其它应用的媒体会话，但没有网易云 —— "
                                  "网易云原版客户端不注册 SMTC（官方限制），"
                                  "程序将自动使用兼容模式。", "warn")
                else:
                    self._no_session_hint()
            except Exception as e:
                self._log(f"GSMTC 初始化失败：{e!r}", "err")
            self._refresh_awake()   # 若勾选防休眠，从启动起即保持系统唤醒（长期挂机）
            self._idle_ui()
        except Exception as e:
            self._log(f"初始化异常：{e!r}", "err")

    # ---------- 兼容模式（无 SMTC 时自动接管） ----------
    async def _startup_unfreeze(self):
        """启动保险：解冻上次可能遗留的冻结状态（无冻结时静默无日志）。"""
        n = await asyncio.get_running_loop().run_in_executor(
            None, self.compat.suspend_resume_all, False)
        if n:
            self._log(f"启动自检：已对 {n} 个网易云进程执行解冻确认"
                      f"（清除上次可能遗留的冻结）。", "ok")

    async def _compat_ensure_playing(self, max_tries: int = 3):
        """兼容模式：确保网易云正在播放（UIA 播放栏按钮 + 电平验证）。
        返回 True=已在播放/已恢复播放。音量为 0 时无法用电平验证，不盲发。"""
        loop = asyncio.get_running_loop()
        for _ in range(max_tries):
            snap = await loop.run_in_executor(None, self.compat.snapshot)
            if snap["n"] == 0:
                self._log("兼容模式：未找到网易云音频会话（客户端可能已关闭）。", "warn")
                return False
            if snap["volume"] is not None and snap["volume"] <= 0.001:
                self._log("兼容模式：网易云会话音量为 0，无法验证播放状态"
                          "——请在网易云里调高音量。", "warn")
                return False
            if (snap["peak"] or 0.0) > PLAY_THRESHOLD:
                return True
            # 静默 → 二次加长采样确认（避免切歌间隙误判），再点击播放按钮
            snap2 = await loop.run_in_executor(None, self.compat.snapshot, 16, 0.1)
            if (snap2["peak"] or 0.0) > PLAY_THRESHOLD:
                return True
            await loop.run_in_executor(None, self.compat.press_play_pause)
            await asyncio.sleep(2.5)
        snap = await loop.run_in_executor(None, self.compat.snapshot)
        return (snap["peak"] or 0.0) > PLAY_THRESHOLD

    async def _compat_start(self):
        """倒计时开始：按「开声音」选项设置网易云会话静音状态。
        默认静音挂机（音乐继续走时长、不出声）；勾选开声音则解除静音。"""
        self._muted_by_tool = False
        self._silence_since = None
        self._no_sess_streak = 0
        loop = asyncio.get_running_loop()
        want_sound = self.flags.play_sound
        resumed = await loop.run_in_executor(None, self.compat.suspend_resume_all, False)
        if resumed:
            self._log(f"已解冻网易云进程（{resumed} 个）——从上次真暂停处继续。", "ok")
        n = await loop.run_in_executor(None, self.compat.set_mute, not want_sound)
        self._muted_by_tool = (not want_sound) and n > 0
        if n:
            if want_sound:
                self._log(f"兼容模式：开声音——已解除网易云静音（{n} 个会话）。", "ok")
            else:
                self._log(f"兼容模式：静音挂机——已静音网易云（{n} 个会话），"
                          f"音乐继续走时长、不出声。", "ok")
        else:
            self._log("兼容模式：暂无网易云音频会话——网易云播放过音乐后会自动接管。", "warn")
        ok = await self._compat_ensure_playing()
        snap = await loop.run_in_executor(None, self.compat.snapshot)
        title = snap.get("title") or ""
        if ok:
            self._log("兼容模式：" + ("音乐播放中" if want_sound else "静音挂机刷时长中")
                      + (f" ｜ {title}" if title else ""), "ok")
            self._pill("兼容模式 · " + ("音乐放声中" if want_sound else "静音挂机刷时长中"),
                       PILL_OK)
        else:
            self._log("兼容模式：未能确认网易云在播放（可能已关闭或音量为 0）——"
                      "请打开网易云并播放一次。", "warn")
            self._pill("兼容模式 · 未能确认播放", PILL_WARN)

    # ---------- 模式①：立即开始计时 ----------
    def request_immediate(self, minutes: int) -> bool:
        if self._countdown_active():
            self._log("已有倒计时任务在运行，请先【停止全部任务】再重新开始。", "warn")
            return False
        self._busy = True   # 同步置位，防止连续点击在任务真正创建前重复进入
        # 注意：run_coroutine_threadsafe 返回 concurrent.futures.Future，
        # 不能当 asyncio.Task 使用；真正的 Task 在 _start_countdown 内创建。
        self.worker.submit(self._start_countdown(minutes))
        return True

    async def _start_countdown(self, minutes: int):
        """在事件循环内创建真正的 asyncio.Task（可 cancel / await / gather）。"""
        self._countdown_task = asyncio.get_running_loop().create_task(
            self._countdown(minutes))

    def _countdown_active(self) -> bool:
        if self._busy:
            return True
        task = self._countdown_task
        return task is not None and not task.done()

    # ---------- 模式②：每日定时任务 ----------
    def request_daily(self, hour: int, minute: int, minutes: int) -> bool:
        if schedule is None:
            self._log("缺少 schedule 库，无法启用每日定时（pip install schedule）。", "err")
            return False
        self.worker.submit(self._arm_daily(hour, minute, minutes))
        return True

    async def _arm_daily(self, hour: int, minute: int, minutes: int):
        self._sched_gen += 1
        gen = self._sched_gen
        self._sched = schedule.Scheduler()
        # 每天固定 HH:MM 触发一次（本机本地时间）
        self._sched.every().day.at(f"{hour:02d}:{minute:02d}").do(
            self._on_daily_fire, minutes)
        self._pump_task = asyncio.get_event_loop().create_task(self._daily_pump(gen))
        self._refresh_awake()   # 每日驻留期间按开关保持系统唤醒
        nxt = sched_next_run(self._sched)
        self._log(f"每日定时任务已启用：每天 {hour:02d}:{minute:02d} 自动开始 "
                  f"{minutes} 分钟倒计时（触发时若已有倒计时在跑则跳过）。", "ok")
        if nxt is not None:
            self._log(f"下次自动触发时间：{nxt:%Y-%m-%d %H:%M:%S}")

    async def _daily_pump(self, gen: int):
        """后台驻留：周期性驱动 schedule 检查是否到达每日触发时间，
        并在空闲时把“距下次触发”的倒计时刷新到界面（每 3 秒一次，纯本地计算）。
        长时间挂机保障：每小时心跳日志 + 系统睡眠唤醒检测（唤醒后调度照常生效）。"""
        started = time.time()
        last_beat = time.time()
        self._log("每日定时调度已启动，软件后台驻留等待中（可最小化窗口）。")
        while gen == self._sched_gen:
            try:
                self._sched.run_pending()
            except Exception as e:
                self._log(f"定时调度异常：{e!r}", "err")
            t0 = time.time()
            await asyncio.sleep(3)
            gap = time.time() - t0
            if gap > 60:
                # 正常只休眠 3 秒：超出很多说明系统刚从睡眠/休眠中恢复
                self._log(f"检测到系统从睡眠/休眠恢复（暂停了约 {gap / 60:.0f} 分钟），"
                          f"每日定时继续有效，错过的触发点将立即补跑。", "warn")
                last_beat = time.time()
            if time.time() - last_beat >= 3600:
                days = (time.time() - started) / 86400.0
                self._log(f"💓 挂机心跳：程序运行正常，已连续驻留 {days:.1f} 天，"
                          f"累计自动执行 {self._fire_count} 次。", "ok")
                last_beat = time.time()
            if not self._countdown_active() and self._sched is not None:
                try:
                    nxt = sched_next_run(self._sched)
                    if nxt is not None:
                        remain = (nxt - dt_now()).total_seconds()
                        if remain > 0:
                            self._timer(self._fmt(remain))
                            self._progress(0)
                            self._pill("每日定时等待中（可最小化窗口）", PILL_BLUE)
                            self._eta(f"下次自动执行：{nxt:%Y-%m-%d %H:%M:%S}")
                except Exception:
                    pass
        self._log("每日定时调度已停止。")
        self._refresh_awake()

    def _on_daily_fire(self, minutes: int):
        """schedule 到点回调（运行在事件循环线程内）。"""
        if self._countdown_active():
            self._log("每日定时触发：检测到倒计时仍在进行，本次跳过。", "warn")
            return
        self._busy = True
        self._fire_count += 1
        self._log(f"每日定时时间到，自动开始第 {self._fire_count} 次倒计时。", "ok")
        self._countdown_task = self._loop.create_task(
            self._countdown(minutes, scheduled=True))

    # ---------- 倒计时主流程 ----------
    async def _countdown(self, minutes: float, scheduled: bool = False):
        # scheduled 作为协程局部参数：各次倒计时状态互不干扰，杜绝共享状态覆盖
        total_sec = max(5.0, float(minutes) * 60.0)
        end_ts = time.time() + total_sec
        self._refresh_awake()   # 按防休眠开关保持系统唤醒（倒计时期间不睡）
        try:
            # --- 起始播报与网易云会话定位 ---
            if not find_netease_processes():
                self._log("提示：未检测到网易云进程，请先手动打开客户端并开始播放。", "warn")
            try:
                aumids, how = await self.bridge.discover_netease()
                if aumids:
                    self._mode = "gsmtc"
                    self.bridge.set_primary_aumid(aumids[0])
                    self._log(f"已定位网易云媒体会话：{aumids[0]}（匹配方式：{how}）", "ok")
                    desc = await self.bridge.describe_track()
                    if desc:
                        self._log(f"当前曲目（py-now-playing 读取）：{desc}")
                else:
                    self._mode = "compat"
                    self._log("⚠ 未找到网易云媒体会话 —— 自动进入兼容模式"
                              "（音频电平检测 + 会话静音控制）。", "warn")
                    await self._compat_start()
            except Exception as e:
                self._log(f"定位网易云会话失败：{e!r}", "err")

            self._log(f"⏱ 倒计时开始：本次目标播放 {int(round(total_sec / 60))} 分钟，"
                      f"预计 {time.strftime('%H:%M:%S', time.localtime(end_ts))} 结束。"
                      f"轮询策略：{POLL_MIN_SEC:.0f}~{POLL_MAX_SEC:.0f} 秒随机低频。", "ok")
            self._eta(f"预计 {time.strftime('%H:%M:%S', time.localtime(end_ts))} 结束"
                      f" · 本次共 {int(round(total_sec / 60))} 分钟")

            # --- 主循环：界面 1 秒节拍刷新（纯本地计时），
            #     GSMTC 媒体查询保持 8~12 秒随机低频（风控稳定关键） ---
            poll_at = time.time()   # 立即做第一次状态检查
            self._silence_since = None   # 兼容模式：连续静默起点（一次性提示）
            while True:
                remain = end_ts - time.time()
                if remain <= 0:
                    break
                self._timer(self._fmt(remain))
                self._progress((total_sec - remain) / total_sec * 100.0)
                if time.time() >= poll_at:
                    poll_at = time.time() + random.uniform(POLL_MIN_SEC, POLL_MAX_SEC)
                    try:
                        st = await self.bridge.read_status()
                        if st["found"] == 0:
                            # ---- 兼容模式：无 SMTC 会话时的状态检测与静音守卫 ----
                            loop = asyncio.get_running_loop()
                            snap = await loop.run_in_executor(None, self.compat.snapshot)
                            if snap["n"] == 0:
                                self._no_sess_streak += 1
                                self._pill("兼容模式 · 等待网易云开始播放", PILL_WARN)
                                if self._no_sess_streak == 1:
                                    self._log(f"剩余 {self._fmt(remain)} ｜ 兼容模式："
                                              f"暂无网易云音频会话（播放一次音乐后自动接管）",
                                              "warn")
                                elif self._no_sess_streak == 15:
                                    self._log("兼容模式提示：网易云长时间无音频会话——"
                                              "客户端可能已被关闭。", "warn")
                            else:
                                self._no_sess_streak = 0
                                if self.flags.play_sound:
                                    if snap["mute"] and not self._muted_by_tool:
                                        # 意外静音守卫（开声音模式）：非本程序静音 → 恢复声音
                                        await loop.run_in_executor(None,
                                                                   self.compat.set_mute,
                                                                   False)
                                        self._log(f"剩余 {self._fmt(remain)} ｜ 兼容模式："
                                                  f"检测到网易云被静音，已自动恢复声音。",
                                                  "warn")
                                        snap = await loop.run_in_executor(
                                            None, self.compat.snapshot)
                                elif (not snap["mute"]) and self._muted_by_tool:
                                    # 静音挂机守卫：本程序执行静音后被解除 → 重新静音
                                    await loop.run_in_executor(None,
                                                               self.compat.set_mute,
                                                               True)
                                    self._log(f"剩余 {self._fmt(remain)} ｜ 兼容模式："
                                              f"检测到网易云声音被打开，已重新静音"
                                              f"（静音挂机）。", "warn")
                                    snap = await loop.run_in_executor(
                                        None, self.compat.snapshot)
                                title = snap.get("title") or ""
                                short = title if len(title) <= 24 else title[:23] + "…"
                                play = (snap["peak"] or 0.0) > PLAY_THRESHOLD
                                if self.flags.play_sound and snap["volume"] is not None \
                                        and snap["volume"] <= 0.001:
                                    self._pill("兼容模式 · 网易云音量为 0", PILL_WARN)
                                    self._log(f"剩余 {self._fmt(remain)} ｜ 兼容模式："
                                              f"网易云会话音量为 0，无法出声——"
                                              f"请在网易云里调高音量。", "warn")
                                elif play:
                                    tag = "兼容模式 · 音乐播放中" if self.flags.play_sound \
                                        else "静音挂机 · 刷时长中"
                                    self._pill(tag + (f" · {short}" if short else ""), PILL_OK)
                                    if time.time() - self._last_play_log >= 120 \
                                            or not self._last_play_log:
                                        self._log(f"剩余 {self._fmt(remain)} ｜ "
                                                  + ("兼容模式：音乐播放中" if self.flags.play_sound
                                                     else "静音挂机：刷时长中（无声音）")
                                                  + (f" ｜ {title}" if title else ""))
                                        self._last_play_log = time.time()
                                    self._silence_since = None
                                else:
                                    self._pill("兼容模式 · 网易云静默（可能已暂停）", PILL_WARN)
                                    if self._silence_since is None:
                                        self._silence_since = time.time()
                                        self._log(f"剩余 {self._fmt(remain)} ｜ 兼容模式："
                                                  f"网易云静默（可能已暂停）"
                                                  + (f" ｜ {title}" if title else ""), "warn")
                                    elif time.time() - self._silence_since >= 15:
                                        # 确认静默 → 通过播放栏按钮自动恢复播放
                                        self._silence_since = time.time()
                                        if await self._compat_ensure_playing():
                                            self._log("检测到网易云已暂停——已通过播放栏"
                                                      "按钮自动恢复播放。", "ok")
                                            self._silence_since = None
                                        else:
                                            self._log("自动恢复播放未成功——若网易云已关闭"
                                                      "请重新打开并播放一次。", "warn")
                        elif st["status"] == GSMT_PAUSED:
                            # 唯一会发送控制指令的异常场景：意外暂停 → 系统媒体“播放”恢复
                            self._pill("已暂停 → 正在发送播放指令恢复", PILL_WARN)
                            self._log(f"剩余 {self._fmt(remain)} ｜ 检测到网易云已暂停："
                                      f"{'；'.join(st['tracks'])} → 发送系统“播放”指令恢复",
                                      "warn")
                            await self.bridge.send_play()
                        elif st["status"] == GSMT_PLAYING:
                            # 正常播放：零操作零打扰，交给网易云原生循环（风控稳定关键）
                            self._pill("正在播放 · 无需干预", PILL_OK)
                            if time.time() - self._last_play_log >= 120 \
                                    or not self._last_play_log:
                                self._log(f"剩余 {self._fmt(remain)} ｜ 正在播放 ｜ "
                                          f"{'；'.join(st['tracks'])}")
                                self._last_play_log = time.time()
                        else:
                            self._pill(f"状态：{GSMT_STATUS_ZH.get(st['status'], st['status'])}",
                                       PILL_IDLE)
                            self._log(f"剩余 {self._fmt(remain)} ｜ 状态："
                                      f"{GSMT_STATUS_ZH.get(st['status'], st['status'])}")
                    except asyncio.CancelledError:
                        raise
                    except Exception as e:
                        self._pill("轮询异常 · 下个周期自动重试", PILL_ERR)
                        self._log(f"轮询/控制异常：{e!r}（下个周期自动重试）", "err")
                # 1 秒 UI 节拍（不产生任何媒体查询）
                await asyncio.sleep(1)

            # --- 倒计时自然结束：按当前模式执行“到点停声” ---
            self._timer("00:00:00")
            self._progress(100)
            self._log("⏰ 倒计时结束：执行到点停声。", "ok")
            if self._mode == "compat":
                loop = asyncio.get_running_loop()
                if self.flags.true_pause:
                    n = await loop.run_in_executor(None, self.compat.suspend_resume_all,
                                                   True)
                    if n:
                        self._log(f"已冻结网易云进程（{n} 个）——真暂停：音乐立即停止、"
                                  f"歌曲不再推进；次日定时或停止任务时自动恢复。", "ok")
                        self._pill("任务完成 · 网易云已真暂停", PILL_OK)
                    else:
                        n2 = await loop.run_in_executor(None, self.compat.set_mute, True)
                        self._muted_by_tool = True
                        self._log(f"进程冻结未生效，已静音网易云兜底（{n2} 个会话）——"
                                  f"到点停声完成。", "warn")
                        self._pill("任务完成 · 已静音网易云（兜底）", PILL_WARN)
                else:
                    n = await loop.run_in_executor(None, self.compat.set_mute, True)
                    self._muted_by_tool = True
                    if n:
                        self._log(f"已将网易云音频会话静音（共 {n} 个会话）——"
                                  f"到点停声完成，音乐继续走、无声。", "ok")
                        self._pill("任务完成 · 网易云已静音", PILL_OK)
                    else:
                        self._log("未找到网易云音频会话，到点停声跳过"
                                  "（网易云可能已关闭）。", "warn")
                        self._pill("任务完成 · 未找到网易云音频会话", PILL_WARN)
            else:
                self._pill("任务完成 · 已发送暂停指令", PILL_OK)
                try:
                    await self.bridge.send_pause()
                except Exception as e:
                    self._log(f"发送暂停指令失败：{e!r}", "err")
                if self._muted_by_tool:
                    # 静音挂机：暂停后恢复网易云正常音量状态，方便用户之后正常听歌
                    loop = asyncio.get_running_loop()
                    await loop.run_in_executor(None, self.compat.set_mute, False)
                    self._muted_by_tool = False
                    self._log("已恢复网易云正常音量状态（静音挂机收尾）。", "ok")
            if scheduled:
                nxt = sched_next_run(self._sched) if self._sched is not None else None
                tip = f"下次自动执行：{nxt:%Y-%m-%d %H:%M:%S}" if nxt is not None \
                    else "每日定时未在运行"
                self._log(f"📅 今日定时播放已完成（第 {self._fire_count} 次执行），"
                          f"音乐已停声；{tip}。", "ok")
        except asyncio.CancelledError:
            # 被手动停止：GSMTC 模式补发“暂停”；
            # 兼容模式的静音保持由 _stop_all 统一收尾（避免重复日志）
            self._log("任务被手动停止，正在执行收尾…", "warn")
            self._pill("正在停止 · 执行收尾中", PILL_WARN)
            try:
                if self._mode != "compat":
                    await asyncio.wait_for(self.bridge.send_pause(), timeout=3)
            except BaseException:
                pass   # 即使遭遇二次取消也要继续收尾，保证状态与日志完整
            raise
        finally:
            self._busy = False
            self._refresh_awake()   # 若每日定时仍在驻留，继续保持防休眠

    # ---------- 模式③：停止全部任务 ----------
    def request_stop_all(self):
        self.worker.submit(self._stop_all())

    async def _stop_all(self):
        self._sched_gen += 1                     # 让每日调度泵自动退出
        began = self._busy                       # 快照：进入时是否有倒计时在跑
        for t in (self._pump_task, self._countdown_task):
            if t is not None and not t.done():
                t.cancel()
                # gather(return_exceptions=True) 绝不抛异常：无论子任务以正常完成、
                # CancelledError 还是其他异常收尾，都能安全等待其清理结束
                await asyncio.gather(t, return_exceptions=True)
        self._pump_task = None
        self._countdown_task = None
        self._busy = False
        try:
            n = await asyncio.get_running_loop().run_in_executor(
                None, self.compat.suspend_resume_all, False)
            if n:
                self._log(f"已解冻网易云进程（{n} 个）——真暂停状态随任务停止一并恢复。",
                          "ok")
        except Exception:
            pass
        self._refresh_awake()   # 全部任务已停止 → 释放系统唤醒标记
        if self._muted_by_tool:
            # 停止任务 = 交还控制权：解除本程序设置的静音，恢复声音。
            # （只解除本程序设置的静音；用户自己在音量合成器静音的不动）
            self._muted_by_tool = False
            try:
                await asyncio.get_running_loop().run_in_executor(
                    None, self.compat.set_mute, False)
                self._log("已解除网易云静音、恢复声音——任务已停止，"
                          "控制权交还给你。", "ok")
            except Exception:
                pass
        if began:
            if self._mode == "compat":
                self._log("已停止全部任务（倒计时已终止，兼容模式收尾完成）。", "ok")
            else:
                self._log("已停止全部任务（倒计时已终止，暂停指令已发送）。", "ok")
        else:
            self._log("已停止全部任务（当前没有正在运行的倒计时）。", "ok")
        self._idle_ui()

    # ---------- 手动检测：媒体会话 / 播放状态 ----------
    def request_detect(self):
        self.worker.submit(self._detect())

    async def _detect(self):
        """手动触发一次媒体会话检测并输出诊断日志（不影响任何计时任务）。"""
        self._log("—— 手动检测开始 ——")
        try:
            procs = find_netease_processes()
            if procs:
                self._log(f"网易云进程：{procs[0][1]} (pid={procs[0][0]})，共 {len(procs)} 个。",
                          "ok")
            else:
                self._log("未检测到网易云进程 —— 请先手动打开网易云客户端。", "warn")
            manager = await self.bridge.ensure_manager()
            all_aumids = [s.source_app_user_model_id or "" for s in manager.get_sessions()]
            if all_aumids:
                self._log(f"系统媒体会话 {len(all_aumids)} 个：{'；'.join(all_aumids)}")
            else:
                self._log("系统当前没有任何媒体会话。", "warn")
                self._no_session_hint()
            aumids, how = await self.bridge.discover_netease()
            if aumids:
                self._log(f"已定位网易云媒体会话：{aumids[0]}（匹配方式：{how}）", "ok")
                st = await self.bridge.read_status()
                zh = GSMT_STATUS_ZH.get(st["status"], st["status"]) if st["status"] is not None \
                    else "未知"
                self._log(f"网易云当前状态：{zh}｜曲目：{'；'.join(st['tracks']) or '未知'}",
                          "ok" if st["status"] == GSMT_PLAYING else "warn")
                self._log("结论：将使用 GSMTC 完整控制模式（真暂停/恢复 + 曲目识别）。", "ok")
            else:
                self._log("未发现网易云媒体会话 —— 原因与解决方法见上方「诊断结论」。", "warn")
            # ---- 兼容模式自检（无 SMTC 时的替代通道） ----
            loop = asyncio.get_running_loop()
            snap = await loop.run_in_executor(None, self.compat.snapshot)
            if snap["n"]:
                state = "播放中" if (snap["peak"] or 0.0) > PLAY_THRESHOLD else "静默/暂停"
                self._log(f"兼容模式自检：音频会话 {snap['n']} 个 ｜ 状态：{state} ｜ "
                          f"峰值电平={snap['peak']:.4f} ｜ 静音={snap['mute']} ｜ "
                          f"会话音量={snap['volume']}", "ok")
                if snap["title"]:
                    self._log(f"兼容模式自检：窗口标题曲目：{snap['title']}")
                if snap["volume"] is not None and snap["volume"] <= 0.001:
                    self._log("自检提示：网易云会话音量为 0，请调高音量，否则到点也无法出声。",
                              "warn")
                self._log("结论：将使用兼容模式（音频电平检测 + 会话静音控制）。", "ok")
            else:
                self._log("兼容模式自检：暂无网易云音频会话（播放一次音乐后可用）。", "warn")
        except Exception as e:
            self._log(f"检测失败：{e!r}", "err")
        self._log("—— 检测结束 ——")

    # ---------- 退出清理（供窗口关闭时同步调用） ----------
    def shutdown(self):
        try:
            fut = self.worker.submit(self._stop_all())
            fut.result(timeout=8)
        except BaseException:
            pass   # 退出路径必须万无一失（含 CancelledError）
        self.worker.stop()

    @staticmethod
    def _fmt(sec: float) -> str:
        """剩余秒数 → h:mm:ss。"""
        sec = max(0, int(sec))
        h, rem = divmod(sec, 3600)
        m, s = divmod(rem, 60)
        return f"{h:d}:{m:02d}:{s:02d}"


def dt_now():
    """当前本地时间（供每日定时剩余时间计算）。"""
    import datetime
    return datetime.datetime.now()


def sched_next_run(sched):
    """兼容各版本 schedule 库取“下次运行时间”：
    1.2.x 中 Scheduler.next_run 是 property（datetime），旧版为方法。"""
    try:
        nr = sched.next_run
        if callable(nr):
            nr = nr()
        return nr
    except Exception:
        return None


# ============================== tkinter GUI（卡片式现代界面） ==============================
class App(tk.Tk):
    """主窗口：顶部品牌条 + 免责横幅 + 参数卡片 + 状态卡片（大字倒计时/进度条/状态灯）
    + 日志卡片（导出/清空/手动检测）。固定窗口大小，全部控件运行在主线程。"""

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("780x680")
        self.resizable(False, False)
        self.configure(bg=C_BG)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._cfg = load_config()
        self.q = queue.Queue()
        self.flags = RuntimeFlags()
        self.flags.play_sound = bool(self._cfg.get("play_sound", False))
        self.flags.true_pause = bool(self._cfg.get("true_pause", True))
        self.controller = TaskController(self.q, self.flags)
        # 保险：若上次运行在“真暂停”后异常退出，启动时先解冻网易云
        # （延迟到主事件循环跑起来后再提交，避免 worker 线程尚未就绪）
        self.after(1000, lambda: self.controller.worker.submit(
            self.controller._startup_unfreeze()))

        self._build_ui()
        self._apply_window_icon()
        self._setup_ttk_style()

        self.controller.start()
        self.after(150, self._drain)   # 主线程定时取队列，刷新日志与状态

    # ---------- 窗口图标（打包时附带 app.ico） ----------
    def _apply_window_icon(self):
        try:
            ico = os.path.join(app_dir(), "app.ico")
            if os.path.exists(ico):
                self.iconbitmap(ico)
        except Exception:
            pass

    def _setup_ttk_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Accent.Horizontal.TProgressbar",
                        troughcolor="#ECEEF3", bordercolor=C_CARD,
                        background=C_ACCENT, lightcolor=C_ACCENT,
                        darkcolor=C_ACCENT, thickness=10)

    # ---------- 小工具：扁平彩色按钮 ----------
    def _flat_button(self, parent, text, bg, fg, hover_bg, hover_fg,
                     command, font=("Microsoft YaHei UI", 11, "bold"), pady=9):
        btn = tk.Button(parent, text=text, command=command, bg=bg, fg=fg,
                        activebackground=hover_bg, activeforeground=hover_fg,
                        relief="flat", bd=0, cursor="hand2",
                        font=font, pady=pady, takefocus=0)
        btn.bind("<Enter>", lambda e: btn.configure(bg=hover_bg, fg=hover_fg))
        btn.bind("<Leave>", lambda e: btn.configure(bg=bg, fg=fg))
        return btn

    def _mini_button(self, parent, text, command):
        btn = tk.Button(parent, text=text, command=command,
                        bg=C_LOG_BG, fg=C_SUB, activebackground="#EEF1F7",
                        activeforeground=C_TEXT, relief="flat", bd=0,
                        cursor="hand2", font=("Microsoft YaHei UI", 9), padx=8, pady=2)
        btn.bind("<Enter>", lambda e: btn.configure(fg=C_ACCENT))
        btn.bind("<Leave>", lambda e: btn.configure(fg=C_SUB))
        return btn

    # ---------- 界面搭建 ----------
    def _build_ui(self):
        # ===== 顶部品牌条 =====
        head = tk.Frame(self, bg=C_ACCENT, height=64)
        head.pack(fill="x")
        head.pack_propagate(False)
        tk.Label(head, text="网易云定时播放助手", bg=C_ACCENT, fg="#FFFFFF",
                 font=("Microsoft YaHei UI", 16, "bold")).pack(side="left", padx=(18, 8))
        tk.Label(head, text="GSMTC 系统媒体会话 · 仅技术学习演示 " + APP_VERSION,
                 bg=C_ACCENT, fg="#F5C8C4",
                 font=("Microsoft YaHei UI", 9)).pack(side="left", pady=(16, 0))

        # ===== 免责横幅 =====
        tk.Label(self, bg="#FFF4E5", fg="#8A5A00", justify="left",
                 font=("Microsoft YaHei UI", 9), anchor="w",
                 text="⚠ 自动化挂机存在平台风控风险，仅供技术学习演示，请勿用于违反平台条款的用途。"
                      "使用前请手动打开网易云、选好歌单开启列表循环并点击播放（本程序不代劳）。"
        ).pack(fill="x", ipady=6, padx=12, pady=(10, 2))

        # ===== 参数卡片 =====
        card1 = tk.Frame(self, bg=C_CARD, highlightbackground=C_BORDER,
                         highlightthickness=1)
        card1.pack(fill="x", padx=12, pady=(4, 6))
        inner1 = tk.Frame(card1, bg=C_CARD)
        inner1.pack(fill="x", padx=14, pady=10)

        def field_label(parent, text):
            return tk.Label(parent, text=text, bg=C_CARD, fg=C_SUB,
                            font=("Microsoft YaHei UI", 10))

        def num_entry(parent, width, key):
            ent = tk.Entry(parent, width=width, justify="center", relief="flat",
                           bg="#F7F8FB", fg=C_TEXT, font=("Consolas", 12, "bold"),
                           highlightthickness=1, highlightbackground=C_BORDER,
                           highlightcolor=C_ACCENT, insertbackground=C_TEXT)
            ent.insert(0, self._cfg.get(key, "60" if key == "minutes" else "0"))
            return ent

        field_label(inner1, "播放时长").grid(row=0, column=0, padx=(0, 6))
        self.ent_minutes = num_entry(inner1, 7, "minutes")
        self.ent_minutes.grid(row=0, column=1, ipady=5)
        tk.Label(inner1, text="分钟", bg=C_CARD, fg=C_SUB,
                 font=("Microsoft YaHei UI", 10)).grid(row=0, column=2, padx=(4, 18))

        field_label(inner1, "每日定时").grid(row=0, column=3, padx=(0, 6))
        self.ent_hour = num_entry(inner1, 4, "hour")
        self.ent_hour.grid(row=0, column=4, ipady=5)
        tk.Label(inner1, text=":", bg=C_CARD, fg=C_SUB,
                 font=("Consolas", 13, "bold")).grid(row=0, column=5)
        self.ent_minute = num_entry(inner1, 4, "minute")
        self.ent_minute.grid(row=0, column=6, ipady=5)
        tk.Label(inner1, text="分", bg=C_CARD, fg=C_SUB,
                 font=("Microsoft YaHei UI", 10)).grid(row=0, column=7, padx=(4, 0))

        btns = tk.Frame(card1, bg=C_CARD)
        btns.pack(fill="x", padx=14, pady=(0, 12))
        self.btn_now = self._flat_button(btns, "▶  立即开始计时", C_ACCENT, "#FFFFFF",
                                         C_ACCENT2, "#FFFFFF", self.on_start_now)
        self.btn_now.pack(side="left", expand=True, fill="x", padx=(0, 8))
        self.btn_daily = self._flat_button(btns, "⟳  启用每日定时任务", "#FDECEA", C_ACCENT,
                                           "#F8D7D3", C_ACCENT2, self.on_start_daily)
        self.btn_daily.pack(side="left", expand=True, fill="x", padx=(0, 8))
        self.btn_stop = self._flat_button(btns, "■  停止全部任务", "#4A4E59", "#FFFFFF",
                                          "#3A3D46", "#FFFFFF", self.on_stop_all)
        self.btn_stop.pack(side="left", expand=True, fill="x")

        # 长时间挂机选项行
        opts = tk.Frame(card1, bg=C_CARD)
        opts.pack(fill="x", padx=14, pady=(0, 10))
        self.var_sleep = tk.BooleanVar(value=True)
        tk.Checkbutton(opts, text="挂机期间防止系统休眠（推荐，显示器仍可正常关闭）",
                       variable=self.var_sleep, command=self.on_toggle_sleep,
                       bg=C_CARD, fg=C_TEXT, activebackground=C_CARD,
                       font=("Microsoft YaHei UI", 9), cursor="hand2").pack(side="left")
        self.var_autostart = tk.BooleanVar(value=get_autostart())
        tk.Checkbutton(opts, text="开机自动启动（配合网易云开机自启，重启后恢复挂机）",
                       variable=self.var_autostart, command=self.on_toggle_autostart,
                       bg=C_CARD, fg=C_TEXT, activebackground=C_CARD,
                       font=("Microsoft YaHei UI", 9), cursor="hand2").pack(side="left", padx=(16, 0))
        opts2 = tk.Frame(card1, bg=C_CARD)
        opts2.pack(fill="x", padx=14, pady=(0, 10))
        self.var_play_sound = tk.BooleanVar(value=bool(self._cfg.get("play_sound", False)))
        self.flags.play_sound = self.var_play_sound.get()
        tk.Checkbutton(opts2, text="倒计时期间开声音（默认静音挂机刷时长，到点自动真暂停）",
                       variable=self.var_play_sound, command=self.on_toggle_play_sound,
                       bg=C_CARD, fg=C_TEXT, activebackground=C_CARD,
                       font=("Microsoft YaHei UI", 9), cursor="hand2").pack(side="left")
        self.var_true_pause = tk.BooleanVar(value=bool(self._cfg.get("true_pause", True)))
        tk.Checkbutton(opts2, text="到点后真暂停（冻结网易云，次日自动恢复）",
                       variable=self.var_true_pause, command=self.on_toggle_true_pause,
                       bg=C_CARD, fg=C_TEXT, activebackground=C_CARD,
                       font=("Microsoft YaHei UI", 9), cursor="hand2").pack(side="left", padx=(16, 0))

        # ===== 状态卡片：状态灯 + 大字倒计时 + 进度条 =====
        card2 = tk.Frame(self, bg=C_CARD, highlightbackground=C_BORDER,
                         highlightthickness=1)
        card2.pack(fill="x", padx=12, pady=6)
        inner2 = tk.Frame(card2, bg=C_CARD)
        inner2.pack(fill="x", padx=14, pady=(10, 12))

        self.pill = tk.Label(inner2, text="  ● 初始化中…  ", bg=PILL_IDLE[0], fg=PILL_IDLE[1],
                             font=("Microsoft YaHei UI", 10, "bold"))
        self.pill.pack(anchor="w")

        row2 = tk.Frame(inner2, bg=C_CARD)
        row2.pack(fill="x", pady=(6, 0))
        tk.Label(row2, text="剩余时间", bg=C_CARD, fg=C_SUB,
                 font=("Microsoft YaHei UI", 10)).pack(side="left")
        self.lbl_eta = tk.Label(row2, text="", bg=C_CARD, fg=C_SUB,
                                font=("Microsoft YaHei UI", 9))
        self.lbl_eta.pack(side="right")

        self.lbl_timer = tk.Label(inner2, text="--:--:--", bg=C_CARD, fg=C_TEXT,
                                  font=("Consolas", 34, "bold"), anchor="w")
        self.lbl_timer.pack(fill="x")

        self.bar = ttk.Progressbar(inner2, style="Accent.Horizontal.TProgressbar",
                                   maximum=100, value=0)
        self.bar.pack(fill="x", pady=(6, 0))

        # ===== 日志卡片 =====
        card3 = tk.Frame(self, bg=C_CARD, highlightbackground=C_BORDER,
                         highlightthickness=1)
        card3.pack(fill="both", expand=True, padx=12, pady=(6, 12))
        head3 = tk.Frame(card3, bg=C_CARD)
        head3.pack(fill="x", padx=14, pady=(10, 2))
        tk.Label(head3, text="运行日志", bg=C_CARD, fg=C_TEXT,
                 font=("Microsoft YaHei UI", 11, "bold")).pack(side="left")
        tk.Label(head3, text="等待定时 / 剩余时长 / 播放状态 / 异常", bg=C_CARD, fg=C_SUB,
                 font=("Microsoft YaHei UI", 9)).pack(side="left", padx=(8, 0), pady=(2, 0))
        self.btn_detect = self._mini_button(head3, "检测会话", self.on_detect)
        self.btn_detect.pack(side="right")
        self.btn_clear = self._mini_button(head3, "清空日志", self.on_clear_log)
        self.btn_clear.pack(side="right", padx=(0, 6))
        self.btn_export = self._mini_button(head3, "导出日志", self.on_export_log)
        self.btn_export.pack(side="right", padx=(0, 6))

        log_wrap = tk.Frame(card3, bg=C_CARD)
        log_wrap.pack(fill="both", expand=True, padx=14, pady=(2, 12))
        self.txt = scrolledtext.ScrolledText(
            log_wrap, height=14, state="disabled", wrap="word", relief="flat",
            bg=C_LOG_BG, fg="#33363D", font=("Consolas", 9),
            highlightthickness=1, highlightbackground=C_BORDER)
        self.txt.pack(fill="both", expand=True)
        for tag, color in (("info", "#33363D"), ("ok", "#157A47"),
                           ("warn", "#B26A00"), ("err", "#C02929")):
            self.txt.tag_configure(tag, foreground=color)

    # ---------- 输入校验 ----------
    def _parse_int(self, widget, lo, hi, name):
        raw = widget.get().strip()
        try:
            value = int(raw)
        except ValueError:
            messagebox.showerror(APP_TITLE, f"【{name}】必须是整数（当前输入：{raw!r}）。")
            return None
        if not (lo <= value <= hi):
            messagebox.showerror(APP_TITLE, f"【{name}】必须在 {lo} ~ {hi} 之间。")
            return None
        return value

    def _persist(self, minutes, hour, minute):
        save_config(str(minutes), str(hour), str(minute),
                    bool(self.flags.play_sound), bool(self.flags.true_pause))

    # ---------- 按钮回调 ----------
    def on_start_now(self):
        minutes = self._parse_int(self.ent_minutes, 1, MAX_MINUTES, "播放时长（分钟）")
        if minutes is None:
            return
        if self.controller.request_immediate(minutes):
            self._persist(minutes, self.ent_hour.get(), self.ent_minute.get())
            self._ui_log(f"已请求立即开始：{minutes} 分钟倒计时。", "ok")

    def on_start_daily(self):
        minutes = self._parse_int(self.ent_minutes, 1, MAX_MINUTES, "播放时长（分钟）")
        if minutes is None:
            return
        hour = self._parse_int(self.ent_hour, 0, 23, "定时小时")
        if hour is None:
            return
        minute = self._parse_int(self.ent_minute, 0, 59, "定时分钟")
        if minute is None:
            return
        if self.controller.request_daily(hour, minute, minutes):
            self._persist(minutes, hour, minute)
            self._ui_log(f"已请求启用每日定时：每天 {hour:02d}:{minute:02d}，"
                         f"{minutes} 分钟。", "ok")

    def on_stop_all(self):
        self._ui_log("已请求停止全部任务。", "warn")   # 先入队，保证日志顺序
        self.controller.request_stop_all()

    def on_detect(self):
        self._ui_log("手动检测媒体会话…")
        self.controller.request_detect()

    def on_toggle_sleep(self):
        enabled = self.var_sleep.get()
        self.controller.set_sleep_guard(enabled)
        if enabled:
            self._ui_log("防休眠已开启：挂机期间系统将保持唤醒（显示器仍可关闭）。", "ok")
        else:
            self._ui_log("防休眠已关闭：系统可能按电源计划进入睡眠，唤醒后定时仍会补跑。", "warn")

    def on_toggle_autostart(self):
        enable = self.var_autostart.get()
        try:
            set_autostart(enable)
            if enable:
                self._ui_log("已开启开机自动启动（写入 HKCU Run）。"
                             "注意：重启后仍需手动打开网易云并播放一次。", "ok")
            else:
                self._ui_log("已关闭开机自动启动。", "ok")
        except Exception as e:
            self.var_autostart.set(not enable)
            messagebox.showerror(APP_TITLE, f"设置开机自启失败：{e}")

    def on_toggle_play_sound(self):
        enable = self.var_play_sound.get()
        self.flags.play_sound = enable
        try:
            self._persist(self.ent_minutes.get(), self.ent_hour.get(),
                          self.ent_minute.get())
        except Exception:
            pass
        if enable:
            self._ui_log("已开启倒计时期间开声音：兼容模式将解除网易云静音。", "ok")
        else:
            self._ui_log("已关闭声音：静音挂机刷时长（音乐继续走、不出声）。", "ok")

    def on_toggle_true_pause(self):
        enable = self.var_true_pause.get()
        self.flags.true_pause = enable
        try:
            self._persist(self.ent_minutes.get(), self.ent_hour.get(),
                          self.ent_minute.get())
        except Exception:
            pass
        if enable:
            self._ui_log("已开启到点真暂停：时长结束后冻结网易云进程（音乐立即停止，"
                         "次日定时或停止任务时自动恢复）。", "ok")
        else:
            self._ui_log("已关闭真暂停：到点仅静音网易云（音乐继续走、无声）。", "ok")

    def on_clear_log(self):
        self.txt.configure(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.configure(state="disabled")

    def on_export_log(self):
        from tkinter import filedialog
        default = "运行日志_" + time.strftime("%Y%m%d_%H%M%S") + ".txt"
        path = filedialog.asksaveasfilename(
            title="导出运行日志", initialfile=default,
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.txt.get("1.0", "end"))
            self._ui_log(f"日志已导出：{path}", "ok")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"导出失败：{e}")

    # ---------- 日志（主线程直写） ----------
    def _ui_log(self, msg, level="info"):
        self.q.put(("log", time.strftime("%H:%M:%S"), msg, level))

    # ---------- 队列消费：刷新全部界面元素（主线程，150ms 节拍） ----------
    def _drain(self):
        try:
            while True:
                item = self.q.get_nowait()
                kind = item[0]
                if kind == "log":
                    _, ts, msg, level = item
                    self.txt.configure(state="normal")
                    self.txt.insert("end", f"[{ts}] {msg}\n", level)
                    # 日志自动瘦身，防止长时间挂机内存膨胀
                    line_count = int(self.txt.index("end-1c").split(".")[0])
                    if line_count > LOG_MAX_LINES:
                        self.txt.delete("1.0", f"{line_count - LOG_MAX_LINES // 2}.0")
                    self.txt.see("end")
                    self.txt.configure(state="disabled")
                elif kind == "timer":
                    self.lbl_timer.configure(text=item[1])
                elif kind == "progress":
                    self.bar.configure(value=item[1])
                elif kind == "pill":
                    self.pill.configure(text=f"  ● {item[1]}  ", bg=item[2], fg=item[3])
                elif kind == "eta":
                    self.lbl_eta.configure(text=item[1])
        except queue.Empty:
            pass
        self.after(150, self._drain)

    # ---------- 关闭窗口：结束全部任务并退出 ----------
    def _on_close(self):
        if not messagebox.askokcancel(
                APP_TITLE,
                "关闭窗口将停止全部任务并退出程序。\n"
                "（若倒计时进行中，会先向网易云发送系统暂停指令）\n\n确定退出？"):
            return
        try:
            self.controller.shutdown()   # 同步收尾（最多约 8 秒）
        except Exception:
            pass
        self.destroy()


# ============================== 程序入口 ==============================
def main():
    if sys.platform != "win32":
        print("本工具仅支持 Windows（依赖系统 GSMTC 媒体会话 API）。")
        return
    # DPI 感知：让界面在高分屏上不模糊（必须在创建窗口之前）
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    if not acquire_single_instance():
        root = tk.Tk()
        root.withdraw()
        messagebox.showwarning(APP_TITLE, "程序已在运行中，请勿重复启动。")
        root.destroy()
        return
    if MediaManager is None:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            APP_TITLE,
            "未找到 winsdk / winrt 库，无法访问系统媒体会话 API。\n\n"
            "请先在终端执行：\n  pip install py-now-playing schedule psutil pywin32")
        root.destroy()
        return
    app = App()
    app.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        # 打包成无窗口 exe 后若发生意外错误，把堆栈写入文件便于排查
        import traceback
        err = traceback.format_exc()
        try:
            with open(os.path.join(app_dir(), "crash.log"), "a", encoding="utf-8") as f:
                f.write("[" + time.strftime("%Y-%m-%d %H:%M:%S") + "]\n" + err + "\n")
        except Exception:
            pass
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(APP_TITLE, "程序发生未处理异常，详情见 crash.log：\n\n" + err[-800:])
            root.destroy()
        except Exception:
            print(err)
        sys.exit(1)
