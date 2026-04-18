"""
iMouseXP 自动化工具 V2.0 - GUI版
支持功能：
  1. 上传图片/视频到手机相册
  2. 上传文件到手机文件系统（iOS 15+）
  3. 从手机相册下载图片/视频
  4. 从手机文件系统下载文件（iOS 15+）
  5. 查看手机相册/文件列表
  6. 一键发布到所有在线设备（TikTok）
  7. Excel批量导入发布任务

使用方法: python transfer_gui.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
import time
import glob

import json as _json
import base64 as _base64

try:
    import imouse
except ImportError:
    messagebox.showerror("错误", "未找到 imouse 模块，请先安装：pip install imouse-py")
    sys.exit(1)


class XpAPI:
    """iMouseXP SDK 封装（与 ProAPI 接口兼容）"""

    def __init__(self, host="localhost"):
        self._api = imouse.api(host=host)
        self._helper = imouse.helper(self._api)
        self._api.start()
        self.host = host

    def _dev(self, deviceid):
        return self._helper.device(deviceid)

    def is_connected(self):
        try:
            return self._api.is_connected()
        except Exception:
            return False

    def get_device_list(self):
        try:
            devs = self._helper.console.device.list_by_id()
        except Exception:
            return {}
        if not devs:
            return {}
        result = {}
        for dev in devs:
            did = getattr(dev, 'device_id', '') or getattr(dev, 'deviceid', '')
            result[did] = {
                "name": getattr(dev, 'name', '') or '',
                "username": getattr(dev, 'user_name', '') or '',
                "device_name": getattr(dev, 'device_name', '') or '',
                "model": '',
                "ip": getattr(dev, 'ip', '') or '',
                "state": 1,
            }
        return result

    def click(self, deviceid, x, y, button="left"):
        self._dev(deviceid).mouse.click(int(x), int(y), delay=0.5)
        return {"status": 0}

    def swipe(self, deviceid, direction, length=0.6, sx=None, sy=None):
        from imouse.types import MouseSwipeParams
        sx = int(sx) if sx is not None else 200
        sy = int(sy) if sy is not None else 500
        if direction == "up":
            ex, ey = sx, max(50, int(sy - 600 * length))
        elif direction == "down":
            ex, ey = sx, min(1050, int(sy + 600 * length))
        elif direction == "left":
            ex, ey = max(10, int(sx - 400 * length)), sy
        elif direction == "right":
            ex, ey = min(450, int(sx + 400 * length)), sy
        else:
            ex, ey = sx, sy
        params = MouseSwipeParams(
            button="left", direction=direction, len=length,
            step_sleep=5, steping=20, brake=False,
            sx=sx, sy=sy, ex=ex, ey=ey
        )
        self._dev(deviceid).mouse.swipe(params)
        return {"status": 0}

    def send_key(self, deviceid, fn_key=None, key=None):
        from imouse.helper.device.keyboard import FunctionKeys
        if fn_key:
            key_map = {"WIN+h": FunctionKeys.HOME, "WIN+v": FunctionKeys.PASTE}
            fk = key_map.get(fn_key)
            if fk:
                self._dev(deviceid).key_board.send_fn_key(fk, delay=0.5)
        return {"status": 0}

    def home(self, deviceid):
        return self.send_key(deviceid, fn_key="WIN+h")

    def paste(self, deviceid):
        return self.send_key(deviceid, fn_key="WIN+v")

    def screenshot(self, deviceid):
        import time as _t
        for _ in range(3):
            try:
                data = self._dev(deviceid).image.screenshot()
                if data:
                    if isinstance(data, str):
                        return _base64.b64decode(data)
                    return data
            except Exception:
                pass
            _t.sleep(1)
        return None

    def find_image(self, deviceid, img_b64, similarity=0.7):
        try:
            results = self._dev(deviceid).image.find_image_cv([img_b64], similarity=similarity)
            if results:
                return [results[0].centre[0], results[0].centre[1]]
        except Exception:
            pass
        return None

    def find_image_ex(self, deviceid, img_list, similarity=0.7):
        try:
            results = self._dev(deviceid).image.find_image_cv(img_list, similarity=similarity)
            if results:
                return [results[0].centre[0], results[0].centre[1]]
        except Exception:
            pass
        return None

    def ocr(self, deviceid, rect=None):
        try:
            results = self._dev(deviceid).image.ocr()
            if not results:
                return []
            return [{"txt": getattr(r, 'text', '') or '',
                     "confidence": 0.9,
                     "result": [r.centre[0], r.centre[1]]}
                    for r in results]
        except Exception:
            return []

    def find_text(self, deviceid, text_list, similarity=0.7):
        try:
            results = self._dev(deviceid).image.find_text(text_list, similarity=similarity)
            if not results:
                return []
            return [{"txt": getattr(r, 'text', '') or '',
                     "result": [r.centre[0], r.centre[1]]}
                    for r in results]
        except Exception:
            return []

    def shortcut(self, deviceid, shortcut_id, parameter=None, outtime=15000):
        return {"status": 1, "message": "XP shortcut not directly supported"}

    def album_upload(self, deviceid, file_paths, album_name=""):
        try:
            result = self._dev(deviceid).shortcut.album_update(
                file_paths, album_name=album_name, outtime=60)
            return {"status": 0 if result else 1}
        except Exception as e:
            return {"status": 1, "message": str(e)}

    def clipboard_set(self, deviceid, text):
        try:
            self._dev(deviceid).shortcut.clipboard_set(text, outtime=10)
            return {"status": 0}
        except Exception as e:
            return {"status": 1, "message": str(e)}

    def exec_url(self, deviceid, url):
        try:
            self._dev(deviceid).shortcut.exec_url(url, outtime=10)
            return {"status": 0}
        except Exception as e:
            return {"status": 1, "message": str(e)}

    def album_list(self, deviceid, num=20, date=""):
        try:
            results = self._dev(deviceid).shortcut.album_get(
                album_name="", num=num, outtime=30)
            if not results:
                return []
            return [{"name": getattr(r, 'name', ''), "ext": getattr(r, 'ext', ''),
                     "size": getattr(r, 'size', ''), "time": getattr(r, 'create_time', '')}
                    for r in results]
        except Exception:
            return []

    def album_down(self, deviceid, file_names):
        from imouse.types import AlbumFileParams
        try:
            params = []
            for fn in file_names:
                parts = fn.rsplit('.', 1)
                params.append(AlbumFileParams(
                    album_name="", name=parts[0],
                    ext=parts[1] if len(parts) > 1 else ''))
            self._dev(deviceid).shortcut.album_down(params, outtime=120)
            return {"status": 0}
        except Exception as e:
            return {"status": 1, "message": str(e)}

    def file_list(self, deviceid, path="/"):
        try:
            results = self._dev(deviceid).shortcut.file_get(path=path, outtime=30)
            if not results:
                return []
            return [{"name": getattr(r, 'name', ''), "ext": getattr(r, 'ext', ''),
                     "size": getattr(r, 'size', ''), "time": getattr(r, 'create_time', '')}
                    for r in results]
        except Exception:
            return []

    def file_down(self, deviceid, file_paths):
        from imouse.types import PhoneFileParams
        try:
            params = []
            for fp in file_paths:
                fn = fp.rsplit('/', 1)[-1]
                parts = fn.rsplit('.', 1)
                params.append(PhoneFileParams(
                    name=parts[0], ext=parts[1] if len(parts) > 1 else ''))
            self._dev(deviceid).shortcut.file_down(path="/", files=params, outtime=120)
            return {"status": 0}
        except Exception as e:
            return {"status": 1, "message": str(e)}

    def file_upload(self, deviceid, file_paths, target_path="/"):
        try:
            self._dev(deviceid).shortcut.file_upload(
                file_paths, path=target_path, outtime=60)
            return {"status": 0}
        except Exception as e:
            return {"status": 1, "message": str(e)}

    def stop(self):
        try:
            self._api.stop()
        except Exception:
            pass


def _file_to_base64(path):
    with open(path, "rb") as f:
        return _base64.b64encode(f.read()).decode()


class DeviceInfo:
    def __init__(self, raw, deviceid):
        self.deviceid = deviceid
        self.device_id = deviceid
        self.name = (raw.get("name") or "").strip()
        self.user_name = (raw.get("username") or "").strip()
        self.device_name = (raw.get("device_name") or "").strip()
        self.ip = (raw.get("ip") or "").strip()
        self.state = raw.get("state", 0)

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ─────────────────────────── 常量 ───────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ICON_DIR = os.path.join(SCRIPT_DIR, "icon")

# ─────────────────────────── 自动更新 ───────────────────────────
CURRENT_VERSION = "2.0.3"
UPDATE_REPO = "qiguaizhiru/qihao---xp"
UPDATE_BRANCH = "main"
UPDATE_VERSION_URL = f"https://raw.githubusercontent.com/{UPDATE_REPO}/{UPDATE_BRANCH}/version.txt"
UPDATE_ZIP_URL = f"https://codeload.github.com/{UPDATE_REPO}/zip/refs/heads/{UPDATE_BRANCH}"

BG_MAIN = "#F0F0F0"
BG_HEADER = "#2B579A"
BG_SIDEBAR = "#FFFFFF"
BG_LOG = "#1E1E1E"
FG_LOG = "#CCCCCC"
FG_HEADER = "#FFFFFF"
ACCENT = "#2B579A"
SUCCESS_COLOR = "#4CAF50"
ERROR_COLOR = "#F44336"
WARNING_COLOR = "#FF9800"
PUBLISH_COLOR = "#E91E63"

MEDIA_EXTS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.heic', '.heif',
              '.mp4', '.mov', '.avi', '.mkv', '.m4v', '.3gp', '.wmv', '.webp'}

# 发布流程中图像识别用的icon路径
ICONS = {
    "tiktok":       os.path.join(ICON_DIR, "tiktok.bmp"),
    "plus_black":   os.path.join(ICON_DIR, "+black.bmp"),
    "plus_white":   os.path.join(ICON_DIR, "+white.bmp"),
    "record":       os.path.join(ICON_DIR, "record.bmp"),
    "record2":      os.path.join(ICON_DIR, "record2.bmp"),
    "next":         os.path.join(ICON_DIR, "next.bmp"),
    "post":         os.path.join(ICON_DIR, "post.bmp"),
    "old_post":     os.path.join(ICON_DIR, "旧版post.bmp"),
    "title":        os.path.join(ICON_DIR, "Add a catchy title.bmp"),
    "title2":       os.path.join(ICON_DIR, "Add a catchy title2.bmp"),
    "desc":         os.path.join(ICON_DIR, "Add description.bmp"),
    "long_desc":    os.path.join(ICON_DIR, "Writing a long description.bmp"),
    "usesound":     os.path.join(ICON_DIR, "usesound.bmp"),
    "three":        os.path.join(ICON_DIR, "three.png"),
    "settings":     os.path.join(ICON_DIR, "settings.png"),
    "switch_account": os.path.join(ICON_DIR, "switch_account.png"),
    "checkmark":      os.path.join(ICON_DIR, "checkmark.png"),
    "home_icon":      os.path.join(ICON_DIR, "home.png"),
    "heart":          os.path.join(ICON_DIR, "heart.png"),
}


class TransferApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"iMouseXP 自动化工具 v{CURRENT_VERSION}")
        self.root.geometry("1280x800")
        self.root.minsize(1050, 650)
        self.root.configure(bg=BG_MAIN)

        # 状态变量
        self.xp_api = None
        self.connected = False
        self.devices = []
        self.selected_files = []
        self.album_files = []
        self.phone_files = []
        self.publish_tasks = []      # 发布任务列表
        self.publishing = False      # 是否正在发布中
        self.stop_publish = False    # 停止发布标记
        self.nurturing = False       # 是否正在养号中
        self.stop_nurture = False    # 停止养号标记
        self.scheduled_timer = None  # 定时发布线程
        self.scheduled_stop = False  # 取消定时发布标记
        self.scheduled_target = None # 目标时间戳
        self.scheduled_mode = None   # "all" 或 "selected"

        self._build_ui()
        self._auto_connect()

    # ═══════════════════════════ UI 构建 ═══════════════════════════

    def _build_ui(self):
        # ── 顶部标题栏 ──
        header = tk.Frame(self.root, bg=BG_HEADER, height=56)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        tk.Label(header, text="iMouseXP 自动化工具 V2.0",
                 font=("Microsoft YaHei UI", 16, "bold"),
                 bg=BG_HEADER, fg=FG_HEADER).pack(side=tk.LEFT, padx=15)

        self.lbl_status = tk.Label(header, text="未连接",
                                   font=("Microsoft YaHei UI", 12),
                                   bg=BG_HEADER, fg="#FFD54F")
        self.lbl_status.pack(side=tk.RIGHT, padx=15)

        tk.Button(header, text="重新连接", font=("Microsoft YaHei UI", 11),
                  bg="#3A6BC5", fg="white", bd=0, padx=10, pady=2,
                  command=self._auto_connect).pack(side=tk.RIGHT, padx=5)
        tk.Button(header, text=f"检查更新 v{CURRENT_VERSION}", font=("Microsoft YaHei UI", 11),
                  bg="#4CAF50", fg="white", bd=0, padx=10, pady=2,
                  command=self._check_update).pack(side=tk.RIGHT, padx=5)

        # ── 主体区域 ──
        body = tk.Frame(self.root, bg=BG_MAIN)
        body.pack(fill=tk.BOTH, expand=True, padx=8, pady=5)

        # 左面板 - 设备列表
        left = tk.LabelFrame(body, text=" 设备列表 ", font=("Microsoft YaHei UI", 12, "bold"),
                             bg=BG_SIDEBAR, padx=5, pady=5)
        left.pack(side=tk.LEFT, fill=tk.BOTH, padx=(0, 5))

        btn_frame_top = tk.Frame(left, bg=BG_SIDEBAR)
        btn_frame_top.pack(fill=tk.X, pady=(0, 5))

        tk.Button(btn_frame_top, text="刷新设备", font=("Microsoft YaHei UI", 11),
                  command=self._refresh_devices, bg=ACCENT, fg="white", bd=0,
                  padx=8, pady=3).pack(side=tk.LEFT)
        tk.Button(btn_frame_top, text="一键切号", font=("Microsoft YaHei UI", 11),
                  command=self._switch_account_all, bg="#E91E63", fg="white", bd=0,
                  padx=8, pady=3).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame_top, text="选设备切号", font=("Microsoft YaHei UI", 11),
                  command=self._switch_account_selected, bg="#9C27B0", fg="white", bd=0,
                  padx=8, pady=3).pack(side=tk.LEFT)
        tk.Button(btn_frame_top, text="全选", font=("Microsoft YaHei UI", 11),
                  command=self._select_all_devices, bg="#757575", fg="white", bd=0,
                  padx=8, pady=3).pack(side=tk.RIGHT)
        tk.Button(btn_frame_top, text="反选", font=("Microsoft YaHei UI", 11),
                  command=self._invert_selection, bg="#757575", fg="white", bd=0,
                  padx=8, pady=3).pack(side=tk.RIGHT, padx=3)

        # 养号按钮行
        btn_frame_nurture = tk.Frame(left, bg=BG_SIDEBAR)
        btn_frame_nurture.pack(fill=tk.X, pady=(0, 5))
        self.btn_nurture = tk.Button(btn_frame_nurture, text="一键养号",
                                      font=("Microsoft YaHei UI", 11, "bold"),
                                      command=self._nurture_all,
                                      bg="#FF6F00", fg="white", bd=0,
                                      padx=12, pady=3)
        self.btn_nurture.pack(side=tk.LEFT)
        self.btn_stop_nurture = tk.Button(btn_frame_nurture, text="停止养号",
                                           font=("Microsoft YaHei UI", 11),
                                           command=self._stop_nurturing,
                                           bg=ERROR_COLOR, fg="white", bd=0,
                                           padx=8, pady=3, state=tk.DISABLED)
        self.btn_stop_nurture.pack(side=tk.LEFT, padx=5)

        columns = ("check", "custom_name", "device_name", "model", "ip")
        self.tree_devices = ttk.Treeview(left, columns=columns, show="headings",
                                          height=18, selectmode="extended")
        self.tree_devices.heading("check", text="V")
        self.tree_devices.heading("custom_name", text="自定义名")
        self.tree_devices.heading("device_name", text="设备名")
        self.tree_devices.heading("model", text="型号")
        self.tree_devices.heading("ip", text="IP")
        self.tree_devices.column("check", width=35, anchor="center")
        self.tree_devices.column("custom_name", width=90)
        self.tree_devices.column("device_name", width=90)
        self.tree_devices.column("model", width=90)
        self.tree_devices.column("ip", width=120)
        self.tree_devices.pack(fill=tk.BOTH, expand=True)
        self.tree_devices.bind("<ButtonRelease-1>", self._toggle_device_check)
        self.device_checked = {}

        # 右面板
        right = tk.Frame(body, bg=BG_MAIN)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.notebook = ttk.Notebook(right)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self._build_publish_tab()    # 一键发布（第一个tab）
        self._build_upload_tab()
        self._build_download_tab()
        self._build_browse_tab()

        # ── 底部日志区 ──
        log_frame = tk.LabelFrame(self.root, text=" 操作日志 ",
                                   font=("Microsoft YaHei UI", 9, "bold"),
                                   bg=BG_MAIN, padx=5, pady=3)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 5))

        self.txt_log = tk.Text(log_frame, height=5, bg=BG_LOG, fg=FG_LOG,
                               font=("Consolas", 11), wrap=tk.WORD, bd=0,
                               insertbackground=FG_LOG)
        self.txt_log.pack(fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(self.txt_log, command=self.txt_log.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_log.config(yscrollcommand=scrollbar.set)
        self.txt_log.tag_config("ok", foreground=SUCCESS_COLOR)
        self.txt_log.tag_config("err", foreground=ERROR_COLOR)
        self.txt_log.tag_config("warn", foreground=WARNING_COLOR)
        self.txt_log.tag_config("info", foreground="#64B5F6")

        # ── 底部状态栏 ──
        status_bar = tk.Frame(self.root, bg="#E0E0E0", height=22)
        status_bar.pack(fill=tk.X)
        status_bar.pack_propagate(False)
        self.lbl_bottom = tk.Label(status_bar, text="就绪",
                                    font=("Microsoft YaHei UI", 10),
                                    bg="#E0E0E0", fg="#555")
        self.lbl_bottom.pack(side=tk.LEFT, padx=10)
        self.progress = ttk.Progressbar(status_bar, mode="determinate", length=200)
        self.progress.pack(side=tk.RIGHT, padx=10, pady=2)

    # ─────────────────── 一键发布选项卡 ───────────────────
    def _build_publish_tab(self):
        tab = tk.Frame(self.notebook, bg=BG_SIDEBAR, padx=10, pady=10)
        self.notebook.add(tab, text="  一键发布  ")

        # 发布内容输入区
        content_frame = tk.LabelFrame(tab, text=" 发布内容 ", font=("Microsoft YaHei UI", 11),
                                       bg=BG_SIDEBAR, padx=10, pady=8)
        content_frame.pack(fill=tk.X, pady=(0, 8))

        # 内容类型
        type_row = tk.Frame(content_frame, bg=BG_SIDEBAR)
        type_row.pack(fill=tk.X, pady=(0, 5))
        tk.Label(type_row, text="内容类型:", font=("Microsoft YaHei UI", 11),
                 bg=BG_SIDEBAR).pack(side=tk.LEFT)
        self.pub_type = tk.StringVar(value="picture")
        tk.Radiobutton(type_row, text="图片", variable=self.pub_type, value="picture",
                       font=("Microsoft YaHei UI", 11), bg=BG_SIDEBAR).pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(type_row, text="视频", variable=self.pub_type, value="video",
                       font=("Microsoft YaHei UI", 11), bg=BG_SIDEBAR).pack(side=tk.LEFT, padx=5)

        # 素材文件
        file_row = tk.Frame(content_frame, bg=BG_SIDEBAR)
        file_row.pack(fill=tk.X, pady=(0, 5))
        tk.Label(file_row, text="素材文件:", font=("Microsoft YaHei UI", 11),
                 bg=BG_SIDEBAR).pack(side=tk.LEFT)
        self.entry_pub_file = tk.Entry(file_row, font=("Microsoft YaHei UI", 11), width=50)
        self.entry_pub_file.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        tk.Button(file_row, text="选择", font=("Microsoft YaHei UI", 11),
                  command=self._pick_pub_file, bg=ACCENT, fg="white", bd=0,
                  padx=8, pady=2).pack(side=tk.LEFT)

        # 素材文件夹（按自定义名分子文件夹）
        folder_row = tk.Frame(content_frame, bg=BG_SIDEBAR)
        folder_row.pack(fill=tk.X, pady=(0, 5))
        tk.Label(folder_row, text="素材文件夹:", font=("Microsoft YaHei UI", 11),
                 bg=BG_SIDEBAR).pack(side=tk.LEFT)
        self.entry_pub_folder = tk.Entry(folder_row, font=("Microsoft YaHei UI", 11), width=50)
        self.entry_pub_folder.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        # 默认路径
        default_media = r"D:\iMouseXP\Shortcut\Media"
        if os.path.isdir(default_media):
            self.entry_pub_folder.insert(0, default_media)
        tk.Button(folder_row, text="选择", font=("Microsoft YaHei UI", 11),
                  command=self._pick_pub_folder, bg=ACCENT, fg="white", bd=0,
                  padx=8, pady=2).pack(side=tk.LEFT)
        tk.Label(folder_row, text="(子文件夹名=自定义名)", font=("Microsoft YaHei UI", 10),
                 bg=BG_SIDEBAR, fg="#999").pack(side=tk.LEFT, padx=3)

        # 音乐URL
        url_row = tk.Frame(content_frame, bg=BG_SIDEBAR)
        url_row.pack(fill=tk.X, pady=(0, 5))
        tk.Label(url_row, text="音乐URL:", font=("Microsoft YaHei UI", 11),
                 bg=BG_SIDEBAR).pack(side=tk.LEFT)
        self.entry_pub_url = tk.Entry(url_row, font=("Microsoft YaHei UI", 11), width=60)
        self.entry_pub_url.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        tk.Label(url_row, text="(可选)", font=("Microsoft YaHei UI", 10),
                 bg=BG_SIDEBAR, fg="#999").pack(side=tk.LEFT)

        # 标题
        title_row = tk.Frame(content_frame, bg=BG_SIDEBAR)
        title_row.pack(fill=tk.X, pady=(0, 5))
        tk.Label(title_row, text="标    题:", font=("Microsoft YaHei UI", 11),
                 bg=BG_SIDEBAR).pack(side=tk.LEFT)
        self.entry_pub_title = tk.Entry(title_row, font=("Microsoft YaHei UI", 11), width=60)
        self.entry_pub_title.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # 描述/标签
        desc_row = tk.Frame(content_frame, bg=BG_SIDEBAR)
        desc_row.pack(fill=tk.X)
        tk.Label(desc_row, text="描述标签:", font=("Microsoft YaHei UI", 11),
                 bg=BG_SIDEBAR).pack(side=tk.LEFT, anchor=tk.N)
        self.text_pub_desc = tk.Text(desc_row, font=("Microsoft YaHei UI", 11),
                                      height=3, width=60, wrap=tk.WORD)
        self.text_pub_desc.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # 发布参数
        param_frame = tk.LabelFrame(tab, text=" 发布参数 ", font=("Microsoft YaHei UI", 11),
                                     bg=BG_SIDEBAR, padx=10, pady=5)
        param_frame.pack(fill=tk.X, pady=(0, 8))

        p_row1 = tk.Frame(param_frame, bg=BG_SIDEBAR)
        p_row1.pack(fill=tk.X, pady=2)

        tk.Label(p_row1, text="步骤间隔(秒):", font=("Microsoft YaHei UI", 11),
                 bg=BG_SIDEBAR).pack(side=tk.LEFT)
        self.entry_step_delay = tk.Entry(p_row1, font=("Microsoft YaHei UI", 11), width=5)
        self.entry_step_delay.insert(0, "3")
        self.entry_step_delay.pack(side=tk.LEFT, padx=5)

        tk.Label(p_row1, text="识别等待(秒):", font=("Microsoft YaHei UI", 11),
                 bg=BG_SIDEBAR).pack(side=tk.LEFT, padx=(20, 0))
        self.entry_find_timeout = tk.Entry(p_row1, font=("Microsoft YaHei UI", 11), width=5)
        self.entry_find_timeout.insert(0, "15")
        self.entry_find_timeout.pack(side=tk.LEFT, padx=5)

        tk.Label(p_row1, text="设备间隔(秒):", font=("Microsoft YaHei UI", 11),
                 bg=BG_SIDEBAR).pack(side=tk.LEFT, padx=(20, 0))
        self.entry_dev_delay = tk.Entry(p_row1, font=("Microsoft YaHei UI", 11), width=5)
        self.entry_dev_delay.insert(0, "2")
        self.entry_dev_delay.pack(side=tk.LEFT, padx=5)

        # Excel导入区
        excel_frame = tk.LabelFrame(tab, text=" Excel批量导入 ", font=("Microsoft YaHei UI", 11),
                                     bg=BG_SIDEBAR, padx=10, pady=5)
        excel_frame.pack(fill=tk.X, pady=(0, 8))

        ex_row = tk.Frame(excel_frame, bg=BG_SIDEBAR)
        ex_row.pack(fill=tk.X)
        tk.Button(ex_row, text="导入Excel任务", font=("Microsoft YaHei UI", 11),
                  command=self._import_excel, bg=ACCENT, fg="white", bd=0,
                  padx=12, pady=4).pack(side=tk.LEFT)
        self.lbl_excel_info = tk.Label(ex_row, text="格式: devices | file | type | url | title | description | status | scheduled_time",
                                        font=("Microsoft YaHei UI", 10), bg=BG_SIDEBAR, fg="#999")
        self.lbl_excel_info.pack(side=tk.LEFT, padx=10)

        # 存草稿选项
        drafts_frame = tk.Frame(tab, bg=BG_SIDEBAR)
        drafts_frame.pack(fill=tk.X, pady=(0, 5))
        self.save_drafts_var = tk.BooleanVar(value=False)
        tk.Checkbutton(drafts_frame, text="存草稿 (不发布，保存为草稿)",
                       variable=self.save_drafts_var,
                       font=("Microsoft YaHei UI", 11),
                       bg=BG_SIDEBAR, fg="#E65100").pack(side=tk.LEFT)

        # 定时发布
        sched_frame = tk.LabelFrame(tab, text=" 定时发布 ", font=("Microsoft YaHei UI", 11),
                                     bg=BG_SIDEBAR, padx=10, pady=5)
        sched_frame.pack(fill=tk.X, pady=(0, 5))

        sr1 = tk.Frame(sched_frame, bg=BG_SIDEBAR)
        sr1.pack(fill=tk.X, pady=2)
        tk.Label(sr1, text="定时时间:", font=("Microsoft YaHei UI", 11),
                 bg=BG_SIDEBAR).pack(side=tk.LEFT)
        self.entry_sched_time = tk.Entry(sr1, font=("Microsoft YaHei UI", 11), width=22)
        # 默认值：当前时间 + 1 小时
        _default_sched = time.strftime("%Y-%m-%d %H:%M:%S",
                                        time.localtime(time.time() + 3600))
        self.entry_sched_time.insert(0, _default_sched)
        self.entry_sched_time.pack(side=tk.LEFT, padx=5)
        tk.Label(sr1, text="格式: YYYY-MM-DD HH:MM:SS",
                 font=("Microsoft YaHei UI", 10),
                 bg=BG_SIDEBAR, fg="#999").pack(side=tk.LEFT, padx=5)

        sr2 = tk.Frame(sched_frame, bg=BG_SIDEBAR)
        sr2.pack(fill=tk.X, pady=2)
        self.btn_sched_all = tk.Button(sr2, text="定时发布到所有在线设备",
                                        font=("Microsoft YaHei UI", 11, "bold"),
                                        command=lambda: self._schedule_publish("all"),
                                        bg="#FF9800", fg="white", bd=0,
                                        padx=12, pady=4)
        self.btn_sched_all.pack(side=tk.LEFT)
        self.btn_sched_selected = tk.Button(sr2, text="定时发布到选中设备",
                                             font=("Microsoft YaHei UI", 11),
                                             command=lambda: self._schedule_publish("selected"),
                                             bg="#795548", fg="white", bd=0,
                                             padx=10, pady=4)
        self.btn_sched_selected.pack(side=tk.LEFT, padx=5)
        self.btn_sched_cancel = tk.Button(sr2, text="取消定时",
                                           font=("Microsoft YaHei UI", 11),
                                           command=self._cancel_scheduled_publish,
                                           bg="#757575", fg="white", bd=0,
                                           padx=10, pady=4,
                                           state=tk.DISABLED)
        self.btn_sched_cancel.pack(side=tk.LEFT, padx=5)
        self.lbl_sched_status = tk.Label(sr2, text="状态: 未设定",
                                          font=("Microsoft YaHei UI", 10),
                                          bg=BG_SIDEBAR, fg="#666")
        self.lbl_sched_status.pack(side=tk.LEFT, padx=10)

        # 发布按钮区
        action_frame = tk.Frame(tab, bg=BG_SIDEBAR)
        action_frame.pack(fill=tk.X, pady=(5, 0))

        self.btn_publish = tk.Button(action_frame,
                                      text="  一键发布到所有在线设备  ",
                                      font=("Microsoft YaHei UI", 15, "bold"),
                                      command=self._do_publish_click,
                                      bg=PUBLISH_COLOR, fg="white",
                                      bd=0, padx=30, pady=10,
                                      activebackground="#C2185B")
        self.btn_publish.pack(side=tk.LEFT, padx=(0, 15))

        self.btn_publish_selected = tk.Button(action_frame,
                                               text="  发布到选中设备  ",
                                               font=("Microsoft YaHei UI", 13, "bold"),
                                               command=self._do_publish_selected_click,
                                               bg=ACCENT, fg="white",
                                               bd=0, padx=20, pady=8)
        self.btn_publish_selected.pack(side=tk.LEFT, padx=(0, 15))

        self.btn_stop_publish = tk.Button(action_frame,
                                           text="  停止发布  ",
                                           font=("Microsoft YaHei UI", 13, "bold"),
                                           command=self._stop_publishing,
                                           bg=ERROR_COLOR, fg="white",
                                           bd=0, padx=20, pady=8,
                                           state=tk.DISABLED)
        self.btn_stop_publish.pack(side=tk.LEFT)

    # ─── 上传选项卡 ───
    def _build_upload_tab(self):
        tab = tk.Frame(self.notebook, bg=BG_SIDEBAR, padx=10, pady=10)
        self.notebook.add(tab, text="  上传  ")

        mode_frame = tk.LabelFrame(tab, text=" 上传目标 ", font=("Microsoft YaHei UI", 11),
                                   bg=BG_SIDEBAR, padx=10, pady=5)
        mode_frame.pack(fill=tk.X, pady=(0, 8))
        self.upload_mode = tk.StringVar(value="album")
        tk.Radiobutton(mode_frame, text="上传到相册", variable=self.upload_mode,
                       value="album", font=("Microsoft YaHei UI", 12),
                       bg=BG_SIDEBAR).pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(mode_frame, text="上传到文件系统 (iOS 15+)", variable=self.upload_mode,
                       value="file", font=("Microsoft YaHei UI", 12),
                       bg=BG_SIDEBAR).pack(side=tk.LEFT, padx=10)

        param_frame = tk.Frame(tab, bg=BG_SIDEBAR)
        param_frame.pack(fill=tk.X, pady=(0, 8))
        tk.Label(param_frame, text="相册/目标路径:", font=("Microsoft YaHei UI", 11),
                 bg=BG_SIDEBAR).pack(side=tk.LEFT)
        self.entry_album = tk.Entry(param_frame, font=("Microsoft YaHei UI", 11), width=30)
        self.entry_album.pack(side=tk.LEFT, padx=5)
        tk.Label(param_frame, text="(留空=默认相册/根目录)", font=("Microsoft YaHei UI", 10),
                 bg=BG_SIDEBAR, fg="#999").pack(side=tk.LEFT)

        # 按设备分文件夹上传
        folder_frame = tk.LabelFrame(tab, text=" 按自定义名分文件夹上传 ", font=("Microsoft YaHei UI", 11),
                                      bg=BG_SIDEBAR, padx=10, pady=5)
        folder_frame.pack(fill=tk.X, pady=(0, 8))

        uf_row = tk.Frame(folder_frame, bg=BG_SIDEBAR)
        uf_row.pack(fill=tk.X)
        tk.Label(uf_row, text="素材文件夹:", font=("Microsoft YaHei UI", 11),
                 bg=BG_SIDEBAR).pack(side=tk.LEFT)
        self.entry_upload_folder = tk.Entry(uf_row, font=("Microsoft YaHei UI", 11), width=50)
        self.entry_upload_folder.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        # 默认路径
        default_media = r"D:\iMouseXP\Shortcut\Media"
        if os.path.isdir(default_media):
            self.entry_upload_folder.insert(0, default_media)
        tk.Button(uf_row, text="选择", font=("Microsoft YaHei UI", 11),
                  command=self._pick_upload_folder, bg=ACCENT, fg="white", bd=0,
                  padx=8, pady=2).pack(side=tk.LEFT)
        tk.Label(uf_row, text="(子文件夹名=自定义名)", font=("Microsoft YaHei UI", 10),
                 bg=BG_SIDEBAR, fg="#999").pack(side=tk.LEFT, padx=3)

        # 统一文件上传
        file_frame = tk.LabelFrame(tab, text=" 统一文件上传 (所有设备相同文件) ", font=("Microsoft YaHei UI", 11),
                                   bg=BG_SIDEBAR, padx=5, pady=5)
        file_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        btn_row = tk.Frame(file_frame, bg=BG_SIDEBAR)
        btn_row.pack(fill=tk.X, pady=(0, 5))
        tk.Button(btn_row, text="选择文件", font=("Microsoft YaHei UI", 11),
                  command=self._pick_files, bg=ACCENT, fg="white", bd=0,
                  padx=12, pady=4).pack(side=tk.LEFT)
        tk.Button(btn_row, text="选择文件夹", font=("Microsoft YaHei UI", 11),
                  command=self._pick_folder, bg=ACCENT, fg="white", bd=0,
                  padx=12, pady=4).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_row, text="清空列表", font=("Microsoft YaHei UI", 11),
                  command=self._clear_files, bg="#757575", fg="white", bd=0,
                  padx=12, pady=4).pack(side=tk.RIGHT)
        self.lbl_file_count = tk.Label(btn_row, text="已选: 0 个文件",
                                        font=("Microsoft YaHei UI", 11),
                                        bg=BG_SIDEBAR, fg="#666")
        self.lbl_file_count.pack(side=tk.RIGHT, padx=10)

        self.listbox_files = tk.Listbox(file_frame, font=("Consolas", 11),
                                         height=6, selectmode=tk.EXTENDED)
        self.listbox_files.pack(fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(self.listbox_files, command=self.listbox_files.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox_files.config(yscrollcommand=sb.set)

        action_frame = tk.Frame(tab, bg=BG_SIDEBAR)
        action_frame.pack(fill=tk.X)
        self.btn_upload = tk.Button(action_frame, text="上传到选中设备",
                                    font=("Microsoft YaHei UI", 13, "bold"),
                                    command=self._do_upload, bg=ACCENT, fg="white",
                                    bd=0, padx=20, pady=8)
        self.btn_upload.pack(side=tk.LEFT, padx=(0, 10))
        self.btn_one_click = tk.Button(action_frame, text="一键上传到所有设备",
                                        font=("Microsoft YaHei UI", 13, "bold"),
                                        command=self._do_one_click_upload,
                                        bg=SUCCESS_COLOR, fg="white",
                                        bd=0, padx=20, pady=8)
        self.btn_one_click.pack(side=tk.LEFT)

    # ─── 下载选项卡 ───
    def _build_download_tab(self):
        tab = tk.Frame(self.notebook, bg=BG_SIDEBAR, padx=10, pady=10)
        self.notebook.add(tab, text="  下载  ")

        mode_frame = tk.LabelFrame(tab, text=" 下载来源 ", font=("Microsoft YaHei UI", 11),
                                   bg=BG_SIDEBAR, padx=10, pady=5)
        mode_frame.pack(fill=tk.X, pady=(0, 8))
        self.download_mode = tk.StringVar(value="album")
        tk.Radiobutton(mode_frame, text="从相册下载", variable=self.download_mode,
                       value="album", font=("Microsoft YaHei UI", 12),
                       bg=BG_SIDEBAR).pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(mode_frame, text="从文件系统下载 (iOS 15+)", variable=self.download_mode,
                       value="file", font=("Microsoft YaHei UI", 12),
                       bg=BG_SIDEBAR).pack(side=tk.LEFT, padx=10)

        param_frame = tk.Frame(tab, bg=BG_SIDEBAR)
        param_frame.pack(fill=tk.X, pady=(0, 8))
        tk.Label(param_frame, text="相册名/路径:", font=("Microsoft YaHei UI", 11),
                 bg=BG_SIDEBAR).pack(side=tk.LEFT)
        self.entry_dl_path = tk.Entry(param_frame, font=("Microsoft YaHei UI", 11), width=25)
        self.entry_dl_path.pack(side=tk.LEFT, padx=5)
        tk.Label(param_frame, text="数量:", font=("Microsoft YaHei UI", 11),
                 bg=BG_SIDEBAR).pack(side=tk.LEFT, padx=(15, 0))
        self.entry_dl_num = tk.Entry(param_frame, font=("Microsoft YaHei UI", 11), width=6)
        self.entry_dl_num.insert(0, "20")
        self.entry_dl_num.pack(side=tk.LEFT, padx=5)
        tk.Button(param_frame, text="获取文件列表", font=("Microsoft YaHei UI", 11),
                  command=self._do_list_remote, bg=ACCENT, fg="white", bd=0,
                  padx=12, pady=4).pack(side=tk.LEFT, padx=10)

        list_frame = tk.LabelFrame(tab, text=" 手机端文件 ", font=("Microsoft YaHei UI", 11),
                                   bg=BG_SIDEBAR, padx=5, pady=5)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        cols_dl = ("name", "ext", "size", "time", "album")
        self.tree_remote = ttk.Treeview(list_frame, columns=cols_dl, show="headings",
                                         height=8, selectmode="extended")
        for col, txt, w in [("name","文件名",200),("ext","格式",50),("size","大小",80),
                            ("time","创建时间",150),("album","相册",100)]:
            self.tree_remote.heading(col, text=txt)
            self.tree_remote.column(col, width=w)
        self.tree_remote.pack(fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(list_frame, command=self.tree_remote.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_remote.config(yscrollcommand=sb.set)

        action_frame = tk.Frame(tab, bg=BG_SIDEBAR)
        action_frame.pack(fill=tk.X)
        self.btn_download = tk.Button(action_frame, text="下载选中文件",
                                      font=("Microsoft YaHei UI", 13, "bold"),
                                      command=self._do_download, bg=ACCENT, fg="white",
                                      bd=0, padx=20, pady=8)
        self.btn_download.pack(side=tk.LEFT)
        tk.Button(action_frame, text="全选", font=("Microsoft YaHei UI", 11),
                  command=lambda: self._select_all_tree(self.tree_remote),
                  bg="#757575", fg="white", bd=0, padx=10, pady=4).pack(side=tk.LEFT, padx=10)

    # ─── 浏览选项卡 ───
    def _build_browse_tab(self):
        tab = tk.Frame(self.notebook, bg=BG_SIDEBAR, padx=10, pady=10)
        self.notebook.add(tab, text="  浏览/管理  ")

        tk.Label(tab, text="选择设备后可查看相册和文件系统内容",
                 font=("Microsoft YaHei UI", 12), bg=BG_SIDEBAR, fg="#666").pack(anchor=tk.W, pady=(0, 10))

        btn_frame = tk.Frame(tab, bg=BG_SIDEBAR)
        btn_frame.pack(fill=tk.X, pady=(0, 8))
        tk.Button(btn_frame, text="查看相册", font=("Microsoft YaHei UI", 11),
                  command=lambda: self._do_browse("album"), bg=ACCENT, fg="white",
                  bd=0, padx=12, pady=4).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="查看文件系统", font=("Microsoft YaHei UI", 11),
                  command=lambda: self._do_browse("file"), bg=ACCENT, fg="white",
                  bd=0, padx=12, pady=4).pack(side=tk.LEFT, padx=5)

        param_frame = tk.Frame(tab, bg=BG_SIDEBAR)
        param_frame.pack(fill=tk.X, pady=(0, 8))
        tk.Label(param_frame, text="相册名/路径:", font=("Microsoft YaHei UI", 11),
                 bg=BG_SIDEBAR).pack(side=tk.LEFT)
        self.entry_browse_path = tk.Entry(param_frame, font=("Microsoft YaHei UI", 11), width=25)
        self.entry_browse_path.pack(side=tk.LEFT, padx=5)
        tk.Label(param_frame, text="数量:", font=("Microsoft YaHei UI", 11),
                 bg=BG_SIDEBAR).pack(side=tk.LEFT, padx=(15, 0))
        self.entry_browse_num = tk.Entry(param_frame, font=("Microsoft YaHei UI", 11), width=6)
        self.entry_browse_num.insert(0, "20")
        self.entry_browse_num.pack(side=tk.LEFT, padx=5)

        cols_br = ("name", "ext", "size", "time", "info")
        self.tree_browse = ttk.Treeview(tab, columns=cols_br, show="headings",
                                         height=12, selectmode="extended")
        for col, txt, w in [("name","文件名",200),("ext","格式",60),("size","大小",80),
                            ("time","创建时间",150),("info","附加信息",120)]:
            self.tree_browse.heading(col, text=txt)
            self.tree_browse.column(col, width=w)
        self.tree_browse.pack(fill=tk.BOTH, expand=True)

    # ═══════════════════════════ 连接与设备 ═══════════════════════════

    def _auto_connect(self):
        self.log("正在连接 iMouse XP内核 (localhost:9911) ...", "info")
        self._set_status("连接中...", "#FFD54F")
        threading.Thread(target=self._connect_thread, daemon=True).start()

    def _connect_thread(self):
        try:
            self.xp_api = XpAPI(host="localhost")
            for i in range(15):
                time.sleep(1)
                if self.xp_api.is_connected():
                    self.connected = True
                    self.root.after(0, lambda: self._set_status("已连接", SUCCESS_COLOR))
                    self.root.after(0, lambda: self.log("已成功连接到 iMouse XP内核 (9911)", "ok"))
                    self.root.after(100, self._refresh_devices)
                    return
            self.root.after(0, lambda: self._set_status("连接超时", ERROR_COLOR))
            self.root.after(0, lambda: self.log("连接超时，请确认 iMouse XP内核已启动", "err"))
        except Exception as e:
            self.root.after(0, lambda: self._set_status("连接失败", ERROR_COLOR))
            self.root.after(0, lambda: self.log(f"连接失败: {e}", "err"))

    def _refresh_devices(self):
        if not self.connected:
            self.log("未连接内核，无法刷新设备", "warn")
            return
        threading.Thread(target=self._refresh_devices_thread, daemon=True).start()

    def _refresh_devices_thread(self):
        try:
            raw = self.xp_api.get_device_list()
            # 调试：打印第一台设备的原始字段
            if raw:
                first_key = next(iter(raw))
                first_val = raw[first_key]
                if isinstance(first_val, dict):
                    self.root.after(0, lambda f=dict(first_val), k=first_key:
                        self.log(f"[DEBUG] 设备 {k} 全部字段: {f}", "info"))
            self.devices = [DeviceInfo(v, k) for k, v in raw.items()] if raw else []
            self.root.after(0, self._update_device_tree)
            self.root.after(0, lambda: self.log(f"已刷新设备列表，共 {len(self.devices)} 台设备", "info"))
        except Exception as e:
            self.root.after(0, lambda: self.log(f"刷新设备失败: {e}", "err"))

    def _update_device_tree(self):
        self.tree_devices.delete(*self.tree_devices.get_children())
        self.device_checked.clear()
        for dev in self.devices:
            # name = 自定义名（用户在控制台设置的别名）
            custom_name = getattr(dev, 'name', '') or ''
            # user_name = 设备名（如 476, 184 等数字编号）
            device_name = getattr(dev, 'user_name', '') or ''
            # device_name = 手机型号（如 iPhone 11）
            model = getattr(dev, 'device_name', '') or ''
            ip = getattr(dev, 'ip', '') or ''
            iid = self.tree_devices.insert("", tk.END, values=("[ ]", custom_name, device_name, model, ip))
            self.device_checked[iid] = False

    def _toggle_device_check(self, event):
        iid = self.tree_devices.identify_row(event.y)
        if not iid:
            return
        checked = not self.device_checked.get(iid, False)
        self.device_checked[iid] = checked
        vals = list(self.tree_devices.item(iid, "values"))
        vals[0] = "[x]" if checked else "[ ]"
        self.tree_devices.item(iid, values=vals)

    def _select_all_devices(self):
        for iid in self.tree_devices.get_children():
            self.device_checked[iid] = True
            vals = list(self.tree_devices.item(iid, "values"))
            vals[0] = "[x]"
            self.tree_devices.item(iid, values=vals)

    def _invert_selection(self):
        for iid in self.tree_devices.get_children():
            checked = not self.device_checked.get(iid, False)
            self.device_checked[iid] = checked
            vals = list(self.tree_devices.item(iid, "values"))
            vals[0] = "[x]" if checked else "[ ]"
            self.tree_devices.item(iid, values=vals)

    def _get_checked_devices(self):
        checked = []
        children = self.tree_devices.get_children()
        for i, iid in enumerate(children):
            if self.device_checked.get(iid, False) and i < len(self.devices):
                checked.append(self.devices[i])
        return checked

    def _get_first_checked_device(self):
        devs = self._get_checked_devices()
        if not devs:
            self.log("请先在左侧勾选至少一台设备", "warn")
            return None
        return devs[0]

    def _select_all_tree(self, tree):
        tree.selection_set(tree.get_children())

    # ═══════════════════════════ 发布功能 ═══════════════════════════

    def _pick_pub_file(self):
        filetypes = [
            ("图片和视频", "*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.heic;*.heif;*.mp4;*.mov;*.avi;*.mkv;*.m4v;*.webp"),
            ("所有文件", "*.*")
        ]
        path = filedialog.askopenfilename(title="选择要发布的素材文件", filetypes=filetypes)
        if path:
            self.entry_pub_file.delete(0, tk.END)
            self.entry_pub_file.insert(0, os.path.abspath(path))

    def _pick_pub_folder(self):
        folder = filedialog.askdirectory(title="选择素材文件夹（子文件夹名=设备自定义名）")
        if folder:
            self.entry_pub_folder.delete(0, tk.END)
            self.entry_pub_folder.insert(0, os.path.abspath(folder))
            # 扫描子文件夹，提示用户
            subs = [d for d in os.listdir(folder) if os.path.isdir(os.path.join(folder, d))]
            if subs:
                self.log(f"已选择素材文件夹，发现 {len(subs)} 个子文件夹: {', '.join(subs[:10])}" +
                         ("..." if len(subs) > 10 else ""), "info")
            else:
                self.log("该文件夹下没有子文件夹，发布时将对所有设备使用该文件夹中的文件", "warn")

    def _get_folder_file_for_device(self, folder, dev_custom_name):
        """根据设备自定义名查找对应子文件夹中的第一个媒体文件"""
        dev_folder = os.path.join(folder, dev_custom_name)
        if not os.path.isdir(dev_folder):
            return None
        for f in os.listdir(dev_folder):
            ext = os.path.splitext(f)[1].lower()
            if ext in MEDIA_EXTS:
                return os.path.abspath(os.path.join(dev_folder, f))
        return None

    def _get_folder_files_for_device(self, folder, dev_custom_name):
        """根据设备自定义名查找对应子文件夹中的所有媒体文件"""
        dev_folder = os.path.join(folder, dev_custom_name)
        if not os.path.isdir(dev_folder):
            return []
        result = []
        for root, dirs, files in os.walk(dev_folder):
            for f in files:
                ext = os.path.splitext(f)[1].lower()
                if ext in MEDIA_EXTS:
                    result.append(os.path.abspath(os.path.join(root, f)))
        return result

    def _pick_upload_folder(self):
        folder = filedialog.askdirectory(title="选择素材文件夹（子文件夹名=设备自定义名）")
        if folder:
            self.entry_upload_folder.delete(0, tk.END)
            self.entry_upload_folder.insert(0, os.path.abspath(folder))
            subs = [d for d in os.listdir(folder) if os.path.isdir(os.path.join(folder, d))]
            if subs:
                self.log(f"已选择上传文件夹，发现 {len(subs)} 个子文件夹: {', '.join(subs[:10])}" +
                         ("..." if len(subs) > 10 else ""), "info")
            else:
                self.log("该文件夹下没有子文件夹", "warn")

    def _import_excel(self):
        if not HAS_OPENPYXL:
            messagebox.showerror("错误", "需要 openpyxl 模块\n请运行: pip install openpyxl")
            return
        path = filedialog.askopenfilename(title="选择Excel任务文件",
                                           filetypes=[("Excel文件", "*.xlsx")])
        if not path:
            return
        try:
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            tasks = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or not row[0]:
                    continue
                task = {
                    "devices": str(row[0]) if row[0] else "",
                    "file": str(row[1]) if len(row) > 1 and row[1] else "",
                    "type": str(row[2]) if len(row) > 2 and row[2] else "picture",
                    "url": str(row[3]) if len(row) > 3 and row[3] else "",
                    "title": str(row[4]) if len(row) > 4 and row[4] else "",
                    "description": str(row[5]) if len(row) > 5 and row[5] else "",
                    "status": str(row[6]) if len(row) > 6 and row[6] else "",
                    "scheduled_time": str(row[7]) if len(row) > 7 and row[7] else "",
                }
                tasks.append(task)
            self.publish_tasks = tasks
            self.log(f"已导入 {len(tasks)} 条发布任务", "ok")
            self.lbl_excel_info.config(text=f"已导入 {len(tasks)} 条任务 - {os.path.basename(path)}")
        except Exception as e:
            self.log(f"导入Excel失败: {e}", "err")

    def _do_publish_click(self):
        """一键发布到所有在线设备 - 弹出确认"""
        if not self.connected:
            self.log("未连接内核", "warn")
            return
        if not self.devices:
            self.log("没有在线设备", "warn")
            return

        pub_file = self.entry_pub_file.get().strip()
        pub_folder = self.entry_pub_folder.get().strip()
        if not pub_file and not pub_folder and not self.publish_tasks:
            self.log("请选择素材文件、素材文件夹或导入Excel任务", "warn")
            return

        n = len(self.devices)
        msg = f"确认要发布到所有 {n} 台在线设备吗？\n\n"
        if pub_folder:
            msg += f"素材文件夹: {os.path.basename(pub_folder)}\n(按自定义名匹配子文件夹)\n"
        elif pub_file:
            msg += f"素材: {os.path.basename(pub_file)}\n"
            title = self.entry_pub_title.get().strip()
            if title:
                msg += f"标题: {title}\n"
        if self.publish_tasks:
            msg += f"Excel任务: {len(self.publish_tasks)} 条\n"
        msg += "\n点击 [是] 开始发布"

        if messagebox.askyesno("确认发布", msg, icon="warning"):
            self._start_publish(self.devices)

    def _do_publish_selected_click(self):
        """发布到选中设备"""
        devs = self._get_checked_devices()
        if not devs:
            self.log("请先勾选要发布的设备", "warn")
            return

        pub_file = self.entry_pub_file.get().strip()
        pub_folder = self.entry_pub_folder.get().strip()
        if not pub_file and not pub_folder and not self.publish_tasks:
            self.log("请选择素材文件、素材文件夹或导入Excel任务", "warn")
            return

        msg = f"确认要发布到选中的 {len(devs)} 台设备吗？"
        if messagebox.askyesno("确认发布", msg, icon="warning"):
            self._start_publish(devs)

    def _capture_publish_snapshot(self):
        """捕获当前UI上所有发布参数的快照"""
        return {
            "step_delay": float(self.entry_step_delay.get() or "3"),
            "find_timeout": float(self.entry_find_timeout.get() or "15"),
            "dev_delay": float(self.entry_dev_delay.get() or "2"),
            "pub_file": self.entry_pub_file.get().strip(),
            "pub_folder": self.entry_pub_folder.get().strip(),
            "pub_url": self.entry_pub_url.get().strip(),
            "pub_title": self.entry_pub_title.get().strip(),
            "pub_desc": self.text_pub_desc.get("1.0", tk.END).strip(),
            "pub_type": self.pub_type.get(),
            "save_drafts": self.save_drafts_var.get(),
            "publish_tasks": list(self.publish_tasks),  # 深拷贝列表
        }

    def _start_publish(self, devices, snapshot=None):
        self.publishing = True
        self.stop_publish = False
        self.btn_publish.config(state=tk.DISABLED)
        self.btn_publish_selected.config(state=tk.DISABLED)
        self.btn_stop_publish.config(state=tk.NORMAL)
        threading.Thread(target=self._publish_thread,
                         args=(devices, snapshot), daemon=True).start()

    def _stop_publishing(self):
        self.stop_publish = True
        self.log("正在停止发布...", "warn")

    # ═══════════════════════════ 定时发布 ═══════════════════════════

    def _schedule_publish(self, mode):
        """设置定时发布。mode: 'all' 或 'selected'"""
        if self.scheduled_timer is not None:
            self.log("已有定时任务在等待，请先取消", "warn")
            return
        if not self.connected:
            self.log("未连接内核", "warn")
            return

        # 解析时间
        time_str = self.entry_sched_time.get().strip()
        try:
            target_ts = time.mktime(time.strptime(time_str, "%Y-%m-%d %H:%M:%S"))
        except Exception:
            messagebox.showerror("时间格式错误",
                "请使用格式: YYYY-MM-DD HH:MM:SS\n例如: 2026-04-19 10:30:00")
            return

        now = time.time()
        if target_ts <= now:
            messagebox.showerror("时间错误", "定时时间必须晚于当前时间")
            return

        # 校验素材
        pub_file = self.entry_pub_file.get().strip()
        pub_folder = self.entry_pub_folder.get().strip()
        if not pub_file and not pub_folder and not self.publish_tasks:
            messagebox.showwarning("提示", "请先选择素材文件、素材文件夹或导入Excel任务")
            return

        # 目标设备
        if mode == "all":
            devices = list(self.devices)
            if not devices:
                self.log("没有在线设备", "warn")
                return
        else:
            devices = self._get_checked_devices()
            if not devices:
                messagebox.showwarning("提示", "请先勾选要发布的设备")
                return

        # ★ 关键：此刻捕获UI快照，定时触发时用这些冻结的值，不再读UI
        snapshot = self._capture_publish_snapshot()

        # 确认（显示快照内容）
        delta = int(target_ts - now)
        h, m, s = delta // 3600, (delta % 3600) // 60, delta % 60
        summary_lines = [
            f"定时时间: {time_str}",
            f"距离现在: {h}小时{m}分{s}秒",
            f"目标设备: {len(devices)} 台",
            f"内容类型: {'图片' if snapshot['pub_type']=='picture' else '视频'}",
        ]
        if snapshot["pub_folder"]:
            summary_lines.append(f"素材文件夹: {os.path.basename(snapshot['pub_folder'])}")
        elif snapshot["pub_file"]:
            summary_lines.append(f"素材: {os.path.basename(snapshot['pub_file'])}")
        if snapshot["pub_title"]:
            summary_lines.append(f"标题: {snapshot['pub_title'][:30]}")
        if snapshot["save_drafts"]:
            summary_lines.append("模式: 存草稿")
        if snapshot["publish_tasks"]:
            summary_lines.append(f"Excel任务: {len(snapshot['publish_tasks'])} 条")
        summary_lines.append("")
        summary_lines.append("★ 所有参数将在此刻冻结")
        summary_lines.append("★ 定时触发时即使您修改了输入框也不会影响")

        if not messagebox.askyesno("确认定时发布", "\n".join(summary_lines), icon="info"):
            return

        # 启动定时线程
        self.scheduled_target = target_ts
        self.scheduled_mode = mode
        self.scheduled_stop = False
        self.btn_sched_all.config(state=tk.DISABLED)
        self.btn_sched_selected.config(state=tk.DISABLED)
        self.btn_sched_cancel.config(state=tk.NORMAL)
        self.entry_sched_time.config(state=tk.DISABLED)

        self.log(f"✓ 定时发布已设定: {time_str}（共 {len(devices)} 台设备，参数已冻结）", "ok")
        self.scheduled_timer = threading.Thread(
            target=self._scheduled_watcher, args=(devices, snapshot), daemon=True)
        self.scheduled_timer.start()

    def _cancel_scheduled_publish(self):
        """取消定时发布"""
        if self.scheduled_timer is None:
            return
        self.scheduled_stop = True
        self.log("定时发布已取消", "warn")
        self._finish_scheduled()

    def _finish_scheduled(self):
        """清理定时状态"""
        self.scheduled_timer = None
        self.scheduled_target = None
        self.scheduled_mode = None
        try:
            self.btn_sched_all.config(state=tk.NORMAL)
            self.btn_sched_selected.config(state=tk.NORMAL)
            self.btn_sched_cancel.config(state=tk.DISABLED)
            self.entry_sched_time.config(state=tk.NORMAL)
            self.lbl_sched_status.config(text="状态: 未设定")
        except Exception:
            pass

    def _scheduled_watcher(self, devices, snapshot):
        """每秒检查是否到点，到点用冻结的 snapshot 触发发布"""
        while not self.scheduled_stop:
            now = time.time()
            remaining = int(self.scheduled_target - now)
            if remaining <= 0:
                break
            # 更新倒计时显示
            h, m, s = remaining // 3600, (remaining % 3600) // 60, remaining % 60
            text = f"状态: 等待中 (还有 {h:02d}:{m:02d}:{s:02d})"
            self.root.after(0, lambda t=text: self.lbl_sched_status.config(text=t))
            time.sleep(1)

        if self.scheduled_stop:
            return

        # 到点触发发布（用冻结的 snapshot）
        self.root.after(0, lambda: self.log("⏰ 定时到达，开始发布（使用冻结参数）...", "ok"))
        self.root.after(0, lambda: self.lbl_sched_status.config(text="状态: 正在发布"))
        self.scheduled_timer = None
        self.scheduled_target = None
        self.scheduled_mode = None
        self.root.after(0, lambda: self._start_publish(devices, snapshot=snapshot))
        self.root.after(0, self._finish_scheduled)

    def _publish_thread(self, devices, snapshot=None):
        """发布主线程。snapshot 为 None 时实时读取UI；否则使用快照（定时发布）"""
        if snapshot is None:
            snapshot = self._capture_publish_snapshot()

        step_delay = snapshot["step_delay"]
        find_timeout = snapshot["find_timeout"]
        dev_delay = snapshot["dev_delay"]

        pub_file = snapshot["pub_file"]
        pub_folder = snapshot["pub_folder"]
        pub_url = snapshot["pub_url"]
        pub_title = snapshot["pub_title"]
        pub_desc = snapshot["pub_desc"]
        pub_type = snapshot["pub_type"]

        # ═══ Excel 任务模式 ═══
        if snapshot.get("publish_tasks"):
            self._publish_excel_tasks(devices, step_delay, find_timeout, dev_delay,
                                       excel_tasks=snapshot["publish_tasks"],
                                       save_drafts=snapshot.get("save_drafts"))
            return

        # ═══ 普通模式 ═══
        total = len(devices)
        self.root.after(0, lambda: self.progress.config(maximum=total, value=0))

        self.root.after(0, lambda: self.log("=" * 50, "info"))
        mode_str = "文件夹模式(按自定义名匹配)" if pub_folder else "统一素材模式"
        self.root.after(0, lambda m=mode_str: self.log(f"开始发布 - 共 {total} 台设备 - {m}", "info"))
        self.root.after(0, lambda: self.log("=" * 50, "info"))

        success_count = 0
        skip_count = 0
        for idx, dev in enumerate(devices):
            if self.stop_publish:
                self.root.after(0, lambda: self.log("用户停止了发布", "warn"))
                break

            dev_id = getattr(dev, 'device_id', '') or getattr(dev, 'deviceid', '')
            dev_custom_name = getattr(dev, 'name', '') or ''
            dev_name = dev_custom_name or getattr(dev, 'user_name', '') or dev_id

            # 文件夹模式：按自定义名查找对应子文件夹
            if pub_folder:
                dev_file = self._get_folder_file_for_device(pub_folder, dev_custom_name)
                if not dev_file:
                    skip_count += 1
                    self.root.after(0, lambda n=dev_name, cn=dev_custom_name, i=idx: self.log(
                        f"\n--- [{i+1}/{total}] 设备: {n} - 跳过(未找到子文件夹 '{cn}') ---", "warn"))
                    continue
                self.root.after(0, lambda n=dev_name, f=dev_file: self.log(
                    f"  匹配文件: {os.path.basename(f)}", "info"))
            else:
                dev_file = pub_file

            self.root.after(0, lambda n=dev_name, i=idx: self.log(
                f"\n--- [{i+1}/{total}] 设备: {n} ---", "info"))
            self.root.after(0, lambda i=idx: self.progress.config(value=i))
            self.root.after(0, lambda n=dev_name: self._set_bottom(f"正在发布: {n}"))

            try:
                ok = self._publish_single_device(
                    dev_id, dev_name,
                    file_path=dev_file,
                    music_url=pub_url,
                    title=pub_title,
                    description=pub_desc,
                    content_type=pub_type,
                    step_delay=step_delay,
                    find_timeout=find_timeout,
                    task_index=idx + 1,
                    save_drafts=snapshot.get("save_drafts")
                )
                if ok:
                    success_count += 1
                    self.root.after(0, lambda n=dev_name: self.log(f"  [OK] {n} 发布成功", "ok"))
                else:
                    self.root.after(0, lambda n=dev_name: self.log(f"  [FAIL] {n} 发布失败", "err"))

            except Exception as e:
                self.root.after(0, lambda n=dev_name, e=e: self.log(f"  [ERROR] {n}: {e}", "err"))

            if idx < total - 1 and not self.stop_publish:
                time.sleep(dev_delay)

        self.root.after(0, lambda: self.progress.config(value=total))
        self.root.after(0, lambda: self.log("=" * 50, "info"))
        skip_msg = f", 跳过 {skip_count}" if skip_count > 0 else ""
        self.root.after(0, lambda s=success_count, t=total, sm=skip_msg: self.log(
            f"发布完成: 成功 {s}/{t}{sm}", "ok" if s == t else "warn"))
        self.root.after(0, lambda: self.log("=" * 50, "info"))
        self.root.after(0, lambda: self._set_bottom("就绪"))
        self.root.after(0, self._finish_publish)

    def _finish_publish(self):
        self.publishing = False
        self.btn_publish.config(state=tk.NORMAL)
        self.btn_publish_selected.config(state=tk.NORMAL)
        self.btn_stop_publish.config(state=tk.DISABLED)

    # ═══════════════════════════ 一键切号 ═══════════════════════════

    def _switch_account_all(self):
        """一键切号 - 所有在线设备切换TikTok账号"""
        if not self.connected:
            self.log("未连接内核", "warn")
            return
        if not self.devices:
            self.log("没有在线设备", "warn")
            return
        n = len(self.devices)
        msg = f"确认要对所有 {n} 台在线设备执行切号吗？\n\n流程：个人页 → 设置 → Switch account → 切换到另一个账号"
        if not messagebox.askyesno("确认切号", msg, icon="warning"):
            return
        self._disable_buttons()
        threading.Thread(target=self._switch_account_thread, args=(self.devices,), daemon=True).start()

    def _switch_account_selected(self):
        """选设备切号 - 只对勾选设备切换TikTok账号"""
        if not self.connected:
            self.log("未连接内核", "warn")
            return
        checked = self._get_checked_devices()
        if not checked:
            messagebox.showwarning("提示", "请先勾选要切号的设备")
            return
        n = len(checked)
        names = "、".join(getattr(d, 'name', '') or getattr(d, 'user_name', '') for d in checked[:5])
        if n > 5:
            names += f" 等{n}台"
        msg = f"确认要对以下 {n} 台设备执行切号吗？\n\n{names}\n\n流程：个人页 → 设置 → Switch account → 切换到另一个账号"
        if not messagebox.askyesno("确认切号", msg, icon="warning"):
            return
        self._disable_buttons()
        threading.Thread(target=self._switch_account_thread, args=(checked,), daemon=True).start()

    def _switch_account_thread(self, devices):
        """切号线程 - 所有设备并发执行"""
        from concurrent.futures import ThreadPoolExecutor, as_completed

        step_delay = float(self.entry_step_delay.get() or "3")
        total = len(devices)
        self.root.after(0, lambda: self.progress.config(maximum=total, value=0))
        self.root.after(0, lambda: self.log("=" * 50, "info"))
        self.root.after(0, lambda t=total: self.log(f"开始并发切号 - 共 {t} 台设备", "info"))
        self.root.after(0, lambda: self.log("=" * 50, "info"))

        success_count = 0
        done_count = 0
        lock = threading.Lock()

        def switch_one(idx, dev):
            nonlocal success_count, done_count
            dev_id = getattr(dev, 'device_id', '') or getattr(dev, 'deviceid', '')
            dev_custom_name = getattr(dev, 'name', '') or ''
            dev_name = dev_custom_name or getattr(dev, 'user_name', '') or dev_id

            self.root.after(0, lambda n=dev_name, i=idx: self.log(
                f"[{i+1}/{total}] {n} 开始切号...", "info"))
            try:
                ok = self._switch_single_device(dev_id, dev_name, step_delay)
                with lock:
                    done_count += 1
                    if ok:
                        success_count += 1
                    self.root.after(0, lambda v=done_count: self.progress.config(value=v))
                if ok:
                    self.root.after(0, lambda n=dev_name: self.log(f"  [OK] {n} 切号成功", "ok"))
                else:
                    self.root.after(0, lambda n=dev_name: self.log(f"  [FAIL] {n} 切号失败", "err"))
            except Exception as e:
                with lock:
                    done_count += 1
                    self.root.after(0, lambda v=done_count: self.progress.config(value=v))
                self.root.after(0, lambda n=dev_name, e=e: self.log(f"  [ERROR] {n}: {e}", "err"))

        with ThreadPoolExecutor(max_workers=total) as executor:
            futures = [executor.submit(switch_one, idx, dev) for idx, dev in enumerate(devices)]
            for f in as_completed(futures):
                pass  # 结果已在 switch_one 中处理

        self.root.after(0, lambda: self.progress.config(value=total))
        self.root.after(0, lambda: self.log("=" * 50, "info"))
        self.root.after(0, lambda s=success_count, t=total: self.log(
            f"切号完成: 成功 {s}/{t}", "ok" if s == t else "warn"))
        self.root.after(0, lambda: self.log("=" * 50, "info"))
        self.root.after(0, lambda: self._set_bottom("就绪"))
        self.root.after(0, self._enable_buttons)

    def _switch_single_device(self, deviceid, dev_name, step_delay=3):
        """单台设备切号流程 (专业版API)"""
        api = self.xp_api

        def log(msg, tag="info"):
            self.root.after(0, lambda: self.log(f"  {msg}", tag))

        def click_at(x, y):
            api.click(deviceid, x, y)
            time.sleep(0.5)

        # Step 1: 回到主屏幕
        log("Step 1: 回到主屏幕...")
        api.home(deviceid)
        time.sleep(2)

        # Step 2: 打开 TikTok
        log("Step 2: 打开 TikTok...")
        r = api.exec_url(deviceid, "tiktok://")
        if r.get("status") != 0:
            log(f"  无法打开 TikTok: {r.get('message','')}", "err")
            return False
        time.sleep(step_delay + 3)

        # Step 2.5: 检测横屏，如果横屏先点返回
        try:
            from io import BytesIO
            from PIL import Image as PILImage
            img_bytes = api.screenshot(deviceid)
            if img_bytes:
                img = PILImage.open(BytesIO(img_bytes)).convert('RGB')
                w, h = img.size
                if w > h:
                    log(f"  检测到横屏 ({w}x{h})，点击返回 (834, 39)", "warn")
                    click_at(834, 39)
                    time.sleep(2)
        except Exception:
            pass

        # Step 3: 点击个人页（Profile）
        log("Step 3: 进入个人页...")
        click_at(390, 1080)
        time.sleep(step_delay)

        # Step 4: 点击右上角三条杠菜单（≡）
        log("Step 4: 点击菜单 ≡ ...")
        pos = self._find_and_click(deviceid, "three", timeout=5)
        if pos:
            log(f"  识图找到 ≡ ({pos[0]}, {pos[1]})", "ok")
        else:
            log("  识图未匹配，使用固定坐标 (466, 85)", "warn")
            click_at(466, 85)
        time.sleep(step_delay)

        # Step 5: 点击 Settings and privacy（识图 + OCR 双重检测）
        log("Step 5: 点击 Settings and privacy...")
        found = False

        # 方式1: 识图（齿轮图标）
        if not found:
            pos = self._find_and_click(deviceid, "settings", timeout=5)
            if pos:
                log(f"  识图找到 settings ({pos[0]}, {pos[1]})", "ok")
                found = True

        # 方式2: OCR 精确匹配 "Settings and privacy"
        if not found:
            matches = api.find_text(deviceid, ["Settings and privacy"])
            if matches:
                r = matches[0]
                cx, cy = r['result'][0], r['result'][1]
                click_at(cx, cy)
                log(f"  OCR 找到 Settings and privacy ({cx}, {cy})", "ok")
                found = True

        # 方式3: OCR 全文扫描找含 "privacy" 的条目
        if not found:
            all_ocr = api.ocr(deviceid)
            for item in all_ocr:
                if 'privacy' in (item.get('txt') or '').lower():
                    cx, cy = item['result'][0], item['result'][1]
                    click_at(cx, cy)
                    log(f"  OCR(全文) 找到 privacy ({cx}, {cy})", "ok")
                    found = True
                    break

        if not found:
            click_at(200, 460)
            log("  使用固定坐标 (200, 460)", "warn")
        time.sleep(step_delay)

        # Step 6: 向下滚动找到 Switch account
        log("Step 6: 向下滚动...")
        for _ in range(5):
            api.swipe(deviceid, "up", length=0.6, sx=200, sy=800)
            time.sleep(0.3)
        time.sleep(1)

        # Step 7: 点击 Switch account（纯 OCR，识别不到用固定坐标）
        log("Step 7: 点击 Switch account...")
        found = False

        # 方式1: OCR 精确匹配 → 多结果取 y 最大
        matches = api.find_text(deviceid, ["Switch account"])
        if matches:
            r = max(matches, key=lambda x: x['result'][1])
            cx, cy = r['result'][0], r['result'][1]
            click_at(cx, cy)
            log(f"  OCR 找到 Switch account ({cx}, {cy})", "ok")
            found = True

        # 方式2: OCR 全文扫描含 switch+account
        if not found:
            all_ocr = api.ocr(deviceid)
            candidates = [r for r in all_ocr
                          if 'switch' in (r.get('txt') or '').lower()
                          and 'account' in (r.get('txt') or '').lower()]
            if candidates:
                r = max(candidates, key=lambda x: x['result'][1])
                cx, cy = r['result'][0], r['result'][1]
                click_at(cx, cy)
                log(f"  OCR(全文) 找到: {r.get('txt','')} ({cx}, {cy})", "ok")
                found = True

        # 方式3: 固定坐标
        if not found:
            log("  OCR 未找到，等待后点击固定坐标...", "warn")
            time.sleep(1.0)
            click_at(175, 810)
            log("  使用固定坐标 (175, 810)", "warn")
        time.sleep(step_delay)

        # Step 8: 找活跃账号(有✓)，点击另一个账号（像素分析）
        log("Step 8: 选择另一个账号...")
        time.sleep(2)

        switched = False

        def _get_img():
            from io import BytesIO
            from PIL import Image as PILImage
            for _ in range(3):
                try:
                    data = api.screenshot(deviceid)
                    if data:
                        return PILImage.open(BytesIO(data)).convert('RGB')
                except Exception:
                    pass
                time.sleep(1)
            return None

        def _find_checkmark_y():
            """找红色✓的y坐标：识图优先，失败用右侧红色像素聚类"""
            # 方式1: 识图
            p = ICONS.get("checkmark", "")
            if p and os.path.exists(p):
                try:
                    result = api.find_image(deviceid, _file_to_base64(p), similarity=0.6)
                    if result:
                        log(f"  识图找到✓ y={result[1]}", "info")
                        return result[1]
                except Exception:
                    pass
            # 方式2: 扫描右侧红色像素
            img = _get_img()
            if not img:
                return None
            w, h = img.size
            red_count = {}
            for y in range(int(h * 0.05), int(h * 0.95)):
                cnt = sum(1 for x in range(int(w * 0.68), w - 2, 2)
                          if img.getpixel((x, y))[0] > 160
                          and img.getpixel((x, y))[0] > img.getpixel((x, y))[1] + 50
                          and img.getpixel((x, y))[0] > img.getpixel((x, y))[2] + 30)
                if cnt > 0:
                    red_count[y] = cnt
            if red_count:
                best = max(red_count, key=red_count.get)
                log(f"  像素扫描找到✓ y={best} (红色像素={red_count[best]})", "info")
                return best
            return None

        def _find_account_rows(ck_y=None):
            """扫描左侧头像区域，若知道ck_y只扫±200px范围"""
            img = _get_img()
            if not img:
                return []
            w, h = img.size
            x0 = max(1, int(w * 0.03))
            x1 = min(w - 1, int(w * 0.22))
            y_min = max(0, ck_y - 200) if ck_y is not None else int(h * 0.03)
            y_max = min(h, ck_y + 200) if ck_y is not None else int(h * 0.97)

            rows_content = []
            for y in range(y_min, y_max, 2):
                cnt = sum(1 for x in range(x0, x1, 3)
                          if not all(v > 215 for v in img.getpixel((x, y))))
                rows_content.append((y, cnt))

            centers = []
            seg_start = seg_end = None
            for y, cnt in rows_content:
                if cnt >= 2:
                    if seg_start is None:
                        seg_start = y
                    seg_end = y
                else:
                    if seg_start is not None:
                        if seg_end - seg_start > 25:
                            centers.append((seg_start + seg_end) // 2)
                        seg_start = seg_end = None
            if seg_start is not None and seg_end - seg_start > 25:
                centers.append((seg_start + seg_end) // 2)

            log(f"  头像扫描找到 {len(centers)} 个账号行: {centers}", "info")
            return centers

        for attempt in range(3):
            try:
                ck_y = _find_checkmark_y()
                row_ys = _find_account_rows(ck_y=ck_y)

                if ck_y is not None and row_ys:
                    others = [y for y in row_ys if 40 < abs(y - ck_y) < 250]
                    if others:
                        target_y = min(others, key=lambda y: abs(y - ck_y))
                        click_at(200, target_y)
                        log(f"  ✓在y={ck_y}，点击非活跃账号 y={target_y}", "ok")
                        switched = True
                        break

                elif ck_y is None and len(row_ys) >= 2:
                    log("  像素未找到✓，尝试OCR辅助...", "warn")
                    ocr_all = api.ocr(deviceid)
                    CK = {"✓", "√", "✔", "☑"}
                    marks = [r for r in ocr_all if (r.get('txt') or '').strip() in CK]
                    if marks:
                        ck_y2 = marks[0]['result'][1]
                        others = [y for y in row_ys if abs(y - ck_y2) > 40]
                        if others:
                            target_y = max(others, key=lambda y: abs(y - ck_y2))
                            click_at(200, target_y)
                            log(f"  OCR✓ y={ck_y2}，点击 y={target_y}", "ok")
                            switched = True
                            break

                log(f"  未能确定目标账号，重试 ({attempt+1}/3)...", "warn")
                time.sleep(2)

            except Exception as e:
                log(f"  账号检测异常: {e}", "warn")
                time.sleep(1)

        if not switched:
            log("  无法识别账号列表，切号失败", "err")
            return False

        time.sleep(step_delay)
        log("切号完成", "ok")
        return True

    # ═══════════════════════════ 一键养号 ═══════════════════════════

    def _nurture_all(self):
        """一键养号 - 所有在线设备"""
        if not self.connected:
            self.log("未连接内核", "warn")
            return
        if not self.devices:
            self.log("没有在线设备", "warn")
            return
        if self.nurturing:
            self.log("养号正在进行中", "warn")
            return
        n = len(self.devices)
        msg = (f"确认要对所有 {n} 台在线设备执行养号吗？\n\n"
               "流程：打开TikTok For You → 无限循环滑视频+点赞\n"
               "点击[停止养号]可随时终止")
        if not messagebox.askyesno("确认养号", msg, icon="warning"):
            return
        self.nurturing = True
        self.stop_nurture = False
        self.btn_nurture.config(state=tk.DISABLED)
        self.btn_stop_nurture.config(state=tk.NORMAL)
        threading.Thread(target=self._nurture_thread, args=(list(self.devices),), daemon=True).start()

    def _stop_nurturing(self):
        self.stop_nurture = True
        self.log("正在停止养号...", "warn")

    def _nurture_thread(self, devices):
        """养号主线程"""
        import random as _random
        from concurrent.futures import ThreadPoolExecutor, as_completed

        api = self.xp_api
        total = len(devices)

        self.root.after(0, lambda: self.log("=" * 50, "info"))
        self.root.after(0, lambda t=total: self.log(f"开始养号 - 共 {t} 台设备", "info"))
        self.root.after(0, lambda: self.log("=" * 50, "info"))

        live_mode = {}  # deviceid → bool（直播模式跟踪）

        # ── Phase 1: 所有设备初始化（并发） ──
        self.root.after(0, lambda: self.log("Phase 1: 初始化所有设备...", "info"))

        def init_device(dev):
            dev_id = dev.deviceid
            dev_name = dev.name or dev.user_name or dev_id
            try:
                # 1. 回到主屏幕
                api.home(dev_id)
                time.sleep(2)
                self.root.after(0, lambda n=dev_name: self.log(
                    f"  {n}: 已回到主屏幕", "ok"))

                # 2. 打开 TikTok
                api.exec_url(dev_id, "tiktok://")
                time.sleep(5)
                self.root.after(0, lambda n=dev_name: self.log(
                    f"  {n}: 已打开 TikTok", "ok"))

                # 3. 检测横屏 → 切换直播模式
                try:
                    from io import BytesIO
                    from PIL import Image as PILImage
                    img_bytes = api.screenshot(dev_id)
                    if img_bytes:
                        img = PILImage.open(BytesIO(img_bytes)).convert('RGB')
                        w, h = img.size
                        if w > h:
                            api.click(dev_id, 71, 387)
                            time.sleep(1.5)
                            live_mode[dev_id] = True
                            self.root.after(0, lambda n=dev_name:
                                self.log(f"  {n}: 检测到横屏，切换直播模式", "warn"))
                except Exception:
                    pass

                # 4. 点击 Home 图标（TikTok底部的For You/Home）
                pos = self._find_and_click(dev_id, "home_icon", timeout=3)
                if pos:
                    self.root.after(0, lambda n=dev_name, p=pos: self.log(
                        f"  {n}: 识图Home ({p[0]},{p[1]})", "ok"))
                else:
                    api.click(dev_id, 41, 830)
                    time.sleep(0.5)
                    self.root.after(0, lambda n=dev_name: self.log(
                        f"  {n}: 固定坐标Home (41,830)", "warn"))
                time.sleep(2)

                # 5. 向右滑到 For You
                for _ in range(3):
                    if self.stop_nurture:
                        return
                    api.swipe(dev_id, "right", length=0.5, sx=100, sy=50)
                    time.sleep(0.5)
                time.sleep(2)
                self.root.after(0, lambda n=dev_name: self.log(
                    f"  {n}: 已滑到 For You", "ok"))
            except Exception as e:
                self.root.after(0, lambda n=dev_name, e=e: self.log(
                    f"  {n}: 初始化异常: {e}", "err"))

        with ThreadPoolExecutor(max_workers=min(total, 10)) as executor:
            futures = [executor.submit(init_device, dev) for dev in devices]
            for f in as_completed(futures):
                pass

        if self.stop_nurture:
            self.root.after(0, lambda: self.log("用户停止了养号", "warn"))
            self._finish_nurture()
            return

        # ── Phase 2: 无限批次循环 ──
        batch_num = 0

        while not self.stop_nurture:
            batch_num += 1
            self.root.after(0, lambda b=batch_num: self.log(
                f"\n--- 养号批次 {b} ---", "info"))

            # 随机分 2-4 组
            shuffled = list(devices)
            _random.shuffle(shuffled)
            num_groups = _random.randint(2, min(4, len(shuffled)))
            groups = []
            base_size = len(shuffled) // num_groups
            remainder = len(shuffled) % num_groups
            idx = 0
            for g in range(num_groups):
                size = base_size + (1 if g < remainder else 0)
                groups.append(shuffled[idx:idx + size])
                idx += size

            self.root.after(0, lambda ng=num_groups, gs=[len(g) for g in groups]:
                self.log(f"  随机分为 {ng} 组: {gs}", "info"))

            for group_idx, group in enumerate(groups):
                if self.stop_nurture:
                    break

                self.root.after(0, lambda gi=group_idx, gl=len(group):
                    self.log(f"  组 {gi+1}: {gl} 台设备", "info"))

                def nurture_one_device(dev, live_mode_dict):
                    dev_id = dev.deviceid
                    dev_name = dev.name or dev.user_name or dev_id

                    # 随机延迟 0-20 秒
                    delay = _random.uniform(0, 20)
                    time.sleep(delay)

                    if self.stop_nurture:
                        return

                    try:
                        # 1. 截图检测横屏
                        try:
                            from io import BytesIO
                            from PIL import Image as PILImage
                            img_bytes = api.screenshot(dev_id)
                            if img_bytes:
                                img = PILImage.open(BytesIO(img_bytes)).convert('RGB')
                                w, h = img.size
                                if w > h:
                                    api.click(dev_id, 71, 387)
                                    time.sleep(1)
                                    live_mode_dict[dev_id] = True
                                    self.root.after(0, lambda n=dev_name:
                                        self.log(f"    {n}: 检测到横屏，切换直播模式", "warn"))
                        except Exception:
                            pass

                        # 2. 从屏幕中心往上滑（下一个视频）
                        api.swipe(dev_id, "up", length=0.5, sx=200, sy=550)
                        time.sleep(1.5)

                        if self.stop_nurture:
                            return

                        # 3. 点赞或双击
                        is_live = live_mode_dict.get(dev_id, False)
                        if is_live:
                            # 直播模式：双击屏幕中心
                            api.click(dev_id, 200, 500)
                            time.sleep(0.15)
                            api.click(dev_id, 200, 500)
                            self.root.after(0, lambda n=dev_name:
                                self.log(f"    {n}: 双击(直播)", "ok"))
                        else:
                            # 视频模式：点赞
                            pos = self._find_and_click(dev_id, "heart", timeout=3)
                            if pos:
                                self.root.after(0, lambda n=dev_name, p=pos:
                                    self.log(f"    {n}: 点赞 ({p[0]},{p[1]})", "ok"))
                            else:
                                api.click(dev_id, 386, 469)
                                time.sleep(0.5)
                                self.root.after(0, lambda n=dev_name:
                                    self.log(f"    {n}: 点赞 (386,469)", "ok"))

                    except Exception as e:
                        self.root.after(0, lambda n=dev_name, e=e:
                            self.log(f"    {n}: 异常: {e}", "err"))

                # 组内并发
                with ThreadPoolExecutor(max_workers=min(len(group), 10)) as executor:
                    futures = [executor.submit(nurture_one_device, dev, live_mode)
                               for dev in group]
                    for f in as_completed(futures):
                        pass

            if self.stop_nurture:
                break

            self.root.after(0, lambda b=batch_num:
                self.log(f"  批次 {b} 完成，等待10秒...", "info"))

            # 等待10秒，期间检查停止标志
            for _ in range(20):
                if self.stop_nurture:
                    break
                time.sleep(0.5)

        self.root.after(0, lambda b=batch_num: self.log(
            f"养号结束，共执行 {b} 个批次", "ok"))
        self.root.after(0, lambda: self.log("=" * 50, "info"))
        self._finish_nurture()

    def _finish_nurture(self):
        self.nurturing = False
        self.root.after(0, lambda: self.btn_nurture.config(state=tk.NORMAL))
        self.root.after(0, lambda: self.btn_stop_nurture.config(state=tk.DISABLED))
        self.root.after(0, lambda: self._set_bottom("就绪"))

    def _publish_excel_tasks(self, devices, step_delay, find_timeout, dev_delay,
                              excel_tasks=None, save_drafts=None):
        """执行 Excel 导入的任务列表。excel_tasks=None时读取self.publish_tasks"""
        tasks = list(excel_tasks if excel_tasks is not None else self.publish_tasks)
        total_tasks = len(tasks)
        if save_drafts is None:
            save_drafts = self.save_drafts_var.get()

        self.root.after(0, lambda: self.progress.config(maximum=total_tasks, value=0))
        self.root.after(0, lambda: self.log("=" * 50, "info"))
        self.root.after(0, lambda t=total_tasks: self.log(f"开始执行 Excel 任务 - 共 {t} 条", "info"))
        self.root.after(0, lambda: self.log("=" * 50, "info"))

        # 建立自定义名 -> 设备对象的映射
        dev_map = {}
        for dev in devices:
            cn = getattr(dev, 'name', '') or ''
            if cn:
                dev_map[cn] = dev
            # 也用设备名做映射
            un = getattr(dev, 'user_name', '') or ''
            if un and un not in dev_map:
                dev_map[un] = dev

        success_count = 0
        for task_idx, task in enumerate(tasks):
            if self.stop_publish:
                self.root.after(0, lambda: self.log("用户停止了发布", "warn"))
                break

            task_devices_str = task.get("devices", "")
            task_file = task.get("file", "")
            task_type = task.get("type", "picture")
            task_url = task.get("url", "")
            task_title = task.get("title", "")
            task_desc = task.get("description", "")

            self.root.after(0, lambda i=task_idx, t=total_tasks, d=task_devices_str: self.log(
                f"\n=== 任务 [{i+1}/{t}] 设备: {d} ===", "info"))
            self.root.after(0, lambda i=task_idx: self.progress.config(value=i))

            # 解析目标设备（逗号分隔的自定义名）
            target_names = [n.strip() for n in task_devices_str.split(",") if n.strip()]
            if not target_names:
                self.root.after(0, lambda: self.log("  跳过：未指定设备", "warn"))
                continue

            task_success = 0
            for dev_idx, target_name in enumerate(target_names):
                if self.stop_publish:
                    break

                dev = dev_map.get(target_name)
                if not dev:
                    self.root.after(0, lambda n=target_name: self.log(
                        f"  设备 '{n}' 未找到，跳过", "warn"))
                    continue

                dev_id = getattr(dev, 'device_id', '') or getattr(dev, 'deviceid', '')
                dev_name = target_name

                self.root.after(0, lambda n=dev_name, i=dev_idx, t=len(target_names): self.log(
                    f"\n--- 设备 [{i+1}/{t}]: {n} ---", "info"))
                self.root.after(0, lambda n=dev_name: self._set_bottom(f"正在发布: {n}"))

                try:
                    ok = self._publish_single_device(
                        dev_id, dev_name,
                        file_path=task_file,
                        music_url=task_url,
                        title=task_title,
                        description=task_desc,
                        content_type=task_type,
                        step_delay=step_delay,
                        find_timeout=find_timeout,
                        task_index=task_idx + 1,
                        save_drafts=save_drafts
                    )
                    if ok:
                        task_success += 1
                        self.root.after(0, lambda n=dev_name: self.log(
                            f"  [OK] {n} 发布成功", "ok"))
                    else:
                        self.root.after(0, lambda n=dev_name: self.log(
                            f"  [FAIL] {n} 发布失败", "err"))
                except Exception as e:
                    self.root.after(0, lambda n=dev_name, e=e: self.log(
                        f"  [ERROR] {n}: {e}", "err"))

                if dev_idx < len(target_names) - 1 and not self.stop_publish:
                    time.sleep(dev_delay)

            if task_success == len(target_names):
                success_count += 1

        self.root.after(0, lambda: self.progress.config(value=total_tasks))
        self.root.after(0, lambda: self.log("=" * 50, "info"))
        self.root.after(0, lambda s=success_count, t=total_tasks: self.log(
            f"Excel 任务完成: {s}/{t} 条任务全部成功", "ok" if s == t else "warn"))
        self.root.after(0, lambda: self.log("=" * 50, "info"))
        self.root.after(0, lambda: self._set_bottom("就绪"))
        self.root.after(0, self._finish_publish)

    # ───────── 辅助：找图标并点击 ─────────
    def _find_and_click(self, deviceid, icon_keys, timeout=15, offset_x=0, offset_y=0):
        """通过图标匹配查找并点击，返回坐标或 None"""
        if isinstance(icon_keys, str):
            icon_keys = [icon_keys]
        img_b64_list = []
        for k in icon_keys:
            p = ICONS.get(k, k)
            if os.path.exists(p):
                try:
                    img_b64_list.append(_file_to_base64(p))
                except Exception as e:
                    self.root.after(0, lambda e=e, k=k: self.log(f"  [图标加载失败] {k}: {e}", "err"))
        if not img_b64_list:
            return None
        start = time.time()
        while time.time() - start < timeout:
            if self.stop_publish:
                return None
            try:
                result = self.xp_api.find_image_ex(deviceid, img_b64_list, similarity=0.7)
                if result:
                    x = result[0] + offset_x
                    y = result[1] + offset_y
                    self.xp_api.click(deviceid, x, y)
                    time.sleep(0.5)
                    return (x, y)
            except Exception:
                pass
            time.sleep(1)
        return None

    # ───────── 辅助：截图后按网格位置点击第N张图 ─────────
    def _click_nth_thumbnail(self, deviceid, n, log_fn):
        """点击相册网格中第 n 张图片 (n从1开始，从左到右从上到下)"""
        from io import BytesIO
        from PIL import Image as PILImage
        try:
            img_bytes = self.xp_api.screenshot(deviceid)
            if img_bytes:
                img = PILImage.open(BytesIO(img_bytes)).convert('RGB')
                w, h = img.size
            else:
                w, h = 390, 844
        except Exception:
            w, h = 390, 844

        grid_top = int(h * 0.14)
        col_count = 4
        col_w = w // col_count
        row_h = col_w  # 正方形缩略图

        col = (n - 1) % col_count
        row = (n - 1) // col_count
        x = col * col_w + col_w // 2
        y = grid_top + row * row_h + row_h // 2

        log_fn(f"  点击第{n}张 col={col} row={row} → ({x}, {y})", "info")
        self.xp_api.click(deviceid, x, y)
        time.sleep(0.5)
        return (x, y)

    # ───────── 辅助：用素材文件做模板找图（保留，不再用于发布主流程）─────────
    def _find_image_in_gallery(self, deviceid, file_path, timeout=15):
        """用上传的素材文件作为模板，在屏幕上找到对应缩略图"""
        if not file_path or not os.path.exists(file_path):
            return None
        try:
            img_b64 = _file_to_base64(file_path)
        except Exception:
            return None
        start = time.time()
        while time.time() - start < timeout:
            if self.stop_publish:
                return None
            try:
                result = self.xp_api.find_image(deviceid, img_b64, similarity=0.7)
                if result:
                    return (result[0], result[1])
            except Exception:
                pass
            time.sleep(1)
        return None

    def _find_last_thumbnail(self, deviceid, file_path):
        """截图分析相册网格，从右下角往左上角扫描，找到最后一个有实际图片内容的缩略图"""
        try:
            from io import BytesIO
            from PIL import Image as PILImage

            # 截图重试3次
            img_bytes = None
            for retry in range(3):
                try:
                    img_bytes = self.xp_api.screenshot(deviceid)
                    if img_bytes:
                        break
                except Exception:
                    pass
                time.sleep(1)

            if not img_bytes:
                self.root.after(0, lambda: self.log("  [截图] 重试3次仍返回空数据", "err"))
                return None

            img = PILImage.open(BytesIO(img_bytes)).convert('RGB')
            w, h = img.size
            self.root.after(0, lambda w=w, h=h: self.log(f"  [截图] 尺寸: {w}x{h}", "info"))

            # 相册网格参数
            grid_top = int(h * 0.14)     # 跳过顶部标签栏
            grid_bottom = int(h * 0.85)  # Next 按钮上方
            col_count = 4
            col_w = w // col_count
            row_h = col_w  # 缩略图通常是正方形

            # 计算网格行数和每个格子的中心坐标
            cells = []
            y = grid_top + row_h // 2
            while y < grid_bottom:
                for col in range(col_count):
                    cx = col * col_w + col_w // 2
                    cells.append((cx, y))
                y += row_h

            self.root.after(0, lambda n=len(cells): self.log(f"  [截图] 共 {n} 个格子", "info"))

            # 从最后一格往前扫描，检测是否是真实缩略图
            for cx, cy in reversed(cells):
                if self._cell_has_image(img, cx, cy, col_w):
                    return (cx, cy)

            self.root.after(0, lambda: self.log("  [截图] 所有格子都未检测到图片", "warn"))
            return None
        except Exception as e:
            self.root.after(0, lambda e=e: self.log(f"  [截图分析异常] {e}", "err"))
            return None

    def _cell_has_image(self, img, cx, cy, cell_size):
        """检测某个网格位置是否有真实图片内容"""
        w, h = img.size
        sample_r = cell_size // 3  # 采样半径（大一些）
        step = max(sample_r // 4, 2)  # 采样步长
        pixels = []
        for dx in range(-sample_r, sample_r + 1, step):
            for dy in range(-sample_r, sample_r + 1, step):
                px = min(max(cx + dx, 0), w - 1)
                py = min(max(cy + dy, 0), h - 1)
                pixels.append(img.getpixel((px, py)))

        if len(pixels) < 9:
            return False

        rs = [p[0] for p in pixels]
        gs = [p[1] for p in pixels]
        bs = [p[2] for p in pixels]

        def std(vals):
            avg = sum(vals) / len(vals)
            return (sum((v - avg) ** 2 for v in vals) / len(vals)) ** 0.5

        color_std = std(rs) + std(gs) + std(bs)

        # 计算平均亮度
        avg_brightness = sum(r + g + b for r, g, b in zip(rs, gs, bs)) / (len(pixels) * 3)

        # 空白格特征：
        # 1. TikTok 暗色主题：接近纯黑 (brightness < 30) 且无变化 (std < 10)
        # 2. 亮色主题：接近纯白 (brightness > 230) 且无变化 (std < 10)
        # 3. 灰色空格：纯灰 (brightness 100-200) 且无变化 (std < 10)
        # 真实图片：有颜色变化 (std > 15) 或中间亮度有内容
        if color_std < 10:
            return False  # 完全单色 = 空白格
        if color_std < 18 and (avg_brightness < 30 or avg_brightness > 230):
            return False  # 接近纯黑/纯白 + 很少变化 = 空白格
        return True

    # ───────── 单台设备发布流程 ─────────
    def _publish_single_device(self, deviceid, dev_name, file_path="", music_url="",
                                title="", description="", content_type="picture",
                                step_delay=3, find_timeout=15, task_index=1,
                                save_drafts=None):
        """
        TikTok 发布流程 (专业版API):
        图片：HOME → TikTok → 音乐URL/+ → 相册入口 → 直接点第task_index张图 → Next → title → desc → Post/Drafts
        视频：HOME → TikTok → + → 相册入口 → 直接点第task_index张图 → Next → desc → Post/Drafts
        """
        api = self.xp_api
        # save_drafts 为 None 时实时读取；否则使用传入的快照值
        if save_drafts is None:
            save_drafts = self.save_drafts_var.get()

        def log(msg, tag="info"):
            self.root.after(0, lambda: self.log(f"  {msg}", tag))

        def click_at(x, y):
            api.click(deviceid, x, y)
            time.sleep(0.5)

        def input_text(text):
            """通过剪贴板粘贴文字"""
            if not text:
                return
            api.clipboard_set(deviceid, text)
            time.sleep(0.5)
            api.paste(deviceid)

        # ═══ Step 1: 上传素材到相册 ═══
        if file_path and os.path.exists(file_path):
            log("Step 1: 上传素材到相册...")
            r = api.album_upload(deviceid, [file_path])
            if r.get("status") == 0:
                log(f"  已上传: {os.path.basename(file_path)}", "ok")
            else:
                log(f"  上传失败: {r.get('message','')}", "err")
                return False
            time.sleep(step_delay)
        else:
            log("Step 1: 跳过上传（使用已有素材）")

        # ═══ Step 2: 回到主屏幕 ═══
        log("Step 2: 回到主屏幕...")
        api.home(deviceid)
        time.sleep(2)

        # ═══ Step 3: 打开 TikTok ═══
        log("Step 3: 打开 TikTok...")
        r = api.exec_url(deviceid, "tiktok://")
        if r.get("status") == 0:
            log("  已通过 URL scheme 打开 TikTok", "ok")
        else:
            pos = self._find_and_click(deviceid, "tiktok", timeout=find_timeout)
            if pos:
                log(f"  已通过图标打开 TikTok ({pos[0]}, {pos[1]})", "ok")
            else:
                log("  无法打开 TikTok", "err")
                return False
        time.sleep(step_delay + 3)

        # ═══════════ 图片发布流程 ═══════════
        if content_type == "picture":
            # Step 4: 打开音乐URL 或 点击 +
            if music_url:
                log("Step 4: 打开音乐 URL...")
                r = api.exec_url(deviceid, music_url)
                if r.get("status") != 0:
                    log(f"  URL 打开失败: {r.get('message','')}", "err")
                    return False
                log(f"  URL 打开成功", "ok")
                time.sleep(10)

                # Step 5: 点击 Use Sound
                log("Step 5: 点击 Use Sound...")
                pos = self._find_and_click(deviceid, "usesound", timeout=find_timeout)
                if not pos:
                    log("  图标未匹配，使用固定坐标 (365, 1007)", "warn")
                    click_at(365, 1007)
                    pos = (365, 1007)
                log(f"  Use Sound ({pos[0]}, {pos[1]})", "ok")
                time.sleep(step_delay)
            else:
                log("Step 4: 点击 + 进入创建页...")
                pos = self._find_and_click(deviceid, ["plus_black", "plus_white"], timeout=find_timeout)
                if not pos:
                    log("  未找到 + 按钮", "err")
                    return False
                log(f"  点击 + ({pos[0]}, {pos[1]})", "ok")
                time.sleep(step_delay)

            # Step 6: 进入相册（直接点击相册入口，跳过相册导航）
            log("Step 6: 进入相册...")
            pos = self._find_and_click(deviceid, "record", timeout=find_timeout)
            if pos:
                log(f"  找到录制界面标记 ({pos[0]}, {pos[1]})", "ok")
            else:
                log("  record图标未匹配，继续...", "warn")
            # 点击相册入口按钮（左下角）
            api.click(deviceid, 30, 1010)
            time.sleep(step_delay)

            # Step 7: 选图 - 先用上传的素材模板匹配，选不中就点最后一张真实缩略图
            log(f"Step 7: 选择第 {task_index} 张图片...")
            selected = False
            if file_path and os.path.exists(file_path):
                match_pos = self._find_image_in_gallery(deviceid, file_path, timeout=5)
                if match_pos:
                    api.click(deviceid, match_pos[0], match_pos[1])
                    time.sleep(0.5)
                    log(f"  模板匹配选中 ({match_pos[0]}, {match_pos[1]})", "ok")
                    selected = True
            if not selected:
                last_pos = self._find_last_thumbnail(deviceid, file_path)
                if last_pos:
                    api.click(deviceid, last_pos[0], last_pos[1])
                    time.sleep(0.5)
                    log(f"  选不中，点击最后一张缩略图 ({last_pos[0]}, {last_pos[1]})", "warn")
                    selected = True
            if not selected:
                log("  未找到缩略图，回退固定网格位置", "warn")
                col = (task_index - 1) % 4
                row = (task_index - 1) // 4
                x = col * 97 + 48
                y = 118 + row * 97 + 48
                click_at(x, y)
                log(f"  固定坐标点击第{task_index}张 ({x}, {y})", "warn")
            time.sleep(step_delay)

            # Step 8: 点击 Next（连点两下）
            log("Step 8: 点击 Next...")
            pos = self._find_and_click(deviceid, "next", timeout=5)
            if pos:
                log(f"  Next ({pos[0]}, {pos[1]})", "ok")
                time.sleep(0.5)
                click_at(pos[0], pos[1])
            else:
                log("  图标未匹配，使用固定坐标 (376, 1010) x2", "warn")
                click_at(376, 1010)
                time.sleep(0.5)
                click_at(376, 1010)
            time.sleep(step_delay)

            # Step 9: 输入标题
            if title:
                log("Step 9: 输入标题...")
                pos = self._find_and_click(deviceid, ["title", "title2"], timeout=5, offset_y=50)
                if pos:
                    log(f"  找到标题位置 ({pos[0]}, {pos[1]})", "ok")
                else:
                    log("  图标未匹配，使用固定坐标 (65, 280)", "warn")
                    click_at(65, 280)
                time.sleep(1)
                input_text(title)
                log("  标题已输入", "ok")
                time.sleep(step_delay)
            else:
                log("Step 9: 跳过标题")

            # Step 10: 输入描述标签
            if description:
                log("Step 10: 输入描述标签...")
                pos = self._find_and_click(deviceid, ["desc", "long_desc"], timeout=5, offset_y=50)
                if pos:
                    log(f"  找到描述位置 ({pos[0]}, {pos[1]})", "ok")
                else:
                    log("  图标未匹配，使用固定坐标 (70, 285)", "warn")
                    click_at(70, 285)
                time.sleep(1)
                input_text(description)
                log("  描述标签已输入", "ok")
                time.sleep(step_delay)
            else:
                log("Step 10: 跳过描述标签")

            # Step 11: 点击 Post 或 Drafts
            if save_drafts:
                log("Step 11: 存入草稿...")
                matches = api.find_text(deviceid, ["Drafts"])
                if matches:
                    r = matches[0]
                    click_at(r['result'][0], r['result'][1])
                    log(f"  OCR找到 Drafts ({r['result'][0]}, {r['result'][1]})", "ok")
                else:
                    log("  OCR未找到Drafts，使用固定坐标 (270, 80)", "warn")
                    click_at(270, 80)
            else:
                log("Step 11: 点击 Post...")
                pos = self._find_and_click(deviceid, ["post", "old_post"], timeout=5)
                if pos:
                    log(f"  Post ({pos[0]}, {pos[1]})", "ok")
                else:
                    log("  图标未匹配，使用固定坐标 (450, 80)", "warn")
                    click_at(450, 80)
            time.sleep(step_delay + 2)

        # ═══════════ 视频发布流程 ═══════════
        else:
            # Step 4: 点击 +
            log("Step 4: 点击 + ...")
            pos = self._find_and_click(deviceid, ["plus_black", "plus_white"], timeout=find_timeout)
            if not pos:
                log("  未找到 + 按钮", "err")
                return False
            log(f"  + ({pos[0]}, {pos[1]})", "ok")
            time.sleep(step_delay)

            # Step 5: 进入相册（直接点击入口，跳过相册导航）
            log("Step 5: 进入相册...")
            pos = self._find_and_click(deviceid, "record2", timeout=find_timeout)
            if pos:
                log(f"  找到录制界面标记 ({pos[0]}, {pos[1]})", "ok")
            else:
                log("  record2图标未匹配，继续...", "warn")
            api.click(deviceid, 30, 1010)
            time.sleep(step_delay)

            # Step 6: 选视频 - 先用上传的素材模板匹配，选不中就点最后一张真实缩略图
            log(f"Step 6: 选择第 {task_index} 张视频...")
            selected = False
            if file_path and os.path.exists(file_path):
                match_pos = self._find_image_in_gallery(deviceid, file_path, timeout=5)
                if match_pos:
                    api.click(deviceid, match_pos[0], match_pos[1])
                    time.sleep(0.5)
                    log(f"  模板匹配选中 ({match_pos[0]}, {match_pos[1]})", "ok")
                    selected = True
            if not selected:
                last_pos = self._find_last_thumbnail(deviceid, file_path)
                if last_pos:
                    api.click(deviceid, last_pos[0], last_pos[1])
                    time.sleep(0.5)
                    log(f"  选不中，点击最后一张缩略图 ({last_pos[0]}, {last_pos[1]})", "warn")
                    selected = True
            if not selected:
                log("  未找到缩略图，回退固定网格位置", "warn")
                col = (task_index - 1) % 4
                row = (task_index - 1) // 4
                x = col * 97 + 48
                y = 118 + row * 97 + 48
                click_at(x, y)
                log(f"  固定坐标点击第{task_index}张 ({x}, {y})", "warn")
            time.sleep(step_delay)

            # Step 7: 点击 Next（连点两下）
            log("Step 7: 点击 Next...")
            pos = self._find_and_click(deviceid, "next", timeout=5)
            if pos:
                log(f"  Next ({pos[0]}, {pos[1]})", "ok")
                time.sleep(0.5)
                click_at(pos[0], pos[1])
            else:
                log("  图标未匹配，使用固定坐标 (376, 1010) x2", "warn")
                click_at(376, 1010)
                time.sleep(0.5)
                click_at(376, 1010)
            time.sleep(step_delay)

            # Step 8: 输入描述
            if description:
                log("Step 8: 输入描述...")
                pos = self._find_and_click(deviceid, ["desc", "long_desc"], timeout=5, offset_y=50)
                if pos:
                    log(f"  找到描述位置 ({pos[0]}, {pos[1]})", "ok")
                else:
                    log("  图标未匹配，使用固定坐标 (65, 280)", "warn")
                    click_at(65, 280)
                time.sleep(1)
                input_text(description)
                log("  描述已输入", "ok")
                time.sleep(step_delay)
            else:
                log("Step 8: 跳过描述")

            # Step 9: 点击 Post 或 Drafts
            if save_drafts:
                log("Step 9: 存入草稿...")
                matches = api.find_text(deviceid, ["Drafts"])
                if matches:
                    r = matches[0]
                    click_at(r['result'][0], r['result'][1])
                    log(f"  OCR找到 Drafts ({r['result'][0]}, {r['result'][1]})", "ok")
                else:
                    log("  OCR未找到Drafts，使用固定坐标 (270, 80)", "warn")
                    click_at(270, 80)
            else:
                log("Step 9: 点击 Post...")
                pos = self._find_and_click(deviceid, ["post", "old_post"], timeout=5)
                if pos:
                    log(f"  Post ({pos[0]}, {pos[1]})", "ok")
                else:
                    log("  图标未匹配，使用固定坐标 (450, 80)", "warn")
                    click_at(450, 80)
            time.sleep(step_delay + 2)

        log("发布流程完成!", "ok")
        return True

    # ═══════════════════════════ 文件选择 ═══════════════════════════

    def _pick_files(self):
        filetypes = [
            ("图片和视频", "*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.heic;*.heif;*.mp4;*.mov;*.avi;*.mkv;*.m4v;*.webp"),
            ("图片", "*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.heic;*.heif;*.webp"),
            ("视频", "*.mp4;*.mov;*.avi;*.mkv;*.m4v"),
            ("所有文件", "*.*")
        ]
        paths = filedialog.askopenfilenames(title="选择要上传的文件", filetypes=filetypes)
        if paths:
            for p in paths:
                abp = os.path.abspath(p)
                if abp not in self.selected_files:
                    self.selected_files.append(abp)
            self._update_file_list()

    def _pick_folder(self):
        folder = filedialog.askdirectory(title="选择包含图片/视频的文件夹")
        if folder:
            count = 0
            for root, dirs, files in os.walk(folder):
                for f in files:
                    ext = os.path.splitext(f)[1].lower()
                    if ext in MEDIA_EXTS:
                        full = os.path.abspath(os.path.join(root, f))
                        if full not in self.selected_files:
                            self.selected_files.append(full)
                            count += 1
            self.log(f"从文件夹添加了 {count} 个媒体文件（含子文件夹）", "info")
            self._update_file_list()

    def _clear_files(self):
        self.selected_files.clear()
        self._update_file_list()

    def _update_file_list(self):
        self.listbox_files.delete(0, tk.END)
        for f in self.selected_files:
            size_kb = os.path.getsize(f) / 1024
            name = os.path.basename(f)
            if size_kb >= 1024:
                self.listbox_files.insert(tk.END, f"  {name}  ({size_kb/1024:.1f} MB)")
            else:
                self.listbox_files.insert(tk.END, f"  {name}  ({size_kb:.1f} KB)")
        self.lbl_file_count.config(text=f"已选: {len(self.selected_files)} 个文件")

    # ═══════════════════════════ 上传功能 ═══════════════════════════

    def _do_upload(self):
        upload_folder = self.entry_upload_folder.get().strip()
        if not self.selected_files and not upload_folder:
            self.log("请先选择要上传的文件或素材文件夹", "warn")
            return
        devs = self._get_checked_devices()
        if not devs:
            self.log("请先勾选要上传的设备", "warn")
            return
        self._disable_buttons()
        threading.Thread(target=self._upload_thread, args=(devs,), daemon=True).start()

    def _do_one_click_upload(self):
        upload_folder = self.entry_upload_folder.get().strip()
        if not self.selected_files and not upload_folder:
            self.log("请先选择要上传的文件或素材文件夹", "warn")
            return
        if not self.devices:
            self.log("没有已连接的设备", "warn")
            return
        self._disable_buttons()
        if upload_folder:
            self.log(f"一键上传(文件夹模式): 按自定义名匹配 -> {len(self.devices)} 台设备", "info")
        else:
            self.log(f"一键上传: {len(self.selected_files)} 个文件 -> {len(self.devices)} 台设备", "info")
        threading.Thread(target=self._upload_thread, args=(self.devices,), daemon=True).start()

    def _upload_thread(self, devices):
        mode = self.upload_mode.get()
        target_path = self.entry_album.get().strip()
        upload_folder = self.entry_upload_folder.get().strip()
        files = list(self.selected_files)

        # 文件夹模式优先
        use_folder_mode = bool(upload_folder)

        if not use_folder_mode:
            total_tasks = len(devices) * len(files)
        else:
            total_tasks = len(devices)
        done_count = 0
        skip_count = 0
        self.root.after(0, lambda: self.progress.config(maximum=total_tasks, value=0))

        for dev_idx, dev in enumerate(devices):
            dev_id = getattr(dev, 'device_id', '') or getattr(dev, 'deviceid', '')
            dev_custom_name = getattr(dev, 'name', '') or ''
            dev_name = dev_custom_name or getattr(dev, 'user_name', '') or dev_id

            # 文件夹模式：按自定义名查找子文件夹
            if use_folder_mode:
                dev_files = self._get_folder_files_for_device(upload_folder, dev_custom_name)
                if not dev_files:
                    skip_count += 1
                    done_count += 1
                    self.root.after(0, lambda v=done_count: self.progress.config(value=v))
                    self.root.after(0, lambda n=dev_name, cn=dev_custom_name, i=dev_idx: self.log(
                        f"-- 设备 [{i+1}/{len(devices)}]: {n} - 跳过(未找到子文件夹 '{cn}') --", "warn"))
                    continue
                files = dev_files

            self.root.after(0, lambda n=dev_name, i=dev_idx, fc=len(files): self.log(
                f"-- 设备 [{i+1}/{len(devices)}]: {n} ({fc}个文件) --", "info"))

            # 逐个文件上传（视频上传较慢，不批量以避免整批超时）
            success_count = 0
            for i, fp in enumerate(files):
                fname = os.path.basename(fp)
                self.root.after(0, lambda lbl=f"{i+1}/{len(files)} {fname}":
                    self._set_bottom(f"上传中... {lbl}"))
                self.root.after(0, lambda n=fname, idx=i+1, t=len(files):
                    self.log(f"  [{idx}/{t}] 上传中: {n}", "info"))
                try:
                    if mode == "album":
                        r = self.xp_api.album_upload(dev_id, [fp], album_name=target_path)
                    else:
                        r = self.xp_api.file_upload(dev_id, [fp], target_path=target_path or "/")
                    if r.get("status") == 0:
                        self.root.after(0, lambda n=fname: self.log(f"  [OK] {n}", "ok"))
                        success_count += 1
                    else:
                        self.root.after(0, lambda n=fname, msg=r.get('message',''):
                            self.log(f"  [FAIL] {n}: {msg}", "err"))
                except Exception as e:
                    self.root.after(0, lambda n=fname, e=e:
                        self.log(f"  [异常] {n}: {e}", "err"))
                done_count += 1
                self.root.after(0, lambda v=done_count: self.progress.config(value=v))

            if use_folder_mode:
                done_count += 1
                self.root.after(0, lambda v=done_count: self.progress.config(value=v))
            self.root.after(0, lambda n=dev_name, s=success_count, t=len(files):
                self.log(f"  {n}: 成功 {s}/{t}", "ok" if s == t else "warn"))

        skip_msg = f"，跳过 {skip_count}" if skip_count > 0 else ""
        self.root.after(0, lambda sm=skip_msg: self.log(f"上传任务完成{sm}", "ok"))
        self.root.after(0, lambda: self._set_bottom("就绪"))
        self.root.after(0, self._enable_buttons)

    # ═══════════════════════════ 下载功能 ═══════════════════════════

    def _do_list_remote(self):
        dev = self._get_first_checked_device()
        if not dev:
            return
        threading.Thread(target=self._list_remote_thread, args=(dev,), daemon=True).start()

    def _list_remote_thread(self, dev):
        dev_id = getattr(dev, 'device_id', '') or getattr(dev, 'deviceid', '')
        mode = self.download_mode.get()
        path = self.entry_dl_path.get().strip()
        num_str = self.entry_dl_num.get().strip()
        num = int(num_str) if num_str.isdigit() else 20
        self.root.after(0, lambda: self.log("正在获取文件列表...", "info"))
        try:
            if mode == "album":
                file_list = self.xp_api.album_list(dev_id, num=num)
                self.album_files = file_list if file_list else []
            else:
                file_list = self.xp_api.file_list(dev_id, path=path or "/")
                self.phone_files = file_list if file_list else []
            files = file_list if file_list else []
            self.root.after(0, lambda: self._update_remote_tree(files, mode))
            self.root.after(0, lambda: self.log(f"获取到 {len(files)} 个文件", "ok"))
        except Exception as e:
            self.root.after(0, lambda: self.log(f"获取文件列表失败: {e}", "err"))

    def _update_remote_tree(self, files, mode):
        self.tree_remote.delete(*self.tree_remote.get_children())
        for f in files:
            name = f.get('name', '') if isinstance(f, dict) else getattr(f, 'name', '')
            ext = f.get('ext', '') if isinstance(f, dict) else getattr(f, 'ext', '')
            size = f.get('size', '') if isinstance(f, dict) else getattr(f, 'size', '')
            t = f.get('time', f.get('create_time', '')) if isinstance(f, dict) else getattr(f, 'create_time', '')
            album = f.get('album_name', '') if isinstance(f, dict) else getattr(f, 'album_name', '')
            self.tree_remote.insert("", tk.END, values=(
                name, ext, size, t, album if mode == "album" else ''))

    def _do_download(self):
        dev = self._get_first_checked_device()
        if not dev:
            return
        selected = self.tree_remote.selection()
        if not selected:
            self.log("请先选择要下载的文件", "warn")
            return
        self._disable_buttons()
        threading.Thread(target=self._download_thread, args=(dev, selected), daemon=True).start()

    def _download_thread(self, dev, selected_iids):
        dev_id = getattr(dev, 'device_id', '') or getattr(dev, 'deviceid', '')
        mode = self.download_mode.get()
        path = self.entry_dl_path.get().strip() or "/"
        try:
            if mode == "album":
                source = self.album_files
                indices = [self.tree_remote.index(iid) for iid in selected_iids]
                file_names = []
                for idx in indices:
                    if idx < len(source):
                        f = source[idx]
                        n = f.get('name', '') if isinstance(f, dict) else getattr(f, 'name', '')
                        e = f.get('ext', '') if isinstance(f, dict) else getattr(f, 'ext', '')
                        file_names.append(f"{n}.{e}" if e else n)
                self.root.after(0, lambda: self.log(f"正在下载 {len(file_names)} 个文件...", "info"))
                r = self.xp_api.album_down(dev_id, file_names)
            else:
                source = self.phone_files
                indices = [self.tree_remote.index(iid) for iid in selected_iids]
                file_paths = []
                for idx in indices:
                    if idx < len(source):
                        f = source[idx]
                        n = f.get('name', '') if isinstance(f, dict) else getattr(f, 'name', '')
                        e = f.get('ext', '') if isinstance(f, dict) else getattr(f, 'ext', '')
                        fp = f"{path.rstrip('/')}/{n}.{e}" if e else f"{path.rstrip('/')}/{n}"
                        file_paths.append(fp)
                self.root.after(0, lambda: self.log(f"正在下载 {len(file_paths)} 个文件...", "info"))
                r = self.xp_api.file_down(dev_id, file_paths)

            if r.get("status") == 0:
                self.root.after(0, lambda: self.log("下载完成！文件在 iMouse\\Shortcut\\Media 文件夹", "ok"))
            else:
                self.root.after(0, lambda msg=r.get('message',''): self.log(f"下载失败: {msg}", "err"))
        except Exception as e:
            self.root.after(0, lambda: self.log(f"下载异常: {e}", "err"))
        self.root.after(0, self._enable_buttons)

    # ═══════════════════════════ 浏览功能 ═══════════════════════════

    def _do_browse(self, mode):
        dev = self._get_first_checked_device()
        if not dev:
            return
        threading.Thread(target=self._browse_thread, args=(dev, mode), daemon=True).start()

    def _browse_thread(self, dev, mode):
        dev_id = getattr(dev, 'device_id', '') or getattr(dev, 'deviceid', '')
        path = self.entry_browse_path.get().strip()
        num_str = self.entry_browse_num.get().strip()
        num = int(num_str) if num_str.isdigit() else 20
        try:
            if mode == "album":
                file_list = self.xp_api.album_list(dev_id, num=num)
            else:
                file_list = self.xp_api.file_list(dev_id, path=path or "/")
            files = file_list if file_list else []
            self.root.after(0, lambda: self._update_browse_tree(files, mode))
            self.root.after(0, lambda: self.log(f"浏览: 获取到 {len(files)} 个文件", "ok"))
        except Exception as e:
            self.root.after(0, lambda: self.log(f"浏览失败: {e}", "err"))

    def _update_browse_tree(self, files, mode):
        self.tree_browse.delete(*self.tree_browse.get_children())
        for f in files:
            name = f.get('name', '') if isinstance(f, dict) else getattr(f, 'name', '')
            ext = f.get('ext', '') if isinstance(f, dict) else getattr(f, 'ext', '')
            size = f.get('size', '') if isinstance(f, dict) else getattr(f, 'size', '')
            t = f.get('time', f.get('create_time', '')) if isinstance(f, dict) else getattr(f, 'create_time', '')
            album = f.get('album_name', '') if isinstance(f, dict) else getattr(f, 'album_name', '')
            self.tree_browse.insert("", tk.END, values=(
                name, ext, size, t, album if mode == "album" else ''))

    # ═══════════════════════════ 辅助方法 ═══════════════════════════

    def log(self, msg, tag=None):
        timestamp = time.strftime("%H:%M:%S")
        self.txt_log.insert(tk.END, f"[{timestamp}] ", "info")
        self.txt_log.insert(tk.END, f"{msg}\n", tag)
        self.txt_log.see(tk.END)

    def _set_status(self, text, color):
        self.lbl_status.config(text=text, fg=color)

    def _set_bottom(self, text):
        self.lbl_bottom.config(text=text)

    def _disable_buttons(self):
        for btn in (self.btn_upload, self.btn_one_click, self.btn_download):
            btn.config(state=tk.DISABLED)

    def _enable_buttons(self):
        for btn in (self.btn_upload, self.btn_one_click, self.btn_download):
            btn.config(state=tk.NORMAL)

    # ═══════════════════════════ 自动更新 ═══════════════════════════

    def _check_update(self):
        """点击检查更新按钮"""
        self.log(f"正在检查更新... (当前版本 v{CURRENT_VERSION})", "info")
        threading.Thread(target=self._check_update_thread, daemon=True).start()

    def _check_update_thread(self):
        try:
            # 加 ?_={ts} 绕过 raw.githubusercontent.com 的 max-age=300 CDN 缓存
            resp = requests.get(
                UPDATE_VERSION_URL,
                params={"_": int(time.time())},
                headers={"Cache-Control": "no-cache", "Pragma": "no-cache"},
                timeout=15,
            )
            if resp.status_code != 200:
                self.root.after(0, lambda: self.log(
                    f"检查更新失败: HTTP {resp.status_code}", "err"))
                return
            latest = resp.text.strip()
            self.root.after(0, lambda v=latest: self.log(
                f"最新版本: v{v}", "info"))

            if self._version_tuple(latest) <= self._version_tuple(CURRENT_VERSION):
                self.root.after(0, lambda: self.log("已是最新版本", "ok"))
                self.root.after(0, lambda: messagebox.showinfo(
                    "检查更新", f"当前已是最新版本 v{CURRENT_VERSION}"))
                return

            # 有新版本，弹窗确认
            def prompt():
                if messagebox.askyesno("发现新版本",
                        f"当前版本: v{CURRENT_VERSION}\n最新版本: v{latest}\n\n是否立即更新？\n\n"
                        "更新后需要重启工具。", icon="info"):
                    threading.Thread(target=self._do_update, args=(latest,), daemon=True).start()
            self.root.after(0, prompt)
        except Exception as e:
            self.root.after(0, lambda e=e: self.log(f"检查更新异常: {e}", "err"))

    def _version_tuple(self, v):
        try:
            return tuple(int(x) for x in str(v).split("."))
        except Exception:
            return (0,)

    def _do_update(self, new_version):
        """下载并替换文件"""
        import io
        import zipfile as _zf
        try:
            self.root.after(0, lambda: self.log("正在下载更新包...", "info"))
            resp = requests.get(
                UPDATE_ZIP_URL,
                params={"_": int(time.time())},
                headers={"Cache-Control": "no-cache", "Pragma": "no-cache"},
                timeout=120,
                stream=True,
            )
            if resp.status_code != 200:
                self.root.after(0, lambda: self.log(
                    f"下载失败: HTTP {resp.status_code}", "err"))
                return
            buf = io.BytesIO()
            total = 0
            last_log = 0
            for chunk in resp.iter_content(chunk_size=64 * 1024):
                if chunk:
                    buf.write(chunk)
                    total += len(chunk)
                    if total - last_log >= 200 * 1024:
                        last_log = total
                        self.root.after(0, lambda s=total: self.log(
                            f"  已下载 {s/1024:.1f} KB...", "info"))
            self.root.after(0, lambda s=total: self.log(
                f"下载完成，共 {s/1024:.1f} KB", "ok"))

            buf.seek(0)
            # 验证 zip 文件完整性
            try:
                _test = _zf.ZipFile(buf)
                _test.testzip()
            except _zf.BadZipFile:
                self.root.after(0, lambda: self.log("更新包损坏或不是有效的zip", "err"))
                return
            buf.seek(0)
            self.root.after(0, lambda: self.log("正在解压并替换文件...", "info"))

            # 不应更新的文件/文件夹（跳过）
            skip_patterns = {'packages', '__pycache__', 'logs'}
            updated_count = 0

            with _zf.ZipFile(buf) as zf:
                # GitHub zip 解压后外层有一个文件夹 qihao---xp-main/
                names = zf.namelist()
                if not names:
                    self.root.after(0, lambda: self.log("更新包为空", "err"))
                    return
                root_prefix = names[0].split("/")[0] + "/"

                for name in names:
                    if name.endswith("/") or not name.startswith(root_prefix):
                        continue
                    rel_path = name[len(root_prefix):]
                    if not rel_path:
                        continue
                    # 跳过不需要更新的
                    first_part = rel_path.split("/")[0].split("\\")[0]
                    if first_part in skip_patterns:
                        continue

                    # 目标文件路径（覆盖本地）
                    target = os.path.join(SCRIPT_DIR, rel_path.replace("/", os.sep))
                    target_dir = os.path.dirname(target)
                    if target_dir and not os.path.isdir(target_dir):
                        os.makedirs(target_dir, exist_ok=True)

                    with zf.open(name) as src:
                        try:
                            with open(target, "wb") as dst:
                                dst.write(src.read())
                            updated_count += 1
                        except PermissionError:
                            # 当前文件被占用（如 transfer_gui.py），先写 .new 然后启动脚本替换
                            with open(target + ".new", "wb") as dst:
                                dst.write(src.read())

            self.root.after(0, lambda c=updated_count: self.log(
                f"更新完成，共替换 {c} 个文件", "ok"))

            # 写一个 apply_update.bat 用于替换被占用的 .new 文件并重启
            self._write_apply_update_script()

            self.root.after(0, lambda v=new_version: messagebox.showinfo(
                "更新完成",
                f"已更新到 v{v}\n\n"
                "请关闭工具后重新启动 start.bat 以应用更新。"))
        except Exception as e:
            self.root.after(0, lambda e=e: self.log(f"更新失败: {e}", "err"))

    def _write_apply_update_script(self):
        """生成一个批处理脚本，在下次启动前把 .new 文件覆盖原文件"""
        script_path = os.path.join(SCRIPT_DIR, "_apply_update.bat")
        try:
            with open(script_path, "w", encoding="gbk", errors="ignore") as f:
                f.write("@echo off\r\n")
                f.write("cd /d \"%~dp0\"\r\n")
                f.write("for /r %%f in (*.new) do (\r\n")
                f.write("    move /y \"%%f\" \"%%~dpnf\" >nul\r\n")
                f.write(")\r\n")
        except Exception:
            pass

    def on_close(self):
        if self.xp_api:
            try:
                self.xp_api.stop()
            except:
                pass
        self.root.destroy()


def main():
    root = tk.Tk()

    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Treeview", font=("Microsoft YaHei UI", 11), rowheight=30)
    style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 13, "bold"))
    style.configure("TNotebook.Tab", font=("Microsoft YaHei UI", 12), padding=[18, 6])

    app = TransferApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
