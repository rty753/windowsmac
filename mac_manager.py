#!/usr/bin/env python3
"""
Windows MAC 地址管理工具
通过注册表修改 MAC 地址，支持手动/自动随机更换、恢复出厂、系统托盘常驻。
需要管理员权限运行。
"""

import ctypes
import json
import logging
import os
import random
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk

# ---------- 第三方依赖（系统托盘） ----------
try:
    import PIL.Image
    import PIL.ImageDraw
    import pystray

    HAS_TRAY = True
except ImportError:
    HAS_TRAY = False

# ---------- Windows 专用 ----------
if sys.platform == "win32":
    import winreg
else:
    print("本工具仅支持 Windows 系统。")
    sys.exit(1)

# ---------- 常量 ----------
REG_CLASS_ROOT = r"SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}"
LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mac_change_log.txt")
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mac_manager_config.json")

# ---------- 日志 ----------
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    encoding="utf-8",
)
logger = logging.getLogger("mac_manager")

# ============================================================
#  UAC 提权
# ============================================================

def is_admin() -> bool:
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def run_as_admin():
    """以管理员身份重新启动自身。"""
    params = " ".join(f'"{a}"' for a in sys.argv)
    executable = sys.executable
    # 如果是 pyinstaller 打包后的 exe，sys.executable 就是 exe 本身
    if getattr(sys, "frozen", False):
        executable = sys.executable
        params = " ".join(f'"{a}"' for a in sys.argv[1:])
    else:
        params = f'"{sys.argv[0]}"'
        if len(sys.argv) > 1:
            params += " " + " ".join(f'"{a}"' for a in sys.argv[1:])

    ret = ctypes.windll.shell32.ShellExecuteW(
        None, "runas", executable, params, None, 1
    )
    if ret <= 32:
        messagebox.showerror("错误", "无法获取管理员权限，程序将退出。")
    sys.exit(0)


# ============================================================
#  MAC 地址工具函数
# ============================================================

def normalize_mac(mac: str) -> str:
    """去掉分隔符，返回 12 位大写十六进制字符串。"""
    mac = re.sub(r"[^0-9a-fA-F]", "", mac)
    return mac.upper()


def format_mac(mac: str, sep: str = "-") -> str:
    """将 12 位字符串格式化为 XX-XX-XX-XX-XX-XX。"""
    mac = normalize_mac(mac)
    return sep.join(mac[i : i + 2] for i in range(0, 12, 2))


def is_valid_mac(mac: str) -> bool:
    mac = normalize_mac(mac)
    return bool(re.fullmatch(r"[0-9A-F]{12}", mac))


def generate_random_mac() -> str:
    """
    生成随机 MAC 地址。
    第一字节：bit0=0（单播），bit1=1（本地管理）→ 即 x2, x6, xA, xE 结尾。
    """
    first_byte = random.randint(0, 255)
    first_byte &= 0xFE  # 清除 bit0 → 单播
    first_byte |= 0x02  # 设置 bit1 → 本地管理
    rest = [random.randint(0, 255) for _ in range(5)]
    return "".join(f"{b:02X}" for b in [first_byte] + rest)


# ============================================================
#  网络适配器信息
# ============================================================

class AdapterInfo:
    """封装单个适配器的信息。"""

    def __init__(self, name: str, reg_subkey: str, current_mac: str,
                 original_mac: str, description: str, is_active: bool):
        self.name = name
        self.reg_subkey = reg_subkey  # e.g. "0001"
        self.current_mac = current_mac
        self.original_mac = original_mac
        self.description = description
        self.is_active = is_active

    def __str__(self):
        status = " [活跃]" if self.is_active else ""
        return f"{self.name}{status}"


def get_active_macs() -> dict[str, str]:
    """通过 getmac 获取活跃连接的 MAC 映射 {适配器名: MAC}。"""
    result: dict[str, str] = {}
    try:
        out = subprocess.check_output(
            ["getmac", "/v", "/fo", "csv", "/nh"],
            creationflags=subprocess.CREATE_NO_WINDOW,
            text=True,
            encoding="gbk",
            errors="replace",
            timeout=10,
        )
        for line in out.strip().splitlines():
            parts = line.split('","')
            if len(parts) >= 3:
                conn_name = parts[0].strip('"')
                mac_val = parts[2].strip('"').replace("-", "").upper()
                status = parts[3].strip('"') if len(parts) >= 4 else ""
                if re.fullmatch(r"[0-9A-F]{12}", mac_val):
                    result[conn_name] = mac_val
    except Exception:
        pass
    return result


def get_adapters_from_netsh() -> dict[str, str]:
    """通过 netsh 获取接口名列表及状态。"""
    adapters: dict[str, str] = {}
    try:
        out = subprocess.check_output(
            ["netsh", "interface", "show", "interface"],
            creationflags=subprocess.CREATE_NO_WINDOW,
            text=True,
            encoding="gbk",
            errors="replace",
            timeout=10,
        )
        for line in out.strip().splitlines()[3:]:
            parts = line.split()
            if len(parts) >= 4:
                state = parts[0]
                name = " ".join(parts[3:])
                adapters[name] = state
    except Exception:
        pass
    return adapters


def enum_registry_adapters() -> list[AdapterInfo]:
    """枚举注册表中所有网络适配器。"""
    adapters: list[AdapterInfo] = []
    active_macs = get_active_macs()
    netsh_adapters = get_adapters_from_netsh()

    try:
        class_key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE, REG_CLASS_ROOT, 0, winreg.KEY_READ
        )
    except OSError:
        return adapters

    idx = 0
    while True:
        try:
            subkey_name = winreg.EnumKey(class_key, idx)
        except OSError:
            break
        idx += 1

        # 只处理 4 位数字子键
        if not re.fullmatch(r"\d{4}", subkey_name):
            continue

        try:
            sub = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                f"{REG_CLASS_ROOT}\\{subkey_name}",
                0,
                winreg.KEY_READ,
            )
        except OSError:
            continue

        try:
            # 检查是否为网络适配器 (*MediaType 或 ComponentId 含 "net")
            try:
                driver_desc = winreg.QueryValueEx(sub, "DriverDesc")[0]
            except OSError:
                winreg.CloseKey(sub)
                continue

            # 获取 NetCfgInstanceId（对应 WMI 的 GUID）
            try:
                instance_id = winreg.QueryValueEx(sub, "NetCfgInstanceId")[0]
            except OSError:
                instance_id = ""

            # 获取适配器连接名
            adapter_name = _get_adapter_name(instance_id) or driver_desc

            # 当前 MAC（注册表中设置的）
            try:
                net_addr = winreg.QueryValueEx(sub, "NetworkAddress")[0]
                current_mac = normalize_mac(net_addr)
            except OSError:
                current_mac = ""

            # 原始 MAC（从 WMI 或硬件获取）
            original_mac = _get_original_mac(instance_id)

            # 如果没有通过注册表设置过 MAC，当前 MAC 就是原始 MAC
            if not current_mac and original_mac:
                current_mac = original_mac

            if not current_mac:
                # 尝试从活跃连接中获取
                for conn_name, mac_val in active_macs.items():
                    if conn_name == adapter_name:
                        current_mac = mac_val
                        break

            if not original_mac:
                original_mac = current_mac

            # 判断是否活跃
            is_active = adapter_name in active_macs or netsh_adapters.get(adapter_name) == "已连接"

            # 过滤掉非物理适配器（简单启发式）
            skip_keywords = ["Virtual", "Miniport", "Filter", "WAN", "Bluetooth",
                             "Debug", "Kernel", "Teredo", "ISATAP", "6to4",
                             "Microsoft Wi-Fi Direct", "Microsoft Hosted"]
            if any(kw.lower() in driver_desc.lower() for kw in skip_keywords):
                winreg.CloseKey(sub)
                continue

            adapters.append(
                AdapterInfo(
                    name=adapter_name,
                    reg_subkey=subkey_name,
                    current_mac=current_mac,
                    original_mac=original_mac,
                    description=driver_desc,
                    is_active=is_active,
                )
            )
        finally:
            winreg.CloseKey(sub)

    winreg.CloseKey(class_key)
    return adapters


def _get_adapter_name(instance_id: str) -> str:
    """通过 NetCfgInstanceId 查找连接名称。"""
    if not instance_id:
        return ""
    reg_path = (
        r"SYSTEM\CurrentControlSet\Control\Network"
        r"\{4D36E972-E325-11CE-BFC1-08002BE10318}"
        f"\\{instance_id}\\Connection"
    )
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, winreg.KEY_READ)
        name = winreg.QueryValueEx(key, "Name")[0]
        winreg.CloseKey(key)
        return name
    except OSError:
        return ""


def _get_original_mac(instance_id: str) -> str:
    """通过 WMI 获取原始硬件 MAC。"""
    if not instance_id:
        return ""
    try:
        out = subprocess.check_output(
            [
                "wmic",
                "nic",
                "where",
                f"GUID='{instance_id}'",
                "get",
                "MACAddress",
                "/value",
            ],
            creationflags=subprocess.CREATE_NO_WINDOW,
            text=True,
            encoding="gbk",
            errors="replace",
            timeout=10,
        )
        for line in out.splitlines():
            if "MACAddress=" in line:
                mac_str = line.split("=", 1)[1].strip()
                if mac_str:
                    return normalize_mac(mac_str)
    except Exception:
        pass
    return ""


# ============================================================
#  MAC 地址修改核心
# ============================================================

def set_mac_registry(adapter: AdapterInfo, new_mac: str) -> bool:
    """在注册表中写入 NetworkAddress。"""
    new_mac = normalize_mac(new_mac)
    if not is_valid_mac(new_mac):
        return False
    reg_path = f"{REG_CLASS_ROOT}\\{adapter.reg_subkey}"
    try:
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE, reg_path, 0, winreg.KEY_SET_VALUE
        )
        winreg.SetValueEx(key, "NetworkAddress", 0, winreg.REG_SZ, new_mac)
        winreg.CloseKey(key)
        return True
    except OSError as e:
        logger.error(f"写注册表失败: {e}")
        return False


def delete_mac_registry(adapter: AdapterInfo) -> bool:
    """删除注册表中的 NetworkAddress 键，恢复出厂 MAC。"""
    reg_path = f"{REG_CLASS_ROOT}\\{adapter.reg_subkey}"
    try:
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE, reg_path, 0, winreg.KEY_SET_VALUE
        )
        try:
            winreg.DeleteValue(key, "NetworkAddress")
        except FileNotFoundError:
            pass  # 已经没有这个值
        winreg.CloseKey(key)
        return True
    except OSError as e:
        logger.error(f"删除注册表键失败: {e}")
        return False


def restart_adapter(adapter_name: str) -> bool:
    """禁用再启用适配器使 MAC 更改生效。"""
    try:
        subprocess.check_call(
            ["netsh", "interface", "set", "interface", adapter_name, "disabled"],
            creationflags=subprocess.CREATE_NO_WINDOW,
            timeout=15,
        )
        time.sleep(2)
        subprocess.check_call(
            ["netsh", "interface", "set", "interface", adapter_name, "enabled"],
            creationflags=subprocess.CREATE_NO_WINDOW,
            timeout=15,
        )
        time.sleep(3)
        return True
    except Exception as e:
        logger.error(f"重启适配器 {adapter_name} 失败: {e}")
        return False


def apply_mac_change(adapter: AdapterInfo, new_mac: str) -> tuple[bool, str]:
    """完整流程：写注册表 → 重启适配器。返回 (成功?, 消息)。"""
    old_mac = adapter.current_mac
    new_mac = normalize_mac(new_mac)

    if not set_mac_registry(adapter, new_mac):
        return False, "写入注册表失败。请确认以管理员权限运行。"

    if not restart_adapter(adapter.name):
        return False, "注册表已更新，但重启适配器失败。\n请手动禁用再启用网络适配器。"

    logger.info(f"适配器={adapter.name}  旧MAC={format_mac(old_mac)}  新MAC={format_mac(new_mac)}")
    adapter.current_mac = new_mac
    return True, f"MAC 地址已更改为 {format_mac(new_mac)}"


def restore_original_mac(adapter: AdapterInfo) -> tuple[bool, str]:
    """恢复出厂 MAC。"""
    if not adapter.original_mac:
        return False, "无法获取原始 MAC 地址。"

    old_mac = adapter.current_mac

    if not delete_mac_registry(adapter):
        return False, "删除注册表键失败。"

    if not restart_adapter(adapter.name):
        return False, "注册表已更新，但重启适配器失败。\n请手动禁用再启用网络适配器。"

    logger.info(
        f"恢复出厂  适配器={adapter.name}  旧MAC={format_mac(old_mac)}  "
        f"原始MAC={format_mac(adapter.original_mac)}"
    )
    adapter.current_mac = adapter.original_mac
    return True, f"已恢复出厂 MAC: {format_mac(adapter.original_mac)}"


# ============================================================
#  配置持久化
# ============================================================

def load_config() -> dict:
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_config(cfg: dict):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error(f"保存配置失败: {e}")


# ============================================================
#  系统托盘
# ============================================================

def create_tray_icon(app: "MacManagerApp"):
    """创建系统托盘图标。"""
    if not HAS_TRAY:
        return None

    def _create_image():
        img = PIL.Image.new("RGB", (64, 64), color=(34, 139, 34))
        draw = PIL.ImageDraw.Draw(img)
        draw.text((12, 18), "MAC", fill="white")
        return img

    def on_show(icon, item):
        app.root.after(0, app.show_window)

    def on_quit(icon, item):
        app.root.after(0, app.quit_app)

    menu = pystray.Menu(
        pystray.MenuItem("显示窗口", on_show, default=True),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("退出", on_quit),
    )

    icon = pystray.Icon("mac_manager", _create_image(), "MAC 地址管理器", menu)
    return icon


# ============================================================
#  主界面
# ============================================================

class MacManagerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Windows MAC 地址管理工具")
        self.root.geometry("780x620")
        self.root.resizable(True, True)
        self.root.minsize(700, 550)

        # 样式
        style = ttk.Style()
        style.theme_use("clam")

        self.adapters: list[AdapterInfo] = []
        self.selected_adapter: AdapterInfo | None = None
        self.auto_thread: threading.Thread | None = None
        self.auto_running = False
        self.auto_interval_hours = 1.0
        self.tray_icon = None

        self._build_ui()
        self.refresh_adapters()
        self._load_auto_config()

        # 系统托盘
        if HAS_TRAY:
            self.tray_icon = create_tray_icon(self)
            threading.Thread(target=self._run_tray, daemon=True).start()

        # 窗口关闭 → 最小化到托盘
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------- UI 构建 ----------

    def _build_ui(self):
        # ---- 顶部：适配器列表 ----
        frame_top = ttk.LabelFrame(self.root, text="网络适配器列表", padding=8)
        frame_top.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 5))

        columns = ("name", "mac", "original", "status")
        self.tree = ttk.Treeview(frame_top, columns=columns, show="headings", height=6)
        self.tree.heading("name", text="适配器名称")
        self.tree.heading("mac", text="当前 MAC")
        self.tree.heading("original", text="原始 MAC")
        self.tree.heading("status", text="状态")
        self.tree.column("name", width=220)
        self.tree.column("mac", width=160)
        self.tree.column("original", width=160)
        self.tree.column("status", width=80)

        scrollbar = ttk.Scrollbar(frame_top, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        btn_refresh = ttk.Button(self.root, text="刷新适配器列表", command=self.refresh_adapters)
        btn_refresh.pack(pady=(0, 5))

        # ---- 中部：修改 MAC ----
        frame_mid = ttk.LabelFrame(self.root, text="修改 MAC 地址", padding=8)
        frame_mid.pack(fill=tk.X, padx=10, pady=5)

        row1 = ttk.Frame(frame_mid)
        row1.pack(fill=tk.X, pady=2)
        ttk.Label(row1, text="选择适配器:").pack(side=tk.LEFT)
        self.combo_adapter = ttk.Combobox(row1, state="readonly", width=40)
        self.combo_adapter.pack(side=tk.LEFT, padx=5)
        self.combo_adapter.bind("<<ComboboxSelected>>", self._on_combo_select)

        row2 = ttk.Frame(frame_mid)
        row2.pack(fill=tk.X, pady=2)
        ttk.Label(row2, text="新 MAC 地址:").pack(side=tk.LEFT)
        self.entry_mac = ttk.Entry(row2, width=22)
        self.entry_mac.pack(side=tk.LEFT, padx=5)

        self.lbl_mac_status = ttk.Label(row2, text="", foreground="gray")
        self.lbl_mac_status.pack(side=tk.LEFT, padx=5)

        self.entry_mac.bind("<KeyRelease>", self._validate_mac_input)

        row3 = ttk.Frame(frame_mid)
        row3.pack(fill=tk.X, pady=5)
        ttk.Button(row3, text="随机生成 MAC", command=self._gen_random).pack(side=tk.LEFT, padx=5)
        ttk.Button(row3, text="应用修改", command=self._apply_mac).pack(side=tk.LEFT, padx=5)
        ttk.Button(row3, text="恢复出厂 MAC", command=self._restore_mac).pack(side=tk.LEFT, padx=5)

        # ---- 下部：定时更换 ----
        frame_bot = ttk.LabelFrame(self.root, text="定时自动随机更换", padding=8)
        frame_bot.pack(fill=tk.X, padx=10, pady=5)

        row4 = ttk.Frame(frame_bot)
        row4.pack(fill=tk.X, pady=2)
        ttk.Label(row4, text="更换间隔（小时）:").pack(side=tk.LEFT)
        self.combo_interval = ttk.Combobox(
            row4, state="readonly", values=["0.5", "1", "2", "6", "12", "24"], width=6
        )
        self.combo_interval.set("1")
        self.combo_interval.pack(side=tk.LEFT, padx=5)

        self.btn_auto_start = ttk.Button(row4, text="启动自动更换", command=self._toggle_auto)
        self.btn_auto_start.pack(side=tk.LEFT, padx=10)

        self.lbl_auto_status = ttk.Label(row4, text="状态: 未启动", foreground="gray")
        self.lbl_auto_status.pack(side=tk.LEFT, padx=10)

        # ---- 日志区域 ----
        frame_log = ttk.LabelFrame(self.root, text="操作日志", padding=8)
        frame_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=(5, 10))

        self.text_log = tk.Text(frame_log, height=6, state=tk.DISABLED, wrap=tk.WORD,
                                font=("Consolas", 9))
        log_scroll = ttk.Scrollbar(frame_log, orient=tk.VERTICAL, command=self.text_log.yview)
        self.text_log.configure(yscrollcommand=log_scroll.set)
        self.text_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    # ---------- 适配器操作 ----------

    def refresh_adapters(self):
        self.adapters = enum_registry_adapters()
        # 更新 Treeview
        self.tree.delete(*self.tree.get_children())
        for a in self.adapters:
            status = "活跃" if a.is_active else "断开"
            self.tree.insert(
                "", tk.END,
                values=(a.name, format_mac(a.current_mac), format_mac(a.original_mac), status),
            )
        # 更新下拉框
        names = [str(a) for a in self.adapters]
        self.combo_adapter["values"] = names
        if names:
            self.combo_adapter.current(0)
            self.selected_adapter = self.adapters[0]
        self._log("适配器列表已刷新")

    def _on_tree_select(self, event):
        sel = self.tree.selection()
        if sel:
            idx = self.tree.index(sel[0])
            if 0 <= idx < len(self.adapters):
                self.selected_adapter = self.adapters[idx]
                self.combo_adapter.current(idx)

    def _on_combo_select(self, event):
        idx = self.combo_adapter.current()
        if 0 <= idx < len(self.adapters):
            self.selected_adapter = self.adapters[idx]

    # ---------- MAC 修改 ----------

    def _validate_mac_input(self, event=None):
        val = self.entry_mac.get()
        mac = normalize_mac(val)
        if len(mac) < 12:
            self.lbl_mac_status.config(text=f"已输入 {len(mac)}/12 位", foreground="gray")
        elif is_valid_mac(mac):
            self.lbl_mac_status.config(text="✓ 格式正确", foreground="green")
        else:
            self.lbl_mac_status.config(text="✗ 格式错误", foreground="red")

    def _gen_random(self):
        mac = generate_random_mac()
        self.entry_mac.delete(0, tk.END)
        self.entry_mac.insert(0, format_mac(mac))
        self._validate_mac_input()

    def _apply_mac(self):
        if not self.selected_adapter:
            messagebox.showwarning("提示", "请先选择一个适配器。")
            return
        mac = normalize_mac(self.entry_mac.get())
        if not is_valid_mac(mac):
            messagebox.showwarning("提示", "请输入有效的 MAC 地址（12 位十六进制）。")
            return

        self._set_busy(True)

        def _do():
            ok, msg = apply_mac_change(self.selected_adapter, mac)
            self.root.after(0, lambda: self._apply_done(ok, msg))

        threading.Thread(target=_do, daemon=True).start()

    def _apply_done(self, ok: bool, msg: str):
        self._set_busy(False)
        if ok:
            self._log(msg)
            self.refresh_adapters()
            messagebox.showinfo("成功", msg)
        else:
            self._log(f"失败: {msg}")
            messagebox.showerror("失败", msg)

    def _restore_mac(self):
        if not self.selected_adapter:
            messagebox.showwarning("提示", "请先选择一个适配器。")
            return
        if not messagebox.askyesno("确认", f"确定要将 {self.selected_adapter.name} 恢复出厂 MAC 吗？"):
            return

        self._set_busy(True)

        def _do():
            ok, msg = restore_original_mac(self.selected_adapter)
            self.root.after(0, lambda: self._apply_done(ok, msg))

        threading.Thread(target=_do, daemon=True).start()

    def _set_busy(self, busy: bool):
        state = "disabled" if busy else "normal"
        for w in self.root.winfo_children():
            self._set_widget_state(w, state)

    def _set_widget_state(self, widget, state):
        try:
            widget.configure(state=state)
        except tk.TclError:
            pass
        for child in widget.winfo_children():
            self._set_widget_state(child, state)

    # ---------- 定时自动更换 ----------

    def _toggle_auto(self):
        if self.auto_running:
            self._stop_auto()
        else:
            self._start_auto()

    def _start_auto(self):
        if not self.selected_adapter:
            messagebox.showwarning("提示", "请先选择一个适配器。")
            return

        try:
            self.auto_interval_hours = float(self.combo_interval.get())
        except ValueError:
            self.auto_interval_hours = 1.0

        self.auto_running = True
        self.btn_auto_start.config(text="停止自动更换")
        self.lbl_auto_status.config(
            text=f"状态: 运行中（每 {self.auto_interval_hours} 小时）",
            foreground="green",
        )
        self._update_tray_tooltip()
        self._save_auto_config()

        self.auto_thread = threading.Thread(target=self._auto_worker, daemon=True)
        self.auto_thread.start()
        self._log(f"自动更换已启动，间隔 {self.auto_interval_hours} 小时")

    def _stop_auto(self):
        self.auto_running = False
        self.btn_auto_start.config(text="启动自动更换")
        self.lbl_auto_status.config(text="状态: 已停止", foreground="gray")
        self._update_tray_tooltip()
        self._save_auto_config()
        self._log("自动更换已停止")

    def _auto_worker(self):
        interval_sec = self.auto_interval_hours * 3600
        while self.auto_running:
            adapter = self.selected_adapter
            if adapter:
                new_mac = generate_random_mac()
                ok, msg = apply_mac_change(adapter, new_mac)
                ts = datetime.now().strftime("%H:%M:%S")
                if ok:
                    self.root.after(0, lambda m=msg: self._log(f"[自动] {m}"))
                    self.root.after(0, self.refresh_adapters)
                else:
                    self.root.after(0, lambda m=msg: self._log(f"[自动] 失败: {m}"))

            # 分段等待，方便快速响应停止
            waited = 0
            while waited < interval_sec and self.auto_running:
                time.sleep(1)
                waited += 1

    def _save_auto_config(self):
        cfg = load_config()
        cfg["auto_running"] = self.auto_running
        cfg["auto_interval"] = self.auto_interval_hours
        if self.selected_adapter:
            cfg["auto_adapter"] = self.selected_adapter.name
        save_config(cfg)

    def _load_auto_config(self):
        cfg = load_config()
        if cfg.get("auto_running"):
            interval = cfg.get("auto_interval", 1.0)
            adapter_name = cfg.get("auto_adapter", "")
            # 找到对应适配器
            for i, a in enumerate(self.adapters):
                if a.name == adapter_name:
                    self.selected_adapter = a
                    self.combo_adapter.current(i)
                    break
            self.combo_interval.set(str(interval))
            self._start_auto()

    # ---------- 托盘 ----------

    def _run_tray(self):
        if self.tray_icon:
            self.tray_icon.run()

    def _update_tray_tooltip(self):
        if self.tray_icon:
            status = "自动更换运行中" if self.auto_running else "MAC 地址管理器"
            self.tray_icon.title = status

    def show_window(self):
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def on_close(self):
        if HAS_TRAY and self.tray_icon:
            self.root.withdraw()
            self._log("窗口已最小化到系统托盘")
        else:
            self.quit_app()

    def quit_app(self):
        self.auto_running = False
        if self.tray_icon:
            try:
                self.tray_icon.stop()
            except Exception:
                pass
        self.root.destroy()

    # ---------- 日志 ----------

    def _log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}\n"
        self.text_log.config(state=tk.NORMAL)
        self.text_log.insert(tk.END, line)
        self.text_log.see(tk.END)
        self.text_log.config(state=tk.DISABLED)

    # ---------- 运行 ----------

    def run(self):
        self.root.mainloop()


# ============================================================
#  入口
# ============================================================

def main():
    if not is_admin():
        run_as_admin()
        return

    app = MacManagerApp()
    app.run()


if __name__ == "__main__":
    main()
