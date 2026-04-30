import os
import sys
import time
import json
import threading
from turtle import onclick
import winreg
import ctypes
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime
import pystray


# 检查管理员权限（兼容Windows XP到Win11）
def is_admin():
    try:
        # 直接使用ctypes.windll.shell32，ctypes会自动根据Python解释器位数加载相应的库
        shell32 = ctypes.windll.shell32
        
        # Windows XP及以上通用的管理员权限检查
        result = shell32.IsUserAnAdmin()
        return result != 0
    except Exception as e:
        try:
            # 尝试打开一个需要管理员权限的注册表键
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                                r"SOFTWARE\Microsoft\Windows\CurrentVersion",
                                0, winreg.KEY_READ)
            winreg.CloseKey(key)
            return True
        except:
            return False

# 请求管理员权限（兼容Windows XP到Win11，智能选择API）
def run_as_admin():
    if not is_admin():
        try:
            # 检查Windows版本，智能选择API
            # 获取Windows版本信息
            class OSVERSIONINFOEXW(ctypes.Structure):
                _fields_ = [
                    ("dwOSVersionInfoSize", ctypes.c_ulong),
                    ("dwMajorVersion", ctypes.c_ulong),
                    ("dwMinorVersion", ctypes.c_ulong),
                    ("dwBuildNumber", ctypes.c_ulong),
                    ("dwPlatformId", ctypes.c_ulong),
                    ("szCSDVersion", ctypes.c_wchar * 128),
                    ("wServicePackMajor", ctypes.c_ushort),
                    ("wServicePackMinor", ctypes.c_ushort),
                    ("wSuiteMask", ctypes.c_ushort),
                    ("wProductType", ctypes.c_byte),
                    ("wReserved", ctypes.c_byte)
                ]
            
            os_version = OSVERSIONINFOEXW()
            os_version.dwOSVersionInfoSize = ctypes.sizeof(OSVERSIONINFOEXW)
            
            # 使用GetVersionExW获取Windows版本
            if ctypes.windll.kernel32.GetVersionExW(ctypes.byref(os_version)):
                # Windows 10及以上版本（10.0+）使用ShellExecuteW
                if os_version.dwMajorVersion >= 10:
                    ctypes.windll.shell32.ShellExecuteW(
                        None, "runas", sys.executable, " ".join(sys.argv), None, 1
                    )
                else:
                    # Windows XP到Windows 8.1使用ShellExecuteA
                    ctypes.windll.shell32.ShellExecuteA(
                        None, "runas", sys.executable.encode('utf-8'), 
                        " ".join(sys.argv).encode('utf-8'), None, 1
                    )
            else:
                # 如果获取版本失败，尝试使用ShellExecuteW
                try:
                    ctypes.windll.shell32.ShellExecuteW(
                        None, "runas", sys.executable, " ".join(sys.argv), None, 1
                    )
                except:
                    # 最后尝试ShellExecuteA
                    ctypes.windll.shell32.ShellExecuteA(
                        None, "runas", sys.executable.encode('utf-8'), 
                        " ".join(sys.argv).encode('utf-8'), None, 1
                    )
        except Exception as e:
            # 显示错误信息，使用MessageBoxW兼容现代系统
            try:
                message = "需要管理员权限才能运行此程序！\n请右键点击程序，选择'以管理员身份运行'"
                ctypes.windll.user32.MessageBoxW(None, message, "权限错误", 0x10)
            except:
                # 兼容旧系统
                ctypes.windll.user32.MessageBoxA(None, message.encode('gbk'), 
                                                "权限错误".encode('gbk'), 0x10)
        sys.exit()

# 配置文件路径（确保打包成exe后能正确读写）
import os
if hasattr(sys, '_MEIPASS'):
    # 打包成exe后的路径
    CONFIG_FILE = os.path.join(os.path.dirname(sys.executable), "config.json")
else:
    # 开发环境路径
    CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

# 默认配置
DEFAULT_CONFIG = {
    "ntp_servers": [
        "pool.ntp.org",
        "time.nist.gov",
        "cn.ntp.org.cn",
        "time.windows.com"
        # "1.97.84.242",
        # "1.1.16.71",
        # "13.12.3.253",
        # "8.1.19.253",
        # "21.71.254.151",
        # "24.27.254.253",
        # "54.220.39.253",
        # "49.64.201.25",
        # "51.144.252.253",
        # "51.144.254.253",
        # "32.61.227.69"
    ],
    "auto_sync_interval": 3600,  # 默认60分钟
    "startup": True
}

# 加载配置
def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                config = json.load(f)
                # 确保startup字段默认为True
                if "startup" not in config:
                    config["startup"] = True
                # 如果startup是False，强制改为True
                if not config.get("startup", True):
                    config["startup"] = True
                # 确保auto_sync_interval默认为3600秒（60分钟）
                if "auto_sync_interval" not in config:
                    config["auto_sync_interval"] = 3600
                if config.get("auto_sync_interval", 0) < 60:
                    config["auto_sync_interval"] = 3600
                # 无论是否有变更，都保存一次配置
                save_config(config)
                return config
        except:
            return DEFAULT_CONFIG.copy()
    return DEFAULT_CONFIG.copy()

# 保存配置
def save_config(config):
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=4)

# 获取当前系统时间
def get_system_time():
    return datetime.now()

# 使用socket实现NTP时间同步，返回详细结果
def sync_time(ntp_server, set_system_time=False):
    try:
        import socket
        import struct
        from datetime import datetime, timedelta
        
        # NTP服务器端口
        NTP_PORT = 123
        # 超时时间
        TIMEOUT = 5
        
        # 创建socket
        client = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        client.settimeout(TIMEOUT)
        
        # 构建NTP请求包
        # 0-1: LI, VN, Mode
        # 2: Stratum
        # 3: Poll
        # 4: Precision
        # 5-8: Root Delay
        # 9-12: Root Dispersion
        # 13-16: Reference ID
        # 17-24: Reference Timestamp
        # 25-32: Originate Timestamp (客户端发送时间)
        # 33-40: Receive Timestamp (服务器接收时间)
        # 41-48: Transmit Timestamp (服务器发送时间)
        ntp_packet = bytearray(48)
        ntp_packet[0] = 0x1b  # LI=0, VN=3, Mode=3 (客户端模式)
        
        # 获取当前时间（客户端发送时间）
        now = datetime.now()
        # 转换为NTP时间戳（从1900年开始的秒数）
        ntp_time = time.time() + 2208988800  # Unix时间转NTP时间
        
        # 整数部分和小数部分
        ntp_sec = int(ntp_time)
        ntp_frac = int((ntp_time - ntp_sec) * (2**32))
        
        # 填充Originate Timestamp字段（第25-32字节）
        ntp_packet[24:28] = ntp_sec.to_bytes(4, 'big')
        ntp_packet[28:32] = ntp_frac.to_bytes(4, 'big')
        
        # 记录发送时间
        send_time = datetime.now()
        
        # 发送请求并接收响应
        client.sendto(ntp_packet, (ntp_server, NTP_PORT))
        response, _ = client.recvfrom(48)
        
        # 记录接收时间
        receive_time = datetime.now()
        client.close()
        
        # 解析响应包，获取各时间戳字段
        # 发送时间戳 (第24-31字节)
        orig_timestamp = struct.unpack('!I', response[24:28])[0]
        orig_timestamp_fraction = struct.unpack('!I', response[28:32])[0]
        
        # 接收时间戳 (第32-39字节)
        recv_timestamp = struct.unpack('!I', response[32:36])[0]
        recv_timestamp_fraction = struct.unpack('!I', response[36:40])[0]
        
        # 发送时间戳 (第40-47字节)
        transmit_timestamp = struct.unpack('!I', response[40:44])[0]
        transmit_timestamp_fraction = struct.unpack('!I', response[44:48])[0]
        
        # NTP时间戳转换为Unix时间戳 (NTP时间从1900-01-01开始，Unix时间从1970-01-01开始)
        NTP_EPOCH = 2208988800  # 1970-01-01 00:00:00 UTC对应的NTP时间戳
        
        # 计算完整的时间戳（秒 + 毫秒）
        def ntp_to_unix(seconds, fraction):
            # 转换为Unix时间戳（秒）
            unix_sec = seconds - NTP_EPOCH
            # 小数部分转换为毫秒
            ms = int(fraction * 1000 / (2**32))
            return unix_sec, ms
        
        # 转换所有时间戳
        orig_sec, orig_ms = ntp_to_unix(orig_timestamp, orig_timestamp_fraction)
        recv_sec, recv_ms = ntp_to_unix(recv_timestamp, recv_timestamp_fraction)
        transmit_sec, transmit_ms = ntp_to_unix(transmit_timestamp, transmit_timestamp_fraction)
        
        # 计算网络延迟（往返时间）
        round_trip_delay = (receive_time - send_time).total_seconds() * 1000  # 毫秒
        delay = round(round_trip_delay, 1)
        
        # 计算NTP时间（从服务器返回的发送时间）
        # 转换为Unix时间戳（秒）
        transmit_unix_sec = transmit_timestamp - NTP_EPOCH
        # 小数部分转换为毫秒
        transmit_ms = int(transmit_timestamp_fraction * 1000 / (2**32))
        
        # NTP时间的完整秒数
        ntp_seconds = transmit_unix_sec
        ntp_milliseconds = transmit_ms
        
        # 计算NTP时间的完整Unix时间戳（秒）
        ntp_full_time = ntp_seconds + ntp_milliseconds / 1000
        
        # 本地时间的Unix时间戳（秒）- 兼容Windows XP，使用time.mktime
        local_time = time.mktime(receive_time.timetuple())
        
        # 偏移量：NTP时间 - 本地时间（毫秒）
        # 正偏移表示本地时间比NTP时间慢，需要调快
        # 负偏移表示本地时间比NTP时间快，需要调慢
        offset = round((ntp_full_time - local_time) * 1000, 1)
        
        # 计算NTP时间
        ntp_datetime = datetime.fromtimestamp(transmit_sec) + timedelta(milliseconds=transmit_ms)
        
        # 设置系统时间（如果需要）
        if set_system_time:
            try:
                # 使用Windows API设置系统时间，兼容Windows XP到Win11
                
                # 定义SYSTEMTIME结构体
                class SYSTEMTIME(ctypes.Structure):
                    _fields_ = [
                        ("wYear", ctypes.c_ushort),
                        ("wMonth", ctypes.c_ushort),
                        ("wDayOfWeek", ctypes.c_ushort),
                        ("wDay", ctypes.c_ushort),
                        ("wHour", ctypes.c_ushort),
                        ("wMinute", ctypes.c_ushort),
                        ("wSecond", ctypes.c_ushort),
                        ("wMilliseconds", ctypes.c_ushort)
                    ]
                
                # 关键修复：NTP返回的是UTC时间戳，需要正确转换为东八区时间
                # 1. 首先，将NTP时间戳转换为UTC时间
                try:
                    # Python 3.3+支持utcfromtimestamp
                    utc_datetime = datetime.utcfromtimestamp(transmit_sec + transmit_ms / 1000.0)
                except AttributeError:
                    # Python 2.7也支持utcfromtimestamp
                    utc_datetime = datetime.utcfromtimestamp(transmit_sec + transmit_ms / 1000.0)
                
                # 2. 强制转换为东八区（+8）时间
                # 计算东八区时间：UTC时间 + 8小时
                from datetime import timedelta
                east8_datetime = utc_datetime + timedelta(hours=8)
                
                # 3. 填充东八区时间到SYSTEMTIME结构体
                st_east8 = SYSTEMTIME()
                st_east8.wYear = east8_datetime.year
                st_east8.wMonth = east8_datetime.month
                st_east8.wDay = east8_datetime.day
                st_east8.wHour = east8_datetime.hour
                st_east8.wMinute = east8_datetime.minute
                st_east8.wSecond = east8_datetime.second
                st_east8.wMilliseconds = transmit_ms
                st_east8.wDayOfWeek = east8_datetime.weekday()
                
                # 4. 调用SetLocalTime API直接设置东八区时间
                ctypes.windll.kernel32.SetLocalTime(ctypes.byref(st_east8))
            except Exception as e:
                # 如果API调用失败，尝试使用命令行方式（备用方案）
                try:
                    # 确保使用东八区时间
                    from datetime import timedelta
                    # 计算东八区时间
                    east8_datetime = ntp_datetime + timedelta(hours=8)
                    time_str = east8_datetime.strftime('%H:%M:%S')
                    date_str = east8_datetime.strftime('%Y-%m-%d')
                    
                    # 使用cmd.exe /c执行命令，确保在Win11上能正常工作
                    import subprocess
                    subprocess.run(f"cmd.exe /c time {time_str}", shell=True, check=False)
                    subprocess.run(f"cmd.exe /c date {date_str}", shell=True, check=False)
                except Exception as cmd_e:
                    # 记录错误，但不影响返回结果
                    pass
        
        # 返回详细结果
        return {
            'success': True,
            'server': ntp_server,
            'status': 'Good',
            'delay': delay,
            'offset': offset,
            'ntp_time': ntp_datetime,
            'error': ''
        }
    except Exception as e:
        return {
            'success': False,
            'server': ntp_server,
            'status': 'Error',
            'delay': None,
            'offset': None,
            'ntp_time': None,
            'error': str(e)
        }

# 设置开机自启（兼容Windows XP）
def set_startup(enable):
    try:
        # Windows XP兼容的注册表操作
        key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
        
        # 使用HKEY_CURRENT_USER，不需要管理员权限
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, 
                           winreg.KEY_ALL_ACCESS)
        
        if enable:
            exe_path = sys.executable
            # 使用ANSI字符串，兼容Windows XP
            winreg.SetValueEx(key, "SystemTimeSyncTool", 0, winreg.REG_SZ, exe_path)
        else:
            try:
                winreg.DeleteValue(key, "SystemTimeSyncTool")
            except Exception:
                pass
        
        winreg.CloseKey(key)
        return True
    except Exception as e:
        return False, str(e)

# 主应用类
class TimeSyncTool:
    def __init__(self):
        self.config = load_config()
        self.root = tk.Tk()
        self.root.title("系统时间同步工具")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
        # 同步状态标记
        self.syncing = False
        
        # 自动同步定时器
        self.auto_sync_timer = None
        
        # 创建系统托盘图标
        self.create_tray_icon()
        
        # 拦截窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.setup_ui()
        
        # 确保开机自启生效（根据配置设置注册表）
        set_startup(self.config["startup"])
        
        self.start_auto_sync()
    
    def create_tray_icon(self):
        """创建系统托盘图标"""
        # 创建一个简单的图标
        from PIL import Image, ImageDraw
        
        # 创建一个64x64的图标
        image = Image.new('RGB', (64, 64), color=(0, 120, 215))
        draw = ImageDraw.Draw(image)
        
        # 绘制一个时钟图标
        draw.ellipse([8, 8, 56, 56], fill=(255, 255, 255), outline=(255, 255, 255), width=2)
        draw.line([32, 32, 32, 16], fill=(0, 120, 215), width=3)
        draw.line([32, 32, 44, 32], fill=(0, 120, 215), width=3)
        
        # 创建托盘菜单，设置"显示窗口"为默认项
        show_item = pystray.MenuItem("显示窗口", self.show_window, default=True)
        quit_item = pystray.MenuItem("退出程序", self.quit_app)
        menu = pystray.Menu(show_item, quit_item)

        # 创建托盘图标
        self.icon = pystray.Icon("TimeSyncTool", image, "系统时间同步工具", menu)

        # 在单独的线程中运行托盘图标
        threading.Thread(target=self.icon.run, daemon=True).start()

    def show_window(self, icon=None, item=None):
        """显示主窗口"""
        print("显示窗口被调用")
        self.root.deiconify()
        self.root.lift()
        self.root.state('normal')
        print("窗口状态:", self.root.state())
    
    def quit_app(self, icon=None, item=None):
        """退出程序"""
        self.icon.stop()
        self.root.quit()
        self.root.destroy()
    
    def on_closing(self):
        """窗口关闭事件处理"""
        # 隐藏窗口而不是退出
        self.root.withdraw()
    
    def setup_ui(self):
        # 调整窗口大小，确保有足够空间显示所有内容
        self.root.geometry("600x500")
        
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 1. 系统时间显示 - 顶部
        time_frame = ttk.LabelFrame(main_frame, text="当前系统时间", padding="10")
        time_frame.pack(fill=tk.X, pady=5)
        
        self.time_var = tk.StringVar()
        self.update_time()
        time_label = ttk.Label(time_frame, textvariable=self.time_var, font=("Arial", 24))
        time_label.pack()
        
        # 3. 按钮框架 - 底部（先创建，确保显示在最下面）
        button_frame = ttk.Frame(main_frame, padding="10")
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=5)
        
        # 一键同步按钮和loading效果
        sync_frame = ttk.Frame(button_frame)
        sync_frame.pack(side=tk.LEFT, padx=5)
        
        # 一键同步按钮 - 改为实例变量，以便在其他方法中控制
        self.sync_btn = ttk.Button(sync_frame, text="一键同步", command=self.on_sync)
        self.sync_btn.pack(side=tk.LEFT)
        
        # 设置按钮
        settings_btn = ttk.Button(button_frame, text="自定义同步", command=self.open_settings)
        settings_btn.pack(side=tk.LEFT, padx=5)
        
        # 作者按钮
        author_btn = ttk.Button(button_frame, text="关于", command=self.show_about)
        author_btn.pack(side=tk.LEFT, padx=5)
        
        # 开机自启复选框
        self.startup_var = tk.BooleanVar(value=self.config["startup"])
        startup_check = ttk.Checkbutton(button_frame, text="开机自启", variable=self.startup_var, command=self.on_startup_change)
        startup_check.pack(side=tk.RIGHT, padx=5)
        
        # 2. NTP服务器列表 - 中间
        server_frame = ttk.LabelFrame(main_frame, text="NTP服务器", padding="10")
        server_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 创建表格
        self.server_tree = ttk.Treeview(server_frame, columns=('address', 'status', 'offset', 'delay', 'error'), show='headings')
        
        # 设置列标题
        self.server_tree.heading('address', text='服务器地址')
        self.server_tree.heading('status', text='状态')
        self.server_tree.heading('offset', text='偏移量')
        self.server_tree.heading('delay', text='延迟')
        self.server_tree.heading('error', text='最近一次报错')
        
        # 设置列宽
        self.server_tree.column('address', width=150)
        self.server_tree.column('status', width=80)
        self.server_tree.column('offset', width=80)
        self.server_tree.column('delay', width=80)
        self.server_tree.column('error', width=200)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(server_frame, orient=tk.VERTICAL, command=self.server_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.server_tree.configure(yscrollcommand=scrollbar.set)
        
        # 表格布局
        self.server_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        # 初始化服务器数据 - 初始为空列表，不同步显示任何服务器
        self.server_data = []
    
    def update_time(self):
        current_time = get_system_time()
        self.time_var.set(current_time.strftime("%Y-%m-%d %H:%M:%S"))
        self.root.after(1000, self.update_time)
    
    def perform_sync(self, show_message=True, show_loading=True):
        """执行时间同步操作
        
        Args:
            show_message: 是否显示同步结果消息框
            show_loading: 是否显示loading动画
        """
        results = []
        
        # 清空当前表格内容
        for item in self.server_tree.get_children():
            self.server_tree.delete(item)
        
        # 清空服务器数据
        self.server_data.clear()
        
        # 遍历所有配置的服务器进行同步测试
        for server in self.config["ntp_servers"]:
            # 不设置系统时间，只获取同步信息
            result = sync_time(server, set_system_time=False)
            results.append(result)
            
            # 添加到服务器数据列表
            server_info = {
                'address': server,
                'status': result['status'],
                'offset': f"{result['offset']}ms" if result['offset'] is not None else "",
                'delay': f"{result['delay']}ms" if result['delay'] is not None else "",
                'error': result['error'][:50] + '...' if len(result['error']) > 50 else result['error']
            }
            self.server_data.append(server_info)
            
            # 根据同步结果添加到表格
            if result['success']:
                # 连接成功的服务器显示完整信息
                self.server_tree.insert('', tk.END, 
                                      values=(server, result['status'], 
                                              f"{result['offset']}ms" if result['offset'] is not None else "", 
                                              f"{result['delay']}ms" if result['delay'] is not None else "", 
                                              ""))
            else:
                # 连接失败的服务器只显示错误信息
                self.server_tree.insert('', tk.END, 
                                      values=(server, result['status'], 
                                              "", "", 
                                              result['error'][:50] + '...' if len(result['error']) > 50 else result['error']))
        
        # 选择延迟最小的可用服务器
        valid_servers = [r for r in results if r['success']]
        if not valid_servers:
            if show_message:
                messagebox.showerror("错误", "所有NTP服务器均无法同步")
            return
        
        # 按延迟排序，选择最小的
        best_server = min(valid_servers, key=lambda x: x['delay'])
        
        # 使用最佳服务器设置系统时间
        final_result = sync_time(best_server['server'], set_system_time=True)
        if not final_result['success']:
            # if show_message:
                #messagebox.showinfo("成功", f"时间同步成功！\n使用服务器: {best_server['server']}\n延迟: {best_server['delay']}ms\n偏移量: {best_server['offset']}ms")
     
            if show_message:
                messagebox.showerror("错误", f"同步失败: {final_result['error']}")
    
    def _start_loading(self):
        """启动loading效果"""
        # 按钮防呆：禁用同步按钮
        self.sync_btn.config(state=tk.DISABLED)
        
        # 设置同步状态
        self.syncing = True
    
    def _stop_loading(self):
        """停止loading动画"""
        # 恢复按钮状态和停止loading动画
        self.syncing = False
        self.sync_btn.config(state=tk.NORMAL)
        self.loading_var.set("")
    
    def on_sync(self):
        # 启动loading动画
        self._start_loading()
        
        # 在新线程中执行同步，避免UI冻结
        def sync_thread():
            try:
                self.perform_sync(show_message=True)
            finally:
                # 同步完成后，停止loading动画
                self._stop_loading()
        
        threading.Thread(target=sync_thread).start()
    
    def sync_with_loading(self, show_message=True):
        """带loading效果的同步方法，供外部调用"""
        # 启动loading动画
        self._start_loading()
        
        # 在新线程中执行同步，避免UI冻结
        def sync_thread():
            try:
                self.perform_sync(show_message=show_message)
            finally:
                # 同步完成后，停止loading动画
                self._stop_loading()
        
        threading.Thread(target=sync_thread).start()
    
    def open_settings(self):
        SettingsWindow(self)
    
    def on_startup_change(self):
        enable = self.startup_var.get()
        result = set_startup(enable)
        if isinstance(result, tuple):
            messagebox.showerror("错误", f"设置开机自启失败: {result[1]}")
            self.startup_var.set(not enable)
        else:
            self.config["startup"] = enable
            save_config(self.config)
    
    def show_about(self):
        """显示关于弹窗，位置在主窗口正中"""
        # 创建自定义弹窗
        about_window = tk.Toplevel(self.root)
        about_window.title("关于")
        about_window.geometry("300x120")
        about_window.resizable(False, False)
        
        # 先隐藏窗口，等所有配置完成后再显示
        about_window.withdraw()
        
        # 计算主窗口中心位置
        self.root.update_idletasks()
        main_x = self.root.winfo_x()
        main_y = self.root.winfo_y()
        main_w = self.root.winfo_width()
        main_h = self.root.winfo_height()
        
        # 设置弹窗在主窗口正中
        popup_w = 300
        popup_h = 120
        pos_x = main_x + (main_w - popup_w) // 2
        pos_y = main_y + (main_h - popup_h) // 2
        about_window.geometry(f"{popup_w}x{popup_h}+{pos_x}+{pos_y}")
        
        # 设置为模态窗口
        about_window.transient(self.root)
        about_window.grab_set()
        
        # 添加内容
        content_frame = ttk.Frame(about_window, padding="20")
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(content_frame, text="系统时间同步工具").pack(pady=5)
        ttk.Label(content_frame, text="作者：李灿").pack(pady=5)
        ttk.Button(content_frame, text="确定", command=about_window.destroy).pack(pady=10)
        
        # 配置完成后再显示窗口
        about_window.update_idletasks()
        about_window.deiconify()
        
        # 等待窗口关闭
        about_window.wait_window()
    
    def start_auto_sync(self):
        # 清除之前的定时器
        if self.auto_sync_timer:
            self.root.after_cancel(self.auto_sync_timer)
        
        # 设置新的定时器
        def auto_sync():
            if not self.config["ntp_servers"]:
                self.start_auto_sync()
                return
            
            # 遍历所有服务器，选择最佳服务器
            results = []
            for server in self.config["ntp_servers"]:
                result = sync_time(server, set_system_time=False)
                results.append(result)
            
            # 选择延迟最小的可用服务器
            valid_servers = [r for r in results if r['success']]
            if valid_servers:
                best_server = min(valid_servers, key=lambda x: x['delay'])
                sync_time(best_server['server'], set_system_time=True)
            
            self.start_auto_sync()
        
        self.auto_sync_timer = self.root.after(self.config["auto_sync_interval"] * 1000, auto_sync)
    
    def update_server_list(self):
        # 清空表格
        for item in self.server_tree.get_children():
            self.server_tree.delete(item)
        
        # 清空服务器数据
        self.server_data.clear()
        
        # 添加服务器到表格
        for server in self.config["ntp_servers"]:
            self.server_data.append({
                'address': server,
                'status': '',
                'offset': '',
                'delay': '',
                'error': ''
            })
            self.server_tree.insert('', tk.END, values=(server, '', '', '', ''))
    
    def run(self):
        self.root.mainloop()

# 设置窗口类
class SettingsWindow:
    def __init__(self, parent):
        self.parent = parent
        self.settings_window = tk.Toplevel(parent.root)
        self.settings_window.title("设置")
        self.settings_window.geometry("400x400")  # 增加窗口高度
        self.settings_window.resizable(False, False)
        
        # 复制当前配置
        self.temp_config = self.parent.config.copy()
        
        self.setup_ui()
    
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.settings_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # NTP服务器设置
        server_frame = ttk.LabelFrame(main_frame, text="NTP服务器设置", padding="10")
        server_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 服务器列表
        self.server_listbox = tk.Listbox(server_frame, height=5)
        self.server_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(server_frame, orient=tk.VERTICAL, command=self.server_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.server_listbox.config(yscrollcommand=scrollbar.set)
        
        # 添加服务器到列表
        for server in self.temp_config["ntp_servers"]:
            self.server_listbox.insert(tk.END, server)
        
        # 服务器操作按钮
        server_btn_frame = ttk.Frame(server_frame, padding="5")
        server_btn_frame.pack(side=tk.RIGHT, fill=tk.Y)
        
        add_btn = ttk.Button(server_btn_frame, text="添加", command=self.add_server)
        add_btn.pack(fill=tk.X, pady=2)
        
        edit_btn = ttk.Button(server_btn_frame, text="编辑", command=self.edit_server)
        edit_btn.pack(fill=tk.X, pady=2)
        
        delete_btn = ttk.Button(server_btn_frame, text="删除", command=self.delete_server)
        delete_btn.pack(fill=tk.X, pady=2)
        
        # 自动同步设置
        auto_sync_frame = ttk.LabelFrame(main_frame, text="自动同步设置", padding="10")
        auto_sync_frame.pack(fill=tk.X, pady=5)
        
        # 同步周期标签
        interval_label = ttk.Label(auto_sync_frame, text="同步周期（分钟）:")
        interval_label.pack(side=tk.LEFT, padx=5)
        
        # 同步周期输入
        self.interval_var = tk.IntVar(value=self.temp_config["auto_sync_interval"] // 60)
        interval_entry = ttk.Entry(auto_sync_frame, textvariable=self.interval_var, width=10)
        interval_entry.pack(side=tk.LEFT, padx=5)
        
        # 按钮框架
        btn_frame = ttk.Frame(main_frame, padding="10")
        btn_frame.pack(fill=tk.X, pady=5)
        
        # 保存按钮
        save_btn = ttk.Button(btn_frame, text="保存并同步", command=self.save_settings)
        save_btn.pack(side=tk.RIGHT, padx=5)
        
        # 取消按钮
        cancel_btn = ttk.Button(btn_frame, text="取消", command=self.settings_window.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=5)
    
    def add_server(self):
        server = tk.simpledialog.askstring("添加服务器", "请输入NTP服务器地址:")
        if server and server not in self.temp_config["ntp_servers"]:
            self.temp_config["ntp_servers"].append(server)
            self.update_server_list()
    
    def edit_server(self):
        selected_index = self.server_listbox.curselection()
        if not selected_index:
            messagebox.showwarning("警告", "请选择一个服务器进行编辑")
            return
        
        old_server = self.server_listbox.get(selected_index[0])
        new_server = tk.simpledialog.askstring("编辑服务器", "请输入新的NTP服务器地址:", initialvalue=old_server)
        
        if new_server and new_server != old_server:
            if new_server in self.temp_config["ntp_servers"]:
                messagebox.showwarning("警告", "该服务器已存在")
                return
            
            index = self.temp_config["ntp_servers"].index(old_server)
            self.temp_config["ntp_servers"][index] = new_server
            self.update_server_list()
    
    def delete_server(self):
        selected_index = self.server_listbox.curselection()
        if not selected_index:
            messagebox.showwarning("警告", "请选择一个服务器进行删除")
            return
        
        if len(self.temp_config["ntp_servers"]) <= 1:
            messagebox.showwarning("警告", "至少需要保留一个NTP服务器")
            return
        
        server = self.server_listbox.get(selected_index[0])
        self.temp_config["ntp_servers"].remove(server)
        self.update_server_list()
    
    def update_server_list(self):
        self.server_listbox.delete(0, tk.END)
        for server in self.temp_config["ntp_servers"]:
            self.server_listbox.insert(tk.END, server)
    
    def save_settings(self):
        # 验证同步周期
        interval = self.interval_var.get()
        if interval < 1:
            messagebox.showwarning("警告", "同步周期不能小于1分钟")
            return
        
        # 更新配置
        self.temp_config["auto_sync_interval"] = interval * 60
        
        # 保存配置
        self.parent.config = self.temp_config.copy()
        save_config(self.parent.config)
        
        # 更新主窗口
        self.parent.update_server_list()
        self.parent.start_auto_sync()
        
        # 保存设置后执行一次同步，调用主窗口的sync_with_loading方法，显示loading效果
        def sync_after_save():
            self.parent.sync_with_loading(show_message=False)
        
        # 在新线程中执行同步，避免阻塞UI
        threading.Thread(target=sync_after_save).start()
        
        # 关闭设置窗口
        self.settings_window.destroy()

# 主函数
if __name__ == "__main__":
    # 确保time模块已导入
    import time
    
    # 程序启动时检查管理员权限
    if not is_admin():
        # 没有管理员权限，给出提示
        message = "当前没有管理员权限，程序将以管理员身份重新启动，以便使用完整功能。"
        try:
            ctypes.windll.user32.MessageBoxW(None, message, "权限提示", 0x40)  # 0x40 = MB_ICONINFORMATION
        except:
            ctypes.windll.user32.MessageBoxA(None, message.encode('gbk'), "权限提示".encode('gbk'), 0x40)
        
        # 以管理员身份重新启动程序
        run_as_admin()
        # 退出当前实例
        sys.exit()
    
    # 直接运行应用
    try:
        app = TimeSyncTool()
        app.run()
    except Exception as e:
        # 显示错误信息，兼容Windows XP
        error_msg = f"程序运行出错: {str(e)}\n请确保系统已安装Python 2.7或3.4+"
        ctypes.windll.user32.MessageBoxA(None, error_msg.encode('gbk'), 
                                        "运行错误".encode('gbk'), 0x10)
