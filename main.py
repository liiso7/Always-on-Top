import win32gui
import win32con
import win32api
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
from ctypes import windll, byref, c_int, c_void_p, Structure
import json
import os
import sys

def get_resource_path(relative_path):
    """获取资源的绝对路径，兼容 PyInstaller"""
    try:
        # PyInstaller 创建临时文件夹并存储路径在 _MEIPASS 中
        base_path = sys._MEIPASS  # pyright: ignore[reportAttributeAccessIssue]
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- 核心功能模块 ---

class WindowUtils:
    @staticmethod
    def is_window_topmost(hwnd):
        """检测窗口是否已经处于置顶状态"""
        try:
            ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            return (ex_style & win32con.WS_EX_TOPMOST) != 0
        except Exception:
            return False

    @staticmethod
    def set_window_topmost(hwnd, is_topmost):
        """设置窗口置顶/取消置顶"""
        try:
            if is_topmost:
                # 置顶：HWND_TOPMOST(=-1) + 不改变位置 + 保持大小
                win32gui.SetWindowPos(
                    hwnd,
                    win32con.HWND_TOPMOST,
                    0, 0, 0, 0,
                    win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
                )
            else:
                # 取消置顶：HWND_NOTOPMOST(=-2)
                win32gui.SetWindowPos(
                    hwnd,
                    win32con.HWND_NOTOPMOST,
                    0, 0, 0, 0,
                    win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
                )
            return True
        except Exception as e:
            print(f"设置置顶状态失败: {e}")
            return False

    @staticmethod
    def get_all_windows():
        """获取所有可见的窗口（排除无标题/系统窗口）"""
        windows = []
        
        def callback(hwnd, extra):
            # 过滤条件：可见 + 有标题 + 不是工具条/对话框等
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                title = win32gui.GetWindowText(hwnd)
                # 排除过短的标题（系统弹窗）
                if len(title) > 1:
                    windows.append((hwnd, title))
            return True
        
        win32gui.EnumWindows(callback, None)
        return windows

    @staticmethod
    def get_foreground_window():
        """获取当前激活窗口"""
        try:
            hwnd = win32gui.GetForegroundWindow()
            if hwnd and windll.user32.IsWindow(hwnd) and windll.user32.IsWindowVisible(hwnd):
                return hwnd
            return None
        except Exception:
            return None

    @staticmethod
    def show_effect(hwnd, mode='pin'):
        """显示置顶窗口的提醒效果"""
        try:
            rect = win32gui.GetWindowRect(hwnd)
            x, y = rect[0], rect[1]
            
            def move_effect():
                try:
                    offsets = [5, 10, 5, 0] if mode == 'pin' else [-5, -10, -5, 0]
                    for offset in offsets:
                        current_x = x + offset if mode == 'pin' else x + offset # 简化处理，实际上应该分别处理
                        # 简单的晃动效果
                        win32gui.SetWindowPos(hwnd, 0, x + offset, y + offset, 0, 0,
                                            win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)
                        time.sleep(0.05)
                except Exception:
                    pass
            
            # 在单独线程中运行效果，以免阻塞
            threading.Thread(target=move_effect, daemon=True).start()
            
            # 闪烁效果
            # for i in range(2):
            #     win32gui.FlashWindow(hwnd, True)
            #     time.sleep(0.1)
            #     win32gui.FlashWindow(hwnd, False)
            #     time.sleep(0.1)
                
        except Exception:
            pass

# --- 配置管理模块 ---

class ConfigManager:
    def __init__(self, config_file='./config/config.json'):
        self.config_file = config_file
        self.config = self.load_config()
    
    def load_config(self):
        """加载配置文件"""
        default_config = {
            'hotkey_modifier': 0x0002,  # Ctrl
            'hotkey_key': ord('Q'),     # Q
            'modifier1_name': 'Ctrl',
            'modifier2_name': '无',
            'key_name': 'Q'
        }
        
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    default_config.update(loaded_config)
                    # 兼容旧配置，如果缺少 modifier2_name，设置为 '无'
                    if 'modifier2_name' not in loaded_config:
                        default_config['modifier2_name'] = '无'
            except Exception as e:
                print(f"加载配置文件失败: {e}")
        
        return default_config
    
    def save_config(self):
        """保存配置文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"保存配置文件失败: {e}")
            return False
    
    def get_hotkey(self):
        """获取快捷键配置"""
        return self.config['hotkey_modifier'], self.config['hotkey_key']
    
    def get_hotkey_display(self):
        """获取快捷键的显示文本"""
        modifier1 = self.config['modifier1_name']
        modifier2 = self.config.get('modifier2_name', '无')
        key = self.config['key_name']
        
        if modifier2 == '无':
            return f"{modifier1} + {key}"
        else:
            return f"{modifier1} + {modifier2} + {key}"
    
    def set_hotkey(self, modifier, key, modifier1_name, modifier2_name, key_name):
        """设置快捷键配置"""
        self.config['hotkey_modifier'] = modifier
        self.config['hotkey_key'] = key
        self.config['modifier1_name'] = modifier1_name
        self.config['modifier2_name'] = modifier2_name
        self.config['key_name'] = key_name
        return self.save_config()

# --- 快捷键监听模块 ---

class HotkeyListener(threading.Thread):
    def __init__(self, callback, config_manager):
        super().__init__(daemon=True)
        self.callback = callback
        self.config_manager = config_manager
        self.running = True
        self.modifier, self.key = config_manager.get_hotkey()
        self.hotkey_id = 1
        self.hotkey_registered = False
    def run(self):
        # 注册快捷键
        try:
            modifier_name = self.config_manager.config.get('modifier_name', 'Ctrl')
            key_name = self.config_manager.config.get('key_name', 'Q')
            print(f"正在注册快捷键: {modifier_name} + {key_name} (modifier={self.modifier}, key={self.key})")
            
            if not windll.user32.RegisterHotKey(None, self.hotkey_id, self.modifier, self.key):
                print(f"快捷键 {modifier_name}+{key_name} 注册失败")
                print(f"可能原因：快捷键已被其他程序占用")
                return
            self.hotkey_registered = True
            print(f"快捷键 {modifier_name}+{key_name} 注册成功!")
        except Exception as e:
            print(f"注册快捷键异常: {e}")
            return

        # 消息循环
        class MSG(Structure):
            _fields_ = [
                ("hwnd", c_void_p),
                ("message", c_int),
                ("wParam", c_int),
                ("lParam", c_int),
                ("time", c_int),
                ("pt", c_int * 2),
            ]
        
        msg = MSG()
        while self.running:
            try:
                # GetMessage 是阻塞的，所以不需要 sleep
                # 但为了能响应停止信号，我们可以使用 PeekMessage 或者发送一个模拟消息来唤醒
                # 这里简单起见，使用带超时的 GetMessage (不直接支持) 或者 PostQuitMessage
                # 实际上，只要主线程结束，daemon 线程就会被杀掉
                result = windll.user32.GetMessageA(byref(msg), None, 0, 0)
                if result == 0:
                    break
                
                if msg.message == win32con.WM_HOTKEY:
                    if self.callback:
                        self.callback()
                
                windll.user32.TranslateMessage(byref(msg))
                windll.user32.DispatchMessageA(byref(msg))
            except Exception:
                break
        
        # 清理
        if self.hotkey_registered:
            windll.user32.UnregisterHotKey(None, self.hotkey_id)

    def stop(self):
        self.running = False
        # 先尝试注销快捷键，这样 GetMessage 会返回 0
        if self.hotkey_registered:
            windll.user32.UnregisterHotKey(None, self.hotkey_id)
            self.hotkey_registered = False
        # 发送一个空消息来打破 GetMessage 的阻塞
        try:
            windll.user32.PostThreadMessageA(self.ident, win32con.WM_NULL, 0, 0)
        except:
            pass

# --- 设置对话框模块 ---

class SettingsDialog:
    """设置对话框"""
    def __init__(self, parent, config_manager, on_hotkey_changed):
        self.parent = parent
        self.config_manager = config_manager
        self.on_hotkey_changed = on_hotkey_changed
        
        # 初始化实例变量
        self.modifier1_var = None
        self.modifier2_var = None
        self.key_var = None
        self.hotkey_preview_label = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("设置")
        
        # 先隐藏窗口，防止闪烁
        self.dialog.withdraw()
        
        # 设置窗口图标
        icon_path = get_resource_path(os.path.join("icon", "app_icon.ico"))
        if os.path.exists(icon_path):
            try:
                self.dialog.iconbitmap(icon_path)
            except Exception as e:
                print(f"设置对话框图标失败: {e}")

        # 计算居中位置并先设置窗口大小和位置
        dialog_width = 400
        dialog_height = 350
        x = (self.dialog.winfo_screenwidth() // 2) - (dialog_width // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (dialog_height // 2)
        self.dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        self.dialog.resizable(True, True)
        self.dialog.minsize(400, 350)

        # 设置界面
        self.setup_ui()

        # 显示窗口
        self.dialog.deiconify()

        # 设置为模态对话框
        self.dialog.transient(parent)
        self.dialog.grab_set()
    
    def setup_ui(self):
        """设置界面"""
        # 标题
        title_frame = ttk.Frame(self.dialog, padding="5")
        title_frame.pack(fill=tk.X)
        ttk.Label(title_frame, text="设置", font=("微软雅黑", 14, "bold")).pack()
        
        # 创建选项卡
        self.notebook = ttk.Notebook(self.dialog)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0,5))
        
        # 快捷键设置选项卡
        self.hotkey_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.hotkey_tab, text="快捷键")
        
        # 设置快捷键选项卡内容
        self.setup_hotkey_tab()
        
        # 其他设置选项卡（预留）
        self.other_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.other_tab, text="其他设置")
        
        # 设置其他选项卡内容
        self.setup_other_tab()
    
    def setup_hotkey_tab(self):
        """设置快捷键选项卡内容"""
        # 创建一个容器框架，用于垂直布局
        container = ttk.Frame(self.hotkey_tab)
        container.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        # 标题说明
        # title_label = ttk.Label(container, text="设置窗口置顶快捷键", font=("微软雅黑", 11, "bold"))
        # title_label.pack(anchor=tk.W, pady=(0, 12))
        
        # 修饰键1选择
        modifier1_frame = ttk.Frame(container)
        modifier1_frame.pack(fill=tk.X, pady=3)
        
        ttk.Label(modifier1_frame, text="修饰键1:", font=("微软雅黑", 10), width=15).pack(side=tk.LEFT, padx=5)
        
        self.modifier1_var = tk.StringVar(value=self.config_manager.config.get('modifier1_name', 'Ctrl'))
        modifier1_combo = ttk.Combobox(modifier1_frame, textvariable=self.modifier1_var, 
                                     values=['Ctrl', 'Alt', 'Shift', 'Win'], 
                                     state="readonly", width=15)
        modifier1_combo.pack(side=tk.LEFT, padx=5)
        modifier1_combo.bind('<<ComboboxSelected>>', lambda e: self.update_hotkey_preview())
        
        # 修饰键2选择
        modifier2_frame = ttk.Frame(container)
        modifier2_frame.pack(fill=tk.X, pady=3)
        
        ttk.Label(modifier2_frame, text="修饰键2:", font=("微软雅黑", 10), width=15).pack(side=tk.LEFT, padx=5)
        
        modifier2_default = self.config_manager.config.get('modifier2_name', '无')
        self.modifier2_var = tk.StringVar(value=modifier2_default)
        modifier2_combo = ttk.Combobox(modifier2_frame, textvariable=self.modifier2_var, 
                                     values=['无', 'Ctrl', 'Alt', 'Shift', 'Win'], 
                                     state="readonly", width=15)
        modifier2_combo.pack(side=tk.LEFT, padx=5)
        modifier2_combo.bind('<<ComboboxSelected>>', lambda e: self.update_hotkey_preview())
        
        # 主键选择
        key_frame = ttk.Frame(container)
        key_frame.pack(fill=tk.X, pady=3)
        
        ttk.Label(key_frame, text="按键:", font=("微软雅黑", 10), width=15).pack(side=tk.LEFT, padx=5)
        
        self.key_var = tk.StringVar(value=self.config_manager.config['key_name'])
        key_combo = ttk.Combobox(key_frame, textvariable=self.key_var,
                                values=['Q', 'W', 'E', 'R', 'T', 'Y', 'U', 'I', 'O', 'P',
                                       'A', 'S', 'D', 'F', 'G', 'H', 'J', 'K', 'L',
                                       'Z', 'X', 'C', 'V', 'B', 'N', 'M',
                                       'F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8', 'F9', 'F10', 'F11', 'F12'],
                                state="readonly", width=15)
        key_combo.pack(side=tk.LEFT, padx=5)
        key_combo.bind('<<ComboboxSelected>>', lambda e: self.update_hotkey_preview())
        
        # 分隔线
        ttk.Separator(container, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        # 快捷键预览
        preview_frame = ttk.Frame(container)
        preview_frame.pack(fill=tk.X, pady=5)
        ttk.Label(preview_frame, text="当前设置:", font=("微软雅黑", 10, "bold")).pack(anchor=tk.W)
        
        self.hotkey_preview_label = ttk.Label(preview_frame, text="", 
                                            font=("微软雅黑", 10, "bold"), foreground="blue")
        self.hotkey_preview_label.pack(anchor=tk.W, pady=10)
        
        # 说明文字
        info_label = ttk.Label(container, text="提示: 可以选择两个修饰键来组合使用，如 Ctrl + Shift + Q",
                              font=("微软雅黑", 9), foreground="gray")
        info_label.pack(anchor=tk.W, pady=(5, 0))
        
        # 快捷键选项卡的按钮区域
        btn_frame = ttk.Frame(container)
        btn_frame.pack(fill=tk.X, pady=(15, 0))
        
        ttk.Button(btn_frame, text="应用", command=self.apply_settings, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="恢复默认", command=self.restore_default, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=self.dialog.destroy, width=12).pack(side=tk.LEFT, padx=5)
        
        # 更新预览
        self.update_hotkey_preview()
    
    def setup_other_tab(self):
        """设置其他选项卡内容（预留）"""
        container = ttk.Frame(self.other_tab)
        container.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        label = ttk.Label(container, text="其他设置功能开发中...", font=("微软雅黑", 12))
        label.pack(expand=True)
        
        # 其他选项卡的按钮区域（仅取消按钮）
        btn_frame = ttk.Frame(container)
        btn_frame.pack(fill=tk.X, pady=(15, 0))
        
        ttk.Button(btn_frame, text="关闭", command=self.dialog.destroy, width=12).pack(side=tk.LEFT, padx=5)
    
    def update_hotkey_preview(self):
        """更新快捷键预览"""
        modifier1 = self.modifier1_var.get()  # pyright: ignore[reportOptionalMemberAccess]
        modifier2 = self.modifier2_var.get()  # pyright: ignore[reportOptionalMemberAccess]
        key = self.key_var.get()  # pyright: ignore[reportOptionalMemberAccess]
        
        if modifier2 == '无':
            hotkey_text = f"{modifier1} + {key}"
        else:
            hotkey_text = f"{modifier1} + {modifier2} + {key}"
        
        self.hotkey_preview_label.config(text=hotkey_text)  # pyright: ignore[reportOptionalMemberAccess]
    
    def get_modifier_code(self, modifier_name):
        """根据修饰键名称获取对应的代码"""
        modifier_map = {
            'Ctrl': 0x0002,
            'Alt': 0x0001,
            'Shift': 0x0004,
            'Win': 0x0008
        }
        return modifier_map.get(modifier_name, 0)
    
    def get_key_code(self, key_name):
        """根据按键名称获取对应的代码"""
        # 功能键
        f_keys = {
            'F1': 0x70, 'F2': 0x71, 'F3': 0x72, 'F4': 0x73,
            'F5': 0x74, 'F6': 0x75, 'F7': 0x76, 'F8': 0x77,
            'F9': 0x78, 'F10': 0x79, 'F11': 0x7A, 'F12': 0x7B
        }
        
        if key_name in f_keys:
            return f_keys[key_name]
        
        # 字母键
        return ord(key_name.upper())
    
    def apply_settings(self):
        """应用设置"""
        modifier1_name = self.modifier1_var.get()  # pyright: ignore[reportOptionalMemberAccess]
        modifier2_name = self.modifier2_var.get()  # pyright: ignore[reportOptionalMemberAccess]
        key_name = self.key_var.get()  # pyright: ignore[reportOptionalMemberAccess]

        # 计算修饰键组合值
        modifier1 = self.get_modifier_code(modifier1_name)
        modifier2 = self.get_modifier_code(modifier2_name) if modifier2_name != '无' else 0

        modifier = modifier1 | modifier2
        key = self.get_key_code(key_name)

        # 保存配置
        if self.config_manager.set_hotkey(modifier, key, modifier1_name, modifier2_name, key_name):
            hotkey_text = f"{modifier1_name} + {key_name}" if modifier2_name == '无' else f"{modifier1_name} + {modifier2_name} + {key_name}"
            messagebox.showinfo("成功", f"快捷键已设置为: {hotkey_text}")
            # 通知主程序更新快捷键
            self.on_hotkey_changed()
            self.dialog.destroy()
        else:
            messagebox.showerror("错误", "保存设置失败")

    def apply_settings_silent(self):
        """静默应用设置（用于恢复默认）"""
        modifier1_name = self.modifier1_var.get()  # pyright: ignore[reportOptionalMemberAccess]
        modifier2_name = self.modifier2_var.get()  # pyright: ignore[reportOptionalMemberAccess]
        key_name = self.key_var.get()  # pyright: ignore[reportOptionalMemberAccess]

        # 计算修饰键组合值
        modifier1 = self.get_modifier_code(modifier1_name)
        modifier2 = self.get_modifier_code(modifier2_name) if modifier2_name != '无' else 0

        modifier = modifier1 | modifier2
        key = self.get_key_code(key_name)

        # 保存配置
        if self.config_manager.set_hotkey(modifier, key, modifier1_name, modifier2_name, key_name):
            # 通知主程序更新快捷键
            self.on_hotkey_changed()
            self.dialog.destroy()
        else:
            messagebox.showerror("错误", "保存设置失败")
    
    def restore_default(self):
        """恢复默认快捷键设置"""
        if messagebox.askyesno("确认", "确定要恢复默认快捷键吗？\n\n默认快捷键: Ctrl + Q"):
            # 设置为默认值
            self.modifier1_var.set('Ctrl')  # pyright: ignore[reportOptionalMemberAccess]
            self.modifier2_var.set('无')  # pyright: ignore[reportOptionalMemberAccess]
            self.key_var.set('Q')  # pyright: ignore[reportOptionalMemberAccess]
            # 更新预览
            self.update_hotkey_preview()
            # 静默应用设置（不弹出额外的确认框）
            self.apply_settings_silent()

# --- GUI 界面模块 ---

class TopMostApp:
    def __init__(self, root):
        self.root = root
        self.root.title("窗口置顶工具")
        self.root.geometry("650x400")
        self.root.minsize(650, 400)
        
        # 设置窗口图标
        icon_path = get_resource_path(os.path.join("icon", "app_icon.ico"))
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except Exception as e:
                print(f"设置图标失败: {e}")
        
        # 初始化配置管理器
        self.config_manager = ConfigManager()
        
        # 设置样式
        style = ttk.Style()
        style.configure("Treeview", rowheight=25)
        style.configure("Bold.TLabel", font=("微软雅黑", 10, "bold"))
        
        self.setup_ui()

        # self.center_dialog()
            
        # 启动快捷键监听
        self.hotkey_listener = HotkeyListener(self.on_hotkey_triggered, self.config_manager)
        self.hotkey_listener.start()
        
        # 初始刷新
        self.refresh_list()

    def center_dialog(self):
        """将对话框居中显示"""
        self.root.update()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")

    def setup_ui(self):
        # 主容器
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # 顶部说明区域
        top_frame = ttk.Frame(main_container, padding="10")
        top_frame.pack(fill=tk.X)
        
        ttk.Label(top_frame, text="窗口置顶管理工具", font=("微软雅黑", 16, "bold")).pack(side=tk.LEFT)
        current_hotkey = self.config_manager.get_hotkey_display()
        self.hotkey_label = ttk.Label(top_frame, text=f"支持快捷键: {current_hotkey} (置顶窗口)", foreground="gray")
        self.hotkey_label.pack(side=tk.RIGHT, padx=10)
        
        # 中间列表区域
        list_frame = ttk.Frame(main_container, padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ("hwnd", "title", "status")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="browse")
        
        self.tree.heading("hwnd", text="句柄")
        self.tree.column("hwnd", width=80, anchor="center")
        
        self.tree.heading("title", text="窗口标题")
        self.tree.column("title", width=450)
        
        self.tree.heading("status", text="状态")
        self.tree.column("status", width=100, anchor="center")
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 绑定双击事件
        self.tree.bind("<Double-1>", lambda e: self.toggle_selected())
        
        # 底部按钮区域
        btn_frame = ttk.Frame(main_container, padding="10")
        btn_frame.pack(fill=tk.X)
        
        ttk.Button(btn_frame, text="刷新列表 (F5)", command=self.refresh_list).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="切换置顶状态", command=self.toggle_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="设置", command=self.open_settings).pack(side=tk.LEFT, padx=5)
        ttk.Label(btn_frame, text="提示: 双击列表项也可切换状态", foreground="gray").pack(side=tk.LEFT, padx=20)
        
        # 绑定 F5 刷新
        self.root.bind("<F5>", lambda e: self.refresh_list())

    def refresh_list(self):
        # 记录当前选中的项，以便刷新后恢复
        selected_hwnd = None
        selection = self.tree.selection()
        if selection:
            item = self.tree.item(selection[0])
            selected_hwnd = item['values'][0]

        # 清空列表
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # 获取所有窗口
        windows = WindowUtils.get_all_windows()
        
        for hwnd, title in windows:
            is_top = WindowUtils.is_window_topmost(hwnd)
            status = "📌 已置顶" if is_top else "❌ 未置顶"
            
            # 插入数据
            item_id = self.tree.insert("", "end", values=(hwnd, title, status))
            
            # 恢复选中
            if selected_hwnd and str(hwnd) == str(selected_hwnd):
                self.tree.selection_set(item_id)
                self.tree.see(item_id)

    def toggle_selected(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("提示", "请先选择一个窗口")
            return
            
        item = self.tree.item(selection[0])
        hwnd = int(item['values'][0])
        title = item['values'][1]
        
        self.toggle_window_state(hwnd, title)

    def toggle_window_state(self, hwnd, title=None):
        if not win32gui.IsWindow(hwnd):
            messagebox.showerror("错误", "该窗口已不存在")
            self.refresh_list()
            return

        is_top = WindowUtils.is_window_topmost(hwnd)
        new_state = not is_top
        
        if WindowUtils.set_window_topmost(hwnd, new_state):
            # 播放效果
            WindowUtils.show_effect(hwnd, 'pin' if new_state else 'unpin')
            # 刷新列表显示
            self.refresh_list()
            
            # 在底部状态栏显示结果（可选）
            state_str = "置顶" if new_state else "取消置顶"
            print(f"已{state_str}: {title}")

    def on_hotkey_triggered(self):
        """快捷键触发时的回调"""
        hwnd = WindowUtils.get_foreground_window()
        if hwnd:
            # 这里的操作需要在主线程中更新 GUI 吗？
            # refresh_list 包含 GUI 操作，建议使用 after
            # toggle_window_state 主要是 win32 api 调用，比较安全，但为了刷新列表，还是用 after
            self.root.after(0, lambda: self.toggle_window_state(hwnd, win32gui.GetWindowText(hwnd)))
    
    def open_settings(self):
        """打开设置对话框"""
        SettingsDialog(self.root, self.config_manager, self.on_hotkey_settings_changed)
    
    def on_hotkey_settings_changed(self):
        """快捷键设置改变后的回调"""
        # 停止旧的快捷键监听
        print("正在停止旧的快捷键监听...")
        self.hotkey_listener.stop()
        
        # 等待旧线程结束
        print("等待旧线程结束...")
        time.sleep(0.5)
        
        # 创建新的快捷键监听
        print("启动新的快捷键监听...")
        self.hotkey_listener = HotkeyListener(self.on_hotkey_triggered, self.config_manager)
        self.hotkey_listener.start()
        
        # 更新界面上的快捷键提示
        current_hotkey = self.config_manager.get_hotkey_display()
        self.hotkey_label.config(text=f"支持快捷键: {current_hotkey} (置顶窗口)")
        print(f"新快捷键已设置: {current_hotkey}")
        
        # 找到顶部的标签并更新文本
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Frame):
                for child in widget.winfo_children():
                    if isinstance(child, ttk.Label) and "支持快捷键" in str(child.cget("text")):
                        child.config(text=f"支持快捷键: {current_hotkey} (置顶窗口)")
                        break

def main():
    root = tk.Tk()
    app = TopMostApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
