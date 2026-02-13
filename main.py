import win32gui
import win32con
import win32api
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
from ctypes import windll, byref, c_int, c_void_p, Structure

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

# --- 快捷键监听模块 ---

class HotkeyListener(threading.Thread):
    def __init__(self, callback):
        super().__init__(daemon=True)
        self.callback = callback
        self.running = True
        self.modifier = 0x0002  # Ctrl
        self.key = ord('Q')     # Q

        self.hotkey_id = 1
    def run(self):
        # 注册快捷键
        try:
            if not windll.user32.RegisterHotKey(None, self.hotkey_id, self.modifier, self.key):
                print("快捷键 Ctrl+K 注册失败")
                return
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
        windll.user32.UnregisterHotKey(None, self.hotkey_id)

    def stop(self):
        self.running = False
        # 发送一个空消息来打破 GetMessage 的阻塞
        windll.user32.PostThreadMessageA(self.ident, win32con.WM_NULL, 0, 0)

# --- GUI 界面模块 ---

class TopMostApp:
    def __init__(self, root):
        self.root = root
        self.root.title("窗口置顶工具 v2.0")
        self.root.geometry("800x500")
        
        # 设置样式
        style = ttk.Style()
        style.configure("Treeview", rowheight=25)
        style.configure("Bold.TLabel", font=("微软雅黑", 10, "bold"))
        
        self.setup_ui()
        
        # 启动快捷键监听
        self.hotkey_listener = HotkeyListener(self.on_hotkey_triggered)
        self.hotkey_listener.start()
        
        # 初始刷新
        self.refresh_list()

    def setup_ui(self):
        # 顶部说明区域
        top_frame = ttk.Frame(self.root, padding="10")
        top_frame.pack(fill=tk.X)
        
        ttk.Label(top_frame, text="窗口置顶管理工具", font=("微软雅黑", 16, "bold")).pack(side=tk.LEFT)
        ttk.Label(top_frame, text="支持快捷键: Ctrl +  Q (置顶窗口)", foreground="gray").pack(side=tk.RIGHT, padx=10)
        
        # 中间列表区域
        list_frame = ttk.Frame(self.root, padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ("hwnd", "title", "status")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="browse")
        
        self.tree.heading("hwnd", text="句柄")
        self.tree.column("hwnd", width=80, anchor="center")
        
        self.tree.heading("title", text="窗口标题")
        self.tree.column("title", width=500)
        
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
        btn_frame = ttk.Frame(self.root, padding="10")
        btn_frame.pack(fill=tk.X)
        
        ttk.Button(btn_frame, text="刷新列表 (F5)", command=self.refresh_list).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="切换置顶状态", command=self.toggle_selected).pack(side=tk.LEFT, padx=5)
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

def main():
    root = tk.Tk()
    app = TopMostApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
