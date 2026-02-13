"""
窗口置顶工具 - 使用 Ctrl+K 快捷键
按 Ctrl+K 切换当前窗口的置顶状态
"""

import win32gui
import win32con
import time
import threading
import ctypes
from ctypes import windll, byref, c_int, c_uint, c_void_p, Structure
import win32api

class WindowTopper:
    def __init__(self):
        self.running = True
        self.topmost_windows = set()
        self.lock = threading.Lock()
        # 配置选项
        self.config = {
            'hotkey': {'modifier': 0x0002, 'key': ord('K')},  # Ctrl+K
            'debug': False  # 是否显示调试信息
        }
    
    def set_hotkey(self, modifier, key):
        """设置快捷键
        
        Args:
            modifier: 修饰键 (如 0x0002 表示 Ctrl)
            key: 键码 (如 ord('K') 表示 K 键)
        """
        self.config['hotkey']['modifier'] = modifier
        self.config['hotkey']['key'] = key
        # 重新注册快捷键
        self.unregister_hotkey()
        return self.register_hotkey()
    
    def is_topmost(self, hwnd):
        """检查窗口是否置顶"""
        try:
            exstyle = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            return bool(exstyle & win32con.WS_EX_TOPMOST)
        except Exception as e:
            print(f"[DEBUG] 检查窗口置顶状态错误: {e}")
            return False
    
    def toggle_topmost(self, hwnd):
        """切换窗口置顶状态"""
        try:
            title = win32gui.GetWindowText(hwnd)
            is_topmost = self.is_topmost(hwnd)
            
            if is_topmost:
                # 取消置顶 - 使用向左上晃动
                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST,
                                     0, 0, 0, 0,
                                     win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                with self.lock:
                    if hwnd in self.topmost_windows:
                        self.topmost_windows.remove(hwnd)
                self.show_effect(hwnd, mode='unpin')  # 取消置顶效果
                print(f"✓ [{title[:40]}] 已取消置顶")
                return False
            else:
                # 设置置顶 - 使用向右下晃动
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST,
                                     0, 0, 0, 0,
                                     win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                with self.lock:
                    self.topmost_windows.add(hwnd)
                self.show_effect(hwnd, mode='pin')  # 置顶效果
                print(f"✓ [{title[:40]}] 已置顶")
                return True
        except Exception as e:
            print(f"✗ 切换窗口置顶状态错误: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def get_foreground_window(self):
        """获取当前激活窗口"""
        try:
            hwnd = win32gui.GetForegroundWindow()
            if hwnd and windll.user32.IsWindow(hwnd) and windll.user32.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title and len(title) > 0:
                    return hwnd
            return None
        except Exception as e:
            print(f"[DEBUG] 获取当前激活窗口错误: {e}")
            return None
    
    def register_hotkey(self):
        """注册快捷键"""
        try:
            modifier = self.config['hotkey']['modifier']
            key = self.config['hotkey']['key']
            result = windll.user32.RegisterHotKey(
                None,    # hWnd=None, 消息会发送到线程消息队列
                1,       # id
                modifier,  # 修饰键
                key  # 键码
            )
            success = result != 0
            if not success:
                if self.config['debug']:
                    print(f"[DEBUG] 快捷键注册失败，返回值: {result}")
            return success
        except Exception as e:
            if self.config['debug']:
                print(f"[DEBUG] 注册快捷键错误: {e}")
            return False
    
    def unregister_hotkey(self):
        """注销快捷键"""
        try:
            result = windll.user32.UnregisterHotKey(None, 1)
            if not result:
                print(f"[DEBUG] 快捷键注销失败，返回值: {result}")
        except Exception as e:
            print(f"[DEBUG] 注销快捷键错误: {e}")
    
    def show_effect(self, hwnd, mode='pin'):
        """显示置顶窗口的提醒效果
        
        Args:
            mode: 'pin' - 置顶效果(向右下晃动), 'unpin' - 取消置顶效果(向左上晃动)
        """
        try:
            rect = win32gui.GetWindowRect(hwnd)
            x, y = rect[0], rect[1]
            
            if mode == 'pin':
                def move_effect():
                    try:
                        for offset in [5, 10, 5, 0]:
                            win32gui.SetWindowPos(hwnd, 0, x + offset, y + offset, 0, 0,
                                                win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)
                            time.sleep(0.08)
                    except Exception as e:
                        print(f"[DEBUG] 移动效果错误: {e}")
            else:
                def move_effect():
                    try:
                        for offset in [5, 10, 5, 0]:
                            win32gui.SetWindowPos(hwnd, 0, x - offset, y - offset, 0, 0,
                                                win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)
                            time.sleep(0.08)
                    except Exception as e:
                        print(f"[DEBUG] 移动效果错误: {e}")
            
            move_thread = threading.Thread(target=move_effect, daemon=True)
            move_thread.start()
            
            for i in range(3):
                win32gui.FlashWindow(hwnd, True)
                time.sleep(0.15)
                win32gui.FlashWindow(hwnd, False)
                time.sleep(0.15)
            
        except Exception as e:
            print(f"[DEBUG] 提醒效果错误: {e}")
    
    def message_loop(self):
        """消息循环"""
        # 定义MSG结构
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
            result = windll.user32.GetMessageA(
                byref(msg),
                None,
                0,
                0
            )
            
            if result == 0:
                break
            if result == -1:
                time.sleep(0.1)
                continue
            
            if msg.message == win32con.WM_HOTKEY:
                hwnd = self.get_foreground_window()
                if hwnd:
                    self.toggle_topmost(hwnd)
                else:
                    print("✗ 无法获取当前窗口")
            
            windll.user32.TranslateMessage(byref(msg))
            windll.user32.DispatchMessageA(byref(msg))
    
    def cleanup(self):
        """清理所有资源"""
        print("\n正在清理资源...")
        
        # 注销快捷键
        self.unregister_hotkey()
        
        # 清空集合
        with self.lock:
            self.topmost_windows.clear()
        
        print("✓ 资源清理完成")
    
    def run(self):
        """运行主程序"""
        print("=" * 60)
        print("窗口置顶工具 - Ctrl+K 版本")
        print("=" * 60)
        print()
        print("使用说明:")
        print("  1. 点击任意窗口使其获得焦点")
        print("  2. 按 Ctrl+K 切换该窗口的置顶状态")
        print()
        print("正在注册快捷键 Ctrl+K...")
        
        if not self.register_hotkey():
            print("✗ 快捷键注册失败")
            print("  可能原因:")
            print("  - 其他程序已占用 Ctrl+K")
            print("  - 需要管理员权限")
            input("\n按回车退出...")
            return
        
        print("✓ 快捷键注册成功!")
        print()
        print("程序正在运行中...")
        print("按 Ctrl+C 退出程序")
        print("=" * 60)
        print()
        
        try:
            self.message_loop()
        except KeyboardInterrupt:
            print("\n正在退出...")
            self.cleanup()
            print("✓ 程序已退出")

def main():
    app = WindowTopper()
    app.run()

if __name__ == "__main__":
    main()
