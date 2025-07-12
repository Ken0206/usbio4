import tkinter as tk
from tkinter import messagebox
import sys
import config
import customtkinter as ctk

try:
    import win32com.client
    import pywintypes
    # 導入實現單例應用所需的核心模組
    import win32event
    import win32api
    import winerror
except ImportError:
    messagebox.showerror("缺少函式庫", "找不到 'pywin32' 函式庫。\n請執行 'pip install pywin32' 來安裝。")
    sys.exit()


# =============================================================================
# 0. 短暫通知視窗 (Toast Notification)
#    一個無邊框、會自動消失的 Toplevel 視窗。
# =============================================================================
class ToastNotification(ctk.CTkToplevel):
    def __init__(self, message):
        super().__init__()

        # 隱藏視窗邊框和標題列
        self.overrideredirect(True)
        self.attributes("-topmost", True)  # 確保在最上層顯示

        # 設定外觀
        self.configure(fg_color="#333333")
        label = ctk.CTkLabel(self, text=message, font=ctk.CTkFont(family="Microsoft JhengHei", size=14),
                             padx=20, pady=10)
        label.pack()

        # 計算並置中於螢幕
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width // 2) - (self.winfo_width() // 2)
        y = (screen_height // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")

        # 3秒後自動銷毀自己和父視窗（Tk()）
        self.after(3000, self.master.destroy)


# =============================================================================
# 1. HardwareController Class (維持不變)
# =============================================================================
class HardwareController:
    """專門處理與 USB I/O 硬體溝通的類別"""

    def __init__(self, prog_id):
        self.prog_id = prog_id
        self.usbio = None

    def connect(self):
        try:
            self.usbio = win32com.client.Dispatch(self.prog_id)
            if self.usbio.OpenUsbDevice(4660, 26505): return True
            self.usbio = None
            return False
        except (pywintypes.com_error, ConnectionError):
            self.usbio = None
            return False

    def disconnect(self):
        if self.usbio:
            try:
                self.update_leds([False] * config.NUM_LEDS)
                self.usbio.CloseUsbDevice()
            except (pywintypes.com_error, ConnectionError):
                pass
            finally:
                self.usbio = None

    def is_connected(self):
        if not self.usbio: return False
        try:
            return self.usbio.OpenUsbDevice(4660, 26505)
        except (pywintypes.com_error, ConnectionError):
            return False

    def update_leds(self, led_states):
        if not self.usbio: raise ConnectionError("硬體未連接")
        red_mask = sum(1 << i for i, state in enumerate(led_states[:8]) if state)
        self.usbio.OutDataCtrl(red_mask, 32)
        self.usbio.OutDataCtrl(red_mask, 48)
        green_mask = sum(1 << i for i, state in enumerate(led_states[8:]) if state)
        self.usbio.OutDataCtrl(green_mask, 0)


# =============================================================================
# 2. UI 主應用程式 (LedControlApp Class) (維持不變)
# =============================================================================
class LedControlApp:
    def __init__(self, root):
        self.root = root
        self.hardware = HardwareController(config.PROG_ID)
        self.led_states = [False] * config.NUM_LEDS
        self.is_first_connection = True

        self.marquee_running = False
        self.marquee_pos = 0
        self.marquee_dir = 1
        self.led_states_before_marquee = []

        self.after_id_poll = None
        self.after_id_marquee = None

        self.setup_ui()
        self.poll_hardware_status()

    def setup_ui(self):
        self.root.title("USB I/O LED 控制器 (v1.2)")
        self.root.geometry(f"{config.WINDOW_DEFAULT_WIDTH}x{config.WINDOW_DEFAULT_HEIGHT}")
        self.root.resizable(True, True)

        self.root.grid_rowconfigure(0, weight=0)
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_rowconfigure(2, weight=0)
        self.root.grid_columnconfigure(0, weight=1)

        status_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        status_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(5, 0))
        self.status_label = ctk.CTkLabel(status_frame, text="初始化...",
                                         font=ctk.CTkFont(family="Microsoft JhengHei", size=14, weight="bold"))
        self.status_label.pack(side=tk.LEFT)

        self.led_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        self.led_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        self.create_led_buttons()

        self.func_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        self.func_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(5, 10))
        self.create_function_buttons()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_led_buttons(self):
        self.buttons = []
        for i in range(config.NUM_LEDS):
            #button = ctk.CTkButton(self.led_frame, text=f"LED {i + 1}",
            button = ctk.CTkButton(self.led_frame, text=f"{i + 1}",
                                   command=lambda index=i: self.toggle_led(index),
                                   corner_radius=8, font=ctk.CTkFont(weight="bold"))
            row, col = i // 8, 7 - (i % 8)
            button.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")
            self.buttons.append(button)

        for i in range(8): self.led_frame.grid_columnconfigure(i, weight=1)
        for i in range(2): self.led_frame.grid_rowconfigure(i, weight=1)
        self.update_gui_from_states()
        self.disable_buttons(self.buttons)

    def create_function_buttons(self):
        self.func_buttons = []
        btn_config = {"fg_color": config.COLOR_BUTTON_FUNC, "hover_color": config.COLOR_BUTTON_FUNC_HOVER,
                      "text_color": config.COLOR_BUTTON_FUNC_TEXT, "corner_radius": 8,
                      "font": ctk.CTkFont(family="Microsoft JhengHei", weight="bold")}
        btn_all_on = ctk.CTkButton(self.func_frame, text="全部開啟", **btn_config, command=self.all_on)
        btn_all_off = ctk.CTkButton(self.func_frame, text="全部關閉", **btn_config, command=self.all_off)
        btn_invert = ctk.CTkButton(self.func_frame, text="反向", **btn_config, command=self.invert_state)
        self.btn_marquee = ctk.CTkButton(self.func_frame, text="跑馬燈", **btn_config, command=self.toggle_marquee)
        self.func_buttons = [btn_all_on, btn_all_off, btn_invert, self.btn_marquee]
        for i, btn in enumerate(self.func_buttons):
            self.func_frame.grid_columnconfigure(i, weight=1)
            btn.grid(row=0, column=i, padx=5, sticky="ew")
        self.disable_buttons(self.func_buttons)

    def poll_hardware_status(self):
        if self.hardware.is_connected():
            if self.is_first_connection: self.handle_connection_success()
        else:
            if not self.is_first_connection: self.handle_disconnection()
            if self.hardware.connect(): self.handle_connection_success()
        self.after_id_poll = self.root.after(config.POLL_INTERVAL_MS, self.poll_hardware_status)

    def handle_connection_success(self):
        self.status_label.configure(text="硬體已連接", text_color="lightgreen")
        self.enable_buttons(self.buttons + self.func_buttons)
        if self.is_first_connection:
            self.reset_all_leds()
            self.is_first_connection = False
        else:
            if self.marquee_running:
                self.run_marquee_frame()
            else:
                self.restore_led_state()

    def handle_disconnection(self):
        if self.is_first_connection: return
        print("偵測到硬體斷開！")
        if self.after_id_marquee: self.root.after_cancel(self.after_id_marquee)
        self.is_first_connection = True
        self.hardware.disconnect()
        self.disable_buttons(self.buttons + self.func_buttons)
        self.status_label.configure(text="硬體已中斷，正在嘗試重連...", text_color="orange")

    def disable_buttons(self, button_list):
        for btn in button_list: btn.configure(state=tk.DISABLED)

    def enable_buttons(self, button_list):
        for btn in button_list: btn.configure(state=tk.NORMAL)

    def update_gui_from_states(self):
        for i, state in enumerate(self.led_states):
            color = config.COLOR_LED_ON if state else config.COLOR_LED_OFF
            text_color = "black" if state else "white"
            hover_color = config.COLOR_LED_HOVER_ON if state else config.COLOR_LED_HOVER_OFF
            self.buttons[i].configure(fg_color=color, text_color=text_color, hover_color=hover_color)

    def _update_states_and_hardware(self, new_states):
        self.led_states = new_states
        self.update_gui_from_states()
        try:
            if self.hardware.is_connected():
                self.hardware.update_leds(self.led_states)
        except (pywintypes.com_error, ConnectionError):
            self.handle_disconnection()

    def reset_all_leds(self):
        self._update_states_and_hardware([False] * config.NUM_LEDS)

    def restore_led_state(self):
        self._update_states_and_hardware(self.led_states)

    def _execute_action(self, action):
        if self.marquee_running: self.stop_marquee()
        action()
        self._update_states_and_hardware(self.led_states)

    def toggle_led(self, index):
        self._execute_action(lambda: self.led_states.__setitem__(index, not self.led_states[index]))

    def all_on(self):
        self._execute_action(lambda: self.led_states.__init__([True] * config.NUM_LEDS))

    def all_off(self):
        self._execute_action(lambda: self.led_states.__init__([False] * config.NUM_LEDS))

    def invert_state(self):
        self._execute_action(lambda: self.led_states.__init__([not s for s in self.led_states]))

    def toggle_marquee(self):
        if self.marquee_running:
            self.stop_marquee()
        else:
            self.start_marquee()

    def start_marquee(self):
        print("跑馬燈")
        self.marquee_running = True
        self.btn_marquee.configure(text="停止跑馬", fg_color=config.COLOR_MARQUEE_STOP,
                                   hover_color=config.COLOR_MARQUEE_STOP_HOVER)
        self.marquee_pos = 0;
        self.marquee_dir = 1
        self.disable_buttons([b for b in self.buttons + self.func_buttons if b != self.btn_marquee])
        self.run_marquee_frame()

    def stop_marquee(self):
        if not self.marquee_running: return
        print("停止跑馬")
        self.marquee_running = False
        if self.after_id_marquee: self.root.after_cancel(self.after_id_marquee)

        last_states = [False] * config.NUM_LEDS
        last_pos = self.marquee_pos - self.marquee_dir
        if 0 <= last_pos < config.NUM_LEDS:
            last_states[last_pos] = True
        self.led_states = last_states
        print(f"跑馬燈凍結在狀態: {self.led_states}")

        self.btn_marquee.configure(text="跑馬燈", fg_color=config.COLOR_BUTTON_FUNC,
                                   hover_color=config.COLOR_BUTTON_FUNC_HOVER)
        self.enable_buttons(self.buttons + self.func_buttons)
        self.update_gui_from_states()

    def run_marquee_frame(self):
        if not self.marquee_running or not self.hardware.is_connected(): return
        temp_states = [False] * config.NUM_LEDS
        temp_states[self.marquee_pos] = True
        try:
            self.hardware.update_leds(temp_states)
            for i, state in enumerate(temp_states):
                color = config.COLOR_LED_ON if state else config.COLOR_LED_OFF
                text_color = "black" if state else "white"
                hover_color = config.COLOR_LED_HOVER_ON if state else config.COLOR_LED_HOVER_OFF
                self.buttons[i].configure(fg_color=color, text_color=text_color, hover_color=hover_color)
            if self.marquee_pos >= config.NUM_LEDS - 1 and self.marquee_dir == 1:
                self.marquee_dir = -1
            elif self.marquee_pos <= 0 and self.marquee_dir == -1:
                self.marquee_dir = 1
            self.marquee_pos += self.marquee_dir
            self.after_id_marquee = self.root.after(config.MARQUEE_SPEED_MS, self.run_marquee_frame)
        except (pywintypes.com_error, ConnectionError):
            self.handle_disconnection()

    def on_closing(self):
        if messagebox.askokcancel("退出", "確定要關閉程式嗎？"):
            if self.after_id_poll: self.root.after_cancel(self.after_id_poll)
            if self.after_id_marquee: self.root.after_cancel(self.after_id_marquee)
            self.hardware.disconnect()
            self.root.destroy()


# =============================================================================
# 3. 主程式啟動入口
#    加入了 Mutex 檢查以實現單例應用
# =============================================================================
if __name__ == "__main__":
    if sys.platform != "win32":
        messagebox.showerror("系統不支援", "此程式只能在 Windows 作業系統上運行。")
        sys.exit()

    # 建立一個唯一的 Mutex 名稱，確保所有實例都使用同一個名稱
    # 這個 GUID 是隨機產生的，可以替換成任何你喜歡的唯一字串
    mutex_name = "LEDControllerApp_Mutex_{c28f3435-9B2A-4f3d-8F3E-6B8E7A6B295A}"
    mutex = None
    try:
        # 嘗試建立 Mutex
        mutex = win32event.CreateMutex(None, False, mutex_name)

        # 檢查上一個操作的錯誤碼
        last_error = win32api.GetLastError()

        # 如果錯誤碼是 ERROR_ALREADY_EXISTS，表示已有實例在運行
        if last_error == winerror.ERROR_ALREADY_EXISTS:
            print("偵測到程式已在執行中。")
            # 建立一個臨時的根視窗來顯示通知
            temp_root = tk.Tk()
            temp_root.withdraw()  # 隱藏這個臨時的根視窗
            ToastNotification("程式已在執行中...").mainloop()
            sys.exit()

        # 如果沒有錯誤，代表這是第一個實例，正常啟動主程式
        else:
            print("啟動主程式...")
            ctk.set_appearance_mode(config.UI_THEME)
            ctk.set_default_color_theme("blue")
            root = ctk.CTk()
            app = LedControlApp(root)
            root.mainloop()

    finally:
        # 無論如何，程式結束時都要確保釋放 Mutex 句柄
        if mutex:
            win32api.CloseHandle(mutex)
            print("Mutex 已釋放。")