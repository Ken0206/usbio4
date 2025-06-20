# 程式碼由 2025/06/20 Google AI Studio 生成
# USB I/O board : 電腦硬體裝修乙級術科第一題套件 2023/09/01 12000-102201~210 版本
# 必用 python 32 bit 因為 USBIO4.dll 是 32 bit

import tkinter as tk
from tkinter import messagebox
import sys
import config  # 導入美化版的設定檔
import customtkinter as ctk  # 導入 customtkinter

try:
    import win32com.client
    import pywintypes
except ImportError:
    messagebox.showerror("缺少函式庫", "找不到 'pywin32' 函式庫。\n請執行 'pip install pywin32' 來安裝。")
    sys.exit()


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
# 2. UI 主應用程式 (LedControlApp Class)
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
        # --- 不再需要備份狀態，但保留變數以防未來需要 ---
        self.led_states_before_marquee = []

        self.after_id_poll = None
        self.after_id_marquee = None

        self.setup_ui()
        self.poll_hardware_status()

    def setup_ui(self):
        self.root.title("USB I/O LED 控制器")
        self.root.geometry(f"{config.WINDOW_DEFAULT_WIDTH}x{config.WINDOW_DEFAULT_HEIGHT}")
        self.root.resizable(True, True)

        status_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        status_frame.pack(fill=tk.X, padx=10, pady=5)
        self.status_label = ctk.CTkLabel(status_frame, text="初始化...   請接上USB I/O 板...",
                                         font=ctk.CTkFont(family="Microsoft JhengHei", size=14, weight="bold"))
        self.status_label.pack(side=tk.LEFT)

        self.led_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        self.led_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.create_led_buttons()

        func_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        func_frame.pack(fill=tk.X, pady=(5, 10), padx=10)
        self.create_function_buttons(func_frame)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_led_buttons(self):
        self.buttons = []
        for i in range(config.NUM_LEDS):
            button = ctk.CTkButton(self.led_frame, text=f"LED {i + 1}",
                                   command=lambda index=i: self.toggle_led(index),
                                   corner_radius=8, font=ctk.CTkFont(weight="bold"))
            row, col = i // 8, 7 - (i % 8)
            button.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")
            self.buttons.append(button)

        for i in range(8): self.led_frame.grid_columnconfigure(i, weight=1)
        for i in range(2): self.led_frame.grid_rowconfigure(i, weight=1)
        self.update_gui_from_states()
        self.disable_buttons(self.buttons)

    def create_function_buttons(self, parent_frame):
        self.func_buttons = []
        btn_config = {"fg_color": config.COLOR_BUTTON_FUNC, "hover_color": config.COLOR_BUTTON_FUNC_HOVER,
                      "text_color": config.COLOR_BUTTON_FUNC_TEXT, "corner_radius": 8,
                      "font": ctk.CTkFont(family="Microsoft JhengHei", weight="bold")}
        btn_all_on = ctk.CTkButton(parent_frame, text="全部開啟", **btn_config, command=self.all_on)
        btn_all_off = ctk.CTkButton(parent_frame, text="全部關閉", **btn_config, command=self.all_off)
        btn_invert = ctk.CTkButton(parent_frame, text="反向", **btn_config, command=self.invert_state)
        self.btn_marquee = ctk.CTkButton(parent_frame, text="跑馬燈", **btn_config, command=self.toggle_marquee)
        self.func_buttons = [btn_all_on, btn_all_off, btn_invert, self.btn_marquee]
        for btn in self.func_buttons: btn.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
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
                print("重新連接，恢復跑馬燈...")
                self.run_marquee_frame()
            else:
                print("重新連接，恢復一般狀態...")
                self.restore_led_state()

    def handle_disconnection(self):
        if self.is_first_connection: return
        print("偵測到硬體斷開！")
        if self.after_id_marquee:
            self.root.after_cancel(self.after_id_marquee)
            self.after_id_marquee = None
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
        print(f"從內部狀態恢復硬體顯示: {self.led_states}")
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
        print("啟動跑馬燈")
        # 不需要備份了，因為我們就是要凍結
        # self.led_states_before_marquee = self.led_states.copy()
        self.marquee_running = True
        self.btn_marquee.configure(text="停止跑馬燈", fg_color=config.COLOR_MARQUEE_STOP,
                                   hover_color=config.COLOR_MARQUEE_STOP_HOVER)
        self.marquee_pos = 0;
        self.marquee_dir = 1
        self.disable_buttons([b for b in self.buttons + self.func_buttons if b != self.btn_marquee])
        self.run_marquee_frame()

    def stop_marquee(self):
        if not self.marquee_running: return
        print("停止跑馬燈")

        # --- 核心修改 ---
        # 1. 停止動畫循環
        self.marquee_running = False
        if self.after_id_marquee:
            self.root.after_cancel(self.after_id_marquee)
            self.after_id_marquee = None

        # 2. 恢復按鈕功能和外觀
        self.btn_marquee.configure(text="跑馬燈", fg_color=config.COLOR_BUTTON_FUNC,
                                   hover_color=config.COLOR_BUTTON_FUNC_HOVER)
        self.enable_buttons(self.buttons + self.func_buttons)

        # 3. 將內部的 led_states 更新為跑馬燈的最後一幀狀態
        #    這樣後續的操作（如斷線重連）就會基於這個凍結的狀態
        last_states = [False] * config.NUM_LEDS
        # marquee_pos 已經移動到下一個位置，所以要減回去
        last_pos = self.marquee_pos - self.marquee_dir
        if 0 <= last_pos < config.NUM_LEDS:
            last_states[last_pos] = True
        self.led_states = last_states
        print(f"跑馬燈凍結在狀態: {self.led_states}")

        # 4. 不再從備份恢復，因為就是要停在當下

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


# --- 主程式啟動 ---
if __name__ == "__main__":
    if sys.platform != "win32":
        messagebox.showerror("系統不支援", "此程式只能在 Windows 作業系統上運行。")
    else:
        ctk.set_appearance_mode(config.UI_THEME)
        ctk.set_default_color_theme("blue")
        root = ctk.CTk()
        app = LedControlApp(root)
        root.mainloop()