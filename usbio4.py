import tkinter as tk
import sys
import random
import config
import customtkinter as ctk

try:
    import win32com.client
    import pywintypes
    import win32event
    import win32api
    import winerror
except ImportError:
    root = tk.Tk()
    root.withdraw()
    tk.messagebox.showerror(title="缺少函式庫",
                            message="找不到 'pywin32' 函式庫。\n請在命令提示字元執行 'pip install pywin32' 來安裝。")
    sys.exit()


# =============================================================================
# 輔助 UI 類別 (保持不變)
# =============================================================================
class CustomMessageBox(ctk.CTkToplevel):
    def __init__(self, title, message):
        super().__init__()
        self.title(title);
        self.lift();
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self.on_cancel);
        self.result = False
        self.geometry("300x150");
        self.resizable(False, False)
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(expand=True, fill="both", padx=20, pady=20)
        main_frame.grid_rowconfigure(0, weight=1);
        main_frame.grid_columnconfigure((0, 1), weight=1)
        label = ctk.CTkLabel(main_frame, text=message, font=ctk.CTkFont(family="Microsoft JhengHei", size=14))
        label.grid(row=0, column=0, columnspan=2, sticky="nsew")
        ok_button = ctk.CTkButton(main_frame, text="確定", command=self.on_ok)
        ok_button.grid(row=1, column=0, padx=(0, 5), pady=10, sticky="ew")
        cancel_button = ctk.CTkButton(main_frame, text="取消", fg_color="#555555", hover_color="#777777",
                                      command=self.on_cancel)
        cancel_button.grid(row=1, column=1, padx=(5, 0), pady=10, sticky="ew")
        self.center_window();
        self.grab_set()

    def center_window(self):
        self.update_idletasks()
        try:
            x = self.master.winfo_x() + (self.master.winfo_width() // 2) - (self.winfo_width() // 2)
            y = self.master.winfo_y() + (self.master.winfo_height() // 2) - (self.winfo_height() // 2)
            self.geometry(f"+{x}+{y}")
        except Exception:
            sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
            x, y = (sw // 2) - (self.winfo_width() // 2), (sh // 2) - (self.winfo_height() // 2)
            self.geometry(f"+{x}+{y}")

    def on_ok(self):
        self.result = True; self.destroy()

    def on_cancel(self):
        self.result = False; self.destroy()

    def get(self):
        self.master.wait_window(self); return self.result


class ToastNotification(ctk.CTkToplevel):
    def __init__(self, message, master=None):
        super().__init__(master)
        self.overrideredirect(True);
        self.attributes("-topmost", True);
        self.configure(fg_color="#333333")
        label = ctk.CTkLabel(self, text=message, font=ctk.CTkFont(family="Microsoft JhengHei", size=14), padx=20,
                             pady=10)
        label.pack()
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        x, y = (sw // 2) - (self.winfo_width() // 2), (sh // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}");
        self.after(2500, self.destroy_all)

    def destroy_all(self):
        self.destroy();
        if self.master: self.master.destroy()


# =============================================================================
# 1. HardwareController Class (已重構成更健壯的版本)
# =============================================================================
class HardwareController:
    """封裝所有與硬體 COM 元件的互動，並能智慧地處理連接錯誤。"""

    def __init__(self, prog_id):
        self.prog_id = prog_id
        self.usbio = None  # COM 物件的參考

    def _connect_internal(self):
        """[私有] 嘗試建立 COM 物件並打開 USB 設備。"""
        try:
            # 步驟 1: 建立 COM 物件。如果 DLL 未註冊，這一步會失敗。
            print("正在建立 COM 物件...")
            self.usbio = win32com.client.Dispatch(self.prog_id)

            # 步驟 2: 嘗試打開 USB 設備。如果硬體未插入，這一步會返回 False 或引發錯誤。
            if self.usbio.OpenUsbDevice(4660, 26505):
                print("硬體連接成功 (OpenUsbDevice)。")
                return True
            else:
                # OpenUsbDevice 正常返回 False，表示未找到設備。
                print("未找到硬體 (OpenUsbDevice 返回 False)。")
                self.usbio = None  # 清理掉無用的 COM 物件
                return False
        except pywintypes.com_error:
            # 這是硬體錯誤或 COM 註冊問題。
            print("連接時發生 COM 錯誤。")
            self.usbio = None  # 清理掉無用的 COM 物件
            return False

    def disconnect(self):
        """關閉設備連接並徹底釋放 COM 物件。"""
        if self.usbio:
            try:
                # 關閉前嘗試熄滅所有 LED
                red_mask = green_mask = 0
                self.usbio.OutDataCtrl(red_mask, 32);
                self.usbio.OutDataCtrl(red_mask, 48)
                self.usbio.OutDataCtrl(green_mask, 0)
                self.usbio.CloseUsbDevice()
                print("已安全關閉 USB 設備。")
            except (pywintypes.com_error, ConnectionError):
                print("關閉設備時發生錯誤，將忽略。")
            finally:
                # 這是最關鍵的一步：確保 COM 物件被釋放。
                self.usbio = None
                print("COM 物件已釋放。")

    def is_connected(self):
        """
        檢查硬體是否連接的唯一可靠方法。
        如果未連接，它會自動嘗試重新連接。
        """
        # 如果我們連 COM 物件都沒有，那肯定沒連接。嘗試連接。
        if not self.usbio:
            return self._connect_internal()

        # 如果有 COM 物件，我們需要 "Ping" 一下硬體確認它還在。
        try:
            # 再次呼叫 OpenUsbDevice 作為 Ping。如果硬體被拔掉，這裡會引發錯誤。
            if self.usbio.OpenUsbDevice(4660, 26505):
                return True  # Ping 成功，硬體還在。
            else:
                # Ping 失敗，硬體可能以一種奇怪的方式離線了。
                print("Ping 失敗 (OpenUsbDevice 返回 False)，斷開連接。")
                self.disconnect()
                return False
        except pywintypes.com_error:
            # Ping 引發了 COM 錯誤，這是最常見的斷線情況。
            print("Ping 失敗 (COM 錯誤)，斷開連接。")
            self.disconnect()  # 斷開連接，清理掉已失效的 COM 物件。
            return False

    def update_leds(self, led_states):
        """根據給定的狀態更新硬體 LED。如果失敗，會自動斷開連接。"""
        if not self.usbio:
            raise ConnectionError("硬體未連接，無法更新 LED。")
        try:
            red_mask = sum(1 << i for i, state in enumerate(led_states[:8]) if state)
            self.usbio.OutDataCtrl(red_mask, 32)
            self.usbio.OutDataCtrl(red_mask, 48)
            green_mask = sum(1 << i for i, state in enumerate(led_states[8:]) if state)
            self.usbio.OutDataCtrl(green_mask, 0)
        except pywintypes.com_error:
            print("更新 LED 時發生 COM 錯誤，斷開連接。")
            # 通訊失敗，意味著連接已斷開。
            self.disconnect()
            # 拋出異常，讓上層知道操作未完成。
            raise ConnectionError("硬體通訊失敗。")


# =============================================================================
# 2. LedControlApp Class (已簡化輪詢邏輯)
# =============================================================================
class LedControlApp:
    def __init__(self, root):
        self.root = root
        self.hardware = HardwareController(config.PROG_ID)
        self.led_states = [False] * config.NUM_LEDS
        self.is_first_connection = True
        self.active_animation = None
        self.marquee_pos, self.marquee_dir = 0, 1
        self.led_states_before_animation = []
        self.last_animation_visual_state = []
        self.after_id_poll, self.after_id_animation = None, None
        self.setup_ui()
        self.poll_hardware_status()

    def setup_ui(self):
        self.root.title("USB I/O LED 控制器 v1.5")
        self.root.geometry(f"{config.WINDOW_DEFAULT_WIDTH}x{config.WINDOW_DEFAULT_HEIGHT}")
        self.root.resizable(True, True)
        self.root.grid_rowconfigure(1, weight=1);
        self.root.grid_columnconfigure(0, weight=1)
        status_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        status_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(5, 0))
        self.status_label = ctk.CTkLabel(status_frame, text="正在搜尋硬體...",
                                         font=ctk.CTkFont(family="Microsoft JhengHei", size=14, weight="bold"))
        self.status_label.pack(side=tk.LEFT)
        self.led_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        self.led_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5);
        self.create_led_buttons()
        self.func_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        self.func_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(5, 0));
        self.create_function_buttons()
        self.anim_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        self.anim_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=(5, 10));
        self.create_animation_buttons()
        self.disable_buttons(self.buttons + self.func_buttons + self.anim_buttons)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_led_buttons(self):
        self.buttons = []
        for i in range(config.NUM_LEDS):
            button = ctk.CTkButton(self.led_frame, text=f"{i + 1}", corner_radius=8, font=ctk.CTkFont(weight="bold"),
                                   command=lambda index=i: self.toggle_led(index))
            row, col = i // 8, 7 - (i % 8)
            button.grid(row=row, column=col, padx=5, pady=5, sticky="nsew");
            self.buttons.append(button)
        for i in range(8): self.led_frame.grid_columnconfigure(i, weight=1)
        for i in range(2): self.led_frame.grid_rowconfigure(i, weight=1)
        self.update_gui_from_states()

    def create_function_buttons(self):
        self.func_buttons = []
        btn_config = {"fg_color": config.COLOR_BUTTON_FUNC, "hover_color": config.COLOR_BUTTON_FUNC_HOVER,
                      "text_color": config.COLOR_BUTTON_FUNC_TEXT, "corner_radius": 8,
                      "font": ctk.CTkFont(family="Microsoft JhengHei", weight="bold")}
        btn_all_on = ctk.CTkButton(self.func_frame, text="全部開啟", **btn_config, command=self.all_on)
        btn_all_off = ctk.CTkButton(self.func_frame, text="全部關閉", **btn_config, command=self.all_off)
        btn_invert = ctk.CTkButton(self.func_frame, text="反向", **btn_config, command=self.invert_state)
        self.func_buttons = [btn_all_on, btn_all_off, btn_invert]
        for i, btn in enumerate(self.func_buttons): self.func_frame.grid_columnconfigure(i, weight=1); btn.grid(row=0,
                                                                                                                column=i,
                                                                                                                padx=5,
                                                                                                                sticky="ew")

    def create_animation_buttons(self):
        self.anim_buttons = []
        btn_config = {"fg_color": config.COLOR_BUTTON_FUNC, "hover_color": config.COLOR_BUTTON_FUNC_HOVER,
                      "text_color": config.COLOR_BUTTON_FUNC_TEXT, "corner_radius": 8,
                      "font": ctk.CTkFont(family="Microsoft JhengHei", weight="bold")}
        self.btn_marquee = ctk.CTkButton(self.anim_frame, text="跑馬燈", **btn_config,
                                         command=lambda: self.toggle_animation('marquee'))
        self.btn_random = ctk.CTkButton(self.anim_frame, text="亂數閃爍", **btn_config,
                                        command=lambda: self.toggle_animation('random'))
        self.anim_buttons = [self.btn_marquee, self.btn_random]
        for i, btn in enumerate(self.anim_buttons): self.anim_frame.grid_columnconfigure(i, weight=1); btn.grid(row=0,
                                                                                                                column=i,
                                                                                                                padx=5,
                                                                                                                sticky="ew")

    def poll_hardware_status(self):
        """核心輪詢邏輯 (已簡化)。"""
        # 記錄檢查前的連接狀態
        was_connected = self.hardware.usbio is not None
        # 執行檢查 (如果未連接，is_connected 內部會嘗試重連)
        is_now_connected = self.hardware.is_connected()

        # 情況 1: 之前連著，但剛剛的檢查發現斷了
        if was_connected and not is_now_connected:
            self.handle_disconnection()
        # 情況 2: 之前沒連著，但剛剛的檢查成功連上了
        elif not was_connected and is_now_connected:
            self.handle_connection_success()

        # 安排下一次輪詢
        self.after_id_poll = self.root.after(config.POLL_INTERVAL_MS, self.poll_hardware_status)

    def handle_connection_success(self):
        """處理硬體成功連接的邏輯。"""
        print("UI 層：偵測到硬體已連接。")
        self.status_label.configure(text="硬體已連接", text_color="lightgreen")
        all_buttons = self.buttons + self.func_buttons + self.anim_buttons
        self.enable_buttons(all_buttons)
        if self.is_first_connection:
            self.reset_all_leds()
            self.is_first_connection = False
        else:
            if self.active_animation:
                print(f"重新連接，恢復 {self.active_animation} 動畫...")
                current_button = self.btn_marquee if self.active_animation == 'marquee' else self.btn_random
                self.disable_buttons([b for b in all_buttons if b != current_button])
                self.run_animation_frame()
            else:
                self.restore_led_state()

    def handle_disconnection(self):
        """處理硬體斷開連接的邏輯 (現在只負責更新UI)。"""
        print("UI 層：偵測到硬體已斷開。")
        if self.after_id_animation: self.root.after_cancel(self.after_id_animation); self.after_id_animation = None
        # 硬體層的清理工作已由 HardwareController 自動完成，這裡只需更新UI
        self.disable_buttons(self.buttons + self.func_buttons + self.anim_buttons)
        self.status_label.configure(text="硬體已中斷，正在嘗試重連...", text_color="orange")

    def disable_buttons(self, button_list):
        for btn in button_list: btn.configure(state=tk.DISABLED)

    def enable_buttons(self, button_list):
        for btn in button_list: btn.configure(state=tk.NORMAL)

    def update_gui_from_states(self):
        for i, state in enumerate(self.led_states):
            color = config.COLOR_LED_ON if state else config.COLOR_LED_OFF;
            text_color = "black" if state else "white"
            hover_color = config.COLOR_LED_HOVER_ON if state else config.COLOR_LED_HOVER_OFF
            self.buttons[i].configure(fg_color=color, text_color=text_color, hover_color=hover_color)

    def _update_hardware(self):
        try:
            self.hardware.update_leds(self.led_states)
        except (pywintypes.com_error, ConnectionError):
            # update_leds 內部會自動呼叫 disconnect，輪詢器會在下次循環時偵測到狀態變化
            print("UI 層：更新硬體時捕獲到錯誤。")

    def reset_all_leds(self):
        self.led_states = [False] * config.NUM_LEDS;
        self.update_gui_from_states();
        self._update_hardware()

    def restore_led_state(self):
        print(f"從內部狀態恢復硬體顯示: {self.led_states}");
        self.update_gui_from_states();
        self._update_hardware()

    def _execute_action(self, action):
        if self.active_animation: self.stop_animation()
        action();
        self.update_gui_from_states();
        self._update_hardware()

    def _toggle_led_state(self, index):
        self.led_states[index] = not self.led_states[index]

    def toggle_led(self, index):
        self._execute_action(lambda: self._toggle_led_state(index))

    def all_on(self):
        self._execute_action(lambda: self.led_states.__init__([True] * config.NUM_LEDS))

    def all_off(self):
        self._execute_action(lambda: self.led_states.__init__([False] * config.NUM_LEDS))

    def invert_state(self):
        self._execute_action(lambda: self.led_states.__init__([not s for s in self.led_states]))

    def toggle_animation(self, anim_type):
        if self.active_animation == anim_type:
            self.stop_animation()
        else:
            if self.active_animation: self.stop_animation()
            self.start_animation(anim_type)

    def start_animation(self, anim_type):
        print(f"啟動 {anim_type} 動畫");
        self.led_states_before_animation = self.led_states.copy()
        self.active_animation = anim_type
        current_button = self.btn_marquee if anim_type == 'marquee' else self.btn_random
        btn_text = "跑馬燈" if anim_type == 'marquee' else "亂數閃爍"
        current_button.configure(text=f"停止{btn_text}", fg_color=config.COLOR_ANIMATION_STOP,
                                 hover_color=config.COLOR_ANIMATION_STOP_HOVER)
        if anim_type == 'marquee': self.marquee_pos, self.marquee_dir = 0, 1
        self.disable_buttons([b for b in self.buttons + self.func_buttons + self.anim_buttons if b != current_button])
        self.run_animation_frame()

    def stop_animation(self):
        if not self.active_animation: return
        print(f"停止 {self.active_animation} 動畫");
        if self.after_id_animation: self.root.after_cancel(self.after_id_animation); self.after_id_animation = None
        self.btn_marquee.configure(text="跑馬燈", fg_color=config.COLOR_BUTTON_FUNC,
                                   hover_color=config.COLOR_BUTTON_FUNC_HOVER)
        self.btn_random.configure(text="亂數閃爍", fg_color=config.COLOR_BUTTON_FUNC,
                                  hover_color=config.COLOR_BUTTON_FUNC_HOVER)
        if config.RESTORE_STATE_ON_STOP:
            if self.led_states_before_animation: self.led_states = self.led_states_before_animation.copy()
        else:
            if self.last_animation_visual_state: self.led_states = self.last_animation_visual_state.copy()
        self.active_animation = None
        self.enable_buttons(self.buttons + self.func_buttons + self.anim_buttons);
        self.restore_led_state()

    def run_animation_frame(self):
        if not self.active_animation: return
        try:
            # 在動畫循環開始時就檢查連接，如果失敗就直接停止
            if not self.hardware.is_connected():
                self.handle_disconnection()  # is_connected 內部會清理，這裡只需更新UI
                return
            temp_visual_states = [False] * config.NUM_LEDS;
            delay = 100
            if self.active_animation == 'marquee':
                temp_visual_states[self.marquee_pos] = True;
                delay = config.MARQUEE_SPEED_MS
                if self.marquee_pos >= config.NUM_LEDS - 1 and self.marquee_dir == 1:
                    self.marquee_dir = -1
                elif self.marquee_pos <= 0 and self.marquee_dir == -1:
                    self.marquee_dir = 1
                self.marquee_pos += self.marquee_dir
            elif self.active_animation == 'random':
                random_index = random.randint(0, config.NUM_LEDS - 1);
                temp_visual_states[random_index] = True
                delay = int(max(0.1, config.RANDOM_FLICKER_DELAY_S) * 1000)
            self.last_animation_visual_state = temp_visual_states.copy()
            self.hardware.update_leds(temp_visual_states)  # 直接用硬體物件更新
            for i, state in enumerate(temp_visual_states):
                color = config.COLOR_LED_ON if state else config.COLOR_LED_OFF;
                text_color = "black" if state else "white"
                self.buttons[i].configure(fg_color=color, text_color=text_color)
            self.after_id_animation = self.root.after(delay, self.run_animation_frame)
        except (pywintypes.com_error, ConnectionError):
            # 這裡的錯誤處理現在是一個備用方案，理論上 is_connected() 會先捕獲到問題
            print("動畫幀執行時捕獲到錯誤。")

    def on_closing(self):
        dialog = CustomMessageBox(title="退出程式", message="您確定要關閉程式嗎？")
        if dialog.get():
            if self.after_id_poll: self.root.after_cancel(self.after_id_poll)
            if self.after_id_animation: self.root.after_cancel(self.after_id_animation)
            self.hardware.disconnect();
            self.root.destroy()


# =============================================================================
# 3. 主程式啟動入口 (保持不變)
# =============================================================================
if __name__ == "__main__":
    if sys.platform != "win32":
        root = tk.Tk();
        root.withdraw();
        tk.messagebox.showerror("系統不支援", "此程式只能在 Windows 作業系統上運行。");
        sys.exit()
    mutex_name = "USBIO_LED_Controller_Mutex_{c28f3435-9B2A-4f3d-8F3E-6B8E7A6B295A}"
    mutex = None
    try:
        mutex = win32event.CreateMutex(None, False, mutex_name)
        last_error = win32api.GetLastError()
        if last_error == winerror.ERROR_ALREADY_EXISTS:
            temp_root = ctk.CTk();
            temp_root.withdraw();
            ToastNotification("程式已在執行中，無法重複開啟。", master=temp_root).mainloop();
            sys.exit()
        else:
            print("啟動主程式...")
            ctk.set_appearance_mode(config.UI_THEME);
            ctk.set_default_color_theme("blue")
            root = ctk.CTk();
            app = LedControlApp(root);
            root.mainloop()
    finally:
        if mutex: win32api.CloseHandle(mutex); print("Mutex 已釋放，程式結束。")