# =================================================================
# config.py - 專案設定檔 (v4.7 UI 優化)
# =================================================================

# --- 主題設定 (dark, light, system) ---
UI_THEME = "dark"

# --- 顏色設定 for customtkinter ---
# LED 按鈕顏色
COLOR_LED_ON = "#FFD700"      # 黃色 (Gold)
COLOR_LED_OFF = "#4A4A4A"     # 深灰色
COLOR_LED_HOVER_ON = "#FFA500" # ON 狀態懸停時的橘黃色
# --- 新增這一行 ---
COLOR_LED_HOVER_OFF = "#606060" # OFF 狀態懸停時的亮灰色

# 功能按鈕顏色
COLOR_BUTTON_FUNC = "#2C5F2D"
COLOR_BUTTON_FUNC_HOVER = "#4CAF50"
COLOR_BUTTON_FUNC_TEXT = "white"

# 跑馬燈停止按鈕顏色
COLOR_MARQUEE_STOP = "#990000"
COLOR_MARQUEE_STOP_HOVER = "#D2042D"

# --- 硬體設定 ---
PROG_ID = "Innovati.2"
NUM_LEDS = 16

# --- 程式行為設定 ---
POLL_INTERVAL_MS = 1500
WINDOW_DEFAULT_WIDTH = 620
WINDOW_DEFAULT_HEIGHT = 300

# --- 跑馬燈設定 ---
MARQUEE_SPEED_MS = 100