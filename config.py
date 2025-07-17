# =================================================================
# config.py - 專案設定檔 (v1.3)
# =================================================================

# --- 主題設定 (dark, light, system) ---
UI_THEME = "dark"

# --- 顏色設定 for customtkinter ---
COLOR_LED_ON = "#FFD700"
COLOR_LED_OFF = "dimgray"
COLOR_LED_HOVER_ON = "#FFA500"
COLOR_LED_HOVER_OFF = "#606060"

COLOR_BUTTON_FUNC = "#2C5F2D"
COLOR_BUTTON_FUNC_HOVER = "#4CAF50"
COLOR_BUTTON_FUNC_TEXT = "white"

COLOR_ANIMATION_STOP = "#990000"
COLOR_ANIMATION_STOP_HOVER = "#D2042D"

# --- 硬體設定 ---
PROG_ID = "Innovati.2"
NUM_LEDS = 16

# --- 程式行為設定 ---
# 注意：這個版本的重連穩定性依賴於這個輪詢間隔，
# 如果遇到問題，可以嘗試稍微加長這個時間。
POLL_INTERVAL_MS = 2000 # 背景輪詢間隔（毫秒）

WINDOW_DEFAULT_WIDTH = 350
WINDOW_DEFAULT_HEIGHT = 180

# --- 動畫效果設定 ---
MARQUEE_SPEED_MS = 100
RANDOM_FLICKER_DELAY_S = 0.1