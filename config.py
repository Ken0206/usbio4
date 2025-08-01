# =================================================================
# config.py - 專案設定檔 (v1.5)
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
# 背景輪詢間隔（毫秒），用於偵測硬體斷線與重連。
POLL_INTERVAL_MS = 2000

# 視窗預設尺寸
WINDOW_DEFAULT_WIDTH = 350
WINDOW_DEFAULT_HEIGHT = 180

# --- 動畫效果設定 ---
MARQUEE_SPEED_MS = 100
RANDOM_FLICKER_DELAY_S = 0.1

# 停止動畫時的行為設定:
# True  = 回到動畫開始前的狀態 (Restore to state before animation)
# False = 凍結在動畫的最後一幀 (Freeze at the last frame)
RESTORE_STATE_ON_STOP = True