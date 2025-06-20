# USB I/O LED 控制器

## 專案簡介

本專案程式碼由 2025/06/20 Google AI Studio 生成，使用 Windows 11 和 Python 32 bit 開發的專業級圖形化介面（GUI）應用程式，旨在透過特定的 USB I/O 硬體介面卡（使用 `USBIO4.DLL` 作為 COM 元件）來控制 16 個 LED。程式不僅提供了基礎的單點控制，還具備高度的穩定性和豐富的功能，包括斷線自動重連、狀態恢復、多種快捷操作及現代化的 UI 介面。當然還有優化的空間。

必用 python 32 bit 因為 USBIO4.dll 是 32 bit

USB I/O 硬體介面卡為︰技術士技能檢定電腦硬體裝修職類乙級術科測試應檢參考資料 12000-102201~210 版本
![A](./images/USBIO_board_A_finish_s.png)
![B](./images/USBIO_board_B_finish_s.png)
## 主要功能

*   **現代化 UI 介面**：使用 `customtkinter` 函式庫，提供支援 Windows 11 風格的深色/淺色主題、圓角按鈕和動態懸停效果。
*   **16 個獨立 LED 控制**：每個 LED 對應一個按鈕，可獨立切換開/關狀態。按鈕顏色與文字顏色會隨狀態改變，以提高可讀性。
*   **快捷功能鍵**：
    *   `全部開啟`：一鍵點亮所有 LED。
    *   `全部關閉`：一鍵熄滅所有 LED。
    *   `反向`：一鍵反轉所有 LED 的開關狀態。
*   **動態跑馬燈效果**：
    *   提供從 LED 1 到 16，再從 15 回到 1 的「乒乓來回」跑馬燈效果。
    *   **凍結效果**：當按下「停止跑馬燈」按鈕時，LED 會**停留在當前的最後一幀動畫位置**，而不是恢復到動畫開始前的狀態。
*   **智慧型斷線重連**：
    *   程式在背景**自動輪詢**硬體連接狀態。
    *   當 USB 裝置被拔除時，程式會自動禁用按鈕並進入等待重連模式。
    *   當裝置重新接上後，程式會自動恢復連接。
*   **斷線狀態恢復**：
    *   在**一般模式**下斷線，重連後會**正確恢復**到斷線前的狀態。
    *   在**跑馬燈模式**下斷線，重連後會**繼續執行跑馬燈**。
*   **專業級程式碼架構**：
    *   **設定分離**：所有顏色、硬體 ID 等設定均由 `config.py` 統一管理，方便客製化。
    *   **邏輯分離**：硬體通訊邏輯被封裝在獨立的 `HardwareController` 類別中，與 UI 邏輯完全解耦，易於維護和擴充。

## 環境設置與執行

### 1. 安裝依賴套件

本專案需要安裝兩個關鍵的 Python 函式庫。請打開命令提示字元 (CMD) 或 PowerShell，並執行以下指令：

```bash
# 安裝與 Windows COM 元件溝通的函式庫
pip install pywin32

# 安裝現代化 UI 介面函式庫
pip install customtkinter
```

### 2. 檔案結構

請確保以下兩個檔案放置在**同一個資料夾**中：

```
/你的專案資料夾/
├── config.py       # 專案設定檔
└── usbio4.py       # 主程式檔案
```

### 3. dll 硬體註冊 (僅需執行一次)

這個硬體介面卡是以 COM 元件的形式運作，因此在使用前必須先在 Windows 系統中註冊 `USBIO4.DLL`。

1.  將 `USBIO4.DLL` 檔案放置在一個固定的位置，例如 `C:\`。
2.  打開「命令提示字元(CMD)」，務必**點右鍵選擇「以系統管理員身分執行」**。
3.  在系統管理員的命令提示字元中，輸入以下指令後按 Enter：
    ```cmd
    regsvr32 C:\USBIO4.DLL
    ```
4.  如果成功，會跳出一個註冊成功的訊息。此步驟只需在每台電腦上執行一次即可。

### 4. 執行程式

一切準備就緒後，直接執行 `usbio4.py` 即可啟動程式：

```bash
python usbio4.py
```

## 程式碼架構說明

### `config.py` - 設定檔

這個檔案是整個專案的「控制中心」，所有可調整的參數都集中在此。

*   **UI 主題與顏色**：定義了程式的深色/淺色模式，以及所有按鈕在不同狀態（開啟、關閉、懸停）下的顏色。
*   **硬體設定**：包含了 COM 元件的程式 ID (`ProgID`) 和 LED 的總數。
*   **程式行為設定**：控制背景輪詢的頻率和視窗的預設大小。
*   **跑馬燈設定**：控制跑馬燈動畫的速度。

### `usbio4.py` - 主程式

#### `HardwareController` 類別

這是**硬體驅動層**。它將所有與 `pywin32` 和 COM 物件的複雜互動封裝起來，對主程式提供簡單清晰的介面。它的職責單純，只負責「說」和「聽」。

#### `LedControlApp` 類別

這是**應用程式的核心**，負責處理 UI 邏輯、狀態管理和使用者互動。

*   **`__init__` 和 `setup_ui`**: 初始化視窗，建立 `HardwareController` 實例，並使用 `customtkinter` 建立所有現代化的 UI 元件。
*   **`poll_hardware_status`**: 整個程式最核心的**背景輪詢**函式。它以固定的間隔（在 `config.py` 中設定）持續檢查硬體狀態，並根據情況觸發 `handle_connection_success` 或 `handle_disconnection`。
*   **`handle_connection_success` 和 `handle_disconnection`**: 這兩個函式是狀態恢復的關鍵。
    *   `handle_connection_success`: 透過 `is_first_connection` 旗標區分「首次啟動」（重設狀態）和「斷線重連」（恢復狀態）。在重連時，它還會檢查 `marquee_running` 旗標來決定是恢復一般狀態還是繼續跑馬燈。
    *   `handle_disconnection`: 只負責處理斷線後的清理工作，不再修改核心的 `led_states`。
*   **狀態管理 (`led_states`)**: `self.led_states` 列表是整個應用程式在**一般模式下**唯一可信的狀態來源 (Single Source of Truth)。
*   **跑馬燈邏輯 (`start_marquee`, `stop_marquee`, `run_marquee_frame`)**:
    *   `start_marquee`: 啟動動畫循環。
    *   `run_marquee_frame`: 動畫的每一幀。它產生**臨時的視覺狀態**去更新硬體和 GUI，但**不會**修改 `self.led_states`。
    *   `stop_marquee`: 這是**最新的修改重點**。它會停止動畫循環，然後計算出動畫的最後一幀位置，並用這個位置去**更新** `self.led_states`，從而實現「凍結」效果。

## 如何客製化

*   **想改變顏色或主題？** -> 直接修改 `config.py` 中的 `COLOR_` 和 `UI_THEME` 變數。
*   **想調整跑馬燈速度？** -> 修改 `config.py` 中的 `MARQUEE_SPEED_MS`。
*   **想換回「停止跑馬燈後恢復原狀」的行為？** -> 修改 `usbio4.py` 中的 `start_marquee` 和 `stop_marquee` 函式，恢復使用 `led_states_before_marquee` 變數來備份和恢復狀態的邏輯。
*   **未來想換另一塊不同協議的板子？** -> 只需要重寫 `HardwareController` 類別中的方法，`LedControlApp` 的 UI 和應用邏輯幾乎不用變動。

## 參考資料
原理圖，零件配置參考圖，材料表，請參考 p.31 ~ p.35︰

[電腦硬體裝修職類乙級術科測試應檢參考資料 PDF檔](./data/120002B12_技術士技能檢定電腦硬體裝修職類乙級術科測試應檢參考資料.pdf)