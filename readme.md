# USB I/O LED 控制器

## 專案簡介

本專案是一個使用 Python 開發的專業級圖形化介面（GUI）應用程式，旨在透過特定的 USB I/O 硬體介面卡（一個基於 **Visual C++ 2010** 的原生 COM 元件）來控制 16 個 LED。程式不僅提供了基礎的單點控制，還具備高度的穩定性和豐富的功能，包括斷線自動重連、狀態恢復、多種快捷操作及現代化的 UI 介面。

### 必用 python 32 bit 因為 USBIO4.dll 是 32 bit

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
    *   **回復效果**：當按下「停止跑馬燈」按鈕時，LED 會**回到按跑馬燈之前的狀態**。
*   **亂數閃爍效果**：
    *   亂數閃爍 LED 1 到 16，再從 15 回到 1 的「乒乓來回」跑馬燈效果。
    *   **回復效果**：當按下「停止亂數閃爍」按鈕時，LED 會**回到按亂數閃爍之前的狀態**。
*   **智慧型斷線重連**（不是很穩定，要多試幾次）：
    *   程式在背景**自動輪詢**硬體連接狀態。
    *   當 USB 裝置被拔除時，程式會自動禁用按鈕並進入等待重連模式。
    *   當裝置重新接上後，程式會自動恢復連接。
*   **斷線狀態恢復**：
    *   在**一般模式**下斷線，重連後會**正確恢復**到斷線前的狀態。
    *   在**跑馬燈模式**下斷線，重連後會**繼續執行跑馬燈**。
*   **專業級程式碼架構**：
    *   **設定分離**：所有顏色、硬體 ID 等設定均由 `config.py` 統一管理，方便客製化。
    *   **邏輯分離**：硬體通訊邏輯被封裝在獨立的 `HardwareController` 類別中，與 UI 邏輯完全解耦，易於維護和擴充。
*   **單例應用程式**：防止使用者重複開啟多個程式視窗，確保系統中只有一個實例在運行。

## 環境設置與部署流程

這是在一台**全新的 Windows 電腦**上部署此應用程式的**最終、最精簡**的步驟。

### 1. 安裝程式執行依賴

本專案的 Python 程式碼需要安裝兩個關鍵的函式庫。請打開命令提示字元 (CMD) 或 PowerShell，並執行以下指令：

```bash
# 安裝與 Windows COM 元件溝通的函式庫
pip install pywin32

# 安裝現代化 UI 介面函式庫
pip install customtkinter
```

### 2. 安裝硬體驅動依賴 (最關鍵)

原檢定開發環境是 VB 2010 ，如果不安裝 VB 2010，就只要安裝 Microsoft Visual C++ 2010 SP1 可轉散發套件 (x86)

`USBIO4.dll` 這個硬體驅動需要特定的系統執行階段函式庫才能正常運作。

1.  **下載 Microsoft Visual C++ 2010 SP1 可轉散發套件 (x86)**。
    *   **下載連結**: [https://www.microsoft.com/zh-tw/download/details.aspx?id=26999](https://www.microsoft.com/zh-tw/download/details.aspx?id=26999)
    *   本專案 dll 目錄內也有提供 `vcredist_x86.exe` , `USBIO4.dll`
    *   **注意**: 即使您的作業系統是 64 位元，也**必須**安裝 **x86 (32位元)** 版本，因為 `USBIO4.dll` 是 32 位元的。
2.  執行下載的 `vcredist_x86.exe` 並完成安裝。
3.  **建議**：安裝完成後，最好**重新啟動電腦**以確保系統環境完全更新。

### 3. 註冊硬體 COM 元件

此步驟是為了讓 Windows 系統能夠識別 `USBIO4.dll`。

1.  將 `USBIO4.dll` 檔案放置在一個固定的、易於存取的位置，例如 `C:\`。
2.  打開「命令提示字元(CMD)」，務必**點右鍵選擇「以系統管理員身分執行」**。
3.  在系統管理員的命令提示字元中，輸入以下指令後按 Enter：
    ```cmd
    regsvr32 C:\USBIO4.DLL
    ```
4.  如果成功，會跳出一個「DllRegisterServer in C:\USBIO4.DLL succeeded.」的訊息。此步驟只需在每台電腦上成功執行一次即可。

### 4. 程式檔案結構與執行

1.  請確保以下兩個檔案放置在**同一個資料夾**中：
    ```
    /你的專案資料夾/
    ├── config.py       # 專案設定檔
    └── usbio4.py       # 主程式檔案
    ```
2.  一切準備就緒後，直接執行 `usbio4.py` 即可啟動程式：
    ```bash
    python usbio4.py
    ```

## 程式碼架構說明

### `config.py` - 設定檔

這個檔案是整個專案的「控制中心」，所有可調整的參數都集中在此，方便快速客製化，無需更動主程式邏輯。

*   **想改變顏色或主題？** -> 直接修改 `config.py` 中的 `COLOR_` 和 `UI_THEME` 變數。
*   **想調整跑馬燈速度？** -> 修改 `config.py` 中的 `MARQUEE_SPEED_MS`。
*   **想換回「停止跑馬燈後恢復原狀」的行為？** -> 修改 `usbio4.py` 中的 `start_marquee` 和 `stop_marquee` 函式，恢復使用 `led_states_before_marquee` 變數來備份和恢復狀態的邏輯。
*   **未來想換另一塊不同協議的板子？** -> 只需要重寫 `HardwareController` 類別中的方法，`LedControlApp` 的 UI 和應用邏輯幾乎不用變動。

### `usbio4.py` - 主程式

#### `HardwareController` 類別

這是**硬體驅動層**。它將所有與 `pywin32` 和 COM 物件的複雜互動封裝起來，對主程式提供簡單清晰的介面 (`connect`, `disconnect`, `is_connected`, `update_leds`)。它的職責單純，只負責與硬體溝通。

#### `LedControlApp` 類別

這是**應用程式的核心**，負責處理 UI 邏輯、狀態管理和使用者互動。

*   **UI 建立 (`setup_ui`)**: 使用 `customtkinter` 建立所有現代化的 UI 元件，並採用 `.grid()` 佈局管理器實現可縮放的視窗。
*   **背景輪詢 (`poll_hardware_status`)**: 整個程式最核心的循環，以固定的間隔持續檢查硬體狀態，是實現智慧型斷線重連的基礎。
*   **狀態管理 (`is_first_connection`, `led_states`)**:
    *   `is_first_connection` 旗標被用來**嚴格區分**「程式首次啟動」（此時應重設 LED）和「運行中斷線重連」（此時應恢復狀態），這是確保狀態恢復正確的關鍵。
    *   `self.led_states` 列表是整個應用程式在**一般模式下**唯一可信的狀態來源 (Single Source of Truth)。
*   **跑馬燈邏輯 (`start_marquee`, `stop_marquee`, `run_marquee_frame`)**:
    *   `run_marquee_frame`: 動畫的每一幀。它只產生**臨時的視覺狀態**去更新硬體和 GUI，但**不會**修改 `self.led_states`。
    *   `stop_marquee`: 停止動畫循環後，計算出動畫的最後一幀位置，並用這個位置去**更新** `self.led_states`，從而實現「凍結」效果。
*   **單例應用實現 (在 `if __name__ == "__main__":` 中)**:
    *   使用 Windows 的 `CreateMutex` API 來建立一個全域互斥鎖。
    *   程式啟動時會檢查這個鎖是否存在，如果已存在，則彈出短暫通知並退出，確保了系統中永遠只有一個程式實例在運行。


## 參考資料
[電腦硬體裝修職類乙級術科測試應檢參考資料 PDF檔](./data/120002B12_技術士技能檢定電腦硬體裝修職類乙級術科測試應檢參考資料.pdf)

原理圖
![原理圖](./images/p31.png)

零件配置參考圖
![零件配置參考圖](./images/p32.png)

設備表
![設備表1](./images/p33.png)

材料表
![材料表1](./images/p34.png)
![設備表2](./images/p35.png)

