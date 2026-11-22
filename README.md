# 🚓 交通執法工具磁鐵（跨分局通用版）

專為警察機關交通業務承辦人設計的輕量級自動化工具箱，已移除所有特定分局名稱、真實姓名、雲端試算表連線與自動寄信設定，可安全部署為公開工具磁鐵。

## 📁 專案結構

```text
traffic_tool_magnet/
├── app.py                          # 首頁（零外部連線阻塞，2 秒內啟動）
├── menu.py                         # 共用側邊欄、資安聲明與個人網站導流
├── requirements.txt                # 精簡版 Python 相依套件（避免免費版 OOM）
├── packages.txt                    # Linux 系統套件（poppler-utils，供 PDF 轉圖片使用）
└── pages/
    ├── 1_專案敘獎一鍵精算器.py      # 毒駕、危駕改裝車、無照移置、偽變造車牌四合一精算
    ├── 2_執法簡報直出中心.py        # 支援自訂分局與所隊人數的 PPTX 週報產生器
    └── 3_檔案加工與轉檔中心.py      # 自訂署名角標遮蓋加頁碼、PDF 轉 PPTX/圖、多圖併 PDF
```

## 🚀 部署到 Streamlit Cloud 步驟

1. 在 GitHub 建立一個**新的公開（Public）儲存庫**（請勿與原本含私人資料或 `secrets.toml` 的儲存庫混用）。
2. 將本資料夾內的所有檔案上傳至該新儲存庫根目錄。
3. （選用）若希望「檔案加工與轉檔中心」的 PDF 署名使用特定楷體／黑體，可將開源中文字型（如 `NotoSansTC-Regular.ttf` 或 `kaiu.ttf`）放進專案根目錄或 `fonts/` 資料夾。
4. 前往 [Streamlit Community Cloud](https://share.streamlit.io/) 選擇該公開儲存庫、主程式指定 `app.py`，無須設定任何 `Secrets` 即可直接部署上線。
