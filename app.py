import streamlit as st
from menu import show_sidebar

st.set_page_config(
    page_title="交通執法工具磁鐵（跨分局通用版）",
    page_icon="🚓",
    layout="wide",
)

show_sidebar()

st.title("🚓 警察交通執法與行政自動化工具箱（跨分局通用版）")
st.markdown(
    "專為各分局**交通組、交通分隊及派出所業務承辦人**設計。"
    "**免註冊、無浮水印、不綁定特定分局**，所有檔案僅在本次連線記憶體即時處理後下載，不留存任何公務資料。"
)

st.divider()

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("🏆 1. 專案敘獎一鍵精算器")
    st.markdown(
        "依據全國共通法規與常見專案標準，上傳系統『自選匯出.xlsx』即可 3 秒結算："
        "\n- 🧪 **毒駕專案**：區分大型車／小型車／機車含微電車／毒駕拒測，自動產出具體出力事由"
        "\n- 🚨 **危駕與改裝車**：支援 2 件／4 件群組與滿 9 次門檻倍數遞增"
        "\n- 🚗 **無照移置保管**：支援未成年／成年小型車與跨季保留計算"
        "\n- 🔎 **偽變造車牌**：拆分偽變造與他車號牌，支援滿 9 次獎勵加倍"
    )
    st.page_link("pages/1_專案敘獎一鍵精算器.py", label="👉 立即開啟精算器", icon="🏆", use_container_width=True)

with col2:
    st.subheader("📽️ 2. 執法簡報直出中心")
    st.markdown(
        "無需手調表格與字體，支援**任意分局自訂單位名稱與編制人數**，一鍵產出 PowerPoint 簡報："
        "\n- 機關名稱、所隊清單與人數完全自由編輯"
        "\n- 自動計算**雙欄人均件數與排名**，破除大所小所人數迷思"
        "\n- 支援三項重點違規、A1/A2 事故、重大違規、超載與科技執法"
        "\n- 內建**一鍵載入模擬範例**，沒帶檔案也能 3 秒體驗"
    )
    st.page_link("pages/2_執法簡報直出中心.py", label="👉 立即製作簡報", icon="📽️", use_container_width=True)

with col3:
    st.subheader("🗂️ 3. 檔案加工與轉檔中心")
    st.markdown(
        "公文附件、簡報與巡邏照片日常處理神器："
        "\n- 📝 **角標遮蓋與加頁碼**：可自訂單位署名（如「○○分局交通組製」），支援 PDF／PPTX／圖片"
        "\n- 📄 **PDF 轉檔**：一鍵將長官或上級機關 PDF 簡報轉為 PPTX 或圖片 ZIP"
        "\n- 🖼️ **多圖合併 PDF**：巡邏／蒐證照片按順序合併為單一 PDF 附件"
    )
    st.page_link("pages/3_檔案加工與轉檔中心.py", label="👉 立即加工轉檔", icon="🗂️", use_container_width=True)

st.divider()
st.info("💡 提示：三個工具皆內建「**載入模擬範例資料**」按鈕，首次使用可直接點擊體驗完整運算與匯出效果。")
