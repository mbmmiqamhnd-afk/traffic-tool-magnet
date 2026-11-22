import streamlit as st


def show_sidebar():
    with st.sidebar:
        st.title("🚓 交通執法工具磁鐵")
        st.caption("跨分局通用版｜免註冊・無浮水印・本機即時運算")

        st.success(
            "🔒 **資安與隱私聲明**\n\n"
            "所有上傳檔案僅於本次連線記憶體中即時運算與轉檔，"
            "絕不留存、不上傳資料庫，亦不寄送至任何外部信箱。"
        )

        st.divider()
        st.subheader("🧭 功能導覽")
        st.page_link("app.py", label="首頁與使用說明", icon="🏠")
        st.page_link("pages/1_專案敘獎一鍵精算器.py", label="專案敘獎一鍵精算器", icon="🏆")
        st.page_link("pages/2_執法簡報直出中心.py", label="執法簡報直出中心 (PPTX)", icon="📽️")
        st.page_link("pages/3_檔案加工與轉檔中心.py", label="檔案加工與轉檔中心", icon="🗂️")

        st.divider()
        st.markdown(
            "💡 **想看更多 AI 自動化與第二大腦實戰筆記？**\n\n"
            "👉 [前往作者個人網站](https://mbmmiqamhnd-afk.github.io)"
        )
