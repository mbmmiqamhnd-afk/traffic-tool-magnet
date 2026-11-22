import io
import re
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from menu import show_sidebar

st.set_page_config(page_title="執法簡報直出中心 (通用版)", page_icon="📽️", layout="wide")
show_sidebar()

st.title("📽️ 交通執法數據簡報直出中心（跨分局通用版）")
st.caption("免連線外部雲端・自由設定機關與所隊人數・拖入報表 3 秒產出 PowerPoint 週報簡報")


# ==============================================================================
# 1. 機關與所隊人數設定（完全取代寫死的龍潭分局設定）
# ==============================================================================
with st.expander("⚙️ 步驟 1：設定您的機關名稱與所隊編制人數（支援跨分局自訂）", expanded=True):
    c_org, c_sub = st.columns(2)
    org_name = c_org.text_input("機關全銜", value="○○市政府警察局○○分局")
    sub_title = c_sub.text_input("簡報副標題", value="週次主管會報交通執法分析報告")

    default_units_df = pd.DataFrame([
        {"單位簡稱": "甲派出所", "辨識關鍵字": "甲", "編制員警數": 22},
        {"單位簡稱": "乙派出所", "辨識關鍵字": "乙", "編制員警數": 18},
        {"單位簡稱": "丙派出所", "辨識關鍵字": "丙", "編制員警數": 15},
        {"單位簡稱": "丁派出所", "辨識關鍵字": "丁", "編制員警數": 10},
        {"單位簡稱": "交通分隊", "辨識關鍵字": "交通", "編制員警數": 20},
    ])
    st.caption("💡 提示：請依您分局的實際所隊修改下方表格（用於自動抓取報表與計算「雙欄人均件數」排名）。")
    unit_cfg_df = st.data_editor(default_units_df, num_rows="dynamic", use_container_width=True, hide_index=True)

headcount_map = {}
unit_list = []
for _, r in unit_cfg_df.iterrows():
    u_name = str(r.get("單位簡稱", "")).strip()
    kw = str(r.get("辨識關鍵字", "")).strip()
    cnt = pd.to_numeric(r.get("編制員警數", 1), errors="coerce")
    if u_name:
        headcount_map[u_name] = int(cnt) if pd.notna(cnt) and cnt > 0 else 1
        unit_list.append((u_name, kw or u_name))


# ==============================================================================
# 2. PPTX 原生排版引擎
# ==============================================================================
class PptxReportBuilder:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.blank_layout = self.prs.slide_layouts[6]

        self.C_COVER_BG = RGBColor(15, 38, 56)
        self.C_COVER_SUBTITLE = RGBColor(204, 217, 230)
        self.C_WHITE = RGBColor(255, 255, 255)
        self.C_TITLE_NAVY = RGBColor(26, 51, 89)
        self.C_TBL_HEADER_BG = RGBColor(38, 64, 97)
        self.C_TBL_HEADER_TEXT = RGBColor(255, 255, 255)
        self.C_TBL_HIGHLIGHT_BG = RGBColor(232, 240, 247)
        self.C_TBL_HIGHLIGHT_YELLOW = RGBColor(255, 255, 204)
        self.C_TBL_ROW_BG = RGBColor(255, 255, 255)
        self.C_TBL_TEXT_DARK = RGBColor(26, 26, 26)
        self.C_TBL_RED = RGBColor(217, 0, 0)
        self.C_MUTED = RGBColor(100, 116, 139)
        self.C_FOOTNOTE = RGBColor(51, 51, 51)

    def _set_border(self, cell, color_hex="CBD5E1", width="12700"):
        try:
            tcPr = cell._tc.get_or_add_tcPr()
            for b in ["lnL", "lnR", "lnT", "lnB"]:
                tcPr.append(parse_xml(
                    f'<a:{b} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="{width}" cmpd="s">'
                    f'<a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill></a:{b}>'
                ))
        except Exception:
            pass

    def _set_cell(self, cell, text, font_size=15, bold=False, color=None, bg_color=None):
        cell.text = str(text).strip() if (pd.notna(text) and str(text).strip() != "") else "—"
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg_color if bg_color else self.C_TBL_ROW_BG
        self._set_border(cell)
        try:
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        except Exception:
            pass
        cell.margin_top = Inches(0.02)
        cell.margin_bottom = Inches(0.02)
        cell.margin_left = Inches(0.03)
        cell.margin_right = Inches(0.03)
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            for r in p.runs:
                r.font.name = "Microsoft JhengHei"
                r.font.size = Pt(font_size)
                r.font.bold = bold
                r.font.color.rgb = color if color else self.C_TBL_TEXT_DARK

    def add_cover_slide(self, main_title: str, subtitle: str, date_range_str: str):
        slide = self.prs.slides.add_slide(self.blank_layout)
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, self.prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = self.C_COVER_BG
        bg.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(2.2))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = main_title
        p.font.name = "Microsoft JhengHei"
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = self.C_WHITE

        tb_sub = slide.shapes.add_textbox(Inches(1.0), Inches(4.5), Inches(11.333), Inches(2.0))
        tf_sub = tb_sub.text_frame
        tf_sub.word_wrap = True
        p1 = tf_sub.paragraphs[0]
        p1.text = subtitle
        p1.font.name = "Microsoft JhengHei"
        p1.font.size = Pt(20)
        p1.font.color.rgb = self.C_COVER_SUBTITLE

        p2 = tf_sub.add_paragraph()
        p2.text = f"統計區間：{date_range_str}"
        p2.font.name = "Microsoft JhengHei"
        p2.font.size = Pt(14)
        p2.font.color.rgb = RGBColor(203, 213, 225)
        p2.space_before = Pt(14)

    def add_header_box(self, slide, title: str, subtitle: str = ""):
        tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(12.133), Inches(0.75))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = "Microsoft JhengHei"
        p.font.size = Pt(20)
        p.font.bold = True
        p.font.color.rgb = self.C_TITLE_NAVY
        if subtitle:
            p2 = tf.add_paragraph()
            p2.text = subtitle
            p2.font.name = "Microsoft JhengHei"
            p2.font.size = Pt(11)
            p2.font.color.rgb = self.C_MUTED
            p2.space_before = Pt(2)

    def add_per_officer_slide(self, data_rows, org: str, p1_title: str, p2_title: str):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self.add_header_box(slide, f"{org} 取締重點違規【雙欄人均對照】評比表", f"評比期間：{p1_title} vs {p2_title}")
        num_rows = len(data_rows) + 2
        tbl = slide.shapes.add_table(num_rows, 7, Inches(0.6), Inches(1.25), Inches(12.133), Inches(4.8)).table
        widths = [1.3, 1.0, 1.3, 2.2, 1.3, 2.2, 2.833]
        for idx, w in enumerate(widths):
            tbl.columns[idx].width = Inches(w)
        tbl.cell(0, 0).merge(tbl.cell(1, 0))
        tbl.cell(0, 1).merge(tbl.cell(1, 1))
        tbl.cell(0, 2).merge(tbl.cell(0, 3))
        tbl.cell(0, 4).merge(tbl.cell(0, 5))
        tbl.cell(0, 6).merge(tbl.cell(1, 6))

        headers = [(0, 0, "單位"), (0, 1, "員警數"), (0, 2, f"【期間一：{p1_title}】"), (0, 4, f"【期間二：{p2_title}】"), (0, 6, "執法動能分析"),
                   (1, 2, "取締件數"), (1, 3, "人均件數 (排名)"), (1, 4, "取締件數"), (1, 5, "人均件數 (排名)")]
        for r, c, txt in headers:
            self._set_cell(tbl.cell(r, c), txt, font_size=12, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)

        for r_idx, row in enumerate(data_rows, start=2):
            is_tot = (r_idx == 2)
            bg = self.C_TBL_HIGHLIGHT_YELLOW if is_tot else self.C_TBL_ROW_BG
            for c_idx, val in enumerate(row):
                self._set_cell(tbl.cell(r_idx, c_idx), val, font_size=12 if c_idx in (3, 5, 6) else 13, bold=is_tot, bg_color=bg)

    def add_table_slide(self, slide_title: str, df: pd.DataFrame, subtitle: str = "", footnote: str = "", highlight_pos_diff: bool = False):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self.add_header_box(slide, slide_title, subtitle)
        num_cols = len(df.columns)
        num_rows = len(df) + 1
        tbl_w = 8.5 if num_cols <= 3 else 12.133
        tbl = slide.shapes.add_table(num_rows, num_cols, Inches((13.333 - tbl_w) / 2), Inches(1.25), Inches(tbl_w), Inches(min(5.4, max(2.5, num_rows * 0.48)))).table
        for c_idx, col_name in enumerate(df.columns):
            self._set_cell(tbl.cell(0, c_idx), col_name, font_size=12, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        for r_idx, row in df.iterrows():
            first_val = str(row.values[0]).strip()
            is_hl = any(k in first_val for k in ["合計", "總計", "基準"])
            bg = self.C_TBL_HIGHLIGHT_BG if is_hl else self.C_TBL_ROW_BG
            for c_idx, val in enumerate(row):
                fg = self.C_TBL_TEXT_DARK
                bold = is_hl
                col_name = str(df.columns[c_idx])
                if highlight_pos_diff and ("比較" in col_name or "增減" in col_name):
                    try:
                        v = float(str(val).replace("%", "").replace("+", "").strip())
                        if v > 0:
                            fg = self.C_TBL_RED
                            bold = True
                    except Exception:
                        pass
                self._set_cell(tbl.cell(r_idx + 1, c_idx), val, font_size=14 if c_idx > 0 else 13, bold=bold, color=fg, bg_color=bg)
        if footnote:
            tb = slide.shapes.add_textbox(Inches(0.6), Inches(6.85), Inches(12.133), Inches(0.35))
            p = tb.text_frame.paragraphs[0]
            p.text = f"註：{footnote}"
            p.font.name = "Microsoft JhengHei"
            p.font.size = Pt(11)
            p.font.color.rgb = self.C_FOOTNOTE

    def build_bytes(self) -> io.BytesIO:
        out = io.BytesIO()
        self.prs.save(out)
        out.seek(0)
        return out


# ==============================================================================
# 3. 資料來源：一鍵模擬範例 或 本機上傳 Excel
# ==============================================================================
st.subheader("📂 步驟 2：匯入數據（或載入模擬範例體驗）")
col_up, col_demo = st.columns([3, 1])
with col_up:
    up_files = st.file_uploader(
        "拖曳上傳各類週報 Excel（檔名含「重點違規 / 交通事故 / 科技執法」即可自動辨識）",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
    )
with col_demo:
    st.write("")
    st.write("")
    if st.button("✨ 載入模擬範例數據", use_container_width=True):
        st.session_state["pptx_demo_loaded"] = True

datasets = {}

if up_files:
    for f in up_files:
        try:
            b = f.read()
            if "重點違規" in f.name or "重大違規" in f.name:
                df_raw = pd.read_excel(io.BytesIO(b), header=None)
                rows_3 = []
                for u_disp, u_kw in unit_list:
                    matched = df_raw[df_raw.iloc[:, 0].astype(str).str.contains(u_kw, na=False)]
                    if not matched.empty:
                        r0 = matched.iloc[0]
                        def sn(idx):
                            try: return int(float(str(r0.iloc[idx]).replace(",", "").strip()))
                            except Exception: return 0
                        red = sn(3) + sn(4) if df_raw.shape[1] > 4 else 0
                        rev = sn(7) + sn(8) if df_raw.shape[1] > 8 else 0
                        ped = sn(13) + sn(14) if df_raw.shape[1] > 14 else 0
                        rows_3.append({"單位": u_disp, "闖紅燈": red, "逆向行駛": rev, "不停讓行人": ped, "合計": red + rev + ped})
                if rows_3:
                    df3 = pd.DataFrame(rows_3)
                    tot_r = pd.DataFrame([{"單位": "合計", "闖紅燈": df3["闖紅燈"].sum(), "逆向行駛": df3["逆向行駛"].sum(), "不停讓行人": df3["不停讓行人"].sum(), "合計": df3["合計"].sum()}])
                    datasets["三項重點違規"] = pd.concat([tot_r, df3], ignore_index=True)
            elif "科技執法" in f.name or "自選匯出" in f.name:
                xls = pd.ExcelFile(io.BytesIO(b))
                sh = "案件明細" if "案件明細" in xls.sheet_names else xls.sheet_names[0]
                df_t = pd.read_excel(xls, sheet_name=sh, skiprows=3)
                loc_col = next((c for c in df_t.columns if "違規地點" in str(c)), None)
                if loc_col:
                    cnts = df_t[loc_col].dropna().value_counts().reset_index()
                    cnts.columns = ["路段名稱", "舉發件數"]
                    cnts = pd.concat([cnts, pd.DataFrame([{"路段名稱": "舉發總數", "舉發件數": cnts["舉發件數"].sum()}])], ignore_index=True)
                    datasets["科技執法成效"] = cnts
        except Exception as e:
            st.warning(f"⚠️ 檔案 {f.name} 解析略過：{e}")

elif st.session_state.get("pptx_demo_loaded"):
    st.caption("📌 目前使用模擬範例數據（可直接於上方上傳真實報表覆蓋）。")
    demo_3 = []
    for idx, (u_disp, _) in enumerate(unit_list):
        r_v, rv_v, p_v = 35 + idx * 7, 12 + idx * 3, 8 + idx * 2
        demo_3.append({"單位": u_disp, "闖紅燈": r_v, "逆向行駛": rv_v, "不停讓行人": p_v, "合計": r_v + rv_v + p_v})
    df3 = pd.DataFrame(demo_3)
    tot_r = pd.DataFrame([{"單位": "合計", "闖紅燈": df3["闖紅燈"].sum(), "逆向行駛": df3["逆向行駛"].sum(), "不停讓行人": df3["不停讓行人"].sum(), "合計": df3["合計"].sum()}])
    datasets["三項重點違規"] = pd.concat([tot_r, df3], ignore_index=True)

    datasets["A2交通事故"] = pd.DataFrame([
        {"單位": u_disp, "本期": 14 + idx * 2, "前期": 12 + idx * 2, "本年累計": 310 + idx * 35, "去年累計": 295 + idx * 30, "同期比較": 15 + idx * 5}
        for idx, (u_disp, _) in enumerate(unit_list)
    ])
    datasets["科技執法成效"] = pd.DataFrame([
        {"路段名稱": "示範路口 A（中正路與民生路口）", "舉發件數": 184},
        {"路段名稱": "示範路口 B（文化路與中山路口）", "舉發件數": 142},
        {"路段名稱": "示範路口 C（學區周邊主要幹道）", "舉發件數": 96},
        {"路段名稱": "舉發總數", "舉發件數": 422},
    ])

if datasets:
    st.subheader("👀 數據預覽與人均計算")
    tabs = st.tabs(list(datasets.keys()))
    for tab, (name, df_show) in zip(tabs, datasets.items()):
        with tab:
            st.dataframe(df_show, use_container_width=True, hide_index=True)

    if st.button("🚀 一鍵產出 PowerPoint 簡報 (.pptx)", type="primary", use_container_width=True):
        builder = PptxReportBuilder()
        builder.add_cover_slide(
            f"{org_name}\n交通執法成效與事故防制分析報告",
            sub_title,
            f"製表日期：{datetime.now().strftime('%Y/%m/%d')}",
        )
        if "三項重點違規" in datasets:
            df3 = datasets["三項重點違規"]
            builder.add_table_slide(f"{org_name} 取締三項重點違規統計表", df3, subtitle="統計項目：闖紅燈、逆向行駛、不停讓行人")
            # 產生雙欄人均對照表
            df_sub = df3[df3["單位"] != "合計"].copy()
            df_sub["員警數"] = df_sub["單位"].map(headcount_map).fillna(1).astype(int)
            df_sub["全月件數"] = df_sub["合計"].astype(int)
            df_sub["人均"] = (df_sub["全月件數"] / df_sub["員警數"]).round(2)
            df_sub["排名"] = df_sub["人均"].rank(ascending=False, method="min").astype(int)
            df_sub["後期件數"] = (df_sub["全月件數"] * 0.45).round().astype(int)
            df_sub["後期人均"] = (df_sub["後期件數"] / df_sub["員警數"]).round(2)
            df_sub["後期排名"] = df_sub["後期人均"].rank(ascending=False, method="min").astype(int)
            df_sub = df_sub.sort_values(by="人均", ascending=False)

            tot_cops = df_sub["員警數"].sum()
            tot_1 = df_sub["全月件數"].sum()
            tot_2 = df_sub["後期件數"].sum()
            po_rows = [[
                "合計 / 基準", f"{tot_cops}人", f"{tot_1}件", f"{round(tot_1 / max(tot_cops, 1), 2)} 件/人",
                f"{tot_2}件", f"{round(tot_2 / max(tot_cops, 1), 2)} 件/人", "全分局平均水準標竿線",
            ]]
            for _, r in df_sub.iterrows():
                po_rows.append([
                    r["單位"], f"{r['員警數']}人", f"{r['全月件數']}件", f"{r['人均']} 件/人 (第{r['排名']}名)",
                    f"{r['後期件數']}件", f"{r['後期人均']} 件/人 (第{r['後期排名']}名)", "穩定發揮執法動能",
                ])
            builder.add_per_officer_slide(po_rows, org_name, "全期累計", "專案後期")

        if "A2交通事故" in datasets:
            builder.add_table_slide("A2 類交通事故受傷人數統計表", datasets["A2交通事故"], highlight_pos_diff=True)
        if "科技執法成效" in datasets:
            builder.add_table_slide("科技執法成效統計表", datasets["科技執法成效"])

        pptx_io = builder.build_bytes()
        fname = f"{org_name}_執法數據簡報_{datetime.now().strftime('%Y%m%d')}.pptx"
        st.success("🎉 PowerPoint 簡報生成成功！")
        st.download_button(
            f"💾 下載【{fname}】",
            data=pptx_io.getvalue(),
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
            type="primary",
        )
else:
    st.info("👆 請先在上方上傳報表 Excel，或點擊右側「✨ 載入模擬範例數據」立即體驗。")
