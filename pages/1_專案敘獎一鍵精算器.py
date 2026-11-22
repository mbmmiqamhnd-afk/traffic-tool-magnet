import io
import re
import numpy as np
import pandas as pd
import streamlit as st
from menu import show_sidebar

st.set_page_config(page_title="專案敘獎一鍵精算器", page_icon="🏆", layout="wide")
show_sidebar()

st.title("🏆 警察交通專案敘獎一鍵精算器（跨分局通用版）")
st.caption("免註冊・無浮水印・本機記憶體即時運算，支援上傳系統『自選匯出.xlsx』或點擊範例立即體驗。")

tab_drug, tab_danger, tab_unlic, tab_plate = st.tabs([
    "🧪 毒駕專案敘獎",
    "🚨 危駕與改裝車取締",
    "🚗 無照駕駛移置保管",
    "🔎 偽變造車牌專案",
])


# ==============================================================================
# Tab 1: 毒駕專案敘獎 (通用化自 p31)
# ==============================================================================
def classify_drug_violation(fact_text, law_code=""):
    text = f"{law_code} {fact_text}".strip()
    if not text or text == "nan":
        return "未填寫"
    is_drug = bool(re.search(r"毒|毒品|麻醉|迷幻|藥駕|第1項第2款|35102|35300175|35300178|35300184|35300199|35300211|35300253|35300256|35402002", text))
    is_alcohol = bool(re.search(r"酒|酒精|呼氣|吐氣|第1項第1款|第一項第一款|1項1款|35101|35700|35402001", text))
    is_refusal = bool(re.search(r"拒測|拒絕接受|拒絕|35402", text))
    is_recidivism = bool(re.search(r"累犯|二次以上|第2次|第3次|多次", text))
    is_slow = bool(re.search(r"73|慢車|自行車|腳踏車|微型電動二輪車", text))
    is_impound = bool(re.search(r"第三十五條第一、三、四、五項之情形之一|第35條第1、3、4、5項之情形之一|移置|保管|35900", text))

    if is_refusal:
        return "🧪 毒駕拒測" if is_drug else ("🍺 酒駕拒測" if is_alcohol else "⚠️ 拒測(未指明)")
    if is_recidivism:
        return "🧪 毒駕累犯" if is_drug else ("🍺 酒駕累犯" if is_alcohol else "⚠️ 累犯")
    if is_slow:
        return "🚲 慢車毒駕" if is_drug else ("🚲 慢車酒駕" if is_alcohol else "🚲 慢車其他")
    if is_impound:
        return "🚗 車輛移置保管"
    if is_drug:
        return "🧪 毒駕本體"
    if is_alcohol:
        return "🍺 一般酒駕"
    return "📋 其他違規"


def classify_drug_vehicle(v_name):
    v = str(v_name).strip()
    if any(k in v for k in ["大貨", "大客", "聯結", "曳引", "大型車"]):
        return "大型車"
    if any(k in v for k in ["汽車", "自小客", "自小貨", "小型車", "小貨", "小客"]):
        return "小型車"
    if any(k in v for k in ["微型電動二輪車", "微型電動", "重機", "輕機", "機車"]):
        return "機車"
    if any(k in v for k in ["電動輔助", "自行車", "腳踏", "慢車"]):
        return "其餘慢車(不納入)"
    return "其他車種"


def calc_drug_merits(group):
    group = group.sort_values(by="入案日")
    heavy = car = moto = slow = refusal = 0
    tickets = []
    for _, row in group.iterrows():
        v_type = str(row["車種判定"])
        cat = str(row["案件分類"])
        tickets.append(str(row["單號"]))
        if "拒測" in cat:
            refusal += 1
        elif v_type == "大型車":
            heavy += 1
        elif v_type == "小型車":
            car += 1
        elif v_type == "機車":
            moto += 1
        elif v_type == "其餘慢車(不納入)":
            slow += 1

    merit_cnt = heavy
    total_pts = (car * 2) + moto + refusal
    c2 = total_pts // 2
    c1 = total_pts % 2

    parts = []
    if merit_cnt > 0:
        parts.append(f"記功一次{merit_cnt}次" if merit_cnt > 1 else "記功一次")
    if c2 > 0:
        parts.append(f"嘉獎二次{c2}次" if c2 > 1 else "嘉獎二次")
    if c1 > 0:
        parts.append(f"嘉獎一次{c1}次" if c1 > 1 else "嘉獎一次")

    reasons = []
    if heavy > 0:
        reasons.append(f"大型車毒駕{heavy}件")
    if car > 0:
        reasons.append(f"小型車毒駕{car}件")
    if moto > 0:
        reasons.append(f"機車(含微電車)毒駕{moto}件")
    if refusal > 0:
        reasons.append(f"毒駕拒測{refusal}件")

    return pd.Series({
        "大型車(件)": heavy,
        "小型車(件)": car,
        "機車含微電車(件)": moto,
        "其餘慢車(不納入)": slow,
        "毒駕拒測(件)": refusal,
        "計獎總件數": heavy + car + moto + refusal,
        "記功一次": merit_cnt,
        "嘉獎二次": c2,
        "嘉獎一次": c1,
        "建議獎勵額度": "、".join(parts) if parts else "未達標準",
        "具體出力事由": "執行加強攔查取締毒駕專案工作計畫，查獲" + "、".join(reasons) if reasons else "",
        "舉發單號明細": ", ".join(tickets),
    })


def read_standard_case_sheet(file_obj, required_cols):
    xls = pd.ExcelFile(file_obj)
    target_sheet = next((s for s in xls.sheet_names if "案件明細" in s), xls.sheet_names[0])
    df_raw = pd.read_excel(xls, sheet_name=target_sheet, header=None)
    header_idx = None
    for idx in range(min(15, len(df_raw))):
        row_str = " ".join(df_raw.iloc[idx].dropna().astype(str).tolist())
        if "單號" in row_str and ("舉發員警" in row_str or "員警" in row_str):
            header_idx = idx
            break
    if header_idx is None:
        raise ValueError("找不到包含「單號」與「舉發員警」的標題列，請確認工作表格式。")
    cols = [str(x).strip() for x in df_raw.iloc[header_idx]]
    df = df_raw.iloc[header_idx + 1 :].copy().reset_index(drop=True)
    df.columns = cols
    df = df[df["單號"].notna()]
    df = df[~df["單號"].astype(str).str.contains("列印人員|統計期間|製表人員|總計", na=False)]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"缺少必要欄位：{', '.join(missing)}")
    return df[required_cols].dropna(how="all")


with tab_drug:
    st.subheader("🧪 毒駕專案取締敘獎統計")
    st.info(
        "計獎標準：大型車記功一次｜小型車 2 點（嘉獎二次）｜機車及微電車 1 點（嘉獎一次）｜毒駕拒測 1 點（獨立計算）。"
    )
    c_u1, c_d1 = st.columns([3, 1])
    with c_u1:
        up_drug = st.file_uploader("上傳『自選匯出.xlsx』（含 單號、簡式車種名稱、違規法條1、違規事實1、入案日、舉發員警1）", type=["xlsx"], key="up_drug")
    with c_d1:
        st.write("")
        st.write("")
        if st.button("✨ 載入模擬範例", key="demo_drug_btn", use_container_width=True):
            st.session_state["demo_drug_df"] = pd.DataFrame([
                {"單號": "DEMO001", "簡式車種名稱": "自小客車", "違規法條1": "35102001", "違規事實1": "施用毒品後駕駛小型車", "入案日": "1150911", "舉發員警1": "範例警員甲"},
                {"單號": "DEMO002", "簡式車種名稱": "普通重型機車", "違規法條1": "35102002", "違規事實1": "施用毒品後駕駛機車", "入案日": "1150912", "舉發員警1": "範例警員甲"},
                {"單號": "DEMO003", "簡式車種名稱": "大貨車", "違規法條1": "35102001", "違規事實1": "施用毒品後駕駛大貨車", "入案日": "1150913", "舉發員警1": "範例警員乙"},
                {"單號": "DEMO004", "簡式車種名稱": "普通重型機車", "違規法條1": "35402002", "違規事實1": "拒絕接受毒品測試檢定", "入案日": "1150914", "舉發員警1": "範例警員乙"},
            ])

    df_drug_src = None
    if up_drug is not None:
        try:
            df_drug_src = read_standard_case_sheet(up_drug, ["單號", "簡式車種名稱", "違規法條1", "違規事實1", "入案日", "舉發員警1"])
        except Exception as e:
            st.error(f"❌ 讀取失敗：{e}")
    elif "demo_drug_df" in st.session_state:
        df_drug_src = st.session_state["demo_drug_df"].copy()
        st.caption("📌 目前顯示為模擬範例資料，可直接上傳您的 Excel 覆蓋。")

    if df_drug_src is not None and not df_drug_src.empty:
        df_drug_src["案件分類"] = df_drug_src.apply(lambda r: classify_drug_violation(r["違規事實1"], r["違規法條1"]), axis=1)
        df_drug_src["車種判定"] = df_drug_src["簡式車種名稱"].apply(classify_drug_vehicle)
        df_valid = df_drug_src[df_drug_src["案件分類"].str.contains("🧪")].copy()
        if df_valid.empty:
            st.warning("⚠️ 檔案中未偵測到毒駕專案相關案件。")
        else:
            res = df_valid.groupby("舉發員警1").apply(calc_drug_merits).reset_index()
            res = res.sort_values(by=["記功一次", "嘉獎二次", "嘉獎一次", "計獎總件數"], ascending=False).reset_index(drop=True)
            st.dataframe(res, use_container_width=True, hide_index=True)

            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as w:
                res.to_excel(w, index=False, sheet_name="出力人員敘獎名冊")
                df_valid.to_excel(w, index=False, sheet_name="專案案件明細")
            st.download_button("📥 下載毒駕敘獎統計表 (.xlsx)", out.getvalue(), "毒駕專案敘獎統計表.xlsx", use_container_width=True, type="primary")


# ==============================================================================
# Tab 2: 危駕與改裝車取締 (通用化自 p28)
# ==============================================================================
def calc_tiered_rewards(row, col_a, col_b):
    total_points = row[col_a] * 2 + row[col_b] * 1
    jj = 0
    mult = 1
    while True:
        cost = 4 * mult
        if total_points >= cost:
            total_points -= cost
            jj += 1
            if jj % 9 == 0:
                mult += 1
        else:
            break
    return pd.Series([jj, jj // 9, jj % 9])


with tab_danger:
    st.subheader("🚨 危險駕車與改裝車輛取締獎勵統計")
    st.info("計獎標準：Group A（13-1、18-1、43-3）每 2 件嘉獎 1 次｜Group B（16-1-1 機車、16-1-2 排氣管消音器、43-1）每 4 件嘉獎 1 次；累積滿 9 次嘉獎後門檻加倍。")

    c_u2, c_d2 = st.columns([3, 1])
    with c_u2:
        up_danger = st.file_uploader("上傳舉發違規資料檔（含『list2』工作表）", type=["xlsx"], key="up_danger")
    with c_d2:
        st.write("")
        st.write("")
        if st.button("✨ 載入模擬範例", key="demo_danger_btn", use_container_width=True):
            st.session_state["demo_danger_df"] = pd.DataFrame([
                {"單號": "D01", "車牌": "ABC-1234", "車種": "普通重型機車", "違規日期": "1150901", "條款1": "1810001", "違規事實1": "變更車體重要設備", "舉發員警": "範例警員甲", "舉發單位": "甲派出所"},
                {"單號": "D02", "車牌": "DEF-5678", "車種": "普通重型機車", "違規日期": "1150902", "條款1": "4330001", "違規事實1": "拆除消音器", "舉發員警": "範例警員甲", "舉發單位": "甲派出所"},
                {"單號": "D03", "車牌": "GHI-9012", "車種": "普通重型機車", "違規日期": "1150903", "條款1": "1610201", "違規事實1": "排氣管未依規定裝置", "舉發員警": "範例警員乙", "舉發單位": "乙派出所"},
                {"單號": "D04", "車牌": "JKL-3456", "車種": "普通重型機車", "違規日期": "1150904", "條款1": "4310101", "違規事實1": "蛇行危險駕駛", "舉發員警": "範例警員乙", "舉發單位": "乙派出所"},
                {"單號": "D05", "車牌": "MNO-7890", "車種": "普通重型機車", "違規日期": "1150905", "條款1": "1610101", "違規事實1": "機車改裝排氣管", "舉發員警": "範例警員乙", "舉發單位": "乙派出所"},
                {"單號": "D06", "車牌": "PQR-1122", "車種": "普通重型機車", "違規日期": "1150906", "條款1": "4310301", "違規事實1": "行車速度超過最高速限40公里", "舉發員警": "範例警員乙", "舉發單位": "乙派出所"},
            ])

    df_d_src = None
    if up_danger is not None:
        try:
            xls_d = pd.ExcelFile(up_danger)
            sh_d = "list2" if "list2" in xls_d.sheet_names else xls_d.sheet_names[0]
            df_d_src = pd.read_excel(xls_d, sheet_name=sh_d)
        except Exception as e:
            st.error(f"❌ 讀取失敗：{e}")
    elif "demo_danger_df" in st.session_state:
        df_d_src = st.session_state["demo_danger_df"].copy()
        st.caption("📌 目前顯示為模擬範例資料，可直接上傳您的 Excel 覆蓋。")

    if df_d_src is not None and not df_d_src.empty:
        df_d_src["條款1"] = df_d_src["條款1"].astype(str).str.strip()
        df_d_src["違規事實1"] = df_d_src["違規事實1"].astype(str).fillna("")
        df_d_src["車種"] = df_d_src["車種"].astype(str).fillna("")
        df_d_src["舉發員警"] = df_d_src["舉發員警"].astype(str).str.strip()

        mask_a = df_d_src["條款1"].str.startswith(("131", "181", "433"))
        mask_b = (
            (df_d_src["條款1"].str.startswith("16101") & df_d_src["車種"].str.contains("機車|重型|輕型", na=False))
            | (df_d_src["條款1"].str.startswith("16102") & df_d_src["違規事實1"].str.contains("排氣管|消音器", na=False))
            | df_d_src["條款1"].str.startswith(("43101", "43103", "43104", "43105"))
        )
        df_a, df_b = df_d_src[mask_a], df_d_src[mask_b]
        col_a = "Group A (2件1嘉獎)"
        col_b = "Group B (4件1嘉獎)"
        rw = pd.concat([df_a["舉發員警"].value_counts().rename(col_a), df_b["舉發員警"].value_counts().rename(col_b)], axis=1).fillna(0).astype(int)
        rw = rw.drop(index="nan", errors="ignore")
        if not rw.empty:
            rw[["嘉獎次數", "大功數(倍增指標)", "階梯嘉獎數"]] = rw.apply(lambda r: calc_tiered_rewards(r, col_a, col_b), axis=1)
            rw = rw.sort_values(by=["嘉獎次數", col_a], ascending=False).reset_index().rename(columns={"index": "舉發員警", "舉發員警": "舉發員警"})
            st.dataframe(rw, use_container_width=True, hide_index=True)

            out_d = io.BytesIO()
            with pd.ExcelWriter(out_d, engine="xlsxwriter") as w:
                rw.to_excel(w, index=False, sheet_name="獎勵結算表")
                df_a.to_excel(w, index=False, sheet_name="GroupA_明細")
                df_b.to_excel(w, index=False, sheet_name="GroupB_明細")
            st.download_button("📥 下載危駕取締獎勵報表 (.xlsx)", out_d.getvalue(), "危駕與改裝車取締獎勵結算表.xlsx", use_container_width=True, type="primary")


# ==============================================================================
# Tab 3: 無照駕駛移置保管 (通用化自 p29)
# ==============================================================================
def parse_age_num(age_str):
    try:
        if pd.isna(age_str):
            return np.nan
        return int(re.sub(r"\D", "", str(age_str)))
    except Exception:
        return np.nan


def cat_unlicensed(row):
    fact = str(row["違規事實1"])
    law = str(row["違規法條1"])
    age = row["age_num"]
    if pd.isna(age):
        return "不採計"
    if age < 18:
        return "未成年"
    if "機車" not in fact and ("小型車" in fact or "汽車駕駛人" in fact):
        if law.startswith(("21101", "21104", "21105", "212")):
            return "成年"
    return "不採計"


with tab_unlic:
    st.subheader("🚗 無照駕駛車輛移置保管敘獎統計")
    st.info("計獎標準：未成年無照每 5 件嘉獎 1 次（每季上限 2 次）｜成年小型車無照每 3 件嘉獎 1 次（每季上限 2 次）｜未達標件數保留至下一季。")

    c_u3, c_d3 = st.columns([3, 1])
    with c_u3:
        up_unlic = st.file_uploader("上傳『自選匯出.xlsx』（含『案件明細』工作表）", type=["xlsx"], key="up_unlic")
    with c_d3:
        st.write("")
        st.write("")
        if st.button("✨ 載入模擬範例", key="demo_unlic_btn", use_container_width=True):
            st.session_state["demo_unlic_df"] = pd.DataFrame([
                {"單號": "U01", "違規法條1": "2110101", "違規事實1": "未領有駕駛執照駕駛小型車", "入案日": "1150210", "舉發員警1": "範例警員甲", "違規人年齡": "25歲"},
                {"單號": "U02", "違規法條1": "2110101", "違規事實1": "未領有駕駛執照駕駛小型車", "入案日": "1150315", "舉發員警1": "範例警員甲", "違規人年齡": "30歲"},
                {"單號": "U03", "違規法條1": "2110101", "違規事實1": "未領有駕駛執照駕駛小型車", "入案日": "1150420", "舉發員警1": "範例警員甲", "違規人年齡": "22歲"},
                {"單號": "U04", "違規法條1": "2110101", "違規事實1": "未領有駕駛執照駕駛機車", "入案日": "1150218", "舉發員警1": "範例警員乙", "違規人年齡": "16歲"},
                {"單號": "U05", "違規法條1": "2110101", "違規事實1": "未領有駕駛執照駕駛機車", "入案日": "1150502", "舉發員警1": "範例警員乙", "違規人年齡": "17歲"},
            ])

    df_u_src = None
    if up_unlic is not None:
        try:
            xls_u = pd.ExcelFile(up_unlic)
            sh_u = "案件明細" if "案件明細" in xls_u.sheet_names else xls_u.sheet_names[0]
            raw_u = pd.read_excel(xls_u, sheet_name=sh_u, header=None)
            h_idx = raw_u[raw_u.iloc[:, 1] == "違規法條1"].index[0]
            df_u_src = raw_u.iloc[h_idx + 1 :].copy()
            df_u_src.columns = ["單號", "違規法條1", "違規事實1", "入案日", "舉發員警1", "違規人年齡"]
            df_u_src = df_u_src.dropna(subset=["單號", "舉發員警1"])
        except Exception as e:
            st.error(f"❌ 讀取失敗：{e}")
    elif "demo_unlic_df" in st.session_state:
        df_u_src = st.session_state["demo_unlic_df"].copy()
        st.caption("📌 目前顯示為模擬範例資料，可直接上傳您的 Excel 覆蓋。")

    if df_u_src is not None and not df_u_src.empty:
        df_u_src["age_num"] = df_u_src["違規人年齡"].apply(parse_age_num)
        df_u_src["案件類別"] = df_u_src.apply(cat_unlicensed, axis=1)
        df_uv = df_u_src[df_u_src["案件類別"] != "不採計"].copy()

        def get_q(d_str):
            try:
                m = int(str(d_str)[3:5])
                return 1 if m <= 3 else (2 if m <= 6 else (3 if m <= 9 else 4))
            except Exception:
                return 1

        if not df_uv.empty:
            df_uv["季別"] = df_uv["入案日"].apply(get_q)
            piv = df_uv.pivot_table(index="舉發員警1", columns=["案件類別", "季別"], aggfunc="size", fill_value=0)
            rows_u = []
            for officer, r in piv.iterrows():
                jq1, aq1 = r.get(("未成年", 1), 0), r.get(("成年", 1), 0)
                jm1, j_carry = min(jq1 // 5, 2), jq1 - min(jq1 // 5, 2) * 5
                am1, a_carry = min(aq1 // 3, 2), aq1 - min(aq1 // 3, 2) * 3
                jq2, aq2 = r.get(("未成年", 2), 0) + j_carry, r.get(("成年", 2), 0) + a_carry
                jm2, am2 = min(jq2 // 5, 2), min(aq2 // 3, 2)
                rows_u.append({
                    "舉發員警": officer,
                    "Q1未成年(件)": jq1, "Q1成年(件)": aq1, "Q1嘉獎": jm1 + am1,
                    "Q2未成年(含保留)": jq2, "Q2成年(含保留)": aq2, "Q2嘉獎": jm2 + am2,
                    "總嘉獎數": jm1 + am1 + jm2 + am2,
                })
            res_u = pd.DataFrame(rows_u)
            st.dataframe(res_u, use_container_width=True, hide_index=True)

            out_u = io.BytesIO()
            with pd.ExcelWriter(out_u, engine="xlsxwriter") as w:
                res_u.to_excel(w, index=False, sheet_name="敘獎結算表")
                df_uv.to_excel(w, index=False, sheet_name="採計案件明細")
            st.download_button("📥 下載無照移置敘獎統計表 (.xlsx)", out_u.getvalue(), "無照駕駛移置敘獎統計表.xlsx", use_container_width=True, type="primary")
        else:
            st.warning("⚠️ 無符合敘獎資格之案件。")


# ==============================================================================
# Tab 4: 偽變造車牌專案 (通用化自 p30)
# ==============================================================================
def calc_fake_plate(group):
    group = group.sort_values(by="入案日")
    fp_cases = fp_merits = ot_cases = ot_merits = 0
    tickets = []
    for _, row in group.iterrows():
        viol = str(row["違規事實1"])
        veh = str(row["簡式車種名稱"])
        tickets.append(str(row["單號"]))
        mult = 2 if (fp_merits + ot_merits) >= 9 else 1
        if "偽造" in viol or "變造" in viol:
            fp_cases += 1
            fp_merits += (2 if "汽車" in veh else 1) * mult
        elif "他車" in viol:
            ot_cases += 1
            if ot_cases % 2 == 0:
                ot_merits += 1 * mult
    return pd.Series({
        "【偽變造車牌】件數": fp_cases,
        "【偽變造車牌】嘉獎數": fp_merits,
        "【懸掛他車號牌】件數": ot_cases,
        "【懸掛他車號牌】嘉獎數": ot_merits,
        "總件數合計": fp_cases + ot_cases,
        "總嘉獎合計": fp_merits + ot_merits,
        "舉發單號明細": ", ".join(tickets),
    })


with tab_plate:
    st.subheader("🔎 偽變造車牌專案敘獎統計")
    st.info("計獎標準：偽變造車牌汽車嘉獎 2 次、機車 1 次｜懸掛他車號牌每 2 件嘉獎 1 次｜累計滿 9 次嘉獎後獎勵加倍。")

    c_u4, c_d4 = st.columns([3, 1])
    with c_u4:
        up_plate = st.file_uploader("上傳『自選匯出.xlsx』（含 單號、簡式車種名稱、違規法條1、違規事實1、入案日、舉發員警1）", type=["xlsx"], key="up_plate")
    with c_d4:
        st.write("")
        st.write("")
        if st.button("✨ 載入模擬範例", key="demo_plate_btn", use_container_width=True):
            st.session_state["demo_plate_df"] = pd.DataFrame([
                {"單號": "P01", "簡式車種名稱": "自用小客車", "違規法條1": "1210301", "違規事實1": "使用偽造之牌照", "入案日": "1150901", "舉發員警1": "範例警員甲"},
                {"單號": "P02", "簡式車種名稱": "普通重型機車", "違規法條1": "1210501", "違規事實1": "牌照借供他車使用或使用他車牌照", "入案日": "1150902", "舉發員警1": "範例警員甲"},
                {"單號": "P03", "簡式車種名稱": "普通重型機車", "違規法條1": "1210501", "違規事實1": "牌照借供他車使用或使用他車牌照", "入案日": "1150903", "舉發員警1": "範例警員甲"},
                {"單號": "P04", "簡式車種名稱": "普通重型機車", "違規法條1": "1210301", "違規事實1": "使用變造之牌照", "入案日": "1150904", "舉發員警1": "範例警員乙"},
            ])

    df_p_src = None
    if up_plate is not None:
        try:
            df_p_src = read_standard_case_sheet(up_plate, ["單號", "簡式車種名稱", "違規法條1", "違規事實1", "入案日", "舉發員警1"])
        except Exception as e:
            st.error(f"❌ 讀取失敗：{e}")
    elif "demo_plate_df" in st.session_state:
        df_p_src = st.session_state["demo_plate_df"].copy()
        st.caption("📌 目前顯示為模擬範例資料，可直接上傳您的 Excel 覆蓋。")

    if df_p_src is not None and not df_p_src.empty:
        res_p = df_p_src.groupby("舉發員警1").apply(calc_fake_plate).reset_index()
        res_p = res_p.sort_values(by=["總嘉獎合計", "總件數合計"], ascending=False).reset_index(drop=True)
        st.dataframe(res_p, use_container_width=True, hide_index=True)

        out_p = io.BytesIO()
        with pd.ExcelWriter(out_p, engine="openpyxl") as w:
            res_p.to_excel(w, index=False, sheet_name="敘獎統計表")
            df_p_src.sort_values(by=["舉發員警1", "入案日"]).to_excel(w, index=False, sheet_name="案件明細")
        st.download_button("📥 下載偽變造車牌敘獎統計表 (.xlsx)", out_p.getvalue(), "偽變造車牌專案敘獎統計.xlsx", use_container_width=True, type="primary")
