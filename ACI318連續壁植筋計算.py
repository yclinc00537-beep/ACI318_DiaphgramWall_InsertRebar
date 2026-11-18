import io
import math
from datetime import datetime
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
import streamlit as st

# ==========================================
# 1. 頁面設定與資料庫 (避開中括號確保語法穩定)
# ==========================================
st.set_page_config(
    page_title="連續壁錯位植筋反向檢核與替代方案系統",
    page_icon="🧱",
    layout="wide",
)

REBAR_DB = {
    "#4 (D13)": {"db": 1.27, "Ab": 1.267},
    "#5 (D16)": {"db": 1.59, "Ab": 1.986},
    "#6 (D19)": {"db": 1.91, "Ab": 2.865},
    "#7 (D22)": {"db": 2.22, "Ab": 3.871},
    "#8 (D25)": {"db": 2.54, "Ab": 5.067},
    "#9 (D29)": {"db": 2.87, "Ab": 6.469},
    "#10 (D32)": {"db": 3.22, "Ab": 8.143},
}

STEEL_DB = {
    "SD420W (fy=4,200 / fu=6,300)": {"fy": 4200.0, "fu": 6300.0},
    "SD280W (fy=2,800 / fu=4,200)": {"fy": 2800.0, "fu": 4200.0},
    "SD490W (fy=4,900 / fu=6,860)": {"fy": 4900.0, "fu": 6860.0},
}


# ==========================================
# 2. 核心反算引擎 (第 25 章 & 第 17 章)
# ==========================================
def calc_ch25_required_depth(db, fy, fc, spacing_s, is_lap_splice):
    """依據 ACI 318 第 25 章反算連續壁單排植筋最小所需深度"""
    # 連續壁保護層與間距取半中心距 cb = s / 2 (並扣除淨間距影響)
    cb = max((spacing_s - db) / 2.0, 1.0 * db)
    ktr = 0.0  # 連續壁後置植筋無橫向圍束箍筋
    confinement = min((cb + ktr) / db, 2.5)

    psi_t = 1.0  # 垂直壁面水平植筋
    psi_e = 1.0
    psi_s = 0.8 if db <= 1.91 else 1.0
    psi_g = 1.0 if fy <= 4200 else 1.15

    ld_raw = (
        (0.28 * fy / (1.0 * math.sqrt(fc)))
        * ((psi_t * psi_e * psi_s * psi_g) / confinement)
        * db
    )
    ld = max(ld_raw, 30.0)
    req_depth = ld * 1.3 if is_lap_splice else ld
    return req_depth, confinement


def calc_ch17_group_capacity(hef, db, Ab, fy, fu, fc, tau_k, spacing_s, n_row=10):
    """計算整排連續壁植筋(取10支代表連續排)在指定埋深 hef 下的單支平均設計強度與標稱強度"""
    futa = min(fu, 1.9 * fy, 8800.0)
    phi_Nsa = 0.75 * (Ab * futa)

    # 1. 群組混凝土錐狀破裂 Ncbg (整排連續排列)
    Nb = 16.0 * 1.0 * math.sqrt(fc) * (hef**1.5)
    ANco = 9.0 * (hef**2)
    # 假設連續壁上下方向無邊距截斷 (3*hef)，水平方向受中心距 spacing_s 截斷
    bx = (n_row - 1) * min(spacing_s, 3.0 * hef) + 3.0 * hef
    by = 3.0 * hef
    ANc = bx * by
    Ncbg_group = (ANc / ANco) * 1.0 * 1.0 * Nb  # 開裂混凝土 psi_c,N = 1.0
    Ncbg_per_bar = Ncbg_group / n_row
    phi_Ncb_per_bar = 0.65 * Ncbg_per_bar

    # 2. 群組化學黏結強度 Nag
    Nba = 1.0 * tau_k * math.pi * db * hef
    cNa = 10.0 * db * math.sqrt(tau_k / 77.3)
    ANao = (2.0 * cNa) ** 2
    bxa = (n_row - 1) * min(spacing_s, 2.0 * cNa) + 2.0 * cNa
    bya = 2.0 * cNa
    ANa = bxa * bya
    Nag_group = (ANa / ANao) * 1.0 * Nba
    Nag_per_bar = Nag_group / n_row
    phi_Na_per_bar = 0.65 * Nag_per_bar

    phi_Nn = min(phi_Nsa, phi_Ncb_per_bar, phi_Na_per_bar)
    nominal_non_steel = min(Ncbg_per_bar, Nag_per_bar)
    return phi_Nn, phi_Nsa, phi_Ncb_per_bar, phi_Na_per_bar, nominal_non_steel


def solve_ch17_required_depth(
    db, Ab, fy, fu, fc, tau_k, spacing_s, stress_ratio, require_ductile
):
    """反向二分搜尋法：求出滿足目標受力與延性條件的最小埋深 hef (cm)"""
    target_Nu = (Ab * fy) * stress_ratio
    max_hef_limit = 20.0 * db
    min_hef_limit = 4.0 * db

    # 先檢查在 20*db 極限深度下是否能達標
    phi_Nn_max, _, _, _, nom_ns_max = calc_ch17_group_capacity(
        max_hef_limit, db, Ab, fy, fu, fc, tau_k, spacing_s
    )
    ductile_ok_max = (nom_ns_max >= Ab * fy) if require_ductile else True

    if phi_Nn_max < target_Nu or not ductile_ok_max:
        return None, phi_Nn_max, target_Nu, max_hef_limit

    low = min_hef_limit
    high = max_hef_limit
    best_hef = max_hef_limit

    for _ in range(40):
        mid = (low + high) / 2.0
        phi_Nn, _, _, _, nom_ns = calc_ch17_group_capacity(
            mid, db, Ab, fy, fu, fc, tau_k, spacing_s
        )
        cond_strength = phi_Nn >= target_Nu
        cond_ductile = (nom_ns >= Ab * fy) if require_ductile else True

        if cond_strength and cond_ductile:
            best_hef = mid
            high = mid
        else:
            low = mid

    return best_hef, phi_Nn_max, target_Nu, max_hef_limit


# ==========================================
# 3. 自動產生替代方案演算法 (降號數、雙排錯開)
# ==========================================
def find_alternative_schemes(
    orig_rebar_key, orig_Ab, orig_spacing, max_allow_depth, fy, fu, fc, tau_k, is_lap_splice, stress_ratio, require_ductile
):
    """當原號數超限時，搜尋所有符合 <= max_allow_depth 的小號數加密替代方案"""
    alternatives = []
    # 候選小號數鋼筋清單 (由大到小測試)
    candidate_keys = ("#7 (D22)", "#6 (D19)", "#5 (D16)", "#4 (D13)")

    for cand_key in candidate_keys:
        cand_db = REBAR_DB[cand_key]["db"]
        cand_Ab = REBAR_DB[cand_key]["Ab"]

        # 只考慮比原鋼筋小的號數
        if cand_Ab >= orig_Ab:
            continue

        # 計算每支原鋼筋需要幾支候選小鋼筋替代 (倍數 mult = 2, 3, 4)
        for mult in (2, 3, 4):
            total_new_Ab = mult * cand_Ab
            if total_new_Ab < orig_Ab * 0.96:  # 允許 4% 內工程誤差
                continue

            # 若拆成 mult 支，採用雙排交錯配置，每排等效水平中心距可放大
            rows = 2 if mult >= 2 else 1
            eff_spacing = orig_spacing * (rows / mult)
            # 確保交錯後的實際淨中心距不小於 10 cm (斜向中心距檢核)
            diag_spacing = math.hypot(eff_spacing, 10.0) if rows == 2 else eff_spacing
            if diag_spacing < 9.5:
                continue

            # 檢核第 25 章深度
            d25, _ = calc_ch25_required_depth(cand_db, fy, fc, diag_spacing, is_lap_splice)
            # 檢核第 17 章深度
            d17, _, _, _ = solve_ch17_required_depth(
                cand_db, cand_Ab, fy, fu, fc, tau_k, diag_spacing, stress_ratio, require_ductile
            )

            valid_d25 = d25 <= max_allow_depth
            valid_d17 = (d17 is not None) and (d17 <= max_allow_depth)

            if valid_d25 or valid_d17:
                rec_depth = math.ceil(min(d25 if valid_d25 else 999, d17 if valid_d17 else 999))
                method_tag = (
                    "第 25 章伸展法"
                    if valid_d25 and (not valid_d17 or d25 <= d17 + 3)
                    else "第 17 章高強藥劑法"
                )
                alternatives.append(
                    {
                        "方案說明": f"每 1 支 {orig_rebar_key} ➔ 改用 {mult} 支 {cand_key} (上下雙排錯開)",
                        "斷面積比": f"{total_new_Ab:.2f} / {orig_Ab:.2f} cm² ({total_new_Ab/orig_Ab*100:.0f}%)",
                        "建議最小埋深": f"{rec_depth} cm (≤ {max_allow_depth} cm)",
                        "適用規範路徑": method_tag,
                        "現場配置提醒": f"雙排上下排距 10 cm 梅花錯開，斜向中心距 {diag_spacing:.1f} cm",
                    }
                )
                break  # 該號數找到最小可行支數即跳下一號數

    return alternatives


# ==========================================
# 4. Word 計算書產生函數
# ==========================================
def build_word_doc(inputs_dict, summary_dict, alt_list):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Microsoft JhengHei"
    style.font.size = Pt(10.5)

    h = doc.add_heading("連續壁預埋筋錯位—後置植筋反向檢核與替代方案計算書", level=1)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph(
        f"計算時間：{datetime.now().strftime('%Y-%m-%d %H:%M')}  |  單位：公制 (cm, kgf, kgf/cm²)"
    )
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    def add_dict_table(title_str, col1_name, col2_name, data_map):
        doc.add_heading(title_str, level=2)
        tbl = doc.add_table(rows=1, cols=2)
        tbl.style = "Table Grid"
        for idx, cell in enumerate(next(iter(tbl.rows)).cells):
            cell.text = col1_name if idx == 0 else col2_name
        for k, v in data_map.items():
            row_c = tbl.add_row().cells
            for idx, cell in enumerate(row_c):
                cell.text = str(k) if idx == 0 else str(v)

    add_dict_table("一、 連續壁工程與設計條件", "條件項目", "設定數值", inputs_dict)
    add_dict_table("二、 原設計號數反向埋深檢核結果", "檢核指標", "反算結果與判定", summary_dict)

    doc.add_heading("三、 替代方案建議明細 (符合最大允許深度限制)", level=2)
    if alt_list:
        for i, item in enumerate(alt_list, 1):
            doc.add_paragraph(f"【替代方案 {i}】{item['方案說明']}", style="List Bullet")
            doc.add_paragraph(
                f"   • 斷面積檢核：{item['斷面積比']}\n"
                f"   • 建議植入深度：{item['建議最小埋深']}（依據：{item['適用規範路徑']}）\n"
                f"   • 施工排列方式：{item['現場配置提醒']}"
            )
    else:
        doc.add_paragraph("原設計號數已可直接滿足最大深度限制，或現場條件極度嚴苛需改用機械續接器。")

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# ==========================================
# 5. Streamlit 互動介面
# ==========================================
st.title("🧱 連續壁錯位植筋：雙規範反向檢核與替代方案系統")
st.markdown(
    "針對連續壁壁厚與最大鑽孔深度限制，同步反算 **ACI 318 第 25 章（伸展長度法）** 與 **第 17 章（高強度認證藥劑群錨法）** 所需埋深，並自動生成等斷面積替代方案。"
)

with st.sidebar:
    st.header("🏗️ 1. 連續壁幾何與深度限制")
    wall_thick = st.number_input(
        "連續壁設計壁厚 (cm)", min_value=50.0, max_value=200.0, value=100.0, step=5.0
    )
    bad_conc_layer = st.number_input(
        "表面夾泥/劣質混凝土剔除厚度 (cm)", min_value=0.0, max_value=15.0, value=5.0, step=1.0
    )
    max_allow_depth = st.number_input(
        "最大允許有效植入深度上限 (cm)",
        min_value=15.0,
        max_value=150.0,
        value=50.0,
        step=5.0,
        help="扣除劣質層後，健全混凝土內允許的最大鑽孔深度",
    )
    remain_wall = wall_thick - bad_conc_layer - max_allow_depth
    if remain_wall < 15.0:
        st.error(f"⚠️ 警告：鑽孔後方剩餘阻水壁厚僅 {remain_wall:.1f} cm (< 15 cm)，有湧水風險！")
    else:
        st.success(f"✅ 鑽孔後方保有 **{remain_wall:.1f} cm** 健全阻水壁厚。")

    st.divider()
    st.header("🔩 2. 原設計錯位鋼筋與間距")
    rebar_options = tuple(REBAR_DB.keys())
    orig_rebar = st.selectbox("原設計預埋鋼筋號數", options=rebar_options, index=4)  # 預設 #8 (D25)
    orig_spacing = st.number_input(
        "原設計中心間距 s (cm)", min_value=8.0, max_value=40.0, value=15.0, step=0.5
    )
    splice_mode = st.selectbox(
        "植筋接頭受力型式",
        options=("直接延伸作為主筋 (1.0 ld)", "伸出後與既有主筋搭接 (1.3 ld)"),
        index=0,
    )

    st.divider()
    st.header("🧪 3. 材料與高強度藥劑參數")
    fc_options = (210, 245, 280, 350, 420)
    fc = st.selectbox("連續壁混凝土強度 f'c (kgf/cm²)", options=fc_options, index=3)  # 預設 350
    steel_options = tuple(STEEL_DB.keys())
    steel_type = st.selectbox("鋼筋材質", options=steel_options, index=0)
    tau_k = st.number_input(
        "AC308 認證藥劑特徵黏結應力 τ_k (kgf/cm²)",
        min_value=40.0,
        max_value=200.0,
        value=105.0,
        step=5.0,
    )

    st.divider()
    st.header("🎯 4. 第 17 章反算設計目標")
    stress_ratio = st.slider(
        "該斷面實際設計應力比 (Nu / Ab·fy)",
        min_value=0.50,
        max_value=1.00,
        value=1.00,
        step=0.05,
        help="1.00 代表要求達到 100% 鋼筋降伏強度；若非最大彎矩區可依結構計算調降至 0.70~0.85",
    )
    require_ductile = st.checkbox(
        "強制滿足延性破壞原則 (混凝土/黏結標稱力 ≥ 1.0 fy)", value=True
    )

# ==========================================
# 6. 執行主計算與畫面呈現
# ==========================================
db = REBAR_DB[orig_rebar]["db"]
Ab = REBAR_DB[orig_rebar]["Ab"]
fy = STEEL_DB[steel_type]["fy"]
fu = STEEL_DB[steel_type]["fu"]
is_lap = "1.3" in splice_mode

# 1. 第 25 章反算
req_d25, conf_val = calc_ch25_required_depth(db, fy, fc, orig_spacing, is_lap)
d25_ceil = math.ceil(req_d25)
ok_25 = d25_ceil <= max_allow_depth

# 2. 第 17 章反算
req_d17, phi_Nn_at_20db, target_Nu, limit_20db = solve_ch17_required_depth(
    db, Ab, fy, fu, fc, tau_k, orig_spacing, stress_ratio, require_ductile
)
ok_17 = (req_d17 is not None) and (math.ceil(req_d17) <= max_allow_depth)

st.subheader(f"📊 原設計 【{orig_rebar} @ {orig_spacing} cm】 反向埋深檢核結果")

col1, col2, col3 = st.columns(3)
with col1:
    st.metric(
        "最大允許埋深上限",
        f"{max_allow_depth:.0f} cm",
        f"壁厚 {wall_thick:.0f} cm (餘 {remain_wall:.0f} cm)",
    )

with col2:
    st.metric(
        "第 25 章：伸展長度法所需埋深",
        f"{d25_ceil} cm",
        delta="✅ 符合上限" if ok_25 else f"❌ 超出 {d25_ceil - max_allow_depth:.0f} cm",
        delta_color="normal" if ok_25 else "inverse",
    )

with col3:
    if req_d17 is not None:
        d17_ceil = math.ceil(req_d17)
        st.metric(
            "第 17 章：高強藥劑法所需埋深",
            f"{d17_ceil} cm",
            delta="✅ 符合上限" if ok_17 else f"❌ 超出 {d17_ceil - max_allow_depth:.0f} cm",
            delta_color="normal" if ok_17 else "inverse",
        )
    else:
        st.metric(
            "第 17 章：高強藥劑法所需埋深",
            "無法達成 (群組錐裂控制)",
            delta=f"20db 極限僅 {phi_Nn_at_20db:,.0f} kgf",
            delta_color="inverse",
        )
# 判定總結與原因解析
st.divider()
alt_schemes = []

if ok_25 or ok_17:
    st.success("🎉 **檢核結論：原設計鋼筋號數可直接施工！**")
    if ok_25 and ok_17:
        rec_final = min(d25_ceil, math.ceil(req_d17))
        st.write(
            f"- 同時符合第 25 章（{d25_ceil} cm）與第 17 章（{math.ceil(req_d17)} cm），建議現場採用較經濟且安全之深度 **{max(rec_final, 30)} cm**（由健全混凝土面起算）。"
        )
    elif ok_17:
        st.write(
            f"- 雖然第 25 章伸展長度需 **{d25_ceil} cm**（超過上限），但改採 **ACI 318 第 17 章 + 高強度 AC308 認證純環氧樹脂（τ_k = {tau_k} kgf/cm²）** 後，所需埋深可合法縮短至 **{math.ceil(req_d17)} cm**（≤ {max_allow_depth:.0f} cm），可直接依原號數施工！"
        )
    else:
        st.write(
            f"- 依第 25 章計算所需伸展長度為 **{d25_ceil} cm**（≤ {max_allow_depth:.0f} cm），可直接依原號數施工。"
        )
else:
    st.error(
        f"🚨 **檢核結論：原設計 {orig_rebar} @ {orig_spacing} cm 在 {max_allow_depth:.0f} cm 埋深限制下【無法達標】！**"
    )
    st.write(
        f"- **原因分析**：在間距 {orig_spacing} cm 緊密排列下，第 25 章伸展長度需 **{d25_ceil} cm**；而第 17 章即使採用高強度藥劑（τ_k={tau_k}），在 20db ({limit_20db:.1f} cm) 埋深下受限於「整排群組混凝土錐狀破裂重疊折減」，單支設計容量僅 **{phi_Nn_at_20db:,.0f} kgf**（低於目標需求 **{target_Nu:,.0f} kgf**）。"
    )

    # 自動生成替代方案
    st.subheader("💡 系統自動生成：等斷面積替代方案（符合 ≤ 50 cm 限制）")
    alt_schemes = find_alternative_schemes(
        orig_rebar,
        Ab,
        orig_spacing,
        max_allow_depth,
        fy,
        fu,
        fc,
        tau_k,
        is_lap,
        stress_ratio,
        require_ductile,
    )
    if alt_schemes:
        st.table(alt_schemes)
    else:
        st.warning(
            "在目前極小間距與嚴苛應力條件下，即使降號數加密仍受限於群組錐裂，建議：(1) 調降非塑鉸區實際應力比；(2) 增加允許鑽孔深度；或 (3) 局部敲除連續壁保護層露出既有主筋採機械續接。"
        )

# ==========================================
# 7. Word 計算書匯出
# ==========================================
st.divider()
st.subheader("📄 匯出連續壁替代方案計算書 (Word)")

inputs_summary = {
    "連續壁幾何條件": f"壁厚 {wall_thick:.0f} cm，剔除表層劣質混凝土 {bad_conc_layer:.0f} cm，最大允許埋深 {max_allow_depth:.0f} cm",
    "原設計錯位鋼筋": f"{orig_rebar} (db={db} cm, Ab={Ab} cm²)，中心間距 s = {orig_spacing} cm",
    "材料與藥劑參數": f"f'c = {fc} kgf/cm²，{steel_type}，AC308 藥劑 τ_k = {tau_k} kgf/cm²",
    "受力與延性條件": f"{splice_mode}，設計應力比 = {stress_ratio:.2f}，延性檢核 = {'開啟' if require_ductile else '關閉'}",
}

results_summary = {
    "第 25 章伸展長度法反算": f"所需埋深 {d25_ceil} cm ({'合格' if ok_25 else '超過上限'})",
    "第 17 章高強藥劑法反算": (
        f"所需埋深 {math.ceil(req_d17)} cm ({'合格' if ok_17 else '超過上限'})"
        if req_d17 is not None
        else f"群組錐裂控制，20db 深度下僅達 {phi_Nn_at_20db:,.0f} kgf (< 需求 {target_Nu:,.0f} kgf)"
    ),
    "最終工程判定": "原號數可直接施作" if (ok_25 or ok_17) else "原號數超限，須採用下方替代方案",
}

word_file = build_word_doc(inputs_summary, results_summary, alt_schemes)
rebar_clean = next(iter(orig_rebar.split()))

st.download_button(
    label="⬇️ 下載連續壁植筋反算與替代方案計算書 (.docx)",
    data=word_file,
    file_name=f"連續壁植筋反算_{rebar_clean}_fc{fc}_{datetime.now().strftime('%Y%m%d')}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
)        
