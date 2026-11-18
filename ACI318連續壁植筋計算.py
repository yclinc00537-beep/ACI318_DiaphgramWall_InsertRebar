import io
import math
from datetime import datetime
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
import matplotlib.patches as patches
import matplotlib.pyplot as plt
import streamlit as st

# ==========================================
# 1. 頁面設定與資料庫
# ==========================================
st.set_page_config(
    page_title="連續壁錯位植筋：批次總表與每公尺雙排配置系統",
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
    cb = max((spacing_s - db) / 2.0, 1.0 * db)
    ktr = 0.0
    confinement = min((cb + ktr) / db, 2.5)
    psi_t = 1.0
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
    return req_depth


def calc_ch17_group_capacity(hef, db, Ab, fy, fu, fc, tau_k, spacing_s, n_row=10):
    futa = min(fu, 1.9 * fy, 8800.0)
    phi_Nsa = 0.75 * (Ab * futa)

    Nb = 16.0 * 1.0 * math.sqrt(fc) * (hef**1.5)
    ANco = 9.0 * (hef**2)
    bx = (n_row - 1) * min(spacing_s, 3.0 * hef) + 3.0 * hef
    by = 3.0 * hef
    ANc = bx * by
    Ncbg_per_bar = ((ANc / ANco) * Nb) / n_row
    phi_Ncb_per_bar = 0.65 * Ncbg_per_bar

    Nba = 1.0 * tau_k * math.pi * db * hef
    cNa = 10.0 * db * math.sqrt(tau_k / 77.3)
    ANao = (2.0 * cNa) ** 2
    bxa = (n_row - 1) * min(spacing_s, 2.0 * cNa) + 2.0 * cNa
    bya = 2.0 * cNa
    ANa = bxa * bya
    Nag_per_bar = ((ANa / ANao) * Nba) / n_row
    phi_Na_per_bar = 0.65 * Nag_per_bar

    phi_Nn = min(phi_Nsa, phi_Ncb_per_bar, phi_Na_per_bar)
    nominal_non_steel = min(Ncbg_per_bar, Nag_per_bar)
    return phi_Nn, nominal_non_steel


def solve_ch17_required_depth(
    db, Ab, fy, fu, fc, tau_k, spacing_s, stress_ratio, require_ductile
):
    target_Nu = (Ab * fy) * stress_ratio
    max_hef_limit = 20.0 * db
    min_hef_limit = 4.0 * db

    phi_Nn_max, nom_ns_max = calc_ch17_group_capacity(
        max_hef_limit, db, Ab, fy, fu, fc, tau_k, spacing_s
    )
    ductile_ok_max = (nom_ns_max >= Ab * fy) if require_ductile else True

    if phi_Nn_max < target_Nu or not ductile_ok_max:
        return None, phi_Nn_max, target_Nu

    low = min_hef_limit
    high = max_hef_limit
    best_hef = max_hef_limit

    for _ in range(40):
        mid = (low + high) / 2.0
        phi_Nn, nom_ns = calc_ch17_group_capacity(
            mid, db, Ab, fy, fu, fc, tau_k, spacing_s
        )
        cond_strength = phi_Nn >= target_Nu
        cond_ductile = (nom_ns >= Ab * fy) if require_ductile else True

        if cond_strength and cond_ductile:
            best_hef = mid
            high = mid
        else:
            low = mid

    return best_hef, phi_Nn_max, target_Nu


# ==========================================
# 3. 每公尺寬度配筋換算 & 雙排錯開配置圖繪製
# ==========================================
def calc_per_meter_schemes(
    orig_rebar_key, orig_Ab, orig_spacing, max_allow_depth, fy, fu, fc, tau_k, is_lap, stress_ratio, require_ductile
):
    """計算每 100 cm 寬度的原設計總斷面積，並產生雙排錯開小號數替代方案"""
    bars_per_m_orig = 100.0 / orig_spacing
    As_per_m_orig = bars_per_m_orig * orig_Ab

    schemes = []
    candidate_keys = ("#6 (D19)", "#5 (D16)", "#4 (D13)")
    row_vertical_gap = 10.0  # 上下雙排垂直排距 10 cm

    for cand_key in candidate_keys:
        cand_db = REBAR_DB[cand_key]["db"]
        cand_Ab = REBAR_DB[cand_key]["Ab"]
        if cand_Ab >= orig_Ab:
            continue

        # 計算每公尺至少需要幾支小號數鋼筋 (無條件進位至偶數，方便上下雙排對稱排列)
        min_bars_needed = math.ceil(As_per_m_orig / cand_Ab)
        if min_bars_needed % 2 != 0:
            min_bars_needed += 1

        # 測試每公尺總支數 (min_bars_needed 或 +2)
        for total_bars_m in (min_bars_needed, min_bars_needed + 2):
            bars_each_row = total_bars_m / 2.0
            s_each_row = 100.0 / bars_each_row  # 單排水平中心距 (cm)
            # 梅花錯開後的斜向最小中心距 (cm)
            diag_spacing = math.hypot(s_each_row / 2.0, row_vertical_gap)

            if diag_spacing < 9.5 or s_each_row < 10.0:
                continue

            d25 = calc_ch25_required_depth(cand_db, fy, fc, diag_spacing, is_lap)
            d17, _, _ = solve_ch17_required_depth(
                cand_db, cand_Ab, fy, fu, fc, tau_k, diag_spacing, stress_ratio, require_ductile
            )

            ok25 = math.ceil(d25) <= max_allow_depth
            ok17 = (d17 is not None) and (math.ceil(d17) <= max_allow_depth)

            if ok25 or ok17:
                rec_depth = math.ceil(min(d25 if ok25 else 999, d17 if ok17 else 999))
                As_new_m = total_bars_m * cand_Ab
                schemes.append(
                    {
                        "方案名稱": f"雙排交錯 {cand_key} @ {s_each_row:.1f} cm (上下各一排)",
                        "cand_key": cand_key,
                        "cand_db": cand_db,
                        "total_bars_m": int(total_bars_m),
                        "s_each_row": round(s_each_row, 1),
                        "row_gap": row_vertical_gap,
                        "diag_spacing": round(diag_spacing, 1),
                        "As_orig_m": round(As_per_m_orig, 2),
                        "As_new_m": round(As_new_m, 2),
                        "ratio_pct": round(As_new_m / As_per_m_orig * 100, 1),
                        "rec_depth": max(rec_depth, 30),
                        "method": "第25章伸展法" if ok25 else "第17章高強藥劑法",
                    }
                )
                break

    return schemes, bars_per_m_orig, As_per_m_orig


def plot_staggered_layout(orig_rebar, orig_spacing, scheme):
    """繪製連續壁每公尺寬度 (100 cm) 雙排梅花錯開鑽孔配置示意圖"""
    fig, ax = plt.subplots(figsize=(8.5, 3.6), dpi=150)

    # 畫連續壁 100 cm 寬度背景區塊
    wall_rect = patches.Rectangle(
        (0, 0), 100, 36, linewidth=1.5, edgecolor="#475569", facecolor="#F1F5F9"
    )
    ax.add_patch(wall_rect)

    # 標示原設計錯位鋼筋位置 (灰色虛線圈，作為對照)
    n_orig = int(100 // orig_spacing) + 1
    for i in range(n_orig):
        xo = i * orig_spacing + (orig_spacing / 2.0)
        if xo <= 98:
            c_orig = patches.Circle(
                (xo, 18), radius=1.3, fill=False, linestyle="--", edgecolor="#94A3B8", linewidth=1.2
            )
            ax.add_patch(c_orig)

    # 畫替代方案：上排 (y = 23 cm) 與下排 (y = 13 cm) 梅花交錯
    s_row = scheme["s_each_row"]
    y_top = 23.0
    y_bot = 13.0
    r_bar = max(scheme["cand_db"] * 0.7, 1.1)

    top_xs = []
    bot_xs = []

    x_curr = s_row / 4.0
    while x_curr <= 98.0:
        top_xs.append(x_curr)
        ax.add_patch(patches.Circle((x_curr, y_top), radius=r_bar, color="#DC2626", zorder=4))
        x_curr += s_row

    x_curr_b = (s_row / 4.0) + (s_row / 2.0)
    while x_curr_b <= 98.0:
        bot_xs.append(x_curr_b)
        ax.add_patch(patches.Circle((x_curr_b, y_bot), radius=r_bar, color="#2563EB", zorder=4))
        x_curr_b += s_row

    # 標註尺寸線：單排水平間距 s_row
    if len(top_xs) >= 2:
        x1, x2 = top_xs[0], top_xs[1]
        ax.annotate(
            "",
            xy=(x1, 29),
            xytext=(x2, 29),
            arrowprops=dict(arrowstyle="<->", color="#DC2626", lw=1.3),
        )
        ax.text(
            (x1 + x2) / 2,
            30.5,
            f"Upper Row s = {s_row} cm",
            color="#DC2626",
            ha="center",
            fontsize=8.5,
            fontweight="bold",
        )

    # 標註斜向淨中心距與排距
    if len(top_xs) >= 1 and len(bot_xs) >= 1:
        ax.plot([top_xs[0], bot_xs[0]], [y_top, y_bot], color="#059669", linestyle="-.", lw=1.5, zorder=3)
        ax.text(
            (top_xs[0] + bot_xs[0]) / 2 + 1.5,
            (y_top + y_bot) / 2,
            f"Diag = {scheme['diag_spacing']} cm\n(dy = {scheme['row_gap']:.0f} cm)",
            color="#059669",
            fontsize=8,
            fontweight="bold",
            va="center",
        )

    # 圖例說明 (使用英文避免雲端 Linux 缺少中文字型產生方框亂碼)
    ax.scatter([], [], color="#DC2626", label=f"Upper Rebar: {scheme['cand_key']} @ {s_row} cm")
    ax.scatter([], [], color="#2563EB", label=f"Lower Rebar: {scheme['cand_key']} @ {s_row} cm (Staggered)")
    ax.scatter([], [], facecolors="none", edgecolors="#94A3B8", linestyle="--", label=f"Original Misaligned: {orig_rebar} @ {orig_spacing} cm")

    ax.set_xlim(-2, 102)
    ax.set_ylim(0, 38)
    ax.set_xlabel("Diaphragm Wall Width = 100 cm (1.0 m Unit Strip)", fontsize=9.5, fontweight="bold")
    ax.set_ylabel("Elevation (cm)", fontsize=9)
    ax.set_title(
        f"Staggered Drilling Layout (hef = {scheme['rec_depth']} cm, As = {scheme['As_new_m']} cm2/m)",
        fontsize=10.5,
        fontweight="bold",
    )
    ax.legend(loc="lower right", fontsize=7.8, framealpha=0.9)
    ax.grid(True, linestyle=":", alpha=0.5)
    plt.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=180)
    plt.close(fig)
    buf.seek(0)
    return buf


# ==========================================
# 4. 批次計算 #4 ～ #10 全號數總表
# ==========================================
def run_batch_all_rebars(
    spacing_s, max_allow_depth, fy, fu, fc, tau_k, is_lap, stress_ratio, require_ductile
):
    batch_rows = []
    for r_key, props in REBAR_DB.items():
        r_db = props["db"]
        r_Ab = props["Ab"]

        d25 = math.ceil(calc_ch25_required_depth(r_db, fy, fc, spacing_s, is_lap))
        d17_raw, _, _ = solve_ch17_required_depth(
            r_db, r_Ab, fy, fu, fc, tau_k, spacing_s, stress_ratio, require_ductile
        )
        d17_str = f"{math.ceil(d17_raw)} cm" if d17_raw is not None else "超限 (錐裂控制)"
        d17_ok = (d17_raw is not None) and (math.ceil(d17_raw) <= max_allow_depth)
        d25_ok = d25 <= max_allow_depth

        if d25_ok or d17_ok:
            best_d = min(d25 if d25_ok else 999, math.ceil(d17_raw) if d17_ok else 999)
            status = f"✅ 合格 (建議埋深 {max(best_d, 30)} cm)"
            alt_text = "原號數可直接施作"
        else:
            status = f"❌ 超過 {max_allow_depth:.0f} cm 上限"
            sub_schemes, _, _ = calc_per_meter_schemes(
                r_key,
                r_Ab,
                spacing_s,
                max_allow_depth,
                fy,
                fu,
                fc,
                tau_k,
                is_lap,
                stress_ratio,
                require_ductile,
            )
            if sub_schemes:
                top_s = next(iter(sub_schemes))
                alt_text = f"{top_s['方案名稱']} (埋深 {top_s['rec_depth']} cm)"
            else:
                alt_text = "需加大間距或改用機械續接"

        batch_rows.append(
            {
                "鋼筋號數": r_key,
                "斷面積 Ab (cm²)": f"{r_Ab:.3f}",
                "第25章所需埋深": f"{d25} cm",
                "第17章高強藥劑埋深": d17_str,
                "檢核判定": status,
                "每公尺雙排替代方案建議": alt_text,
            }
        )
    return batch_rows


# ==========================================
# 5. Word 計算書產生函數 (含全號數總表與配置圖)
# ==========================================
def build_comprehensive_word_doc(
    inputs_dict, batch_rows, target_rebar, orig_spacing, per_m_schemes, fig_buf
):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Microsoft JhengHei"
    style.font.size = Pt(10)

    title = doc.add_heading("連續壁預埋筋錯位—植筋反向檢核、批次總表與替代配置計算書", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_meta = doc.add_paragraph(
        f"計算日期：{datetime.now().strftime('%Y-%m-%d %H:%M')}  |  單位系統：公制 (cm, kgf, kgf/cm²)"
    )
    p_meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # 一、基本參數表
    doc.add_heading("一、 連續壁工程條件與設計參數", level=2)
    t1 = doc.add_table(rows=1, cols=2)
    t1.style = "Table Grid"
    for idx, cell in enumerate(next(iter(t1.rows)).cells):
        cell.text = "參數項目" if idx == 0 else "設定數值"
    for k, v in inputs_dict.items():
        rc = t1.add_row().cells
        for idx, cell in enumerate(rc):
            cell.text = str(k) if idx == 0 else str(v)

    # 二、#4 ~ #10 批次檢核總表
    doc.add_heading("二、 #4 ～ #10 全號數反向埋深檢核總表", level=2)
    headers = ("鋼筋號數", "斷面積 Ab", "第25章埋深", "第17章埋深", "檢核判定", "每公尺雙排替代方案")
    t2 = doc.add_table(rows=1, cols=len(headers))
    t2.style = "Table Grid"
    for idx, cell in enumerate(next(iter(t2.rows)).cells):
        cell.text = headers[idx]

    for row_data in batch_rows:
        rc = t2.add_row().cells
        vals = (
            row_data["鋼筋號數"],
            row_data["斷面積 Ab (cm²)"],
            row_data["第25章所需埋深"],
            row_data["第17章高強藥劑埋深"],
            row_data["檢核判定"],
            row_data["每公尺雙排替代方案建議"],
        )
        for idx, cell in enumerate(rc):
            cell.text = str(vals[idx])

    # 三、指定號數每公尺寬度換算與配置圖
    doc.add_heading(f"三、 【{target_rebar} @ {orig_spacing} cm】每公尺寬度替代配筋與鑽孔配置圖", level=2)
    if per_m_schemes:
        for i, sc in enumerate(per_m_schemes, 1):
            doc.add_paragraph(
                f"【方案 {i}】{sc['方案名稱']}（每公尺總支數 {sc['total_bars_m']} 支，"
                f"提供斷面積 {sc['As_new_m']} cm²/m ≥ 原設計 {sc['As_orig_m']} cm²/m [{sc['ratio_pct']}%]，"
                f"建議植入深度 {sc['rec_depth']} cm，斜向中心距 {sc['diag_spacing']} cm）",
                style="List Bullet",
            )
        if fig_buf is not None:
            fig_buf.seek(0)
            doc.add_picture(fig_buf, width=Inches(6.2))
            cap = doc.add_paragraph("圖 1：連續壁每公尺寬度（100 cm）上下雙排梅花錯開鑽孔配置示意圖")
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        doc.add_paragraph(
            f"目前設定之 {target_rebar} @ {orig_spacing} cm 可直接依原號數與原間距施作，無須拆分為雙排小號數替代。"
        )

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio
# ==========================================
# 6. Streamlit 側邊欄與主畫面互動介面
# ==========================================
st.title("🧱 連續壁錯位植筋：批次總表與每公尺雙排配置系統")

with st.sidebar:
    st.header("🏗️ 1. 連續壁幾何與深度上限")
    wall_thick = st.number_input("連續壁設計壁厚 (cm)", 50.0, 200.0, 100.0, 5.0)
    bad_layer = st.number_input("表面夾泥/劣質層剔除厚度 (cm)", 0.0, 15.0, 5.0, 1.0)
    max_allow_depth = st.number_input("最大允許有效植入深度上限 (cm)", 15.0, 150.0, 50.0, 5.0)
    remain_wall = wall_thick - bad_layer - max_allow_depth
    if remain_wall < 15.0:
        st.error(f"⚠️ 剩餘阻水壁厚僅 {remain_wall:.1f} cm (< 15 cm)！")
    else:
        st.success(f"✅ 鑽孔後方保有 **{remain_wall:.1f} cm** 健全阻水壁厚。")

    st.divider()
    st.header("🔩 2. 原設計配筋條件")
    rebar_options = tuple(REBAR_DB.keys())
    orig_rebar = st.selectbox("重點檢核鋼筋號數", options=rebar_options, index=4)  # 預設 #8
    orig_spacing = st.number_input("原設計中心間距 s (cm)", 8.0, 40.0, 15.0, 0.5)
    splice_mode = st.selectbox(
        "植筋接頭受力型式",
        options=("直接延伸作為主筋 (1.0 ld)", "伸出後與既有主筋搭接 (1.3 ld)"),
        index=0,
    )

    st.divider()
    st.header("🧪 3. 材料與認證藥劑參數")
    fc_options = (210, 245, 280, 350, 420)
    fc = st.selectbox("連續壁混凝土強度 f'c (kgf/cm²)", options=fc_options, index=3)  # 預設 350
    steel_type = st.selectbox("鋼筋材質", options=tuple(STEEL_DB.keys()), index=0)
    tau_k = st.number_input("AC308 認證藥劑特徵黏結應力 τ_k (kgf/cm²)", 40.0, 200.0, 105.0, 5.0)

    st.divider()
    st.header("🎯 4. 第 17 章反算目標")
    stress_ratio = st.slider("實際設計應力比 (Nu / Ab·fy)", 0.50, 1.00, 1.00, 0.05)
    require_ductile = st.checkbox("強制滿足延性破壞原則 (非鋼材標稱力 ≥ fy)", value=True)

db = REBAR_DB[orig_rebar]["db"]
Ab = REBAR_DB[orig_rebar]["Ab"]
fy = STEEL_DB[steel_type]["fy"]
fu = STEEL_DB[steel_type]["fu"]
is_lap = "1.3" in splice_mode

# 執行批次運算與每公尺替代運算
batch_table = run_batch_all_rebars(
    orig_spacing, max_allow_depth, fy, fu, fc, tau_k, is_lap, stress_ratio, require_ductile
)
per_m_schemes, bars_orig_m, As_orig_m = calc_per_meter_schemes(
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

tab1, tab2 = st.tabs(
    ["📋 一、 #4 ～ #10 全號數批次計算總表", "📐 二、 每公尺寬度配筋換算與雙排錯開配置圖"]
)

# --- 分頁 1：全號數批次總表 ---
with tab1:
    st.subheader(
        f"📊 #4 ～ #10 全號數反向埋深檢核總表（間距 s = {orig_spacing} cm，上限 {max_allow_depth:.0f} cm）"
    )
    st.dataframe(batch_table, width="stretch", hide_index=True)
    st.caption(
        "註：表列埋深皆從「敲除表層劣質混凝土後的健全混凝土面」起算；當原號數超過上限時，右側欄位自動帶出最佳每公尺雙排替代配置。"
    )

# --- 分頁 2：每公尺配筋換算與鑽孔配置圖 ---
with tab2:
    st.subheader(f"📐 【{orig_rebar} @ {orig_spacing} cm】每公尺寬度配筋換算與雙排配置圖")

    c1, c2, c3 = st.columns(3)
    c1.metric("原設計每公尺支數", f"{bars_orig_m:.2f} 支/m", f"間距 @ {orig_spacing} cm")
    c2.metric("原設計每公尺總斷面積 As", f"{As_orig_m:.2f} cm²/m", f"單支 Ab = {Ab} cm²")
    c3.metric("最大允許鑽孔深度", f"{max_allow_depth:.0f} cm", f"壁厚 {wall_thick:.0f} cm")

    fig_buffer = None
    if per_m_schemes:
        scheme_names = tuple(s["方案名稱"] for s in per_m_schemes)
        selected_name = st.selectbox("選擇要檢視與匯出圖面的替代配置方案：", options=scheme_names)
        chosen_scheme = next(s for s in per_m_schemes if s["方案名稱"] == selected_name)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("替代方案每公尺總支數", f"{chosen_scheme['total_bars_m']} 支/m", "上下雙排各半")
        m2.metric(
            "單排水平中心距",
            f"{chosen_scheme['s_each_row']} cm",
            f"垂直排距 {chosen_scheme['row_gap']:.0f} cm",
        )
        m3.metric("梅花錯開斜向中心距", f"{chosen_scheme['diag_spacing']} cm", "確保 ≥ 10 cm 圍束")
        m4.metric(
            "提供總斷面積 As,new",
            f"{chosen_scheme['As_new_m']} cm²/m",
            f"達原設計 {chosen_scheme['ratio_pct']}%",
        )

        st.markdown("#### 🖼️ 連續壁每公尺寬度（100 cm）雙排梅花錯開鑽孔配置圖")
        fig_buffer = plot_staggered_layout(orig_rebar, orig_spacing, chosen_scheme)
        st.image(fig_buffer, width="content")
    else:
        st.info(
            f"💡 目前條件下，原設計 **{orig_rebar} @ {orig_spacing} cm** 可直接施作（或無須拆分為更小號數雙排），若仍想強制查看小號數雙排配置，可嘗試將左側「重點檢核鋼筋號數」調大為 **#8、#9 或 #10**。"
        )

# ==========================================
# 7. Word 計算書一鍵匯出 (含總表 + 配置圖)
# ==========================================
st.divider()
st.subheader("📄 匯出完整結構計算書 (含 #4～#10 總表與鑽孔配置圖)")

inputs_doc = {
    "連續壁幾何與深度限制": f"壁厚 {wall_thick:.0f} cm，剔除表層劣質混凝土 {bad_layer:.0f} cm，最大允許埋深 {max_allow_depth:.0f} cm",
    "重點檢核配筋": f"{orig_rebar} @ {orig_spacing} cm (每公尺 {bars_orig_m:.2f} 支，As = {As_orig_m:.2f} cm²/m)",
    "材料與認證藥劑": f"f'c = {fc} kgf/cm²，{steel_type}，AC308 藥劑 τ_k = {tau_k} kgf/cm²",
    "受力與延性條件": f"{splice_mode}，設計應力比 = {stress_ratio:.2f}，延性檢核 = {'開啟' if require_ductile else '關閉'}",
}

word_bytes = build_comprehensive_word_doc(
    inputs_doc, batch_table, orig_rebar, orig_spacing, per_m_schemes, fig_buffer
)
rebar_clean = next(iter(orig_rebar.split()))

st.download_button(
    label="⬇️ 下載完整 Word 計算書 (.docx) [含全號數總表 & 配置圖]",
    data=word_bytes,
    file_name=f"連續壁植筋批次總表與雙排配置_{rebar_clean}_fc{fc}_{datetime.now().strftime('%Y%m%d')}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
)
