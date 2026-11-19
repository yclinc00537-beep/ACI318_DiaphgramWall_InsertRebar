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
# 1. 頁面設定與工程資料庫
# ==========================================
st.set_page_config(
    page_title="ACI 318 連續壁植筋：雙軌物理修正、批次總表與雙排配置系統",
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

ADHESIVE_CASE_DB = {
    "Case 1：一般二次補強 / 非結構構件 (CNS 10142)": {
        "min_tau": 61.2,
        "desc": "需符合 CNS 10142 接著強度 ≥ 61.2 kgf/cm²，樹脂抗壓 ≥ 510 kgf/cm²。",
    },
    "Case 2：一般結構錨定 - 未開裂混凝土 / 靜載重": {
        "min_tau": 70.0,
        "desc": "ACI 355.4 未開裂最低門檻 70 kgf/cm²；持續拉力不得超過 0.55 φNba。",
    },
    "Case 3：耐震結構錨定 - 開裂混凝土 / 梁柱基板": {
        "min_tau": 55.0,
        "desc": "ACI 355.4 開裂混凝土最低下限 35 kgf/cm²，結構耐震實務建議 ≥ 55 kgf/cm²。",
    },
    "Case 4：連續壁主筋替代 (伸展長度法 / 高強度 / 潮濕孔)": {
        "min_tau": 85.0,
        "desc": "須通過 ICC-ES AC308 Table 3.8 (60d深埋) 及潮濕孔認證，建議 τ_k ≥ 85 kgf/cm²。",
    },
    "Case 5：仰向鑽孔 (天花板) 或長期高溫環境 (>43°C)": {
        "min_tau": 90.0,
        "desc": "需具仰向注膠與高溫認證，折減前基準黏結應力建議 ≥ 90 kgf/cm²。",
    },
}


# ==========================================
# 2. 核心雙軌物理修正計算引擎
# ==========================================
def calc_ch25_with_bond_physics(db, fy, fc, spacing_s, edge_c, tau_k, is_lap):
    """計算第 25 章伸展長度，並納入低黏結放大與混凝土劈裂上限 (Cap=1.0) 物理修正"""
    cb = min(edge_c, max((spacing_s - db) / 2.0, 1.0 * db))
    confinement = min(cb / db, 2.5)
    psi_s = 0.8 if db <= 1.91 else 1.0
    psi_g = 1.0 if fy <= 4200 else 1.15

    # 預埋竹節鋼筋等效基準握裹應力 f_bd,ref (kgf/cm2)
    fbd_ref = 4.5 * math.sqrt(fc)
    # 低黏結藥劑按比例放大；高黏結藥劑受混凝土保護層抗劈裂強度上限控制，鎖定為 1.0
    alpha_bond = max(1.0, fbd_ref / tau_k)

    ld_base = (0.28 * fy / math.sqrt(fc)) * ((1.0 * 1.0 * psi_s * psi_g) / confinement) * db
    ld_mod = max(ld_base * alpha_bond, 30.0)
    req_depth = math.ceil(ld_mod * 1.3 if is_lap else ld_mod)
    return req_depth, fbd_ref, alpha_bond


def calc_ch17_row_capacity(hef, db, Ab, fy, fu, fc, tau_k, spacing_s, edge_c, n_row=10):
    """計算連續壁整排植筋在指定埋深 hef 下的單支平均設計抗拉強度與控制破壞模式"""
    futa = min(fu, 1.9 * fy, 8800.0)
    phi_Nsa = 0.75 * (Ab * futa)

    # 1. 群組混凝土錐狀破裂 Ncbg
    Nb = 16.0 * math.sqrt(fc) * (hef**1.5)
    ANco = 9.0 * (hef**2)
    bx = (n_row - 1) * min(spacing_s, 3.0 * hef) + 3.0 * hef
    by = min(edge_c, 1.5 * hef) + 1.5 * hef
    ANc = bx * by
    psi_ed_N = 1.0 if edge_c >= 1.5 * hef else (0.7 + 0.3 * (edge_c / (1.5 * hef)))
    Ncbg_per_bar = ((ANc / ANco) * psi_ed_N * Nb) / n_row
    phi_Ncb_per_bar = 0.65 * Ncbg_per_bar

    # 2. 群組化學黏結強度 Nag
    Nba = tau_k * math.pi * db * hef
    cNa = 10.0 * db * math.sqrt(tau_k / 77.3)
    ANao = (2.0 * cNa) ** 2
    bxa = (n_row - 1) * min(spacing_s, 2.0 * cNa) + 2.0 * cNa
    bya = min(edge_c, cNa) + cNa
    ANa = bxa * bya
    psi_ed_Na = 1.0 if edge_c >= cNa else (0.7 + 0.3 * (edge_c / cNa))
    Nag_per_bar = ((ANa / ANao) * psi_ed_Na * Nba) / n_row
    phi_Na_per_bar = 0.65 * Nag_per_bar

    phi_Nn = min(phi_Nsa, phi_Ncb_per_bar, phi_Na_per_bar)
    nom_non_steel = min(Ncbg_per_bar, Nag_per_bar)
    gov_mode = "黏結拉出 (Na)" if Nag_per_bar < Ncbg_per_bar else "混凝土錐裂 (Ncb)"
    return phi_Nn, nom_non_steel, cNa, gov_mode


def solve_ch17_parallel_depth(
    db, Ab, fy, fu, fc, tau_k, spacing_s, edge_c, stress_ratio, require_ductile
):
    """不設 20db 強制分流，反向二分搜尋第 17 章所需最小埋深並標示控制破壞模式"""
    target_Nu = (Ab * fy) * stress_ratio
    low = 4.0 * db
    high = 60.0 * db

    phi_max, nom_max, cNa_val, gov_mode = calc_ch17_row_capacity(
        high, db, Ab, fy, fu, fc, tau_k, spacing_s, edge_c
    )
    if phi_max < target_Nu or (require_ductile and nom_max < Ab * fy):
        return None, phi_max, target_Nu, cNa_val, "群組錐裂超限"

    best_h = high
    best_mode = gov_mode
    for _ in range(45):
        mid = (low + high) / 2.0
        p_nn, n_ns, _, g_mode = calc_ch17_row_capacity(
            mid, db, Ab, fy, fu, fc, tau_k, spacing_s, edge_c
        )
        if p_nn >= target_Nu and (not require_ductile or n_ns >= Ab * fy):
            best_h = mid
            best_mode = g_mode
            high = mid
        else:
            low = mid

    return math.ceil(best_h), phi_max, target_Nu, cNa_val, best_mode


# ==========================================
# 3. 每公尺寬度配筋換算 & 雙排梅花錯開配置圖
# ==========================================
def calc_per_meter_schemes(
    orig_rebar_key,
    orig_Ab,
    orig_spacing,
    edge_c,
    row_vertical_gap,
    max_allow_depth,
    fy,
    fu,
    fc,
    tau_k,
    is_lap,
    stress_ratio,
    require_ductile,
):
    """計算每 100 cm 寬度總斷面積，並搜尋符合雙軌物理修正且 <= 最大允許埋深的雙排替代方案"""
    bars_per_m_orig = 100.0 / orig_spacing
    As_per_m_orig = bars_per_m_orig * orig_Ab

    schemes = []
    candidate_keys = ("#6 (D19)", "#5 (D16)", "#4 (D13)")

    for cand_key in candidate_keys:
        cand_db = REBAR_DB[cand_key]["db"]
        cand_Ab = REBAR_DB[cand_key]["Ab"]
        if cand_Ab >= orig_Ab:
            continue

        min_struct_s = max(6.0 * cand_db, 10.0)
        min_bars_needed = math.ceil(As_per_m_orig / cand_Ab)
        if min_bars_needed % 2 != 0:
            min_bars_needed += 1

        for total_bars_m in (min_bars_needed, min_bars_needed + 2):
            bars_each_row = total_bars_m / 2.0
            s_each_row = 100.0 / bars_each_row
            diag_spacing = math.hypot(s_each_row / 2.0, row_vertical_gap)

            if diag_spacing < min_struct_s or s_each_row < min_struct_s:
                continue

            d25, _, alpha_b = calc_ch25_with_bond_physics(
                cand_db, fy, fc, diag_spacing, edge_c, tau_k, is_lap
            )
            d17, _, _, _, gov_17 = solve_ch17_parallel_depth(
                cand_db, cand_Ab, fy, fu, fc, tau_k, diag_spacing, edge_c, stress_ratio, require_ductile
            )

            ok25 = d25 <= max_allow_depth
            ok17 = (d17 is not None) and (d17 <= max_allow_depth)

            if ok25 or ok17:
                rec_depth = min(d25 if ok25 else 999, d17 if ok17 else 999)
                As_new_m = total_bars_m * cand_Ab
                if ok17 and (not ok25 or d17 <= d25):
                    method_note = f"第17章錨定法 ({gov_17})"
                else:
                    method_note = f"第25章伸展法 (α_bond={alpha_b:.2f})"

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
                        "d25_val": d25,
                        "d17_val": d17,
                        "method": method_note,
                    }
                )
                break

    return schemes, bars_per_m_orig, As_per_m_orig


def plot_staggered_layout(orig_rebar, orig_spacing, scheme):
    """繪製連續壁每公尺寬度 (100 cm) 雙排梅花錯開鑽孔配置示意圖"""
    fig, ax = plt.subplots(figsize=(8.8, 3.8), dpi=150)

    wall_rect = patches.Rectangle(
        (0, 0), 100, 38, linewidth=1.5, edgecolor="#475569", facecolor="#F8FAFC"
    )
    ax.add_patch(wall_rect)

    # 原設計錯位主筋位置 (灰色虛線圈)
    n_orig = int(100 // orig_spacing) + 1
    for i in range(n_orig):
        xo = i * orig_spacing + (orig_spacing / 2.0)
        if xo <= 98:
            ax.add_patch(
                patches.Circle(
                    (xo, 19), radius=1.35, fill=False, linestyle="--", edgecolor="#94A3B8", linewidth=1.2
                )
            )

    s_row = scheme["s_each_row"]
    gap_y = scheme["row_gap"]
    y_top = 19.0 + gap_y / 2.0
    y_bot = 19.0 - gap_y / 2.0
    r_bar = max(scheme["cand_db"] * 0.72, 1.15)

    top_xs = []
    bot_xs = []

    x_curr = s_row / 4.0
    while x_curr <= 98.5:
        top_xs.append(x_curr)
        ax.add_patch(patches.Circle((x_curr, y_top), radius=r_bar, color="#DC2626", zorder=4))
        x_curr += s_row

    x_curr_b = (s_row / 4.0) + (s_row / 2.0)
    while x_curr_b <= 98.5:
        bot_xs.append(x_curr_b)
        ax.add_patch(patches.Circle((x_curr_b, y_bot), radius=r_bar, color="#2563EB", zorder=4))
        x_curr_b += s_row

    # 標註單排水平中心距
    if len(top_xs) >= 2:
        x1 = next(iter(top_xs))
        x2 = x1 + s_row
        y_dim = min(y_top + 6.0, 34.5)
        ax.annotate(
            "",
            xy=(x1, y_dim),
            xytext=(x2, y_dim),
            arrowprops=dict(arrowstyle="<->", color="#DC2626", lw=1.4),
        )
        ax.text(
            (x1 + x2) / 2.0,
            y_dim + 1.2,
            f"Upper Row sx = {s_row} cm",
            color="#DC2626",
            ha="center",
            fontsize=8.5,
            fontweight="bold",
        )

    # 標註斜向中心距與垂直排距
    if len(top_xs) >= 1 and len(bot_xs) >= 1:
        xt0 = next(iter(top_xs))
        xb0 = next(iter(bot_xs))
        ax.plot([xt0, xb0], [y_top, y_bot], color="#059669", linestyle="-.", lw=1.6, zorder=3)
        ax.text(
            (xt0 + xb0) / 2.0 + 1.5,
            19.0,
            f"Diag = {scheme['diag_spacing']} cm\n(sy = {gap_y:.0f} cm)",
            color="#059669",
            fontsize=8.2,
            fontweight="bold",
            va="center",
        )

    ax.scatter([], [], color="#DC2626", label=f"Upper Row: {scheme['cand_key']} @ {s_row} cm")
    ax.scatter([], [], color="#2563EB", label=f"Lower Row: {scheme['cand_key']} @ {s_row} cm (Staggered)")
    ax.scatter(
        [],
        [],
        facecolors="none",
        edgecolors="#94A3B8",
        linestyle="--",
        label=f"Original Misaligned: {orig_rebar} @ {orig_spacing} cm",
    )

    ax.set_xlim(-2, 102)
    ax.set_ylim(0, 39)
    ax.set_xlabel("Diaphragm Wall Unit Width = 100 cm (1.0 m)", fontsize=9.5, fontweight="bold")
    ax.set_ylabel("Elevation (cm)", fontsize=9)
    ax.set_title(
        f"Staggered Drilling Layout | hef = {scheme['rec_depth']} cm | As = {scheme['As_new_m']} cm2/m ({scheme['ratio_pct']}%)",
        fontsize=10.2,
        fontweight="bold",
    )
    ax.legend(loc="lower right", fontsize=7.8, framealpha=0.92)
    ax.grid(True, linestyle=":", alpha=0.5)
    plt.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=180)
    plt.close(fig)
    buf.seek(0)
    return buf


# ==========================================
# 4. #4 ～ #10 全號數批次計算與空間幾何檢核
# ==========================================
def run_batch_all_rebars(
    spacing_s,
    edge_c,
    row_gap,
    max_allow_depth,
    fy,
    fu,
    fc,
    tau_k,
    is_lap,
    stress_ratio,
    require_ductile,
):
    batch_rows = []
    for r_key, props in REBAR_DB.items():
        r_db = props["db"]
        r_Ab = props["Ab"]

        d25, _, alpha_b = calc_ch25_with_bond_physics(
            r_db, fy, fc, spacing_s, edge_c, tau_k, is_lap
        )
        d17, _, _, _, gov_17 = solve_ch17_parallel_depth(
            r_db, r_Ab, fy, fu, fc, tau_k, spacing_s, edge_c, stress_ratio, require_ductile
        )

        d17_str = f"{d17} cm ({gov_17})" if d17 is not None else "無法收斂 (群組錐裂)"
        ok25 = d25 <= max_allow_depth
        ok17 = (d17 is not None) and (d17 <= max_allow_depth)

        if ok25 or ok17:
            best_d = min(d25 if ok25 else 999, d17 if ok17 else 999)
            ctrl_src = "第17章" if (ok17 and (not ok25 or d17 <= d25)) else "第25章"
            status = f"✅ 合格 (建議 {max(best_d, 30)} cm / {ctrl_src})"
            alt_text = "原號數可直接施作"
        else:
            status = f"❌ 超過 {max_allow_depth:.0f} cm 上限"
            sub_schemes, _, _ = calc_per_meter_schemes(
                r_key,
                r_Ab,
                spacing_s,
                edge_c,
                row_gap,
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
                alt_text = "需加大間距/排距或採機械續接"

        batch_rows.append(
            {
                "鋼筋號數": r_key,
                "斷面積 Ab": f"{r_Ab:.3f} cm²",
                "第25章埋深 (含α_bond)": f"{d25} cm (α={alpha_b:.2f})",
                "第17章錨定埋深": d17_str,
                "檢核結論": status,
                "每公尺雙排替代方案建議": alt_text,
            }
        )
    return batch_rows


def build_geometry_check_rows(db, sx, sy, edge_c, hef_ref, cNa):
    min_struct = round(max(6.0 * db, 10.0), 1)
    crit_s = round(max(3.0 * hef_ref, 2.0 * cNa), 1)
    crit_e = round(max(1.5 * hef_ref, cNa), 1)

    def eval_item(val, crit_val, label_over):
        if val >= crit_val:
            return "✅ 達無折減臨界距離"
        elif val >= min_struct:
            return f"⚠️ 合格 ({label_over})"
        else:
            return f"❌ 違規 (< 構造下限 {min_struct} cm)"

    return [
        {
            "幾何項目": "X 向水平中心間距 (sx)",
            "設定數值": f"{sx} cm ({sx/db:.1f} db)",
            "構造最小下限": f"{min_struct} cm (max 6db, 10cm)",
            "無折減臨界距離": f"{crit_s} cm (max 3hef, 2cNa)",
            "判定狀態": eval_item(sx, crit_s, "已計入群組投影折減"),
        },
        {
            "幾何項目": "Y 向雙排垂直排距 (sy)",
            "設定數值": f"{sy} cm ({sy/db:.1f} db)",
            "構造最小下限": f"{min_struct} cm (max 6db, 10cm)",
            "無折減臨界距離": f"{crit_s} cm (max 3hef, 2cNa)",
            "判定狀態": eval_item(sy, crit_s, "雙排錯開需檢核斜向距"),
        },
        {
            "幾何項目": "最小自由邊距 (Edge c_a,min)",
            "設定數值": f"{edge_c} cm ({edge_c/db:.1f} db)",
            "構造最小下限": f"{min_struct} cm (max 6db, 10cm)",
            "無折減臨界距離": f"{crit_e} cm (max 1.5hef, cNa)",
            "判定狀態": eval_item(edge_c, crit_e, "已計入邊距折減 ψ_ed"),
        },
    ]


# ==========================================
# 5. Word 完整計算書產生函數
# ==========================================
def build_full_word_report(
    inputs_dict,
    batch_rows,
    target_rebar,
    orig_spacing,
    per_m_schemes,
    fig_buf,
    geom_rows,
    adhesive_msg,
):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Microsoft JhengHei"
    style.font.size = Pt(9.5)

    title = doc.add_heading("ACI 318 連續壁植筋：雙軌物理修正、批次總表與雙排配置計算書", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_meta = doc.add_paragraph(
        f"計算日期：{datetime.now().strftime('%Y-%m-%d %H:%M')}  |  單位系統：公制 (cm, kgf, kgf/cm²)"
    )
    p_meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.add_heading("一、 連續壁工程條件、物理修正參數與藥劑審查", level=2)
    t1 = doc.add_table(rows=1, cols=2)
    t1.style = "Table Grid"
    for idx, cell in enumerate(next(iter(t1.rows)).cells):
        cell.text = "參數項目" if idx == 0 else "設定數值"
    for k, v in inputs_dict.items():
        rc = t1.add_row().cells
        for idx, cell in enumerate(rc):
            cell.text = str(k) if idx == 0 else str(v)

    doc.add_paragraph(f"• 工程 Case 黏著力合規審查：{adhesive_msg}")

    doc.add_heading("二、 #4 ～ #10 全號數雙軌埋深檢核總表", level=2)
    b_headers = ("鋼筋號數", "斷面積 Ab", "第25章埋深", "第17章埋深", "檢核結論", "每公尺雙排替代方案")
    t2 = doc.add_table(rows=1, cols=len(b_headers))
    t2.style = "Table Grid"
    for idx, cell in enumerate(next(iter(t2.rows)).cells):
        cell.text = b_headers[idx]

    for r_data in batch_rows:
        rc = t2.add_row().cells
        vals = (
            r_data["鋼筋號數"],
            r_data["斷面積 Ab"],
            r_data["第25章埋深 (含α_bond)"],
            r_data["第17章錨定埋深"],
            r_data["檢核結論"],
            r_data["每公尺雙排替代方案建議"],
        )
        for idx, cell in enumerate(rc):
            cell.text = str(vals[idx])

    doc.add_heading(f"三、 【{target_rebar} @ {orig_spacing} cm】每公尺寬度替代配筋與雙排配置圖", level=2)
    if per_m_schemes:
        for i, sc in enumerate(per_m_schemes, 1):
            doc.add_paragraph(
                f"【方案 {i}】{sc['方案名稱']}（每公尺總支數 {sc['total_bars_m']} 支，"
                f"提供斷面積 {sc['As_new_m']} cm²/m ≥ 原設計 {sc['As_orig_m']} cm²/m [{sc['ratio_pct']}%]，"
                f"建議植入深度 {sc['rec_depth']} cm，斜向中心距 {sc['diag_spacing']} cm，依據：{sc['method']}）",
                style="List Bullet",
            )
        if fig_buf is not None:
            fig_buf.seek(0)
            doc.add_picture(fig_buf, width=Inches(6.2))
            cap = doc.add_paragraph("圖 1：連續壁每公尺寬度（100 cm）上下雙排梅花錯開鑽孔配置示意圖")
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        doc.add_paragraph(
            f"目前設定之 {target_rebar} @ {orig_spacing} cm 可直接依原號數施作，無須拆分為雙排小號數替代。"
        )

    doc.add_heading("四、 X、Y 間距與 Edge 邊距空間限制檢核", level=2)
    g_headers = ("幾何項目", "設定數值", "構造最小下限", "無折減臨界距離", "判定狀態")
    t3 = doc.add_table(rows=1, cols=len(g_headers))
    t3.style = "Table Grid"
    for idx, cell in enumerate(next(iter(t3.rows)).cells):
        cell.text = g_headers[idx]
    for gr in geom_rows:
        rc = t3.add_row().cells
        vals = (gr["幾何項目"], gr["設定數值"], gr["構造最小下限"], gr["無折減臨界距離"], gr["判定狀態"])
        for idx, cell in enumerate(rc):
            cell.text = str(vals[idx])

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio
# ==========================================
# 6. Streamlit 側邊欄與主畫面介面
# ==========================================
st.title("🧱 連續壁錯位植筋：雙軌物理修正、批次總表與雙排配置系統")

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
    st.header("🧪 2. 工程 Case 與植筋膠黏著力")
    selected_case = st.selectbox("選擇工程應用 Case", options=tuple(ADHESIVE_CASE_DB.keys()), index=3)
    min_req_tau = ADHESIVE_CASE_DB[selected_case]["min_tau"]
    tau_k = st.number_input("藥劑特徵黏結應力 τ_k (kgf/cm²)", 30.0, 220.0, 105.0, 5.0)

    st.divider()
    st.header("🔩 3. 配筋與 X/Y/Edge 空間設定")
    orig_rebar = st.selectbox("重點檢核鋼筋號數", options=tuple(REBAR_DB.keys()), index=4)
    orig_spacing = st.number_input("原設計水平間距 sx (cm)", 6.0, 50.0, 15.0, 0.5)
    row_vertical_gap = st.number_input("雙排錯開垂直排距 sy (cm)", 6.0, 30.0, 10.0, 0.5)
    edge_c = st.number_input("最小自由邊距 Edge c_a,min (cm)", 5.0, 150.0, 20.0, 1.0)
    splice_mode = st.selectbox(
        "伸展植筋接頭受力型式",
        options=("壁內直接延伸作為主筋 (1.0 ld)", "壁邊直接與主筋重疊搭接 (1.3 ld)"),
        index=0,
    )

    st.divider()
    st.header("🎯 4. 材料強度與反算目標")
    fc = st.selectbox("連續壁混凝土強度 f'c (kgf/cm²)", options=(210, 245, 280, 350, 420), index=3)
    steel_type = st.selectbox("鋼筋材質", options=tuple(STEEL_DB.keys()), index=0)
    stress_ratio = st.slider("設計應力比 (Nu / Ab·fy)", 0.50, 1.00, 1.00, 0.05)
    require_ductile = st.checkbox("強制滿足延性破壞 (非鋼材標稱力 ≥ fy)", value=True)

db = REBAR_DB[orig_rebar]["db"]
Ab = REBAR_DB[orig_rebar]["Ab"]
fy = STEEL_DB[steel_type]["fy"]
fu = STEEL_DB[steel_type]["fu"]
is_lap = "1.3" in splice_mode

d25_single, fbd_ref, alpha_bond = calc_ch25_with_bond_physics(
    db, fy, fc, orig_spacing, edge_c, tau_k, is_lap
)
d17_single, _, _, cNa_single, gov_mode_17 = solve_ch17_parallel_depth(
    db, Ab, fy, fu, fc, tau_k, orig_spacing, edge_c, stress_ratio, require_ductile
)

tau_ok = tau_k >= min_req_tau
tau_msg = (
    f"✅ 藥劑黏著力合格：τ_k = {tau_k} kgf/cm² ≥ {selected_case} 最低規定 ({min_req_tau} kgf/cm²)。"
    if tau_ok
    else f"❌ 藥劑黏著力不足：τ_k = {tau_k} kgf/cm² < {selected_case} 最低規定 ({min_req_tau} kgf/cm²)！"
)
if tau_ok:
    st.success(tau_msg)
else:
    st.error(tau_msg)

if alpha_bond > 1.0:
    st.warning(
        f"⚠️ **低黏結藥劑物理修正啟動**：τ_k ({tau_k} kgf/cm²) < 混凝土基準握裹應力 f_bd,ref ({fbd_ref:.1f} kgf/cm²)，第 25 章伸展長度已自動放大 **{alpha_bond:.2f} 倍**以防止脫膠！"
    )
else:
    st.info(
        f"🛡️ **高黏結藥劑劈裂上限控制**：τ_k ({tau_k} kgf/cm²) ≥ 混凝土基準握裹應力 f_bd,ref ({fbd_ref:.1f} kgf/cm²)，第 25 章伸展長度受「混凝土保護層劈裂上限」控制（α_bond = 1.00），不再因調高 τ_k 而縮短。"
    )

batch_table = run_batch_all_rebars(
    orig_spacing,
    edge_c,
    row_vertical_gap,
    max_allow_depth,
    fy,
    fu,
    fc,
    tau_k,
    is_lap,
    stress_ratio,
    require_ductile,
)
per_m_schemes, bars_orig_m, As_orig_m = calc_per_meter_schemes(
    orig_rebar,
    Ab,
    orig_spacing,
    edge_c,
    row_vertical_gap,
    max_allow_depth,
    fy,
    fu,
    fc,
    tau_k,
    is_lap,
    stress_ratio,
    require_ductile,
)
ref_h_geom = d17_single if (d17_single is not None and d17_single < d25_single) else d25_single
geom_rows = build_geometry_check_rows(
    db, orig_spacing, row_vertical_gap, edge_c, ref_h_geom, cNa_single
)

tab1, tab2, tab3 = st.tabs(
    [
        "📋 一、 #4 ～ #10 全號數批次總表",
        "📐 二、 每公尺配筋換算與雙排配置圖",
        "📏 三、 X/Y/Edge 幾何限制與黏著力規範",
    ]
)

with tab1:
    st.subheader(
        f"📊 #4 ～ #10 全號數雙軌埋深檢核總表（sx = {orig_spacing} cm, Edge = {edge_c} cm, 上限 {max_allow_depth:.0f} cm）"
    )
    st.dataframe(batch_table, width="stretch", hide_index=True)
    st.caption(
        "註：表中第 25 章埋深已納入 α_bond 物理修正；第 17 章埋深同步標示控制破壞模式（黏結拉出 Na 或混凝土錐裂 Ncb）。"
    )

with tab2:
    st.subheader(f"📐 【{orig_rebar} @ {orig_spacing} cm】每公尺寬度配筋換算與雙排配置圖")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("原設計每公尺支數", f"{bars_orig_m:.2f} 支/m", f"間距 @ {orig_spacing} cm")
    c2.metric("原設計每公尺總斷面積", f"{As_orig_m:.2f} cm²/m", f"單支 Ab = {Ab} cm²")
    c3.metric("原號數第25章埋深", f"{d25_single} cm", f"α_bond = {alpha_bond:.2f}")
    c4.metric(
        "原號數第17章埋深",
        f"{d17_single} cm" if d17_single is not None else "群組錐裂超限",
        gov_mode_17,
    )

    fig_buffer = None
    if per_m_schemes:
        scheme_names = tuple(s["方案名稱"] for s in per_m_schemes)
        selected_name = st.selectbox("選擇要檢視與匯出圖面的雙排替代方案：", options=scheme_names)
        chosen_scheme = next(s for s in per_m_schemes if s["方案名稱"] == selected_name)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("替代每公尺總支數", f"{chosen_scheme['total_bars_m']} 支/m", "上下雙排各半")
        m2.metric(
            "單排水平中心距 sx",
            f"{chosen_scheme['s_each_row']} cm",
            f"垂直排距 sy = {chosen_scheme['row_gap']:.0f} cm",
        )
        m3.metric("梅花錯開斜向中心距", f"{chosen_scheme['diag_spacing']} cm", "確保圍束間距")
        m4.metric(
            "建議安全植入深度",
            f"{chosen_scheme['rec_depth']} cm",
            chosen_scheme["method"],
        )

        st.markdown("#### 🖼️ 連續壁每公尺寬度（100 cm）雙排梅花錯開鑽孔配置圖")
        fig_buffer = plot_staggered_layout(orig_rebar, orig_spacing, chosen_scheme)
        st.image(fig_buffer, width="content")
    else:
        st.info(
            f"💡 目前條件下，原設計 **{orig_rebar} @ {orig_spacing} cm** 可直接依原號數施作（無須拆分為小號數雙排）；若想測試雙排配置圖，可將左側「重點檢核鋼筋號數」調整為 **#8、#9 或 #10**。"
        )

with tab3:
    st.subheader(f"📏 【{orig_rebar}】X、Y 間距與 Edge 邊距空間限制檢核表")
    st.dataframe(geom_rows, width="stretch", hide_index=True)

    st.subheader("📚 各種工程 Case 植筋膠最小黏著力規定總表")
    case_ref_rows = [
        {
            "工程應用 Case": k,
            "最低特徵黏結應力規定": f"≥ {v['min_tau']} kgf/cm²",
            "規範依據與備註": v["desc"],
        }
        for k, v in ADHESIVE_CASE_DB.items()
    ]
    st.dataframe(case_ref_rows, width="stretch", hide_index=True)

# ==========================================
# 7. Word 計算書一鍵匯出
# ==========================================
st.divider()
st.subheader("📄 匯出完整結構計算書 (含物理修正、批次總表、配置圖與幾何檢核)")

inputs_doc = {
    "連續壁幾何與深度限制": f"壁厚 {wall_thick:.0f} cm，剔除劣質層 {bad_layer:.0f} cm，最大允許埋深 {max_allow_depth:.0f} cm",
    "重點檢核配筋與間距": f"{orig_rebar} @ sx={orig_spacing} cm，雙排排距 sy={row_vertical_gap} cm，Edge={edge_c} cm",
    "材料強度與物理修正": f"f'c = {fc} kgf/cm² (基準 f_bd,ref = {fbd_ref:.1f} kgf/cm²)，{steel_type}",
    "藥劑黏著力與修正係數": f"τ_k = {tau_k} kgf/cm² (第25章修正係數 α_bond = {alpha_bond:.2f})",
    "受力與延性目標": f"{splice_mode}，設計應力比 = {stress_ratio:.2f}，延性檢核 = {'開啟' if require_ductile else '關閉'}",
}

word_bytes = build_full_word_report(
    inputs_doc,
    batch_table,
    orig_rebar,
    orig_spacing,
    per_m_schemes,
    fig_buffer,
    geom_rows,
    tau_msg,
)
rebar_clean = next(iter(orig_rebar.split()))

st.download_button(
    label="⬇️ 下載完整 Word 計算書 (.docx)",
    data=word_bytes,
    file_name=f"ACI318_連續壁植筋完整計算書_{rebar_clean}_fc{fc}_{datetime.now().strftime('%Y%m%d')}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
)