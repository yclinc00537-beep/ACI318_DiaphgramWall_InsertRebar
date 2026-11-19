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
    page_title="ACI 318 連續壁植筋：全功能結構設計與施工品管系統",
    page_icon="🧱",
    layout="wide",
)

# 鋼筋資料庫 (號數: {直徑 db cm, 斷面積 Ab cm2, 標準鑽孔直徑 d0 cm})
REBAR_DB = {
    "#4 (D13)": {"db": 1.27, "Ab": 1.267, "d0": 1.6},
    "#5 (D16)": {"db": 1.59, "Ab": 1.986, "d0": 2.0},
    "#6 (D19)": {"db": 1.91, "Ab": 2.865, "d0": 2.5},
    "#7 (D22)": {"db": 2.22, "Ab": 3.871, "d0": 2.8},
    "#8 (D25)": {"db": 2.54, "Ab": 5.067, "d0": 3.2},
    "#9 (D29)": {"db": 2.87, "Ab": 6.469, "d0": 3.5},
    "#10 (D32)": {"db": 3.22, "Ab": 8.143, "d0": 4.0},
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
# 2. 核心雙軌物理修正計算引擎 & 壁外伸出長度計算
# ==========================================
def calc_ch25_with_bond_physics(db, fy, fc, spacing_s, edge_c, tau_k, is_lap):
    cb = min(edge_c, max((spacing_s - db) / 2.0, 1.0 * db))
    confinement = min(cb / db, 2.5)
    psi_s = 0.8 if db <= 1.91 else 1.0
    psi_g = 1.0 if fy <= 4200 else 1.15

    fbd_ref = 4.5 * math.sqrt(fc)
    alpha_bond = max(1.0, fbd_ref / tau_k)

    ld_base = (0.28 * fy / math.sqrt(fc)) * ((1.0 * 1.0 * psi_s * psi_g) / confinement) * db
    ld_mod = max(ld_base * alpha_bond, 30.0)
    req_depth = math.ceil(ld_mod * 1.3 if is_lap else ld_mod)
    return req_depth, fbd_ref, alpha_bond


def calc_exterior_lap_and_hook(db, fy, fc_new_slab, spacing_s, cover_new=4.0):
    """計算伸出連續壁外之直線搭接長度、標準彎鉤長度及 SA 級摩擦焊續接器最短預留長度"""
    cb = min(cover_new + db / 2.0, spacing_s / 2.0)
    confinement = min(cb / db, 2.5)
    psi_s = 0.8 if db <= 1.91 else 1.0
    psi_g = 1.0 if fy <= 4200 else 1.15

    ld_bot_raw = (0.28 * fy / math.sqrt(fc_new_slab)) * ((1.0 * psi_s * psi_g) / confinement) * db
    ld_bot = max(ld_bot_raw, 30.0)
    lap_bot = math.ceil(1.3 * ld_bot)

    ld_top = max(ld_bot_raw * 1.3, 30.0)
    lap_top = math.ceil(1.3 * ld_top)

    ldh_raw = (0.075 * fy / math.sqrt(fc_new_slab)) * db
    ldh = math.ceil(max(ldh_raw, 8.0 * db, 15.0))

    # SA 級機械續接器 (現場摩擦焊)：夾持操作空間 + 套筒半長 + 植筋膠高溫隔熱緩衝 >= max(8*db, 15 cm)
    sa_coupler_min = math.ceil(max(8.0 * db, 15.0))

    return {
        "ld_bot": math.ceil(ld_bot),
        "lap_bot": lap_bot,
        "ld_top": math.ceil(ld_top),
        "lap_top": lap_top,
        "ldh": ldh,
        "sa_coupler": sa_coupler_min,
    }


def calc_ch17_row_capacity(hef, db, Ab, fy, fu, fc, tau_k, spacing_s, edge_c, n_row=10):
    futa = min(fu, 1.9 * fy, 8800.0)
    phi_Nsa = 0.75 * (Ab * futa)

    Nb = 16.0 * math.sqrt(fc) * (hef**1.5)
    ANco = 9.0 * (hef**2)
    bx = (n_row - 1) * min(spacing_s, 3.0 * hef) + 3.0 * hef
    by = min(edge_c, 1.5 * hef) + 1.5 * hef
    ANc = bx * by
    psi_ed_N = 1.0 if edge_c >= 1.5 * hef else (0.7 + 0.3 * (edge_c / (1.5 * hef)))
    Ncbg_per_bar = ((ANc / ANco) * psi_ed_N * Nb) / n_row
    phi_Ncb_per_bar = 0.65 * Ncbg_per_bar

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
# 3. 注膠量估算 & 現場拉拔試驗驗收標準計算
# ==========================================
def calc_grout_and_pullout_qa(
    rebar_key, hef_cm, total_bars_job, cartridge_ml, waste_pct, fy, stress_ratio
):
    db = REBAR_DB[rebar_key]["db"]
    Ab = REBAR_DB[rebar_key]["Ab"]
    d0 = REBAR_DB[rebar_key]["d0"]

    # 1. 單孔理論環狀注膠體積 (cm3 = ml) 與含損耗體積
    vol_net_ml = (math.pi / 4.0) * (d0**2 - db**2) * hef_cm
    vol_gross_ml = vol_net_ml * (1.0 + waste_pct / 100.0)
    total_vol_ml = vol_gross_ml * total_bars_job
    cartridges_needed = math.ceil(total_vol_ml / cartridge_ml)

    # 2. 現場非破壞拉拔試驗 (Proof Load) 驗收標準
    yield_load_kgf = Ab * fy
    design_load_kgf = yield_load_kgf * stress_ratio
    # 試驗拉力取 1.25 倍設計拉力，但上限不得超過鋼筋降伏力之 80% (防止鋼筋降伏變形)
    proof_load_kgf = min(1.25 * design_load_kgf, 0.80 * yield_load_kgf)
    # 抽驗支數：總支數 5% 且不少於 3 支 (若總支數 < 3 則全檢)
    sample_count = min(total_bars_job, max(3, math.ceil(total_bars_job * 0.05)))

    return {
        "d0_mm": int(round(d0 * 10)),
        "vol_net_ml": round(vol_net_ml, 1),
        "vol_gross_ml": round(vol_gross_ml, 1),
        "total_vol_ml": round(total_vol_ml, 0),
        "cartridges_needed": cartridges_needed,
        "yield_load_kgf": round(yield_load_kgf, 0),
        "design_load_kgf": round(design_load_kgf, 0),
        "proof_load_kgf": round(proof_load_kgf, 0),
        "sample_count": sample_count,
    }


# ==========================================
# 4. 每公尺寬度配筋換算 & 雙排梅花錯開配置圖
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
                        "method": method_note,
                    }
                )
                break

    return schemes, bars_per_m_orig, As_per_m_orig


def plot_staggered_layout(orig_rebar, orig_spacing, scheme):
    fig, ax = plt.subplots(figsize=(8.8, 3.8), dpi=150)

    wall_rect = patches.Rectangle(
        (0, 0), 100, 38, linewidth=1.5, edgecolor="#475569", facecolor="#F8FAFC"
    )
    ax.add_patch(wall_rect)

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
# 5. #4 ～ #10 全號數批次計算與空間幾何檢核
# ==========================================
def run_batch_all_rebars(
    spacing_s,
    edge_c,
    row_gap,
    max_allow_depth,
    fy,
    fu,
    fc,
    fc_slab,
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
        ext_info = calc_exterior_lap_and_hook(r_db, fy, fc_slab, spacing_s)

        d17_str = f"{d17} cm ({gov_17})" if d17 is not None else "群組錐裂超限"
        ok25 = d25 <= max_allow_depth
        ok17 = (d17 is not None) and (d17 <= max_allow_depth)

        if ok25 or ok17:
            best_d = min(d25 if ok25 else 999, d17 if ok17 else 999)
            ctrl_src = "第17章" if (ok17 and (not ok25 or d17 <= d25)) else "第25章"
            status = f"✅ 合格 ({max(best_d, 30)} cm / {ctrl_src})"
            alt_text = "原號數可直接施作"
        else:
            status = f"❌ 超過 {max_allow_depth:.0f} cm"
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
                alt_text = "需加大間距或採機械續接"

        batch_rows.append(
            {
                "鋼筋號數": r_key,
                "壁內第25章埋深": f"{d25} cm (α={alpha_b:.2f})",
                "壁內第17章埋深": d17_str,
                "壁外直線搭接(下/上)": f"{ext_info['lap_bot']} / {ext_info['lap_top']} cm",
                "壁外彎鉤 ldh": f"{ext_info['ldh']} cm",
                "SA級摩擦焊預留": f"{ext_info['sa_coupler']} cm",
                "壁內檢核結論": status,
                "每公尺雙排替代方案": alt_text,
            }
        )
    return batch_rows


def build_geometry_check_rows(db, sx, sy, edge_c, hef_ref, cNa, d_agg=2.0):
    min_clear_cast_in = round(max(2.5, db, (4.0 / 3.0) * d_agg), 2)
    min_center_cast_in = round(min_clear_cast_in + db, 2)
    min_center_post_5db = round(max(5.0 * db, 10.0), 1)
    min_center_post_6db = round(max(6.0 * db, 10.0), 1)

    crit_s = round(max(3.0 * hef_ref, 2.0 * cNa), 1)
    crit_e = round(max(1.5 * hef_ref, cNa), 1)

    clear_sx = round(sx - db, 2)
    clear_sy = round(sy - db, 2)

    def eval_spacing(center_val, clear_val, label_over):
        if center_val >= crit_s:
            return "✅ 達無折減臨界距離 (100%強度)"
        elif center_val >= min_center_post_6db:
            return f"⚠️ 合格 ({label_over}；淨距 {clear_val} cm ≥ 預埋下限 {min_clear_cast_in} cm)"
        elif center_val >= min_center_post_5db:
            return f"⚠️ 警戒 (介於 5db~6db 鑽孔極限；{label_over})"
        elif clear_val >= min_clear_cast_in:
            return f"❌ 鑽孔過密 (< 植筋下限 {min_center_post_5db} cm，易震碎孔壁)"
        else:
            return f"❌ 雙重違規 (同時小於預埋淨距 {min_clear_cast_in} cm 與植筋中心距 {min_center_post_5db} cm)"

    def eval_edge(edge_val):
        if edge_val >= crit_e:
            return "✅ 達無折減臨界邊距"
        elif edge_val >= min_center_post_6db:
            return "⚠️ 合格 (已計入邊距折減 ψ_ed)"
        else:
            return f"❌ 違規 (< 植筋最小邊距下限 {min_center_post_6db} cm)"

    return [
        {
            "幾何項目": "X 向水平中心間距 (sx)",
            "設定數值 (中心 / 淨距)": f"{sx} cm / 淨距 {clear_sx} cm",
            "預埋鋼筋下限 (25.2節)": f"淨距 ≥ {min_clear_cast_in} cm (中心 {min_center_cast_in} cm)",
            "後置植筋鑽孔下限 (AC308)": f"中心 ≥ {min_center_post_5db} ~ {min_center_post_6db} cm (5db~6db)",
            "無折減臨界距離": f"{crit_s} cm (max 3hef, 2cNa)",
            "雙重標準判定狀態": eval_spacing(sx, clear_sx, "已計入群組投影折減"),
        },
        {
            "幾何項目": "Y 向雙排垂直排距 (sy)",
            "設定數值 (中心 / 淨距)": f"{sy} cm / 淨距 {clear_sy} cm",
            "預埋鋼筋下限 (25.2節)": f"淨距 ≥ {min_clear_cast_in} cm (中心 {min_center_cast_in} cm)",
            "後置植筋鑽孔下限 (AC308)": f"中心 ≥ {min_center_post_5db} ~ {min_center_post_6db} cm (5db~6db)",
            "無折減臨界距離": f"{crit_s} cm (max 3hef, 2cNa)",
            "雙重標準判定狀態": eval_spacing(sy, clear_sy, "雙排錯開需檢核斜向距"),
        },
        {
            "幾何項目": "最小自由邊距 (Edge c_a,min)",
            "設定數值 (中心 / 淨距)": f"{edge_c} cm / 淨保護層 {round(edge_c - db/2, 2)} cm",
            "預埋鋼筋下限 (25.2節)": "淨保護層 ≥ 4.0 ~ 7.5 cm (依環境條件)",
            "後置植筋鑽孔下限 (AC308)": f"中心 ≥ {min_center_post_6db} cm (max 6db, 10cm)",
            "無折減臨界距離": f"{crit_e} cm (max 1.5hef, cNa)",
            "雙重標準判定狀態": eval_edge(edge_c),
        },
    ]
# ==========================================
# 6. Word 完整計算書產生函數
# ==========================================
def build_full_word_report(
    inputs_dict,
    batch_rows,
    target_rebar,
    orig_spacing,
    ext_dict,
    per_m_schemes,
    fig_buf,
    geom_rows,
    qa_dict,
    qa_rebar_label,
    qa_hef,
    qa_total_bars,
    cartridge_ml,
    adhesive_msg,
):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Microsoft JhengHei"
    style.font.size = Pt(9.5)

    title = doc.add_heading("ACI 318 連續壁植筋：結構設計、壁外搭接與施工品管計算書", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_meta = doc.add_paragraph(
        f"計算日期：{datetime.now().strftime('%Y-%m-%d %H:%M')}  |  單位系統：公制 (cm, kgf, kgf/cm²)"
    )
    p_meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.add_heading("一、 連續壁工程條件與藥劑合規審查", level=2)
    t1 = doc.add_table(rows=1, cols=2)
    t1.style = "Table Grid"
    for idx, cell in enumerate(next(iter(t1.rows)).cells):
        cell.text = "參數項目" if idx == 0 else "設定數值"
    for k, v in inputs_dict.items():
        rc = t1.add_row().cells
        for idx, cell in enumerate(rc):
            cell.text = str(k) if idx == 0 else str(v)
    doc.add_paragraph(f"• 工程 Case 黏著力合規審查：{adhesive_msg}")

    doc.add_heading("二、 #4 ～ #10 全號數壁內埋深與壁外伸出長度總表", level=2)
    b_headers = ("鋼筋號數", "壁內第25章", "壁內第17章", "壁外搭接(下/上)", "彎鉤 ldh", "SA級摩擦焊", "結論", "雙排替代方案")
    t2 = doc.add_table(rows=1, cols=len(b_headers))
    t2.style = "Table Grid"
    for idx, cell in enumerate(next(iter(t2.rows)).cells):
        cell.text = b_headers[idx]

    for r_data in batch_rows:
        rc = t2.add_row().cells
        vals = (
            r_data["鋼筋號數"],
            r_data["壁內第25章埋深"],
            r_data["壁內第17章埋深"],
            r_data["壁外直線搭接(下/上)"],
            r_data["壁外彎鉤 ldh"],
            r_data["SA級摩擦焊預留"],
            r_data["壁內檢核結論"],
            r_data["每公尺雙排替代方案"],
        )
        for idx, cell in enumerate(rc):
            cell.text = str(vals[idx])

    doc.add_heading(f"三、 【{target_rebar} @ {orig_spacing} cm】壁外搭接、SA級續接與雙排配置圖", level=2)
    doc.add_paragraph(
        f"• 壁外新澆置梁版內伸出長度（{target_rebar}）：\n"
        f"  1. 下層主筋（ψt=1.0）：基本伸展 ld = {ext_dict['ld_bot']} cm，B 級受拉搭接長度 = {ext_dict['lap_bot']} cm。\n"
        f"  2. 上層主筋（ψt=1.3，下方新混凝土 > 30cm）：伸展 ld = {ext_dict['ld_top']} cm，B 級受拉搭接長度 = {ext_dict['lap_top']} cm。\n"
        f"  3. 若壁外寬度受限改採 90°/180° 標準彎鉤錨定：最小彎鉤伸展長度 ldh = {ext_dict['ldh']} cm。\n"
        f"  4. 若改採 SA 級機械續接器（現場摩擦焊）：壁外最短預留鋼筋長度 = {ext_dict['sa_coupler']} cm（含焊機夾持空間與隔熱緩衝；施焊時植筋根部須包裹濕布或冷卻銅塊，防止高溫傳導損及植筋膠）。"
    )

    if per_m_schemes:
        for i, sc in enumerate(per_m_schemes, 1):
            doc.add_paragraph(
                f"【雙排替代方案 {i}】{sc['方案名稱']}（每公尺 {sc['total_bars_m']} 支，"
                f"提供斷面積 {sc['As_new_m']} cm²/m ≥ 原設計 {sc['As_orig_m']} cm²/m [{sc['ratio_pct']}%]，"
                f"壁內植入深度 {sc['rec_depth']} cm，斜向中心距 {sc['diag_spacing']} cm）",
                style="List Bullet",
            )
        if fig_buf is not None:
            fig_buf.seek(0)
            doc.add_picture(fig_buf, width=Inches(6.2))
            cap = doc.add_paragraph("圖 1：連續壁每公尺寬度（100 cm）上下雙排梅花錯開鑽孔配置示意圖")
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        doc.add_paragraph(f"目前設定之 {target_rebar} @ {orig_spacing} cm 可直接依原號數施作，無須雙排替代。")

    doc.add_heading("四、 X、Y 間距與 Edge 邊距空間限制檢核", level=2)
    g_headers = ("幾何項目", "設定數值", "預埋鋼筋下限", "後置植筋下限", "無折減臨界距離", "雙重標準判定")
    t3 = doc.add_table(rows=1, cols=len(g_headers))
    t3.style = "Table Grid"
    for idx, cell in enumerate(next(iter(t3.rows)).cells):
        cell.text = g_headers[idx]
    for gr in geom_rows:
        rc = t3.add_row().cells
        vals = (
            gr["幾何項目"],
            gr["設定數值 (中心 / 淨距)"],
            gr["預埋鋼筋下限 (25.2節)"],
            gr["後置植筋鑽孔下限 (AC308)"],
            gr["無折減臨界距離"],
            gr["雙重標準判定狀態"],
        )
        for idx, cell in enumerate(rc):
            cell.text = str(vals[idx])

    doc.add_heading("五、 植筋膠用量估算與現場非破壞拉拔試驗驗收標準", level=2)
    qa_map = {
        "施工計算基準規格": f"{qa_rebar_label}，標準鑽孔直徑 d0 = {qa_dict['d0_mm']} mm，有效埋深 hef = {qa_hef} cm",
        "施工總支數與單孔注膠量": f"總計 {qa_total_bars} 支；單孔淨體積 {qa_dict['vol_net_ml']} ml（含損耗 {qa_dict['vol_gross_ml']} ml/孔）",
        "植筋膠總需求量估算": f"總用膠量約 {qa_dict['total_vol_ml']:,.0f} ml ➔ 建議備料 {qa_dict['cartridges_needed']} 支（每支 {cartridge_ml} ml）",
        "現場拉拔抽驗支數規定": f"依規範抽驗 5%（且不少於 3 支）➔ 本批至少抽驗 {qa_dict['sample_count']} 支",
        "現場試驗拉力 (Proof Load)": f"{qa_dict['proof_load_kgf']:,.0f} kgf / 支（取 1.25 倍設計拉力，並以 0.80 Ab·fy = {0.8*qa_dict['yield_load_kgf']:,.0f} kgf 為上限）",
        "現場試驗合格判定基準": "加載至試驗拉力持荷 2 分鐘，植筋膠無滑移剝離、混凝土無放射狀開裂、鋼筋無降伏變形即為合格。",
    }
    t4 = doc.add_table(rows=1, cols=2)
    t4.style = "Table Grid"
    for idx, cell in enumerate(next(iter(t4.rows)).cells):
        cell.text = "品管與備料項目" if idx == 0 else "計算數值與驗收標準"
    for k, v in qa_map.items():
        rc = t4.add_row().cells
        for idx, cell in enumerate(rc):
            cell.text = str(k) if idx == 0 else str(v)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# ==========================================
# 7. Streamlit 側邊欄與主畫面介面
# ==========================================
st.title("🧱 連續壁錯位植筋：全功能結構設計與施工品管系統")

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
        "壁內伸展植筋受力型式",
        options=("壁內直接延伸作為主筋 (1.0 ld)", "壁邊直接與主筋重疊搭接 (1.3 ld)"),
        index=0,
    )

    st.divider()
    st.header("🎯 4. 混凝土強度與反算目標")
    fc = st.selectbox("連續壁混凝土強度 f'c (kgf/cm²)", options=(210, 245, 280, 350, 420), index=3)
    fc_slab = st.selectbox("壁外新澆置梁版強度 f'c,slab (kgf/cm²)", options=(210, 245, 280, 350, 420), index=2)
    steel_type = st.selectbox("鋼筋材質", options=tuple(STEEL_DB.keys()), index=0)
    d_agg = st.selectbox("粗骨材最大粒徑 d_agg (cm)", options=(1.9, 2.0, 2.5, 3.8), index=1)
    stress_ratio = st.slider("設計應力比 (Nu / Ab·fy)", 0.50, 1.00, 1.00, 0.05)
    require_ductile = st.checkbox("強制滿足延性破壞 (非鋼材標稱力 ≥ fy)", value=True)

    st.divider()
    st.header("📦 5. 施工範圍與注膠備料參數")
    job_length_m = st.number_input("錯位施工區段總長度 (m)", 1.0, 500.0, 10.0, 1.0)
    cartridge_ml = st.selectbox("單支植筋膠包裝容量 (ml)", options=(500, 390, 585, 1400), index=0)
    waste_pct = st.slider("現場施工注膠損耗率 (%)", 10, 40, 20, 5)
    
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
ext_single = calc_exterior_lap_and_hook(db, fy, fc_slab, orig_spacing)

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
        f"🛡️ **高黏結藥劑劈裂上限控制**：τ_k ({tau_k} kgf/cm²) ≥ 混凝土基準握裹應力 f_bd,ref ({fbd_ref:.1f} kgf/cm²)，第 25 章伸展長度受「混凝土保護層劈裂上限」控制（α_bond = 1.00）。"
    )

batch_table = run_batch_all_rebars(
    orig_spacing,
    edge_c,
    row_vertical_gap,
    max_allow_depth,
    fy,
    fu,
    fc,
    fc_slab,
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
    db, orig_spacing, row_vertical_gap, edge_c, ref_h_geom, cNa_single, d_agg
)

tab1, tab2, tab3, tab4 = st.tabs(
    [
        "📋 一、 #4～#10 全號數總表",
        "📐 二、 壁外搭接彎鉤與雙排配置圖",
        "📏 三、 X/Y/Edge 幾何與黏著力規範",
        "🧪 四、 注膠量估算與拉拔驗收標準",
    ]
)

with tab1:
    st.subheader(
        f"📊 #4 ～ #10 全號數壁內埋深與壁外伸出長度總表（壁內上限 {max_allow_depth:.0f} cm，壁外 f'c,slab = {fc_slab} kgf/cm²）"
    )
    st.dataframe(batch_table, width="stretch", hide_index=True)

with tab2:
    st.subheader(f"🔗 1. 【{orig_rebar}】壁外新澆置梁版內伸出長度計算（f'c,slab = {fc_slab} kgf/cm²）")
    e1, e2, e3, e4, e5 = st.columns(5)
    e1.metric("下層筋直線搭接", f"{ext_single['lap_bot']} cm", f"基本 ld={ext_single['ld_bot']} cm")
    e2.metric("上層筋直線搭接", f"{ext_single['lap_top']} cm", "含 ψt=1.3 與 B級搭接")
    e3.metric("90°/180° 標準彎鉤", f"{ext_single['ldh']} cm", "規範 ldh 伸展長度")
    e4.metric("SA級摩擦焊最短預留", f"{ext_single['sa_coupler']} cm", "含夾持與隔熱緩衝 (≥8db)")
    e5.metric(
        "壁內建議植入深度",
        f"{min(d25_single, d17_single) if d17_single else d25_single} cm",
        f"壁內上限 {max_allow_depth:.0f} cm",
    )
    st.caption(
        "🔥 現場摩擦焊施工安全提示：摩擦焊介面高溫達 1,000°C 以上，施焊時植筋根部務必包覆濕布或冷卻散熱夾具，避免熱傳導使壁內植筋膠超過熱變形溫度（~60°C）而軟化脫膠。"
    )

    st.divider()
    st.subheader(f"📐 2. 【{orig_rebar} @ {orig_spacing} cm】每公尺寬度雙排替代配置圖")

    fig_buffer = None
    active_scheme = None
    if per_m_schemes:
        scheme_names = tuple(s["方案名稱"] for s in per_m_schemes)
        selected_name = st.selectbox("選擇雙排替代配置方案：", options=scheme_names)
        active_scheme = next(s for s in per_m_schemes if s["方案名稱"] == selected_name)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("替代每公尺總支數", f"{active_scheme['total_bars_m']} 支/m", "上下雙排各半")
        m2.metric("單排水平中心距 sx", f"{active_scheme['s_each_row']} cm", f"垂直排距 sy = {active_scheme['row_gap']:.0f} cm")
        m3.metric("梅花錯開斜向中心距", f"{active_scheme['diag_spacing']} cm", "滿足圍束間距")
        m4.metric("替代壁內植入深度", f"{active_scheme['rec_depth']} cm", active_scheme["method"])

        fig_buffer = plot_staggered_layout(orig_rebar, orig_spacing, active_scheme)
        st.image(fig_buffer, width="content")
    else:
        st.info(
            f"💡 目前條件下，原設計 **{orig_rebar} @ {orig_spacing} cm** 可直接依原號數施作（無須拆分為小號數雙排）。"
        )

with tab3:
    st.subheader(f"📏 【{orig_rebar}】X、Y 間距與 Edge 邊距雙重標準檢核表")
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

with tab4:
    st.subheader("🧪 施工注膠備料估算與現場非破壞拉拔試驗驗收標準表")

    if active_scheme is not None:
        qa_rebar = active_scheme["cand_key"]
        qa_hef = active_scheme["rec_depth"]
        qa_bars_per_m = active_scheme["total_bars_m"]
        qa_note = f"依【替代方案：{active_scheme['方案名稱']}】計算"
    else:
        qa_rebar = orig_rebar
        qa_hef = max(min(d25_single, d17_single) if d17_single else d25_single, 30)
        qa_bars_per_m = bars_orig_m
        qa_note = f"依【原設計：{orig_rebar} @ {orig_spacing} cm】計算"

    total_bars_job = math.ceil(qa_bars_per_m * job_length_m)
    qa_res = calc_grout_and_pullout_qa(
        qa_rebar, qa_hef, total_bars_job, cartridge_ml, waste_pct, fy, stress_ratio
    )

    st.info(f"📌 目前計算基準：**{qa_note}** ｜ 施工區段長度 **{job_length_m} m**（共計 **{total_bars_job} 支**植筋）")

    q1, q2, q3, q4 = st.columns(4)
    q1.metric("標準鑽孔直徑 d0", f"{qa_res['d0_mm']} mm", f"鋼筋 {qa_rebar} / 埋深 {qa_hef} cm")
    q2.metric("單孔含損耗注膠量", f"{qa_res['vol_gross_ml']} ml/孔", f"淨體積 {qa_res['vol_net_ml']} ml (+{waste_pct}%)")
    q3.metric(f"預估植筋膠總支數 ({cartridge_ml}ml)", f"{qa_res['cartridges_needed']} 支", f"總用量 {qa_res['total_vol_ml']:,.0f} ml")
    q4.metric("現場拉拔試驗拉力 (Proof Load)", f"{qa_res['proof_load_kgf']:,.0f} kgf", f"至少抽驗 {qa_res['sample_count']} 支 (5%)")

    qa_table_display = [
        {"品管與備料檢核項目": "施工區段總植筋支數", "計算結果與標準": f"{total_bars_job} 支（每公尺 {qa_bars_per_m:.1f} 支 × {job_length_m} m）"},
        {"品管與備料檢核項目": "鑽孔孔徑與深度規格", "計算結果與標準": f"鑽孔直徑 d0 = {qa_res['d0_mm']} mm，健全混凝土有效埋深 hef = {qa_hef} cm"},
        {"品管與備料檢核項目": "單孔注膠體積（淨量 / 含損耗）", "計算結果與標準": f"淨體積 {qa_res['vol_net_ml']} ml / 含 {waste_pct}% 損耗共 {qa_res['vol_gross_ml']} ml"},
        {"品管與備料檢核項目": "整批工程植筋膠採購支數", "計算結果與標準": f"總用膠量 {qa_res['total_vol_ml']:,.0f} ml ➔ 需準備 {qa_res['cartridges_needed']} 支（{cartridge_ml} ml/支）"},
        {"品管與備料檢核項目": "現場拉拔試驗抽驗數量", "計算結果與標準": f"依規範抽驗總支數 5% 且不少於 3 支 ➔ 至少抽驗 {qa_res['sample_count']} 支"},
        {"品管與備料檢核項目": "現場非破壞試驗拉力 (Proof Load)", "計算結果與標準": f"{qa_res['proof_load_kgf']:,.0f} kgf（1.25倍設計拉力，且不超過 0.80 Ab·fy = {0.8*qa_res['yield_load_kgf']:,.0f} kgf）"},
        {"品管與備料檢核項目": "現場試驗合格判定條件", "計算結果與標準": "加載至試驗拉力並持荷 2 分鐘，無滑移、無混凝土劈裂、鋼筋無降伏變形為合格。"},
    ]
    st.dataframe(qa_table_display, width="stretch", hide_index=True)

# ==========================================
# 8. Word 完整計算書一鍵匯出
# ==========================================
st.divider()
st.subheader("📄 匯出完整結構設計與施工品管計算書 (Word)")

inputs_doc = {
    "連續壁幾何與深度限制": f"壁厚 {wall_thick:.0f} cm，剔除劣質層 {bad_layer:.0f} cm，最大允許埋深 {max_allow_depth:.0f} cm",
    "重點檢核配筋與間距": f"{orig_rebar} @ sx={orig_spacing} cm，雙排排距 sy={row_vertical_gap} cm，Edge={edge_c} cm",
    "材料強度設定": f"連續壁 f'c = {fc} kgf/cm²，壁外新澆置梁版 f'c,slab = {fc_slab} kgf/cm²，{steel_type}",
    "藥劑黏著力與物理修正": f"τ_k = {tau_k} kgf/cm² (基準 f_bd,ref = {fbd_ref:.1f} kgf/cm²，α_bond = {alpha_bond:.2f})",
    "施工區段與備料條件": f"施工長度 {job_length_m} m，單支膠容量 {cartridge_ml} ml，預估損耗 {waste_pct}%",
}

word_bytes = build_full_word_report(
    inputs_doc,
    batch_table,
    orig_rebar,
    orig_spacing,
    ext_single,
    per_m_schemes,
    fig_buffer,
    geom_rows,
    qa_res,
    qa_rebar,
    qa_hef,
    total_bars_job,
    cartridge_ml,
    tau_msg,
)
rebar_clean = next(iter(orig_rebar.split()))

st.download_button(
    label="⬇️ 下載完整 Word 計算書 (.docx) [含壁外搭接、全號數總表、配置圖 & 品管驗收表]",
    data=word_bytes,
    file_name=f"ACI318_連續壁植筋完整設計與品管計算書_{rebar_clean}_fc{fc}_{datetime.now().strftime('%Y%m%d')}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
)
