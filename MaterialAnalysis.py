
import re
from typing import Dict, List, Tuple, Any
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Material Analysis — Money Slide", layout="wide")
st.title("Material Analysis")

# ---------- Sidebar ----------
with st.sidebar:
    st.header("1) Upload workbook")
    file = st.file_uploader("Excel workbook (.xlsx)", type=["xlsx"], key="xlsx")
    st.caption("Workbook must include Run 0 / Run 1 tabs (e.g., A310516_Run 0). 'BladeMap' tab is optional.")
    st.header("2) Options")
    header_row = st.number_input("Header row (Run sheets)", min_value=1, value=6, step=1, key="hdr")
    name_col_label = st.text_input("Name header (Run sheets)", value="Name", key="namehdr")
    dev_col_label = st.text_input("Dev header (Run sheets)", value="Dev.", key="devhdr")
    st.header("3) Go")
    run_btn = st.button("Build", type="primary", use_container_width=True)

if "autorun" not in st.session_state:
    st.session_state.autorun = False
if file is not None and not st.session_state.autorun:
    st.session_state.autorun = True
    run_btn = True

# ---------- Helpers ----------
# Accept "Run 0", "Run0", "Run 0" (NBSP), optional suffixes
RUN_RE = re.compile(r"^(.*?)_Run(?:\s|\u00A0)*([01])\b", flags=re.IGNORECASE)

def parse_sheet_name(name: str) -> Tuple[str, str]:
    m = RUN_RE.search(name)
    if m:
        sn = m.group(1).strip()
        run = m.group(2)
        return sn, run
    return "", ""

def is_valid_base(base: str) -> bool:
    return re.match(r"B\d+_(FRONT|TOP|BACK)$", base.strip().upper()) is not None

def read_run_sheet_points_indexed(df: pd.DataFrame, header_row: int, name_label: str, dev_label: str) -> Dict[str, Dict[str, float]]:
    hdr = df.iloc[header_row-1].astype(str).str.strip().tolist()
    df_cols = {str(c).strip(): i for i, c in enumerate(hdr)}
    name_col = None; dev_col = None
    for col, idx in df_cols.items():
        if col.lower() == name_label.strip().lower(): name_col = idx
        if col.lower().replace(".", "") == dev_label.strip().lower().replace(".", ""): dev_col = idx
    if name_col is None or dev_col is None:
        raise ValueError("Could not find 'Name' and/or 'Dev.' headers on a run sheet.")
    data = df.iloc[header_row:, [name_col, dev_col]].copy()
    data.columns = ["Name", "Dev"]
    out: Dict[str, Dict[str, float]] = {}
    for _, row in data.iterrows():
        name_raw = str(row["Name"]).strip(); v = row["Dev"]
        if name_raw == "" or pd.isna(v): continue
        if ":" not in name_raw: continue
        base, idx = name_raw.split(":", 1); base = base.strip(); idx = idx.strip()
        if not idx.isdigit(): continue
        if not is_valid_base(base): continue
        try: val = float(v)
        except: continue
        out.setdefault(base, {})[idx] = val
    return out

def build_blade_map(xls: pd.ExcelFile):
    try:
        blade_df = xls.parse("BladeMap", header=None, dtype=str)
    except Exception:
        return {}, ["Material"]
    if blade_df.empty:
        return {}, ["Material"]
    materials = [str(x) for x in blade_df.iloc[0, 4:8].tolist()]
    idx: Dict[str, Dict[str, str]] = {}
    for _, row in blade_df.iloc[1:].iterrows():
        sn = str(row.iloc[0]).strip()
        if not sn: continue
        m: Dict[str, str] = idx.setdefault(sn, {})
        for c in range(4, 8):
            blades_str = str(row.iloc[c]).strip()
            if blades_str and blades_str.lower() != "nan":
                parts = [p.strip().upper().replace(" ", "") for p in str(blades_str).split(",") if p.strip()]
                for b in parts: m[b] = str(blade_df.iloc[0, c])
    return idx, materials

def calc_percent(st: Dict[str, Any]) -> float | None:
    if st["Included"] == 0: return None
    dev = st["SumDelta"] / st["Included"]; avg0 = st["SumRun0"] / st["Included"]
    if avg0 == 0: return None
    return (dev / avg0) * 100.0

def order_points(keys: list) -> list:
    def sort_key(p: str):
        p = p.strip().upper()
        blade, face = (p.split("_", 1) + [""])[:2]
        n = int(blade[1:]) if blade[1:].isdigit() else 9999
        pos = {"FRONT":1, "TOP":2, "BACK":3}.get(face, 9)
        return (n, pos)
    return sorted(keys, key=sort_key)

def unique_blades(all_points: list) -> list:
    seen = {}
    for k in all_points: seen[k.split("_")[0]] = True
    return sorted(seen.keys(), key=lambda b: int(b[1:]) if b[1:].isdigit() else 9999)

# ---------- Color helpers ----------
def _interp(a, b, t): return tuple(int(a[i] + (b[i]-a[i])*t) for i in range(3))

def wear_color_matrix(df: pd.DataFrame) -> pd.DataFrame:
    green  = (0,176,80); yellow=(255,255,0); orange=(255,165,0); red=(255,0,0); light_blue=(221,235,247)
    vals = pd.to_numeric(df.values.flatten(), errors="coerce")
    numeric = vals[~np.isnan(vals)]
    if numeric.size == 0:
        return pd.DataFrame([[f"background-color: rgb{light_blue}"]*df.shape[1]]*df.shape[0], index=df.index, columns=df.columns)
    vmin = float(np.min(numeric)); vmax = float(np.max(numeric)); denom = (vmax - vmin) if (vmax - vmin) != 0 else 1.0
    css = []
    for i in range(df.shape[0]):
        row_css = []
        for j in range(df.shape[1]):
            v = df.iat[i, j]
            try: v = float(v)
            except: v = np.nan
            if np.isnan(v): row_css.append(f"background-color: rgb{light_blue}")
            elif v == vmin: row_css.append(f"background-color: rgb{green}")
            elif v == vmax: row_css.append(f"background-color: rgb{red}")
            else:
                t = (v - vmin)/denom
                if t <= 0.5: rgb = _interp(green, yellow, t/0.5)
                elif t <= 0.75: rgb = _interp(yellow, orange, (t-0.5)/0.25)
                else: rgb = _interp(orange, red, (t-0.75)/0.25)
                row_css.append(f"background-color: rgb{rgb}")
        css.append(row_css)
    return pd.DataFrame(css, index=df.index, columns=df.columns)

def color_scale_conf(vals: pd.Series):
    out=[]
    for v in vals:
        if pd.isna(v): out.append(""); continue
        t=float(v)
        if t<=0.5: k=t/0.5; r=255; g=int(k*255); b=0
        else: k=(t-0.5)/0.5; r=int(255-k*255); g=255; b=0
        out.append(f"background-color: rgb({r},{g},{b})")
    return out

# ---------- Renderers ----------
def render_summary(sn_summaries: Dict[str, Dict[str, Any]]):
    st.header("Summary")
    if not sn_summaries:
        st.info("No SN pairs with both Run 0 and Run 1 were found.")
        return
    for sn in sorted(sn_summaries.keys()):
        stats = sn_summaries[sn]
        ordered = order_points(list(stats.keys()))
        rows = {"Included (>0)": [], "Excluded (≤0)": [], "Deviation": [], "Run0 Avg": [], "Result %": []}
        for pt in ordered:
            stx = stats[pt]
            rows["Included (>0)"].append(stx["Included"])
            rows["Excluded (≤0)"].append(stx["Excluded"])
            rows["Deviation"].append(np.nan if stx["Included"]==0 else stx["SumDelta"]/stx["Included"])
            rows["Run0 Avg"].append(np.nan if stx["Included"]==0 else stx["SumRun0"]/stx["Included"])
            p = calc_percent(stx); rows["Result %"].append(np.nan if p is None else p/100.0)
        df = pd.DataFrame(rows, index=ordered).T.reset_index().rename(columns={"index":"Metric \\ Point"})
        st.markdown(f"**SN: {sn}**")
        st.dataframe(df, use_container_width=True)

def make_face_table(sn: str, face: str, columns: list, materials: list,
                    blade_map: Dict[str, Dict[str, str]], values_map: Dict[str, float],
                    is_wear: bool):
    data = []
    for mat in materials:
        row = []
        for pt in columns:
            blade = pt.split("_")[0]
            mat_match = blade_map.get(sn, {}).get(blade.upper())
            if (mat_match or "").strip().upper() == (mat or "").strip().upper():
                v = values_map.get(pt, np.nan)
                row.append(v)
            else:
                row.append(np.nan)
        data.append(row)
    df = pd.DataFrame(data, index=materials, columns=columns)
    if is_wear:
        color_df = wear_color_matrix(df)
        df_disp = df.copy().astype(object)
        for i in range(df_disp.shape[0]):
            for j in range(df_disp.shape[1]):
                v = df.iat[i, j]
                df_disp.iat[i, j] = ("" if pd.isna(v) else f"{float(v):.1%}")
        sty = df_disp.style.apply(lambda _: color_df, axis=None)
        return sty
    else:
        return df.style.format("{:.1%}", na_rep="").apply(color_scale_conf, axis=1)

def render_money_slide(sn_list, money_pct, money_counts, all_points, blade_idx, materials):
    st.header("Money Slide")
    if not sn_list:
        st.info("No SN pairs with both Run 0 and Run 1 were found.")
        return
    st.caption("Left: Wear % per face (global min=green → global max=red; blanks = light blue). Right: Condensed Valid Measurements (Front/Top/Back vs Blades).")
    for sn in sn_list:
        ordered = order_points(all_points[sn])
        blades = unique_blades(ordered)
        faces = ["Front", "Top", "Back"]
        st.markdown(f"### SN: {sn}")
        cols = st.columns([1.4, 0.1, 1.2])  # Wear | gap | condensed
        left, _, mid = cols

        with left:
            for face in faces:
                face_cols = [p for p in ordered if p.upper().endswith(face.upper())]
                if not face_cols: continue
                st.markdown(f"**{face} (Wear %)**")
                sty = make_face_table(sn, face, face_cols, materials, blade_idx, money_pct.get(sn, {}), True)
                st.dataframe(sty, use_container_width=True)

        with mid:
            st.markdown("**Valid Measurements % (Front/Top/Back vs Blades)**")
            rows = []
            for face in faces:
                row = []
                for b in blades:
                    key = f"{b}_{face}"
                    cnt = money_counts.get(sn, {}).get(key)
                    if cnt and (cnt.get('inc',0)+cnt.get('exc',0))>0:
                        frac = cnt['inc']/(cnt['inc']+cnt['exc'])
                    else:
                        frac = np.nan
                    row.append(frac)
                rows.append(row)
            if rows:
                df = pd.DataFrame(rows, index=faces, columns=blades)
                sty = df.style.format("{:.1%}", na_rep="").apply(color_scale_conf, axis=1)
                st.dataframe(sty, use_container_width=True)
            else:
                st.warning("No points available to compute Valid Measurements matrix for this SN.")

# ---------- Pipeline ----------
def run_pipeline(xls: pd.ExcelFile, header_row: int, name_label: str, dev_label: str):
    run0, run1 = {}, {}
    for name in xls.sheet_names:
        sn, rid = parse_sheet_name(name)
        if sn and rid in ("0","1"): (run0 if rid=="0" else run1)[sn] = name

    blade_idx, materials = build_blade_map(xls)

    sn_summaries = {}
    money_pct = {}
    money_counts = {}
    all_points_by_sn = {}

    for sn, sh0 in run0.items():
        if sn not in run1: continue
        sh1 = run1[sn]
        df0 = xls.parse(sh0, header=None); df1 = xls.parse(sh1, header=None)
        pts0 = read_run_sheet_points_indexed(df0, header_row, name_label, dev_label)
        pts1 = read_run_sheet_points_indexed(df1, header_row, name_label, dev_label)

        bases = sorted(set(pts0.keys()) | set(pts1.keys()))
        if not bases: continue

        stats = {}; perPct = {}; counts = {}

        for base in bases:
            inc=0; exc=0; sumD=0.0; sumR0=0.0
            m0 = pts0.get(base, {}); m1 = pts1.get(base, {})
            for idx, v0 in m0.items():
                if idx in m1:
                    d = v0 - m1[idx]
                    if d > 0: inc += 1; sumD += d; sumR0 += v0
                    else: exc += 1
            stx = {"Included": inc, "Excluded": exc, "SumDelta": sumD, "SumRun0": sumR0}
            stats[base] = stx
            p = calc_percent(stx)
            if p is not None: perPct[base] = p/100.0
            blade, face = base.split("_", 1)
            counts[f"{blade}_{face}"] = {"inc":inc, "exc":exc}

        sn_summaries[sn] = stats
        money_pct[sn] = perPct
        money_counts[sn] = counts
        all_points_by_sn[sn] = bases

    tab1, tab2 = st.tabs(["Summary", "Money Slide"])
    with tab1:
        render_summary(sn_summaries)
    with tab2:
        render_money_slide(sorted(all_points_by_sn.keys()), money_pct, money_counts, all_points_by_sn, blade_idx, materials)

if run_btn and file is not None:
    try:
        xls = pd.ExcelFile(file, engine="openpyxl")
        run_pipeline(xls, int(header_row), name_col_label, dev_col_label)
    except Exception as e:
        st.exception(e)
elif run_btn and file is None:
    st.warning("Please upload a workbook (.xlsx) first.")
