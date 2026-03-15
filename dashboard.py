import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import math
import itertools
import os
import io
import time
from datetime import datetime
import streamlit.components.v1 as components

# --- 輔助函式：產生 a 的斜體加下標字元 (用於 DataFrame 與圖表標籤) ---
def get_a_subscript(val):
    sub_map = str.maketrans("0123456789", "₀₁₂₃₄₅₆₇₈₉")
    return f"𝑎{str(val).translate(sub_map)}"

# --- 0. 基本設定 ---
st.set_page_config(page_title="基於生成式AI與網路可靠度於製造系統戰情儀表設計", page_icon="🏭", layout="wide", initial_sidebar_state="expanded")

DEFAULT_EXCEL_PATH = "!!!最新版簡單!!!.xlsx"

# --- 1. 全局 CSS 與 Modal 樣式 ---
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
    .stApp { background: #23395B !important; color: #e6eef6; font-family: 'Inter', sans-serif; }
    .block-container { padding-top: 2rem !important; padding-bottom: 2rem !important; }
    section[data-testid="stSidebar"] { background-color: #0b1626 !important; border-right: 1px solid rgba(255, 255, 255, 0.1); }
    section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] .stMarkdown p { color: #e6eef6 !important; font-weight: 500 !important; }
    section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 { color: #ffffff !important; }
    [data-testid='stFileUploader'] label[data-testid='stWidgetLabel'] { color: #FFFFFF !important; font-size: 1.2rem !important; font-weight: 700 !important; text-shadow: 0 2px 4px rgba(0,0,0,0.5); }
    [data-testid='stFileUploader'] .stMarkdown p { color: #e0e0e0 !important; }
    [data-testid='stFileUploader'] { background-color: rgba(243, 162, 26, 0.15); border: 2px dashed #f3a21a; border-radius: 12px; padding: 20px; }
    [data-testid='stFileUploader'] button { background-color: #f3a21a !important; color: #12223A !important; border: 2px solid #ffffff !important; font-size: 18px !important; font-weight: 900 !important; border-radius: 8px !important; }
    div.stButton > button { border-radius: 8px !important; font-weight: bold !important; font-size: 16px !important; border: none !important; padding: 0.6rem 1.2rem !important; transition: all 0.2s ease !important; width: 100%; }
    div.stButton > button[kind="primary"] { background-color: #3fe6ff !important; color: #000000 !important; box-shadow: 0 4px 10px rgba(63, 230, 255, 0.4); }
    div.stButton > button[kind="primary"]:hover { background-color: #88f2ff !important; transform: translateY(-2px); }
    div.stButton > button:not([kind="primary"]) { background-color: #4cd37a !important; color: #000000 !important; box-shadow: 0 4px 10px rgba(76, 211, 122, 0.4); }
    div.stButton > button:not([kind="primary"]):hover { background-color: #72e89a !important; transform: translateY(-2px); }
    .kpi-row { display:flex; gap:18px; align-items:stretch; width:100%; }
    .kpi-box { flex:1; border-radius:10px; padding:18px; background: linear-gradient(180deg, rgba(255,255,255,0.02), rgba(255,255,255,0.01)); box-shadow: 0 6px 18px rgba(2,8,23,0.5); border: 2px solid rgba(255,255,255,0.06); min-height:92px; transition: transform 0.18s ease; }
    .kpi-label { color:#f3a21a; font-weight:700; font-size:18px; margin-bottom:8px; }
    .kpi-value { color:#3fe6ff; font-weight:800; font-size:26px; letter-spacing:1px; }
    .kpi-border-green { border-color: #4cd37a !important; }
    .kpi-border-yellow { border-color: #ffd86b !important; }
    .kpi-border-red { border-color: #ff6b6b !important; }
    @keyframes kpiPulse { 0% { transform: scale(1); box-shadow: 0 0 0 0 rgba(255, 216, 107, 0.7); } 50% { transform: scale(1.05); box-shadow: 0 0 20px 10px rgba(255, 216, 107, 0); } 100% { transform: scale(1); box-shadow: 0 0 0 0 rgba(255, 216, 107, 0); } }
    .kpi-pulse { animation: kpiPulse 1.5s infinite; z-index: 10; border-color: #ffd86b !important; }
    @keyframes kpiShake { 0% { transform: translateX(0); box-shadow: 0 0 0 rgba(255,107,107,0); } 25% { transform: translateX(-5px) rotate(-1deg); box-shadow: 0 0 15px rgba(255,107,107,0.5); } 50% { transform: translateX(5px) rotate(1deg); box-shadow: 0 0 25px rgba(255,107,107,0.8); } 75% { transform: translateX(-5px) rotate(-1deg); box-shadow: 0 0 15px rgba(255,107,107,0.5); } 100% { transform: translateX(0); box-shadow: 0 0 0 rgba(255,107,107,0); } }
    .kpi-shake { animation: kpiShake 0.5s infinite; border-color: #ff6b6b !important; }
    .topo-container { position: relative; width: 100%; height: 100px; display: flex; align-items: center; justify-content: center; overflow: visible !important; }
    .topo-node { width: 60px; height: 60px; border-radius: 50%; display: flex; align-items: center; justify-content: center; color: #fff; border: 3px solid rgba(255,255,255,0.3); box-shadow: 0 4px 10px rgba(0,0,0,0.3); transition: all 0.3s ease; position: relative; z-index: 2; background: #23395B; }
    .topo-node-content { display: inline-flex; align-items: baseline; justify-content: center; }
    .topo-node i { font-family: 'Times New Roman', serif; font-size: 1.6rem; font-weight: 700; font-style: italic; }
    .topo-node sub { font-size: 0.8rem; font-weight: 900; margin-left: 2px; }
    .node-green { background: linear-gradient(135deg, #4cd37a, #218838); box-shadow: 0 0 15px rgba(76, 211, 122, 0.4); }
    .node-yellow { background: linear-gradient(135deg, #ffd86b, #e0a800); box-shadow: 0 0 15px rgba(255, 216, 107, 0.4); }
    .node-red { background: linear-gradient(135deg, #ff6b6b, #c82333); box-shadow: 0 0 15px rgba(255, 107, 107, 0.6); }
    .node-fail { background: #8B0000 !important; animation: failBlink 0.8s infinite, kpiShake 0.4s infinite !important; box-shadow: 0 0 30px rgba(255, 0, 0, 0.8) !important; z-index: 10; }
    .node-fail::after { content: "FAIL"; position: absolute; top: -25px; color: #ff6b6b; font-weight: 900; font-size: 14px; text-shadow: 0 2px 4px #000; left: 50%; transform: translateX(-50%); }
    .pre-connector-line { position: absolute; top: 50%; right: 50%; width: 100%; height: 2px; background: #cccccc; transform: translateY(-50%); z-index: 1; }
    .pre-connector-line::after { content: ''; position: absolute; top: -4px; width: 0; height: 0; border-top: 5px solid transparent; border-bottom: 5px solid transparent; border-left: 8px solid #cccccc; right: 35px; }
    .input-group, .output-group { position: absolute; top: 50%; transform: translateY(-50%); display: flex; align-items: center; white-space: nowrap; z-index: 5; }
    .input-group { right: 50%; margin-right: 35px; }
    .output-group { left: 50%; margin-left: 35px; }
    .input-label, .output-label { color: #fff; font-weight: 700; font-size: 16px; text-shadow: 0 2px 4px rgba(0,0,0,0.8); }
    .input-label { margin-right: 8px; }
    .output-label { margin-left: 8px; }
    .input-arrow, .output-arrow { width: 40px; height: 2px; background: #cccccc; position: relative; }
    .input-arrow::after, .output-arrow::after { content: ''; position: absolute; right: 0; top: -4px; border-top: 5px solid transparent; border-bottom: 5px solid transparent; border-left: 8px solid #cccccc; }
    .detail-card-highlight { border: 2px solid #3fe6ff; background: rgba(63, 230, 255, 0.1); padding: 15px; border-radius: 10px; margin-top: 10px; margin-bottom: 20px; }
    [data-testid="stPlotlyChart"] { background-color: #ffffff !important; border-radius: 18px; box-shadow: 0 8px 24px rgba(0,0,0,0.20); padding: 10px; margin-bottom: 20px; }
    .success-modal-overlay { position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background: rgba(0, 0, 0, 0.6); display: flex; justify-content: center; align-items: center; backdrop-filter: blur(4px); animation: fadeOutContainer 2.5s forwards; z-index: 999999; }
    .success-modal-content { background: rgba(20, 24, 30, 0.95); border: 2px solid #4cd37a; border-radius: 16px; padding: 40px 60px; text-align: center; box-shadow: 0 0 40px rgba(76, 211, 122, 0.4); }
    @keyframes fadeOutContainer { 0% { opacity: 1; pointer-events: auto; } 70% { opacity: 1; pointer-events: auto; } 100% { opacity: 0; pointer-events: none; z-index: -1; } }
    button[data-baseweb="tab"][aria-selected="false"] { color: #FFFFFF !important; }
    </style>
    """,
    unsafe_allow_html=True
)

# --- 2. 狀態檢查與 Modal 渲染 ---
if "show_success_modal" not in st.session_state: st.session_state.show_success_modal = False

if st.session_state.show_success_modal:
    st.balloons()
    st.toast("✅ 資料已儲存並同步更新！", icon="💾")
    st.markdown("""
        <div class="success-modal-overlay">
            <div class="success-modal-content">
                <div style="font-size: 60px; margin-bottom: 10px;">✅</div>
                <h2 style="color: #4cd37a; margin: 0;">儲存成功</h2>
                <p style="color: #ddd; margin-top: 10px;">Dashboard 已完成同步更新</p>
            </div>
        </div>
    """, unsafe_allow_html=True)
    st.session_state.show_success_modal = False

# --- 3. 輔助函式與核心計算邏輯 ---
def parse_list_from_string(s):
    if isinstance(s, list): return s
    if pd.isna(s) or str(s).strip() == "": return []
    s = str(s).strip().replace('[', '').replace(']', '')
    try: return [float(x.strip()) for x in s.split(',') if x.strip()]
    except: return []

def get_default_data():
    return pd.DataFrame([
        {"Station": 1, "p": 0.96, "power": 200.0, "k": 1.5, "capacities": "[0, 4800, 9600, 14400, 19200, 24000]", "probs": "[0.001, 0.003, 0.005, 0.007, 0.012, 0.972]"},
        {"Station": 2, "p": 0.95, "power": 25.0, "k": 2.0, "capacities": "[0, 5750, 11500, 17250, 23000]", "probs": "[0.001, 0.001, 0.004, 0.005, 0.989]"},
        {"Station": 3, "p": 0.94, "power": 40.0, "k": 3.0, "capacities": "[0, 4400, 8800, 13200, 17600, 22000]", "probs": "[0.001, 0.003, 0.003, 0.005, 0.007, 0.981]"},
        {"Station": 4, "p": 0.93, "power": 30.0, "k": 5.0, "capacities": "[0, 5250, 10500, 15750, 21000]", "probs": "[0.003, 0.005, 0.007, 0.01, 0.975]"},
        {"Station": 5, "p": 0.97, "power": 15.0, "k": 1.0, "capacities": "[0, 8500, 17000, 25500]", "probs": "[0.001, 0.001, 0.003, 0.995]"}
    ])

def load_data_from_excel_authority(file_source=None):
    if file_source is None:
        path = DEFAULT_EXCEL_PATH
        if not os.path.exists(path):
            return get_default_data(), {"d": 10000, "carbon_factor": 0.474, "tb": 1.0} 
        file_source = path

    try:
        if hasattr(file_source, 'name') and file_source.name.endswith('.csv'):
            df_raw = pd.read_csv(file_source, header=None)
            file_source.seek(0)
            df_data = pd.read_csv(file_source)
        else:
            df_raw = pd.read_excel(file_source, header=None)
            df_data = pd.read_excel(file_source)
            
        d_val, co2_val, tb_val = 10000, 0.474, 1.0
        try:
            for r_idx, row in df_raw.iterrows():
                for c_idx, val in enumerate(row):
                    if pd.notna(val):
                        val_str = str(val).strip()
                        if val_str == "d=": d_val = float(df_raw.iloc[r_idx, c_idx + 1])
                        if val_str == "CO2=": co2_val = float(df_raw.iloc[r_idx, c_idx + 1])
                        if val_str == "Tb=": tb_val = float(df_raw.iloc[r_idx, c_idx + 1])
        except Exception:
            pass

        excel_scalars = {"d": d_val, "carbon_factor": co2_val, "tb": tb_val}
        df_data.rename(columns={"Success rate": "Success_Rate", "capacity": "Capacity"}, inplace=True)
        
        req_cols = ["Station", "Power(kW)加工功率", "Capacity", "Capacity_Prob", "Success_Rate"]
        if not all(col in df_data.columns for col in req_cols):
             return get_default_data(), excel_scalars

        stations = []
        grouped = df_data.groupby("Station", sort=True)
        for name, group in grouped:
            first_row = group.iloc[0]
            caps = group["Capacity"].dropna().tolist()
            probs = group["Capacity_Prob"].dropna().tolist()
            stations.append({
                "Station": int(name),
                "p": float(first_row["Success_Rate"]),
                "power": float(first_row["Power(kW)加工功率"]),
                "k": float(first_row.get("k", 1.0) if pd.notna(first_row.get("k")) else 1.0),
                "capacities": str(caps),
                "probs": str(probs)
            })
            
        return pd.DataFrame(stations), excel_scalars

    except Exception as e:
        st.error(f"檔案讀取錯誤: {e}。已載入預設資料。")
        return get_default_data(), {"d": 10000, "carbon_factor": 0.474, "tb": 1.0}

if "df_data" not in st.session_state:
    df_loaded, excel_auth_data = load_data_from_excel_authority()
    st.session_state.df_data = df_loaded
    st.session_state.excel_authority = excel_auth_data 

if "force_tab_index" not in st.session_state:
    st.session_state.force_tab_index = None

# --- 核心運算 ---
def calculate_metrics(demand, carbon_factor, _station_data, tb_value):
    n = len(_station_data)
    
    # 1. 計算 pi
    mu = 1.0
    pi_list = []
    for d in _station_data:
        p_val = d.get('p', 0.96)
        k_val = d.get('k', 1.0)
        pi = p_val * math.exp(-k_val * (tb_value - mu)**2)
        pi_list.append(pi)
    
    # 2. 累乘計算投入量
    product_pi = 1.0
    for pi_val in pi_list: product_pi *= pi_val
    total_input = demand / product_pi
    
    inputs = []
    current_input = total_input
    for i in range(n):
        inputs.append(current_input)
        current_input *= pi_list[i] 
    rounded_inputs = [math.ceil(x) for x in inputs]

    # 3. 計算各站精確的變動 Power 與 Carbon
    selected_caps = []
    energies = []
    carbons = []
    
    for i in range(n):
        req_input = rounded_inputs[i]
        caps = _station_data[i]["capacities"]
        base_power = _station_data[i]["power"]
        max_cap = max(caps) if caps else 0
        
        sel_cap = 0
        for c in caps:
            if c >= req_input:
                sel_cap = c
                break
        if sel_cap == 0: sel_cap = max_cap
        selected_caps.append(sel_cap)
        
        cap_ratio = (sel_cap / max_cap) if max_cap > 0 else 0
        power_ratio = cap_ratio
        
        if i == 4 and n == 5:
            st4_caps = _station_data[3]["capacities"]
            st4_max = max(st4_caps) if st4_caps else 1
            st4_req = rounded_inputs[3]
            st4_sel = next((c for c in st4_caps if c >= st4_req), st4_max)
            st4_ratio = (st4_sel / st4_max) if st4_max > 0 else cap_ratio
            power_ratio = st4_ratio 
        
        actual_power = base_power * power_ratio
        actual_carbon = actual_power * cap_ratio * carbon_factor
        
        energies.append(actual_power)
        carbons.append(actual_carbon)

    calc_total_energy = sum(energies)
    calc_carbon = sum(carbons)

    # 4. 耗損計算
    losses = [inputs[i] * (1 - pi_list[i]) for i in range(n)]
    total_loss = sum(losses)

    # 5. 系統可靠度計算 (排列組合乘積)
    total_probability = 0
    indices_ranges = [range(len(d["capacities"])) for d in _station_data]
    
    limit_count = 0
    for state_indices in itertools.product(*indices_ranges):
        limit_count += 1
        if limit_count > 1000000: break 
        
        current_prob = 1.0
        valid = True
        
        for i, state_idx in enumerate(state_indices):
            cap = _station_data[i]["capacities"][state_idx]
            prob = _station_data[i]["probs"][state_idx]
            if cap < rounded_inputs[i]:
                valid = False
                break
            current_prob *= prob
        if valid:
            total_probability += current_prob

    return {
        "pi_list": pi_list,
        "inputs": inputs,
        "rounded_inputs": rounded_inputs,
        "selected_caps": selected_caps,
        "energies": energies,
        "carbons": carbons,
        "losses": losses, 
        "total_loss": total_loss, 
        "total_energy": calc_total_energy,
        "carbon_emission": calc_carbon,
        "reliability": total_probability,
    }

# --- 4. UI 顯示 ---
st.markdown("""
<div style="padding:14px 10px; border-radius:10px; background: linear-gradient(90deg, rgba(6,21,39,0.6), rgba(8,30,46,0.35)); box-shadow:0 6px 18px rgba(2,8,23,0.6); margin-bottom:12px;">
<h1 style="margin:0;color:#e6f7ff">🏭 基於生成式AI與網路可靠度於製造系統戰情儀表設計</h1>
</div>
""", unsafe_allow_html=True)

tab_dashboard, tab_editor = st.tabs(["📊 戰情儀表板 (Dashboard)", "📝 資料管理 (Excel 編輯)"])

if st.session_state.force_tab_index is not None:
    target_index = st.session_state.force_tab_index
    components.html(
        f"""
        <script>
            setTimeout(function() {{
                var tabs = window.parent.document.querySelectorAll('button[data-baseweb="tab"]');
                if (tabs.length > {target_index}) {{
                    tabs[{target_index}].click();
                }}
            }}, 150);
        </script>
        """,
        height=0, width=0
    )
    st.session_state.force_tab_index = None

# --- TAB 1: Dashboard ---
with tab_dashboard:
    try:
        source_df = st.session_state.df_data
        STATION_DATA = []
        for _, row in source_df.iterrows():
            STATION_DATA.append({
                "name": f"{int(row['Station'])}", 
                "id": int(row['Station']),
                "capacities": parse_list_from_string(row['capacities']),
                "probs": parse_list_from_string(row['probs']),
                "p": float(row['p']),
                "power": float(row['power']),
                "k": float(row.get('k', 1.0))
            })
        FIXED_N = len(STATION_DATA)
    except Exception as e:
        st.error(f"資料結構錯誤: {e}")
        STATION_DATA = []
        FIXED_N = 0

    if not STATION_DATA:
        st.warning("無有效工作站資料")
    else:
        with st.sidebar:
            st.markdown("""<div style='padding:12px 10px; background-color: rgba(255, 255, 255, 0.08); border-radius: 8px; margin-bottom: 15px;'><h3 style='margin:0; color:#ffffff'>系統參數面板</h3></div>""", unsafe_allow_html=True)
            
            auth_data = st.session_state.get("excel_authority", {"d": 10000, "carbon_factor": 0.474, "tb": 1.0})
            demand = st.number_input("輸出量 (𝑑)", min_value=1, value=int(auth_data.get("d", 10000)), step=100)
            carbon_factor = st.number_input("CO₂ 係數 (kg/kWh)", min_value=0.001, value=float(auth_data.get("carbon_factor", 0.474)), step=0.001, format="%.3f")
            tb_val = st.slider("厚度參數 ($T_b$)", min_value=0.8, max_value=1.2, value=float(auth_data.get("tb", 1.0)), step=0.01)
            
            st.divider()
            res = calculate_metrics(demand, carbon_factor, STATION_DATA, tb_val)
            rel_val = res['reliability']
            
            sys_status_sidebar = "green" if rel_val >= 0.95 else "yellow" if rel_val >= 0.9 else "red"
            status_colors = {"green": "#4cd37a", "yellow": "#ffd86b", "red": "#ff6b6b"}
            status_bgs = {"green": "rgba(76, 211, 122, 0.05)", "yellow": "rgba(255, 216, 107, 0.05)", "red": "rgba(255, 107, 107, 0.05)"}
            status_texts = {"green": "可靠度正常", "yellow": "可靠度警告", "red": "可靠度過低"}
            st.markdown(f'<div style="background-color: {status_bgs[sys_status_sidebar]}; padding: 12px; border-radius: 8px; text-align: center; margin-top: 10px;"><span style="color: {status_colors[sys_status_sidebar]}; font-weight: 700; font-size: 16px;">{status_texts[sys_status_sidebar]} : {rel_val:.4f}</span></div>', unsafe_allow_html=True)

            # 更新：斜體化的 Rd 且括號為普通字體
            st.markdown(
"""<div style="padding:15px; background-color: rgba(255, 255, 255, 0.05); border-radius: 8px; margin-top: 25px; border: 1px solid rgba(255,255,255,0.1);">
<h4 style="margin-top:0; color:#e6eef6; font-size: 16px; border-bottom: 1px solid rgba(255,255,255,0.2); padding-bottom: 8px;">🚦 狀態燈號閾值說明</h4>
<div style="font-size: 0.9rem; color: #ddd; margin-top: 10px;">
<div style="margin-bottom: 8px;"><b>系統可靠度 (<span style="font-family: 'Times New Roman', serif; font-style: italic;">R<sub>d</sub></span>)</b></div>
<div style="display: flex; justify-content: space-between; margin-bottom: 4px;"><span style="color:#4cd37a;">🟢 正常 (Green)</span> <span>≥ 0.95</span></div>
<div style="display: flex; justify-content: space-between; margin-bottom: 4px;"><span style="color:#ffd86b;">🟡 警告 (Yellow)</span> <span>0.90 ~ 0.949</span></div>
<div style="display: flex; justify-content: space-between; margin-bottom: 16px;"><span style="color:#ff6b6b;">🔴 危險 (Red)</span> <span>＜ 0.90</span></div>
<div style="margin-bottom: 8px;"><b>總碳排放 (kg)</b></div>
<div style="display: flex; justify-content: space-between; margin-bottom: 4px;"><span style="color:#4cd37a;">🟢 正常 (Green)</span> <span>0 ~ 70</span></div>
<div style="display: flex; justify-content: space-between; margin-bottom: 4px;"><span style="color:#ffd86b;">🟡 警告 (Yellow)</span> <span>71 ~ 100</span></div>
<div style="display: flex; justify-content: space-between;"><span style="color:#ff6b6b;">🔴 危險 (Red)</span> <span>＞ 100</span></div>
</div>
</div>""", unsafe_allow_html=True)

        sys_reliability = res['reliability']
        sys_carbon = res['carbon_emission']
        sys_status = "green" if sys_reliability >= 0.95 else "yellow" if sys_reliability >= 0.9 else "red"
        sys_anim = "kpi-pulse" if sys_status == "yellow" else "kpi-shake" if sys_status == "red" else ""

        node_states = []
        for i, station in enumerate(STATION_DATA):
            is_failed = res["rounded_inputs"][i] > (max(station["capacities"]) if station["capacities"] else 0)
            node_states.append("node-fail" if is_failed else f"node-{sys_status} {sys_anim}")

        st.markdown("### 🕸️ 生產線即時拓樸監控")
        if "selected_node_idx" not in st.session_state: st.session_state.selected_node_idx = None
        station_labels = ["🔽 吹瓶站", "🔽 充填站", "🔽 套標站", "🔽 包裝站", "🔽 疊棧站"]

        topo_cols = st.columns(FIXED_N)
        for i, col in enumerate(topo_cols):
            with col:
                html_content = f"""<div class="topo-container">"""
                if i == 0: html_content += """<div class="input-group"><span class="input-label">Input</span><div class="input-arrow"></div></div>"""
                if i > 0: html_content += '<div class="pre-connector-line"></div>'
                html_content += f"""<div class="topo-node {node_states[i]}"><div class="topo-node-content"><i>a</i><sub>{STATION_DATA[i]["id"]}</sub></div></div>"""
                if i == FIXED_N - 1: html_content += """<div class="output-group"><div class="output-arrow"></div><span class="output-label">Output</span></div>"""
                html_content += "</div>" 
                st.markdown(html_content, unsafe_allow_html=True)
                
                btn_label = station_labels[i] if i < len(station_labels) else f"🔽 工作站 {get_a_subscript(STATION_DATA[i]['id'])}"
                if st.button(btn_label, key=f"btn_node_{i}", type="primary" if st.session_state.selected_node_idx == i else "secondary", use_container_width=True):
                    st.session_state.selected_node_idx = None if st.session_state.selected_node_idx == i else i
                    st.rerun()

        if st.session_state.selected_node_idx is not None:
            idx = st.session_state.selected_node_idx
            if 0 <= idx < len(STATION_DATA):
                d_st = STATION_DATA[idx]
                st_power = res['energies'][idx]
                st_carbon = res['carbons'][idx]
                st_loss = res['losses'][idx]
                
                detail_names = ["吹瓶站", "充填站", "套標站", "包裝站", "疊棧站"]
                st_detail_name = detail_names[idx] if idx < len(detail_names) else f"工作站 {get_a_subscript(d_st['id'])}"
                
                st.markdown(f"""
                <div class="detail-card-highlight">
                <h5 style="margin-bottom: 15px; color: #fff;">🔍 {st_detail_name} 詳細數據</h5>
                <div style="display: flex; justify-content: space-between; text-align: center; gap: 10px;">
                <div style="flex: 1;"><div style="font-size: 0.9rem; color: rgba(255,255,255,0.7); margin-bottom: 4px;">投入量</div><div style="font-size: 1.5rem; font-weight: 700; color: #fff;">{res["rounded_inputs"][idx]}</div></div>
                <div style="flex: 1;"><div style="font-size: 0.9rem; color: rgba(255,255,255,0.7); margin-bottom: 4px;">動態功率 (kW)</div><div style="font-size: 1.5rem; font-weight: 700; color: #fff;">{st_power:.2f}</div></div>
                <div style="flex: 1;"><div style="font-size: 0.9rem; color: rgba(255,255,255,0.7); margin-bottom: 4px;">參數 (𝑘)</div><div style="font-size: 1.5rem; font-weight: 700; color: #fff;">{d_st.get('k', 1.0)}</div></div>
                <div style="flex: 1;"><div style="font-size: 0.9rem; color: rgba(255,255,255,0.7); margin-bottom: 4px;">品質調整後成功率</div><div style="font-size: 1.5rem; font-weight: 700; color: #ffffff;">{res['pi_list'][idx]:.4f}</div></div>
                <div style="flex: 1;"><div style="font-size: 0.9rem; color: rgba(255,255,255,0.7); margin-bottom: 4px;">碳排放 (kg)</div><div style="font-size: 1.5rem; font-weight: 700; color: #fff;">{st_carbon:.3f}</div></div>
                <div style="flex: 1;"><div style="font-size: 0.9rem; color: rgba(255,255,255,0.7); margin-bottom: 4px;">耗損 (qty)</div><div style="font-size: 1.5rem; font-weight: 700; color: #ff6b6b;">{st_loss:.3f}</div></div>
                </div></div>""", unsafe_allow_html=True)

        k1, k2, k3, k4, k5 = st.columns([1,1,1,1,1], gap="large")
        with k1: st.markdown(f'<div class="kpi-box kpi-border-{sys_status} {sys_anim}"><div class="kpi-label">系統可靠度 (<span style="font-family: \'Times New Roman\', serif; font-style: italic;">R<sub>d</sub></span>)</div><div class="kpi-value">{res["reliability"]:.4f}</div></div>', unsafe_allow_html=True)
        with k2: st.markdown(f'<div class="kpi-box"><div class="kpi-label">輸出量 (𝑑)</div><div class="kpi-value">{demand}</div></div>', unsafe_allow_html=True)
        with k3: st.markdown(f'<div class="kpi-box"><div class="kpi-label">動態總功率 (<span style="font-family: \'Times New Roman\', serif;">kW</span>)</div><div class="kpi-value">{res["total_energy"]:.3f}</div></div>', unsafe_allow_html=True)
        
        # 更新：碳排放 KPI 數值框的顏色判定邏輯
        c_color = "green" if sys_carbon <= 70 else "yellow" if sys_carbon <= 100 else "red"
        
        with k4: st.markdown(f'<div class="kpi-box kpi-border-{c_color}"><div class="kpi-label">總碳排放 (kg)</div><div class="kpi-value">{res["carbon_emission"]:.3f}</div></div>', unsafe_allow_html=True)
        with k5: st.markdown(f'<div class="kpi-box kpi-border-red"><div class="kpi-label">總耗損 (qty)</div><div class="kpi-value">{res["total_loss"]:.3f}</div></div>', unsafe_allow_html=True)

        st.divider()
        st.markdown("### 📈 數據視覺化分析")
        plot_stations = [get_a_subscript(d['id']) for d in STATION_DATA]
        
        c1, c2 = st.columns(2)
        with c1:
            fig1 = go.Figure(go.Bar(x=plot_stations, y=res["losses"], marker_color='#60d3ff', name="耗損量"))
            fig1.update_layout(
                title=dict(text="各工作站耗損量", font=dict(size=22, color='black', weight='bold')),
                paper_bgcolor='white', plot_bgcolor='white', height=350, margin=dict(b=0),
                xaxis=dict(title=dict(text='工作站', font=dict(size=18, color='black')), type='category', color='#000000', showline=False, ticks='', ticklen=0, tickfont=dict(size=18, color='#000000', family='Times New Roman')),
                yaxis=dict(title=dict(text='耗損量', font=dict(size=18, color='black')), color='#000000', showline=True, linecolor='#000000', gridcolor='#000000', tickfont=dict(size=16, color='#000000'), range=[0, max(res["losses"])*1.15 if res["losses"] else 1], autorange=False, rangemode='tozero', zeroline=True, zerolinecolor='#000000') 
            )
            st.plotly_chart(fig1, use_container_width=True)
            
        with c2:
            fig2 = go.Figure(go.Bar(x=plot_stations, y=res["energies"], marker_color='#ffcf60', name="動態功率"))
            fig2.update_layout(
                title=dict(text="各工作站動態功率 (kW)", font=dict(size=22, color='black', weight='bold')),
                paper_bgcolor='white', plot_bgcolor='white', height=350, margin=dict(b=0),
                xaxis=dict(title=dict(text='工作站', font=dict(size=18, color='black')), type='category', color='#000000', showline=False, ticks='', ticklen=0, tickfont=dict(size=18, color='#000000', family='Times New Roman')),
                yaxis=dict(title=dict(text='動態功率 (kW)', font=dict(size=18, color='black')), color='#000000', showline=True, linecolor='#000000', gridcolor='#000000', tickfont=dict(size=16, color='#000000'), range=[0, max(res["energies"])*1.15 if res["energies"] else 1], autorange=False, rangemode='tozero', zeroline=True, zerolinecolor='#000000')
            )
            st.plotly_chart(fig2, use_container_width=True)

        st.markdown("### 📉 系統可靠度敏感度分析")
        def get_dynamic_crit_d(_station_data, _tb_val):
            pi_list_local = []
            for st_data in _station_data:
                pi_val = st_data.get('p', 0.96) * math.exp(-st_data.get('k', 1.0) * (_tb_val - 1.0)**2)
                pi_list_local.append(pi_val)
            max_d_limits = []
            for i in range(len(_station_data)):
                prod = 1.0
                for j in range(i, len(_station_data)): prod *= pi_list_local[j]
                max_d_limits.append((max(_station_data[i]["capacities"]) if _station_data[i]["capacities"] else 0) * prod)
            return int(math.floor(min(max_d_limits))) if max_d_limits else 1000

        crit_d = get_dynamic_crit_d(STATION_DATA, tb_val)
        step = max(500, (crit_d // 10 // 500) * 500)
        if step == 0: step = 500
        raw_range = np.arange(10000, crit_d + step + 500, step)
        d_range_vals = [int(v) for v in np.sort(np.unique(np.concatenate((raw_range, [crit_d, crit_d + 1])))) if v <= crit_d + max(step, 1000)]
        y_vals = [calculate_metrics(val, carbon_factor, STATION_DATA, tb_val)['reliability'] for val in d_range_vals]

        fig3 = go.Figure()
        fig3.add_trace(go.Scatter(x=d_range_vals, y=y_vals, mode='lines+markers', name='可靠度曲線', line=dict(color='#3fe6ff', width=3), marker=dict(size=8, color='#3fe6ff'), cliponaxis=False))
        fig3.add_trace(go.Scatter(x=[crit_d], y=[calculate_metrics(crit_d, carbon_factor, STATION_DATA, tb_val)['reliability']], mode='markers', name=f'臨界點 ((𝑑)={crit_d})', marker=dict(symbol='star', size=22, color='#ffd86b', line=dict(width=2, color='#ff0000')), cliponaxis=False))
        fig3.add_trace(go.Scatter(x=[demand], y=[res['reliability']], mode='markers', name=f'當前輸出量 ((𝑑)={demand})', marker=dict(symbol='circle', size=14, color='#4cd37a', line=dict(width=2, color='#ffffff')), cliponaxis=False))
        
        fig3.update_layout(
            title=dict(text="系統可靠度敏感度分析", font=dict(size=22, color='black', weight='bold')),
            xaxis_title=dict(text="輸出量 (𝑑)", font=dict(size=18, color='black')), yaxis_title=dict(text="系統可靠度", font=dict(size=18, color='black')),
            paper_bgcolor='white', plot_bgcolor='white', height=400, margin=dict(l=20, r=20, t=40, b=20),
            legend=dict(yanchor="top", y=0.99, xanchor="right", x=0.99, font=dict(color="black", size=14)),
            xaxis=dict(color='#000000', showline=False, ticks='', ticklen=0, gridcolor='#000000', zeroline=False, tickfont=dict(size=16, color='#000000'), range=[10000, max(d_range_vals + [demand, crit_d]) + max(step, 1000)]),
            yaxis=dict(color='#000000', showline=True, linecolor='#000000', ticks='', ticklen=0, gridcolor='#000000', zeroline=True, zerolinecolor='#000000', tickmode='array', tickvals=[0, 0.2, 0.4, 0.6, 0.8, 1.0], tickfont=dict(size=16, color='#000000'), range=[0, 1.05], rangemode='tozero')
        )
        st.plotly_chart(fig3, use_container_width=True)

        st.markdown("### 📋 工作站狀態表")
        df_res = pd.DataFrame({
            "工作站": [get_a_subscript(d["id"]) for d in STATION_DATA], 
            "參數 (𝑘)": [d.get("k", 1.0) for d in STATION_DATA],
            "品質調整後成功率": [f"{pi:.4f}" for pi in res["pi_list"]],
            "投入量": [f"{x:.2f}" for x in res["inputs"]], 
            "取整輸入量": res["rounded_inputs"],
            "動態功率 (kW)": [f"{p:.2f}" for p in res["energies"]], 
            "碳排放 (kg)": [f"{c:.3f}" for c in res["carbons"]],
            "實際耗損 (qty)": [f"{l:.3f}" for l in res["losses"]],
            "狀態數量": [len(d['capacities']) for d in STATION_DATA]
        })
        st.dataframe(df_res, use_container_width=True)

# --- TAB 2: Editor ---
with tab_editor:
    st.subheader("Excel 資料編輯器 (支援動態長度)")
    col_upload, col_settings = st.columns([2, 1])
    with col_upload:
        uploaded_file = st.file_uploader("📂 上傳 Excel 或 CSV", type=["xlsx", "csv"])

    if uploaded_file:
        file_id = f"{uploaded_file.name}_{uploaded_file.size}"
        if "processed_file_id" not in st.session_state or st.session_state.processed_file_id != file_id:
            try:
                new_df, new_scalars = load_data_from_excel_authority(uploaded_file)
                st.session_state.df_data = new_df
                if new_scalars: st.session_state.excel_authority = new_scalars
                st.session_state.processed_file_id = file_id
                st.session_state.last_uploaded_name = uploaded_file.name
                
                if "editor_table" in st.session_state: del st.session_state["editor_table"]
                st.session_state.force_tab_index = 1
                st.rerun()
            except Exception as e:
                st.error(f"讀取失敗: {e}")

    def maintain_editor_tab(): st.session_state.force_tab_index = 1

    edited_df = st.data_editor(
        st.session_state.df_data[['Station', 'p', 'power', 'k', 'capacities', 'probs']],
        num_rows="dynamic", use_container_width=True, key="editor_table", on_change=maintain_editor_tab, disabled=["k"], 
        column_config={
            "Station": st.column_config.NumberColumn("站號 (𝑎ₙ)", min_value=1, step=1, required=True),
            "p": None, 
            "power": st.column_config.NumberColumn("基礎功率 (kW)", help="此為最大功率，儀表板將依比例動態計算實際功率"),
            "k": st.column_config.NumberColumn("參數 (𝑘)", format="%.2f"),
            "capacities": st.column_config.TextColumn("產能列表 (List)", help="例如 [0, 100, 200]"),
            "probs": st.column_config.TextColumn("機率列表 (List)", help="例如 [0.1, 0.4, 0.5]")
        }
    )

    col_reset, col_save = st.columns([1, 1])
    with col_reset:
        if st.button("🔄 重置為預設資料", use_container_width=True):
            st.session_state.df_data = get_default_data()
            st.session_state.force_tab_index = 1 
            st.rerun()

    with col_save:
        if st.button("💾 儲存並更新", use_container_width=True):
            try:
                validated_rows = []
                for _, row in edited_df.iterrows():
                    caps = parse_list_from_string(row['capacities'])
                    probs = parse_list_from_string(row['probs'])
                    if not isinstance(caps, list) or not isinstance(probs, list): st.error(f"站號 {row['Station']}: 列表格式錯誤"); st.stop()
                    if len(caps) != len(probs): st.error(f"站號 {row['Station']}: 產能({len(caps)})與機率({len(probs)})長度不符"); st.stop()
                    if len(caps) > 1 and not all(x < y for x, y in zip(caps, caps[1:])): st.error(f"站號 {row['Station']}: 產能列表必須嚴格遞增"); st.stop()
                    validated_rows.append((row, caps, probs))

                long_rows = []
                curr_tb = st.session_state.excel_authority.get("tb", 1.0)
                
                for row, caps, probs in validated_rows:
                    for i in range(len(caps)):
                        pi_calc = row['p'] * math.exp(-row.get('k', 1.0) * (curr_tb - 1.0)**2)
                        long_rows.append({
                            "Station": int(row['Station']), "Machine": 1, "Success rate": row['p'],
                            "Power(kW)加工功率": row['power'], "capacity": caps[i], "Capacity_Prob": probs[i],
                            "k": row.get('k', 1.0), "pi(deg)": pi_calc
                        })
                
                df_long = pd.DataFrame(long_rows)
                for i in range(7, 14): 
                    if f"Unnamed: {i}" not in df_long.columns: df_long[f"Unnamed: {i}"] = np.nan
                
                cols = ['Station', 'Machine', 'Success rate', 'Power(kW)加工功率', 'capacity', 'Capacity_Prob', 'k', 'Unnamed: 7', 'Unnamed: 8', 'Unnamed: 9', 'Unnamed: 10', 'Unnamed: 11', 'Unnamed: 12', 'Unnamed: 13', 'pi(deg)']
                df_long = df_long.reindex(columns=cols)
                while len(df_long) < 5: df_long = pd.concat([df_long, pd.DataFrame([np.nan]*df_long.shape[1], columns=df_long.columns)], ignore_index=True)
                
                curr_d = st.session_state.excel_authority.get("d", 10000)
                curr_c = st.session_state.excel_authority.get("carbon_factor", 0.474)
                df_long.iloc[1, 7], df_long.iloc[1, 8] = "d=", curr_d
                df_long.iloc[2, 7], df_long.iloc[2, 8] = "CO2=", curr_c
                df_long.iloc[3, 7], df_long.iloc[3, 8] = "Tb=", curr_tb
                
                save_name = st.session_state.get("last_uploaded_name", "!!!最新版簡單!!!_modified.csv")
                save_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), save_name)
                
                if save_name.endswith('.csv'): df_long.to_csv(save_path, index=False)
                else: df_long.to_excel(save_path, index=False)
                
                st.session_state.df_data = edited_df
                st.session_state.show_success_modal = True
                st.session_state.force_tab_index = 0
                st.rerun()
            except Exception as e:
                st.error(f"儲存失敗: {e}")
#在終端機輸入：python -m streamlit run "C:\Users\user\OneDrive\桌面\dashboard.py"