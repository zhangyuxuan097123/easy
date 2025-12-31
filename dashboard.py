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

# --- 0. 基本設定 ---
st.set_page_config(page_title="基於生成式AI與網路可靠度於製造系統戰情儀表設計", page_icon="🏭", layout="wide", initial_sidebar_state="expanded")

# 預設 Excel 路徑
DEFAULT_EXCEL_PATH = "新版簡單.xlsx"

# --- 1. 全局 CSS 與 Modal 樣式 (完全保留原版面) ---
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');

    /* 主畫面背景 */
    .stApp { background: #23395B !important; color: #e6eef6; font-family: 'Inter', sans-serif; }
    .block-container { padding-top: 2rem !important; padding-bottom: 2rem !important; }

    /* 側邊欄背景 */
    section[data-testid="stSidebar"] { background-color: #0b1626 !important; border-right: 1px solid rgba(255, 255, 255, 0.1); }
    section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] .stMarkdown p { color: #e6eef6 !important; font-weight: 500 !important; }
    section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 { color: #ffffff !important; }

    /* 上傳區塊 */
    [data-testid='stFileUploader'] label[data-testid='stWidgetLabel'] { color: #FFFFFF !important; font-size: 1.2rem !important; font-weight: 700 !important; text-shadow: 0 2px 4px rgba(0,0,0,0.5); }
    [data-testid='stFileUploader'] .stMarkdown p { color: #e0e0e0 !important; }
    [data-testid='stFileUploader'] { background-color: rgba(243, 162, 26, 0.15); border: 2px dashed #f3a21a; border-radius: 12px; padding: 20px; }
    [data-testid='stFileUploader'] button { background-color: #f3a21a !important; color: #12223A !important; border: 2px solid #ffffff !important; font-size: 18px !important; font-weight: 900 !important; border-radius: 8px !important; }

    /* 按鈕 */
    div.stButton > button { border-radius: 8px !important; font-weight: bold !important; font-size: 16px !important; border: none !important; padding: 0.6rem 1.2rem !important; transition: all 0.2s ease !important; width: 100%; }
    div.stButton > button[kind="primary"] { background-color: #3fe6ff !important; color: #000000 !important; box-shadow: 0 4px 10px rgba(63, 230, 255, 0.4); }
    div.stButton > button[kind="primary"]:hover { background-color: #88f2ff !important; transform: translateY(-2px); }
    div.stButton > button:not([kind="primary"]) { background-color: #4cd37a !important; color: #000000 !important; box-shadow: 0 4px 10px rgba(76, 211, 122, 0.4); }
    div.stButton > button:not([kind="primary"]):hover { background-color: #72e89a !important; transform: translateY(-2px); }

    /* KPI Box */
    .kpi-row { display:flex; gap:18px; align-items:stretch; width:100%; }
    .kpi-box { flex:1; border-radius:10px; padding:18px; background: linear-gradient(180deg, rgba(255,255,255,0.02), rgba(255,255,255,0.01)); box-shadow: 0 6px 18px rgba(2,8,23,0.5); border: 2px solid rgba(255,255,255,0.06); min-height:92px; transition: transform 0.18s ease; }
    .kpi-label { color:#f3a21a; font-weight:700; font-size:18px; margin-bottom:8px; }
    .kpi-value { color:#3fe6ff; font-weight:800; font-size:26px; letter-spacing:1px; }
    .kpi-border-green { border-color: #4cd37a !important; }
    .kpi-border-yellow { border-color: #ffd86b !important; }
    .kpi-border-red { border-color: #ff6b6b !important; }

    /* 動畫 */
    @keyframes kpiPulse { 0% { transform: scale(1); box-shadow: 0 0 0 0 rgba(255, 216, 107, 0.7); } 50% { transform: scale(1.05); box-shadow: 0 0 20px 10px rgba(255, 216, 107, 0); } 100% { transform: scale(1); box-shadow: 0 0 0 0 rgba(255, 216, 107, 0); } }
    .kpi-pulse { animation: kpiPulse 1.5s infinite; z-index: 10; border-color: #ffd86b !important; }
    @keyframes kpiShake { 0% { transform: translateX(0); box-shadow: 0 0 0 rgba(255,107,107,0); } 25% { transform: translateX(-5px) rotate(-1deg); box-shadow: 0 0 15px rgba(255,107,107,0.5); } 50% { transform: translateX(5px) rotate(1deg); box-shadow: 0 0 25px rgba(255,107,107,0.8); } 75% { transform: translateX(-5px) rotate(-1deg); box-shadow: 0 0 15px rgba(255,107,107,0.5); } 100% { transform: translateX(0); box-shadow: 0 0 0 rgba(255,107,107,0); } }
    .kpi-shake { animation: kpiShake 0.5s infinite; border-color: #ff6b6b !important; }

    /* --- [修改] 拓樸圖全新樣式 (解決重疊與美觀問題) --- */
    
    /* 1. 容器設定：確保內容可視，不會被裁切 */
    .topo-container {
        position: relative;
        width: 100%;
        height: 100px;
        display: flex;
        align-items: center;
        justify-content: center;
        overflow: visible !important; /* 關鍵：讓 Input/Output 可以凸出去 */
    }

    /* 2. 節點圓圈 (半徑約 30px) */
    .topo-node { 
        width: 60px; height: 60px; 
        border-radius: 50%; 
        display: flex; align-items: center; justify-content: center; 
        font-weight: bold; font-size: 1.2rem; color: #fff; 
        border: 3px solid rgba(255,255,255,0.3); 
        box-shadow: 0 4px 10px rgba(0,0,0,0.3); 
        transition: all 0.3s ease; 
        position: relative; 
        z-index: 2; /* 確保圓圈蓋在線條上 */
        background: #23395B; /* 預設背景 */
    }
    
    /* 狀態顏色 */
    .node-green { background: linear-gradient(135deg, #4cd37a, #218838); box-shadow: 0 0 15px rgba(76, 211, 122, 0.4); }
    .node-yellow { background: linear-gradient(135deg, #ffd86b, #e0a800); box-shadow: 0 0 15px rgba(255, 216, 107, 0.4); }
    .node-red { background: linear-gradient(135deg, #ff6b6b, #c82333); box-shadow: 0 0 15px rgba(255, 107, 107, 0.6); }
    .node-fail { background: #8B0000 !important; animation: failBlink 0.8s infinite, kpiShake 0.4s infinite !important; box-shadow: 0 0 30px rgba(255, 0, 0, 0.8) !important; z-index: 10; }
    .node-fail::after { content: "FAIL"; position: absolute; top: -25px; color: #ff6b6b; font-weight: 900; font-size: 14px; text-shadow: 0 2px 4px #000; left: 50%; transform: translateX(-50%); }

    /* [修改] 3. 節點左側的連接線 (從左側節點中心到目前節點中心) */
    .pre-connector-line {
        position: absolute;
        top: 50%;
        right: 50%; /* 從目前節點中心向左延伸 */
        width: 100%; /* 延伸到上一個節點中心 (Streamlit Columns 等寬) */
        height: 2px;
        background: #cccccc; /* 實心灰色 */
        transform: translateY(-50%);
        z-index: 1;
    }
    /* 連接線中間的箭頭 (實心，靠近節點左側) */
    .pre-connector-line::after {
        content: '';
        position: absolute;
        top: -4px;
        width: 0;
        height: 0;
        border-top: 5px solid transparent;
        border-bottom: 5px solid transparent;
        border-left: 8px solid #cccccc; /* 實心灰色箭頭 */
        /* [關鍵修改] 35px 確保箭頭在圓圈(半徑30px)的外部左側，不會被蓋住 */
        right: 35px; 
    }

    /* 4. Input 區塊 (絕對定位於節點左側，實心) */
    .input-group {
        position: absolute;
        right: 50%; /* 從中心點開始算 */
        margin-right: 35px; /* 向左推：半徑(30) + 間距(5) */
        top: 50%;
        transform: translateY(-50%);
        display: flex;
        align-items: center;
        white-space: nowrap; /* 強制不換行 */
        z-index: 5;
    }
    .input-label {
        color: #fff;
        font-weight: 700;
        font-size: 16px;
        margin-right: 8px;
        text-shadow: 0 2px 4px rgba(0,0,0,0.8);
    }
    .input-arrow {
        width: 40px;
        height: 2px;
        background: #cccccc; /* [修改] 實心灰色 */
        position: relative;
    }
    .input-arrow::after {
        content: '';
        position: absolute;
        right: 0;
        top: -4px;
        border-top: 5px solid transparent;
        border-bottom: 5px solid transparent;
        border-left: 8px solid #cccccc; /* [修改] 實心灰色箭頭 */
    }

    /* 5. Output 區塊 (絕對定位於節點右側，實心) */
    .output-group {
        position: absolute;
        left: 50%; /* 從中心點開始算 */
        margin-left: 35px; /* 向右推：半徑(30) + 間距(5) */
        top: 50%;
        transform: translateY(-50%);
        display: flex;
        align-items: center;
        white-space: nowrap; /* 強制不換行 */
        z-index: 5;
    }
    .output-label {
        color: #fff;
        font-weight: 700;
        font-size: 16px;
        margin-left: 8px;
        text-shadow: 0 2px 4px rgba(0,0,0,0.8);
    }
    .output-arrow {
        width: 40px;
        height: 2px;
        background: #cccccc; /* [修改] 實心灰色 */
        position: relative;
    }
    .output-arrow::after {
        content: '';
        position: absolute;
        right: 0;
        top: -4px;
        border-top: 5px solid transparent;
        border-bottom: 5px solid transparent;
        border-left: 8px solid #cccccc; /* [修改] 實心灰色箭頭 */
    }

    .detail-card-highlight { border: 2px solid #3fe6ff; background: rgba(63, 230, 255, 0.1); padding: 15px; border-radius: 10px; margin-top: 10px; margin-bottom: 20px; }
    [data-testid="stPlotlyChart"] { background-color: #ffffff !important; border-radius: 18px; box-shadow: 0 8px 24px rgba(0,0,0,0.20); padding: 10px; margin-bottom: 20px; }
      
    /* 成功儲存 Modal 樣式 */
    .success-modal-overlay {
        position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
        background: rgba(0, 0, 0, 0.6);
        display: flex; justify-content: center; align-items: center;
        backdrop-filter: blur(4px);
        animation: fadeOutContainer 2.5s forwards; 
        z-index: 999999;
    }
    .success-modal-content {
        background: rgba(20, 24, 30, 0.95); 
        border: 2px solid #4cd37a; border-radius: 16px;
        padding: 40px 60px; text-align: center;
        box-shadow: 0 0 40px rgba(76, 211, 122, 0.4);
    }
    @keyframes fadeOutContainer {
        0% { opacity: 1; pointer-events: auto; }
        70% { opacity: 1; pointer-events: auto; }
        100% { opacity: 0; pointer-events: none; z-index: -1; }
    }
    
    /* Tabs 未選取狀態文字顏色修正 */
    button[data-baseweb="tab"][aria-selected="false"] {
        color: #FFFFFF !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# --- 2. 狀態檢查與 Modal 渲染 ---
if "show_success_modal" not in st.session_state:
    st.session_state.show_success_modal = False

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
    try:
        return [float(x.strip()) for x in s.split(',') if x.strip()]
    except:
        return []

def get_default_data():
    return pd.DataFrame([
        {"Station": 1, "p": 0.96, "power": 28.9, "capacities": "[0, 600, 1200, 1800, 2400, 3000]", "probs": "[0.001, 0.003, 0.005, 0.007, 0.012, 0.972]"},
        {"Station": 2, "p": 0.96, "power": 46.6, "capacities": "[0, 725, 1450, 2175, 2900]", "probs": "[0.001, 0.001, 0.004, 0.005, 0.989]"},
        {"Station": 3, "p": 0.97, "power": 137.0, "capacities": "[0, 570, 1140, 1710, 2280, 2850]", "probs": "[0.001, 0.003, 0.003, 0.005, 0.007, 0.981]"},
        {"Station": 4, "p": 0.97, "power": 17.7, "capacities": "[0, 725, 1450, 2175, 2900]", "probs": "[0.003, 0.005, 0.007, 0.01, 0.975]"},
        {"Station": 5, "p": 0.97, "power": 38.8, "capacities": "[0, 925, 1850, 2775]", "probs": "[0.001, 0.003, 0.003, 0.995]"}
    ])

def load_data_from_excel_authority(file_source=None):
    if file_source is None:
        path = DEFAULT_EXCEL_PATH
        if not os.path.exists(path):
            return get_default_data(), {"d": 2500, "carbon_factor": 0.474} 
        file_source = path

    try:
        df_raw = pd.read_excel(file_source, header=None)
        d_val, co2_val = 2500, 0.474
        try:
            for r_idx, row in df_raw.iterrows():
                for c_idx, val in enumerate(row):
                    if val == "d=":
                        d_val = float(df_raw.iloc[r_idx, c_idx + 1])
                    if val == "CO2=":
                        co2_val = float(df_raw.iloc[r_idx, c_idx + 1])
        except Exception:
            pass

        excel_scalars = {"d": d_val, "carbon_factor": co2_val}
        df_data = pd.read_excel(file_source)
        
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
                "capacities": str(caps),
                "probs": str(probs)
            })
            
        df_processed = pd.DataFrame(stations)
        return df_processed, excel_scalars

    except Exception as e:
        st.error(f"Excel 讀取錯誤: {e}。已載入預設資料。")
        return get_default_data(), {"d": 2500, "carbon_factor": 0.474}

# 初始化 Session
if "df_data" not in st.session_state:
    df_loaded, excel_auth_data = load_data_from_excel_authority()
    st.session_state.df_data = df_loaded
    st.session_state.excel_authority = excel_auth_data 

# [新增] 用來控制分頁鎖定的變數： None=不強制, 0=強制Dashboard, 1=強制Editor
if "force_tab_index" not in st.session_state:
    st.session_state.force_tab_index = None

# 防呆檢查
if st.session_state.excel_authority is None:
    st.session_state.excel_authority = {"d": 2500, "carbon_factor": 0.474}

def calculate_metrics(demand, carbon_factor, _station_data):
    n = len(_station_data)
    p_list = [d.get('p', 0.96) for d in _station_data]
    power_list = [d.get('power', 0.0) for d in _station_data]
    
    product_p = 1.0
    for p_val in p_list: product_p *= p_val
    total_input = demand / product_p
    
    inputs = []
    current_input = total_input
    for i in range(n):
        inputs.append(current_input)
        current_input *= p_list[i]
    rounded_inputs = [math.ceil(x) for x in inputs]

    # 能耗與碳排 (靜態)
    energies = power_list 
    calc_total_energy = sum(energies)
    calc_carbon = calc_total_energy * carbon_factor

    # 耗損 (Loss) 計算: Input * (1 - p)
    losses = []
    for i in range(n):
        losses.append(inputs[i] * (1 - p_list[i]))
    total_loss = sum(losses)

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
        "inputs": inputs,
        "rounded_inputs": rounded_inputs,
        "energies": energies,
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

# [還原] 使用原本的 Tabs 結構
tab_dashboard, tab_editor = st.tabs(["📊 戰情儀表板 (Dashboard)", "📝 資料管理 (Excel 編輯)"])

# [核心功能]：分頁控制器 (JS Injection)
# 如果 force_tab_index 不是 None，則注入 JS 強制點擊該分頁，然後重置變數
if st.session_state.force_tab_index is not None:
    target_index = st.session_state.force_tab_index
    components.html(
        f"""
        <script>
            // 等待一點時間確保 DOM 載入
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
    # 執行一次後，將強制狀態解除，讓使用者可以自由切換，直到下一次特定事件發生
    st.session_state.force_tab_index = None

# --- TAB 1: Dashboard ---
with tab_dashboard:
    try:
        source_df = st.session_state.df_data
        STATION_DATA = []
        
        for _, row in source_df.iterrows():
            caps = parse_list_from_string(row['capacities'])
            probs = parse_list_from_string(row['probs'])
            
            STATION_DATA.append({
                "name": f"{int(row['Station'])}", 
                "id": int(row['Station']),
                "capacities": caps,
                "probs": probs,
                "p": float(row['p']),
                "power": float(row['power'])
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
            
            # 安全讀取參數
            auth_data = st.session_state.get("excel_authority")
            if auth_data is None: auth_data = {"d": 2500, "carbon_factor": 0.474}
            
            def_d = auth_data.get("d", 2500)
            def_c = auth_data.get("carbon_factor", 0.474)

            # [修改] 使用 LaTeX 語法 ($d$) 讓側邊欄的 d 變成斜體
            demand = st.number_input("輸出量 ($d$)", min_value=1, value=int(def_d), step=100)
            carbon_factor = st.number_input("CO₂ 係數 (kg/kWh)", min_value=0.001, value=float(def_c), step=0.001, format="%.3f")
            
            st.divider()
            
            # 執行計算
            res = calculate_metrics(demand, carbon_factor, STATION_DATA)
            
            if res['reliability'] < 0.8: st.error(f"可靠度過低：{res['reliability']:.4f}")
            else: st.success(f"可靠度正常：{res['reliability']:.4f}")

        # KPI & Logic
        sys_reliability = res['reliability']
        sys_carbon = res['carbon_emission']
        sys_status = "green" if sys_reliability >= 0.9 else "yellow" if sys_reliability >= 0.8 else "red"
        sys_anim = "kpi-pulse" if sys_status == "yellow" else "kpi-shake" if sys_status == "red" else ""

        node_states = []
        for i, station in enumerate(STATION_DATA):
            station_input = res["rounded_inputs"][i]
            max_cap = max(station["capacities"]) if station["capacities"] else 0
            is_failed = station_input > max_cap
            node_states.append("node-fail" if is_failed else f"node-{sys_status} {sys_anim}")

        st.markdown("### 🕸️ 生產線即時拓樸監控")
        if "selected_node_idx" not in st.session_state: st.session_state.selected_node_idx = None
        
        topo_cols = st.columns(FIXED_N)
        for i, col in enumerate(topo_cols):
            with col:
                # [修改] 拓樸圖繪製邏輯：優化 Input/Output 與箭頭顯示，防止重疊
                html_content = f"""<div class="topo-container">"""
                
                # 1. 第一個節點前加入 Input Group (絕對定位於左側)
                if i == 0:
                     html_content += """
                        <div class="input-group">
                            <span class="input-label">Input</span>
                            <div class="input-arrow"></div>
                        </div>
                     """
                
                # [修改] 2. 其他節點前加入連接箭頭 (絕對定位於左側，指向目前節點)
                if i > 0:
                    html_content += '<div class="pre-connector-line"></div>'

                # 3. 節點本體
                html_content += f"""<div class="topo-node {node_states[i]}">{STATION_DATA[i]["id"]}</div>"""
                
                # 4. 最後一個節點後加入 Output Group (絕對定位於右側)
                if i == FIXED_N - 1:
                     html_content += """
                        <div class="output-group">
                            <div class="output-arrow"></div>
                            <span class="output-label">Output</span>
                        </div>
                     """

                html_content += "</div>" # 關閉容器 div

                st.markdown(html_content, unsafe_allow_html=True)
                
                if st.button("檢視", key=f"btn_node_{i}", type="primary" if st.session_state.selected_node_idx == i else "secondary", use_container_width=True):
                    st.session_state.selected_node_idx = i
                    st.rerun()

        if st.session_state.selected_node_idx is not None:
            idx = st.session_state.selected_node_idx
            if 0 <= idx < len(STATION_DATA):
                d_st = STATION_DATA[idx]
                st_carbon = d_st['power'] * carbon_factor
                st_loss = res['losses'][idx]
                
                st.markdown(f"""
                <div class="detail-card-highlight">
                <h5 style="margin-bottom: 15px; color: #fff;">🔍 {d_st["name"]} 詳細數據</h5>
                <div style="display: flex; justify-content: space-between; text-align: center; gap: 10px;">
                <div style="flex: 1;"><div style="font-size: 0.9rem; color: rgba(255,255,255,0.7); margin-bottom: 4px;">輸入量</div><div style="font-size: 1.5rem; font-weight: 700; color: #fff;">{res["rounded_inputs"][idx]}</div></div>
                <div style="flex: 1;"><div style="font-size: 0.9rem; color: rgba(255,255,255,0.7); margin-bottom: 4px;">功率 (kW)</div><div style="font-size: 1.5rem; font-weight: 700; color: #fff;">{d_st['power']}</div></div>
                <div style="flex: 1;"><div style="font-size: 0.9rem; color: rgba(255,255,255,0.7); margin-bottom: 4px;">成功率 p</div><div style="font-size: 1.5rem; font-weight: 700; color: #fff;">{d_st.get('p', 0.96)}</div></div>
                <div style="flex: 1;"><div style="font-size: 0.9rem; color: rgba(255,255,255,0.7); margin-bottom: 4px;">碳排放 (kg)</div><div style="font-size: 1.5rem; font-weight: 700; color: #fff;">{st_carbon:.3f}</div></div>
                <div style="flex: 1;"><div style="font-size: 0.9rem; color: rgba(255,255,255,0.7); margin-bottom: 4px;">耗損 (qty)</div><div style="font-size: 1.5rem; font-weight: 700; color: #ff6b6b;">{st_loss:.3f}</div></div>
                </div></div>""", unsafe_allow_html=True)

        k1, k2, k3, k4, k5 = st.columns([1,1,1,1,1], gap="large")
        # [修改] 將 "Rd" 中的 d 改為下標 (R<sub>d</sub>)
        with k1: st.markdown(f'<div class="kpi-box kpi-border-{sys_status} {sys_anim}"><div class="kpi-label">系統可靠度 <span style="font-family: \'Times New Roman\', serif; font-style: italic;">(R<sub>d</sub>)</span></div><div class="kpi-value">{res["reliability"]:.4f}</div></div>', unsafe_allow_html=True)
        # [修改] 將 "d" 改為正式論文格式：Times New Roman + 斜體 (變數)
        with k2: st.markdown(f'<div class="kpi-box"><div class="kpi-label">輸出量 <span style="font-family: \'Times New Roman\', serif; font-style: italic;">d</span></div><div class="kpi-value">{demand}</div></div>', unsafe_allow_html=True)
        # [修改] 將 "kW" 改為正式論文格式：Times New Roman (單位通常不斜體)
        with k3: st.markdown(f'<div class="kpi-box"><div class="kpi-label">總功率 (<span style="font-family: \'Times New Roman\', serif;">kW</span>)</div><div class="kpi-value">{res["total_energy"]:.3f}</div></div>', unsafe_allow_html=True)
        c_color = "green" if sys_carbon < 250 else "yellow" if sys_carbon < 300 else "red"
        
        with k4: st.markdown(f'<div class="kpi-box kpi-border-{c_color}"><div class="kpi-label">總碳排放 (kg)</div><div class="kpi-value">{res["carbon_emission"]:.3f}</div></div>', unsafe_allow_html=True)
        with k5: st.markdown(f'<div class="kpi-box kpi-border-red"><div class="kpi-label">總耗損 (qty)</div><div class="kpi-value">{res["total_loss"]:.3f}</div></div>', unsafe_allow_html=True)

        st.divider()
        st.header("📈 數據視覺化分析")
        stations = [d["name"] for d in STATION_DATA]
        c1, c2 = st.columns(2)
        with c1:
            fig1 = go.Figure(go.Bar(x=stations, y=res["losses"], marker_color='#60d3ff', name="耗損量"))
            fig1.update_layout(
                title=dict(text="各工作站耗損量", font=dict(size=22, color='black', weight='bold')),
                paper_bgcolor='white',
                plot_bgcolor='white',
                height=350,
                # 強制設定字體顏色為黑色，並放大字體
                xaxis=dict(title=dict(text='工作站', font=dict(size=18, color='black')), type='category', color='#000000', linecolor='#000000', tickcolor='#000000', gridcolor='#000000', tickfont=dict(size=16, color='#000000', family='Arial')),
                yaxis=dict(title=dict(text='耗損量', font=dict(size=18, color='black')), color='#000000', linecolor='#000000', tickcolor='#000000', gridcolor='#000000', tickfont=dict(size=16, color='#000000', family='Arial'))
            )
            st.plotly_chart(fig1, use_container_width=True)
        with c2:
            fig2 = go.Figure(go.Bar(x=stations, y=res["energies"], marker_color='#ffcf60', name="功率"))
            fig2.update_layout(
                title=dict(text="各工作站功率 (kW)", font=dict(size=22, color='black', weight='bold')),
                paper_bgcolor='white',
                plot_bgcolor='white',
                height=350,
                # 強制設定字體顏色為黑色，並放大字體
                xaxis=dict(title=dict(text='工作站', font=dict(size=18, color='black')), type='category', color='#000000', linecolor='#000000', tickcolor='#000000', gridcolor='#000000', tickfont=dict(size=16, color='#000000', family='Arial')),
                yaxis=dict(title=dict(text='功率 (kW)', font=dict(size=18, color='black')), color='#000000', linecolor='#000000', tickcolor='#000000', gridcolor='#000000', tickfont=dict(size=16, color='#000000', family='Arial'))
            )
            st.plotly_chart(fig2, use_container_width=True)

        st.markdown("### 📉 系統可靠度敏感度分析")
        
        # 定義臨界點
        crit_d = 2523

        # 修改：生成 X 軸數據點。除了原本的 500 間隔外，強制加入「臨界點」與「臨界點下一點 (crit_d + 1)」。
        raw_range = np.arange(500, 5501, 500)
        d_range_vals = np.sort(np.unique(np.concatenate((raw_range, [crit_d, crit_d + 1]))))

        y_vals = []
        for val in d_range_vals:
             y_vals.append(calculate_metrics(val, carbon_factor, STATION_DATA)['reliability'])

        crit_res = calculate_metrics(crit_d, carbon_factor, STATION_DATA)
        crit_y = crit_res['reliability']

        fig3 = go.Figure()
        
        fig3.add_trace(go.Scatter(
            x=d_range_vals, 
            y=y_vals,
            mode='lines+markers',
            name='可靠度曲線',
            line=dict(color='#3fe6ff', width=3),
            marker=dict(size=8, color='#3fe6ff')
        ))

        fig3.add_trace(go.Scatter(
            x=[crit_d], 
            y=[crit_y],
            mode='markers+text',
            # [修改] 將 Legend 中的 "d" 改為 Times New Roman + 斜體
            name=f'臨界點 (<span style="font-family: Times New Roman; font-style: italic;">d</span>={crit_d})',
            marker=dict(symbol='star', size=22, color='#ffd86b', line=dict(width=2, color='#ff0000')),
            text=['★ 臨界點'],
            textposition="top right",
            textfont=dict(color="black", size=14) # 強制文字標籤為黑色
        ))

        fig3.update_layout(
            title=dict(text="系統可靠度敏感度分析", font=dict(size=22, color='black', weight='bold')),
            # [修改] 將 X 軸標題中的 "d" 改為 Times New Roman + 斜體
            xaxis_title=dict(text="輸出量 (<span style='font-family: Times New Roman; font-style: italic;'>d</span>)", font=dict(size=18, color='black')), 
            yaxis_title=dict(text="系統可靠度", font=dict(size=18, color='black')),
            paper_bgcolor='white',
            plot_bgcolor='white',
            height=400,
            margin=dict(l=20, r=20, t=40, b=20),
            legend=dict(yanchor="top", y=0.99, xanchor="right", x=0.99, font=dict(color="black", size=14)),
            xaxis=dict(
                title_font=dict(size=18, color='#000000', family='Arial'),
                color='#000000',
                linecolor='#000000', linewidth=1,
                tickcolor='#000000', tickwidth=1,
                gridcolor='#000000', gridwidth=1,
                zeroline=False,
                tickfont=dict(size=16, color='#000000', family='Arial')
            ),
            yaxis=dict(
                title_font=dict(size=18, color='#000000', family='Arial'),
                color='#000000',
                linecolor='#000000', linewidth=1,
                tickcolor='#000000', tickwidth=1,
                gridcolor='#000000', gridwidth=1,
                zeroline=False,
                tickmode='linear',
                tick0=0,
                dtick=0.2,
                tickfont=dict(size=16, color='#000000', family='Arial')
            )
        )
        st.plotly_chart(fig3, use_container_width=True)

        st.header("📋 工作站狀態表")
        df_res = pd.DataFrame({
            "工作站": stations, 
            "輸入量": res["inputs"], 
            "取整輸入量": res["rounded_inputs"],
            "功率 (kW)": res["energies"], 
            "耗損 (qty)": res["losses"],
            "狀態數量": [len(d['capacities']) for d in STATION_DATA]
        })
        st.dataframe(df_res, use_container_width=True)

# --- TAB 2: Editor ---
with tab_editor:
    st.subheader("Excel 資料編輯器 (支援動態長度)")
    col_upload, col_settings = st.columns([2, 1])
    with col_upload:
        uploaded_file = st.file_uploader("📂 上傳 Excel", type=["xlsx"])

    if uploaded_file:
        file_id = f"{uploaded_file.name}_{uploaded_file.size}"
        if "processed_file_id" not in st.session_state or st.session_state.processed_file_id != file_id:
            try:
                new_df, new_scalars = load_data_from_excel_authority(uploaded_file)
                st.session_state.df_data = new_df
                if new_scalars: st.session_state.excel_authority = new_scalars
                st.session_state.processed_file_id = file_id
                st.session_state.last_uploaded_name = uploaded_file.name
                
                # [新增] 清除編輯器的快取狀態，強制顯示新上傳的 Excel 內容
                if "editor_table" in st.session_state:
                    del st.session_state["editor_table"]

                # 上傳後也強制保持在編輯頁面
                st.session_state.force_tab_index = 1
                st.rerun()
            except Exception as e:
                st.error(f"讀取失敗: {e}")

    df_source = st.session_state.df_data.copy()
    
    # [Callback] 當數據編輯器發生變更時，強制鎖定分頁 Index 為 1 (Editor)
    def maintain_editor_tab():
        st.session_state.force_tab_index = 1

    edited_df = st.data_editor(
        df_source[['Station', 'p', 'power', 'capacities', 'probs']],
        num_rows="dynamic",
        use_container_width=True,
        key="editor_table", 
        on_change=maintain_editor_tab,  # 綁定 Callback
        column_config={
            "Station": st.column_config.NumberColumn("站號", min_value=1, step=1, required=True),
            "p": st.column_config.NumberColumn("成功率 p", min_value=0.0001, max_value=1.0),
            "power": st.column_config.NumberColumn("功率 (kW)"),
            "capacities": st.column_config.TextColumn("產能列表 (List)", help="例如 [0, 100, 200]"),
            "probs": st.column_config.TextColumn("機率列表 (List)", help="例如 [0.1, 0.4, 0.5]")
        }
    )

    col_reset, col_save = st.columns([1, 1])
    with col_reset:
        if st.button("🔄 重置為預設資料", use_container_width=True):
            st.session_state.df_data = get_default_data()
            st.session_state.force_tab_index = 1  # 重置後還是留在編輯頁
            st.rerun()

    with col_save:
        if st.button("💾 儲存並更新", use_container_width=True):
            try:
                # 1. 驗證
                validated_rows = []
                for _, row in edited_df.iterrows():
                    caps = parse_list_from_string(row['capacities'])
                    probs = parse_list_from_string(row['probs'])
                    
                    if not isinstance(caps, list) or not isinstance(probs, list):
                        st.error(f"站號 {row['Station']}: 列表格式錯誤"); st.stop()
                    if len(caps) != len(probs):
                        st.error(f"站號 {row['Station']}: 產能({len(caps)})與機率({len(probs)})長度不符"); st.stop()
                    if len(caps) > 1 and not all(x < y for x, y in zip(caps, caps[1:])):
                        st.error(f"站號 {row['Station']}: 產能列表必須嚴格遞增"); st.stop()
                    if probs and not math.isclose(sum(probs), 1.0, abs_tol=1e-2):
                        st.warning(f"注意: 站號 {row['Station']} 機率和不為 1 ({sum(probs):.3f})")
                    
                    validated_rows.append((row, caps, probs))

                # 2. 寫入
                long_rows = []
                for row, caps, probs in validated_rows:
                    for i in range(len(caps)):
                        long_rows.append({
                            "Station": int(row['Station']),
                            "Machine": 1,
                            "Success_Rate": row['p'],
                            "Power(kW)加工功率": row['power'],
                            "Capacity": caps[i],
                            "Capacity_Prob": probs[i]
                        })
                
                df_long = pd.DataFrame(long_rows)
                
                for i in range(6, 14): df_long[f"Unnamed: {i}"] = np.nan
                while len(df_long) < 5:
                    df_long = pd.concat([df_long, pd.DataFrame([np.nan]*df_long.shape[1], columns=df_long.columns)], ignore_index=True)
                
                auth_data = st.session_state.get("excel_authority")
                if auth_data is None: auth_data = {"d": 2500, "carbon_factor": 0.474}
                
                curr_d = auth_data.get("d", 2500)
                curr_c = auth_data.get("carbon_factor", 0.474)
                
                df_long.iloc[1, 7] = "d="
                df_long.iloc[1, 8] = curr_d
                df_long.iloc[2, 7] = "CO2="
                df_long.iloc[2, 8] = curr_c
                
                save_name = st.session_state.get("last_uploaded_name", "新版簡單_modified.xlsx")
                save_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), save_name)
                df_long.to_excel(save_path, index=False)
                
                # 3. 更新與跳轉
                st.session_state.df_data = edited_df
                st.session_state.excel_authority = {"d": curr_d, "carbon_factor": curr_c}
                st.session_state.show_success_modal = True
                
                # [關鍵] 儲存成功：強制跳轉回 Dashboard (Index 0)
                st.session_state.force_tab_index = 0
                st.rerun()

            except Exception as e:
                st.error(f"儲存失敗: {e}")
#在終端機輸入：python -m streamlit run "C:\Users\user\OneDrive\桌面\dashboard.py"