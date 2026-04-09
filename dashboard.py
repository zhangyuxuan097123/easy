import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import math
import itertools
import io
import time
from datetime import datetime
import streamlit.components.v1 as components
import google.generativeai as genai
import json
import html as html_lib
import re
import os

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# --- 輔助函式：產生 a 的斜體加下標字元 ---
def get_a_subscript(val):
    sub_map = str.maketrans("0123456789", "₀₁₂₃₄₅₆₇₈₉")
    return f"𝑎{str(val).translate(sub_map)}"

# --- 0. 基本設定 ---
st.set_page_config(
    page_title="基於生成式AI與網路可靠度於製造系統戰情儀表設計",
    page_icon="🏭", layout="wide", initial_sidebar_state="expanded"
)

# ── 內建 Groq API Keys（可填入多組，系統自動輪替）──
GOOGLE_API_KEYS = [
    k for k in [
        os.getenv("GOOGLE_API_KEY_1"),
        os.getenv("GOOGLE_API_KEY_2"),
    ] if k
]
GEMINI_MODEL = "gemini-2.5-flash-lite"
DEFAULT_EXCEL_PATH = "!!!最新版簡單!!!.xlsx"

# ============================================================
# --- 1. 全局 CSS ---
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');

.stApp { background: #23395B !important; color: #e6eef6; font-family: 'Inter', sans-serif; }
.block-container {
    padding-top: 2rem !important; padding-bottom: 2rem !important;
    padding-left: 7rem !important; padding-right: 10rem !important;
    max-width: 100% !important; overflow: visible !important;
}

section[data-testid="stSidebar"] { background-color: #0b1626 !important; border-right: 1px solid rgba(255, 255, 255, 0.1); }
section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] .stMarkdown p { color: #e6eef6 !important; font-weight: 500 !important; }
section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 { color: #ffffff !important; }

[data-testid='stFileUploader'] label[data-testid='stWidgetLabel'] { color: #FFFFFF !important; font-size: 1.2rem !important; font-weight: 700 !important; }
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
@keyframes kpiShake { 0% { transform: translateX(0); } 25% { transform: translateX(-5px) rotate(-1deg); } 50% { transform: translateX(5px) rotate(1deg); } 75% { transform: translateX(-5px) rotate(-1deg); } 100% { transform: translateX(0); } }
.kpi-shake { animation: kpiShake 0.5s infinite; border-color: #ff6b6b !important; }

.line-green { background: #4cd37a !important; box-shadow: 0 0 8px rgba(76, 211, 122, 0.5); }
.line-green .arrow-head { border-left-color: #4cd37a !important; }
.line-yellow { background: #ffd86b !important; }
.line-yellow .arrow-head { border-left-color: #ffd86b !important; }
.line-red { background: #ff6b6b !important; }
.line-red .arrow-head { border-left-color: #ff6b6b !important; }
.line-fail { background: #8B0000 !important; }
.line-fail .arrow-head { border-left-color: #8B0000 !important; }
@keyframes linePulse { 0% { box-shadow: 0 0 5px rgba(255, 216, 107, 0.4); } 50% { box-shadow: 0 0 20px rgba(255, 216, 107, 0.9); } 100% { box-shadow: 0 0 5px rgba(255, 216, 107, 0.4); } }
.line-pulse { animation: linePulse 1.5s infinite; z-index: 5; }
@keyframes lineBlink { 0% { opacity: 1; } 50% { opacity: 0.4; } 100% { opacity: 1; } }
.line-blink { animation: lineBlink 0.6s infinite; z-index: 5; }

.topo-node { width: 55px; height: 55px; border-radius: 50%; display: flex; align-items: center; justify-content: center; color: #fff; border: 3px solid rgba(255,255,255,0.3); box-shadow: 0 4px 10px rgba(0,0,0,0.3); transition: all 0.3s ease; background: #23395B; margin: 0; z-index: 2; flex-shrink: 0; }
.topo-node-content { display: inline-flex; align-items: baseline; justify-content: center; }
.topo-node i { font-family: 'Times New Roman', serif; font-size: 1.6rem; font-weight: 700; font-style: italic; }
.topo-node sub { font-size: 0.55rem; font-weight: 900; margin-left: 2px; }

.arc-label { position: absolute; top: -25px; left: 50%; transform: translateX(-50%); color: #fff; font-size: 1.2rem; font-weight: bold; text-shadow: 0 2px 4px rgba(0,0,0,0.8); white-space: nowrap; z-index: 3; }
.arc-label i { font-family: 'Times New Roman', serif; font-style: italic; }
.arc-label sub { font-size: 0.8rem; }

.detail-card-highlight { border: 2px solid #3fe6ff; background: rgba(63, 230, 255, 0.1); padding: 15px; border-radius: 10px; margin-top: 10px; margin-bottom: 20px; }

.chat-container {
    background: rgba(11, 22, 38, 0.95);
    border: 1.5px solid rgba(63, 230, 255, 0.35);
    border-radius: 14px;
    padding: 20px 20px 10px 20px;
    margin-bottom: 16px;
    max-height: 520px;
    overflow-y: auto;
    box-shadow: 0 8px 28px rgba(0,0,0,0.5);
}
.chat-bubble-user {
    background: linear-gradient(135deg, #1e4a6e, #2a5f8a);
    border-radius: 16px 16px 4px 16px;
    padding: 10px 16px;
    margin: 6px 0 6px 60px;
    color: #e6f4ff;
    font-size: 0.95rem;
    line-height: 1.5;
    box-shadow: 0 2px 8px rgba(0,0,0,0.3);
}
.chat-bubble-ai {
    background: linear-gradient(135deg, #0d2137, #122c40);
    border: 1px solid rgba(63, 230, 255, 0.2);
    border-radius: 16px 16px 16px 4px;
    padding: 10px 16px;
    margin: 6px 60px 6px 0;
    color: #d4f0ff;
    font-size: 0.95rem;
    line-height: 1.6;
    box-shadow: 0 2px 8px rgba(0,0,0,0.3);
}
.chat-label-user { text-align: right; font-size: 0.78rem; color: #7abadb; margin-bottom: 2px; }
.chat-label-ai { text-align: left; font-size: 0.78rem; color: #3fe6ff; margin-bottom: 2px; }
.chat-typing {
    display: inline-block;
    padding: 8px 14px;
    background: rgba(63, 230, 255, 0.08);
    border-radius: 12px;
    color: #3fe6ff;
    font-size: 0.9rem;
    font-style: italic;
    animation: typingPulse 1s infinite;
}
@keyframes typingPulse { 0%,100% { opacity: 0.6; } 50% { opacity: 1; } }

[data-testid="stPlotlyChart"] { background-color: #ffffff !important; border-radius: 18px; box-shadow: 0 8px 24px rgba(0,0,0,0.20); padding: 10px; margin-bottom: 20px; }
.success-modal-overlay { position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background: rgba(0,0,0,0.6); display: flex; justify-content: center; align-items: center; backdrop-filter: blur(4px); animation: fadeOutContainer 2.5s forwards; z-index: 999999; }
.success-modal-content { background: rgba(20, 24, 30, 0.95); border: 2px solid #4cd37a; border-radius: 16px; padding: 40px 60px; text-align: center; box-shadow: 0 0 40px rgba(76, 211, 122, 0.4); }
@keyframes fadeOutContainer { 0% { opacity: 1; pointer-events: auto; } 70% { opacity: 1; pointer-events: auto; } 100% { opacity: 0; pointer-events: none; z-index: -1; } }
button[data-baseweb="tab"][aria-selected="false"] { color: #FFFFFF !important; }
div[data-testid="column"] { overflow: visible !important; }
div[data-testid="stHorizontalBlock"] { overflow: visible !important; }
</style>
""", unsafe_allow_html=True)

# --- 2. 狀態初始化 ---
if "show_success_modal" not in st.session_state:
    st.session_state.show_success_modal = False
if "ai_advice" not in st.session_state:
    st.session_state.ai_advice = None
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

if st.session_state.show_success_modal:
    st.balloons()
    st.toast("✅ 資料已儲存並同步更新！", icon="💾")
    st.markdown('<div class="success-modal-overlay"><div class="success-modal-content"><div style="font-size: 60px; margin-bottom: 10px;">✅</div><h2 style="color: #4cd37a; margin: 0;">儲存成功</h2><p style="color: #ddd; margin-top: 10px;">Dashboard 已完成同步更新</p></div></div>', unsafe_allow_html=True)
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
                "Station": int(name), "p": float(first_row["Success_Rate"]),
                "power": float(first_row["Power(kW)加工功率"]),
                "k": float(first_row.get("k", 1.0) if pd.notna(first_row.get("k")) else 1.0),
                "capacities": str(caps), "probs": str(probs)
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

auth_data = st.session_state.get("excel_authority", {"d": 10000, "carbon_factor": 0.474, "tb": 1.0})
if "sim_d" not in st.session_state: st.session_state.sim_d = int(auth_data.get("d", 10000))
if "sim_cf" not in st.session_state: st.session_state.sim_cf = float(auth_data.get("carbon_factor", 0.474))
if "sim_tb" not in st.session_state: st.session_state.sim_tb = float(auth_data.get("tb", 1.0))

if "pending_ai_updates" in st.session_state:
    updates = st.session_state.pending_ai_updates
    st.session_state.sim_tb = updates.get("tb", st.session_state.sim_tb)
    st.session_state.sim_cf = updates.get("cf", st.session_state.sim_cf)
    if updates.get("d") is not None:
        st.session_state.sim_d = int(updates["d"])
    del st.session_state.pending_ai_updates

# --- 核心運算 ---
def calculate_metrics(demand, carbon_factor, _station_data, tb_value):
    n = len(_station_data)
    if n == 0: return {}
    mu = 1.0
    pi_list = []
    for d in _station_data:
        pi = d['p'] * math.exp(-d.get('k', 1.0) * (tb_value - mu)**2)
        pi_list.append(pi)
    product_pi = 1.0
    for pi_val in pi_list: product_pi *= pi_val
    total_input = demand / product_pi
    inputs = []
    current_input = total_input
    for i in range(n):
        inputs.append(current_input)
        current_input *= pi_list[i]
    rounded_inputs = [math.ceil(x) for x in inputs]
    selected_caps, energies, carbons = [], [], []
    for i in range(n):
        req_input = rounded_inputs[i]
        caps = _station_data[i]["capacities"]
        base_power = _station_data[i]["power"]
        max_cap = max(caps) if caps else 0
        sel_cap = next((c for c in caps if c >= req_input), max_cap)
        selected_caps.append(sel_cap)
        cap_ratio = (sel_cap / max_cap) if max_cap > 0 else 0
        power_ratio = cap_ratio
        if i == 4 and n == 5:
            st4_caps = _station_data[3]["capacities"]
            st4_max = max(st4_caps) if st4_caps else 1
            st4_req = rounded_inputs[3]
            st4_sel = next((c for c in st4_caps if c >= st4_req), st4_max)
            power_ratio = (st4_sel / st4_max) if st4_max > 0 else cap_ratio
        actual_power = base_power * power_ratio
        actual_carbon = actual_power * cap_ratio * carbon_factor
        energies.append(actual_power)
        carbons.append(actual_carbon)
    losses = [inputs[i] * (1 - pi_list[i]) for i in range(n)]
    total_probability = 0
    indices_ranges = [range(len(d["capacities"])) for d in _station_data]
    limit_count = 0
    for state_indices in itertools.product(*indices_ranges):
        limit_count += 1
        if limit_count > 1000000: break
        current_prob, valid = 1.0, True
        for i, state_idx in enumerate(state_indices):
            cap = _station_data[i]["capacities"][state_idx]
            if cap < rounded_inputs[i]: valid = False; break
            current_prob *= _station_data[i]["probs"][state_idx]
        if valid: total_probability += current_prob
    return {
        "pi_list": pi_list, "inputs": inputs, "rounded_inputs": rounded_inputs,
        "selected_caps": selected_caps, "energies": energies, "carbons": carbons,
        "losses": losses, "total_loss": sum(losses), "total_energy": sum(energies),
        "carbon_emission": sum(carbons), "reliability": total_probability,
    }

# ============================================================
# 🤖  AI 函式：使用 Groq API
# ============================================================
def build_combined_prompt(query: str, current_params: dict, current_metrics: dict, chat_history: list) -> str:
    history_text = ""
    for turn in chat_history[-6:]:
        history_text += f"使用者：{turn['user']}\nAI 戰情助理：{turn['ai']}\n"

    return f"""你是台灣智慧工廠的 AI 戰情助理，專精於飲料瓶生產線管理。

【語言規定】
- 必須使用繁體中文（zh-TW），絕對不能出現簡體字。
- 語氣：專業、直接、有建設性，避免過度客套。
- 名稱：請一律自稱「AI 戰情助理」。

【當前系統狀態】
- 輸出量 (𝑑)：{current_params.get('d', 'N/A')}
- 瓶胚厚度參數 (𝑡ᵦ)：{current_params.get('tb', 'N/A')}
- CO₂ 係數 (kg/kWh)：{current_params.get('cf', 'N/A')}
- 系統可靠度 (𝑅ᵈ)：{current_metrics.get('reliability', 0):.4f}
- 總碳排放 (kg)：{current_metrics.get('carbon_emission', 0):.2f}
- 總耗損：{current_metrics.get('total_loss', 0):.2f}
- 各站耗損：{dict(zip(['吹瓶站','充填站','套標站','包裝站','疊棧站'], [round(x,2) for x in current_metrics.get('losses', [])]))}
- 各站動態功率 (kW)：{dict(zip(['吹瓶站','充填站','套標站','包裝站','疊棧站'], [round(x,2) for x in current_metrics.get('energies', [])]))}

【可靠度警戒閾值】
- 🟢 正常：𝑅ᵈ > 0.95
- 🟡 警告：0.90 ≤ 𝑅ᵈ ≤ 0.95
- 🔴 危險：𝑅ᵈ < 0.90

【碳排放警戒閾值】
- 🟢 正常：≤ 70 kg
- 🟡 警告：71 ~ 100 kg
- 🔴 危險：> 100 kg

【對話歷史（近期）】
{history_text if history_text else "（目前為對話開始）"}

【使用者最新訊息】
{query}
【限制規則】
- 若使用者詢問與製造系統、生產線、可靠度、碳排放、工作站、參數調整完全無關的話題（例如：天氣、食物、娛樂、政治等），
  請在 "reply" 欄位回覆：「這個問題目前不在製造系統戰情儀表板的服務範圍內。我可以協助分析生產線狀態、可靠度、碳排放或參數調整等內容，請重新提問。」
  並將 d、tb、cf 全部設為 null。
- 只有與工廠生產、系統參數、可靠度分析相關的問題才進行正式分析。

【任務】
請完成以下兩件事，並嚴格以純 JSON 格式輸出，不含任何 Markdown 標籤或說明文字：
1. 從使用者訊息中提取參數（若未提及則設為 null）：
   - "d"：輸出量、產能、產量、需求量（整數）
   - "tb"：瓶胚厚度、厚度參數（浮點數）
   - "cf"：CO₂ 係數、碳排係數（浮點數）
2. 根據使用者問題的性質，自動選擇最適合的回覆格式（150～300 字）：

   - 若使用者詢問「狀態/是否正常/達標/系統如何」類問題：
     使用【現況分析】【風險評估】【建議措施】三段格式

   - 若使用者詢問「哪個最高/最低/比較/排名」類問題：
     直接點出答案，再補充原因與影響，不需強制分段

   - 若使用者詢問「如果改成XX會怎樣/假設/模擬」類問題：
     先說明模擬結果變化，再評估是否有風險，最後給建議

   - 若使用者詢問「建議/怎麼改善/如何優化」類問題：
     直接給出具體可執行建議，條列 2～3 點

   - 若使用者詢問「為什麼/原因/怎麼造成」類問題：
     用淺顯易懂的方式解釋原因，並舉例說明影響

   - 若使用者詢問「趨勢/預測/未來」類問題：
     根據目前數值推估未來走向，說明樂觀與悲觀情境

   - 若使用者詢問「定義/什麼是/解釋」類問題：
     簡短解釋名詞定義，並結合目前系統數值舉例說明

   - 若使用者詢問「最佳值/最適參數/怎麼設定」類問題：
     給出建議的參數範圍，並說明為何這個範圍最佳

   - 若使用者進行閒聊或打招呼：
     友善回應並引導回儀表板相關話題，50 字以內

   - 若為簡單問答（查詢單一數值）：
     簡短直接回答，不需要分段格式，100 字以內

   - 若問題涉及多個面向（例如同時問可靠度與碳排）：
     分別針對每個面向簡短說明，最後給一個綜合結論

輸出格式範例：
{{"d": null, "tb": 0.9, "cf": null, "reply": "【現況分析】...【風險評估】...【建議措施】...", "chart": "loss"}}

【chart 欄位規則】
- "loss"：當使用者問耗損、哪站損失最高時
- "energy"：當使用者問功率、用電、能耗時  
- "carbon"：當使用者問碳排放、碳排哪站最高時
- "reliability"：當使用者問可靠度敏感度、產量影響時
- null：一般問答不需要圖表時
"""

def get_gemini_client():
    last_err = None
    for key_index, key in enumerate(GOOGLE_API_KEYS):
        try:
            genai.configure(api_key=key)
            
            available_models = [
                m.name for m in genai.list_models() 
                if 'generateContent' in m.supported_generation_methods
            ]
            
            if not available_models:
                continue

            # 優先順序：2.5-flash-lite > 2.5-flash > 2.5-pro > 其他
            preferred = [
                "gemini-2.5-flash-lite",
                "gemini-2.5-flash",
                "gemini-2.5-pro",
            ]
            
            target_model_name = None
            for pref in preferred:
                for m_name in available_models:
                    if pref in m_name:
                        target_model_name = m_name
                        break
                if target_model_name:
                    break
            
            if not target_model_name:
                target_model_name = available_models[0]

            model = genai.GenerativeModel(target_model_name)
            return model
            
        except Exception as e:
            error_msg = str(e)
            last_err = e
            if "429" in error_msg or "Quota" in error_msg:
                st.warning(f"⚠️ 第 {key_index + 1} 把金鑰觸發流量限制，嘗試切換下一把...")
                time.sleep(2)
                continue
            else:
                continue
                
    raise Exception(f"所有金鑰目前皆達到流量上限或失效。請等待後再試！最後錯誤：{last_err}")

def call_ai_single(client, query: str, current_params: dict, current_metrics: dict, chat_history: list):
    """單次 Gemini API 呼叫：支援 JSON 與純文字格式的暴力解析"""
    import re # 確保使用正則表達式
    
    prompt = build_combined_prompt(query, current_params, current_metrics, chat_history)
    generation_config = genai.types.GenerationConfig(
        max_output_tokens=2048,
        temperature=0.2, # 降低溫度，讓 AI 不要亂發揮
    )
    
    try:
        resp = client.generate_content(prompt, generation_config=generation_config)
        raw = resp.text.strip()
        raw = raw.replace("```json", "").replace("```", "").strip()
    except Exception as e:
        return current_params.copy(), f"API 呼叫發生錯誤：{str(e)}", None

    # 第一關：先嘗試標準的 JSON 解析 (如果 AI 乖乖聽話的話)
    try:
        data = json.loads(raw)
        extracted = {
            "d":  int(data["d"])    if data.get("d") is not None else current_params.get("d"),
            "tb": float(data["tb"]) if data.get("tb") is not None else current_params.get("tb"),
            "cf": float(data["cf"]) if data.get("cf") is not None else current_params.get("cf"),
        }
        reply = data.get("reply", "（無回覆內容）")
        chart = data.get("chart", None)
        return extracted, reply, chart
        
    except Exception:
        # 第二關：啟動「暴力萃取」，專治不守規矩的 AI
        
        # 1. 強硬抓取數字參數 (找 d: 10000 這種格式)
        d_match = re.search(r'["\']?d["\']?\s*:\s*([\d\.]+)', raw)
        tb_match = re.search(r'["\']?tb["\']?\s*:\s*([\d\.]+)', raw)
        cf_match = re.search(r'["\']?cf["\']?\s*:\s*([\d\.]+)', raw)
        
        extracted = {
            "d": int(float(d_match.group(1))) if d_match else current_params.get("d"),
            "tb": float(tb_match.group(1)) if tb_match else current_params.get("tb"),
            "cf": float(cf_match.group(1)) if cf_match else current_params.get("cf")
        }
        
        # 2. 強硬抓取對話內容 (找出 reply: 後面的所有中文字)
        # 這段語法會忽略前後引號，直接把回覆本體挖出來
        reply_match = re.search(r'["\']?reply["\']?\s*:\s*["\']?(.*?)(?:,?\s*["\']?chart["\']?\s*:|$)', raw, re.DOTALL | re.IGNORECASE)
        
        if reply_match:
            reply_text = reply_match.group(1).strip()
            # 拔除尾巴可能殘留的 } 括號或引號
            reply_text = re.sub(r'["\'\}]+$', '', reply_text) 
            # 把 \\n 換成真正的換行
            reply_text = reply_text.replace('\\n', '\n').replace('\n', '<br>')
        else:
            # 最壞情況：連 reply 都找不到，就把英文字段全部刪除，只留中文
            reply_text = re.sub(r'["\']?[dtbcfchart]+["\']?\s*:\s*[A-Za-z0-9\.\-\_]*\s*,?', '', raw, flags=re.IGNORECASE)
            reply_text = re.sub(r'[{}]', '', reply_text).strip()
            
        # 3. 嘗試抓取 chart 標籤
        chart_match = re.search(r'["\']?chart["\']?\s*:\s*["\']?([a-zA-Z]+)', raw)
        chart_val = chart_match.group(1) if chart_match else None
        if chart_val == "null": chart_val = None
            
        return extracted, reply_text, chart_val


# ============================================================
# --- 4. UI 顯示 ---
# ============================================================
st.markdown('<div style="padding:14px 10px; border-radius:10px; background: linear-gradient(90deg, rgba(6,21,39,0.6), rgba(8,30,46,0.35)); box-shadow:0 6px 18px rgba(2,8,23,0.6); margin-bottom:12px;"><h1 style="margin:0;color:#e6f7ff">🏭 基於生成式AI與網路可靠度於製造系統戰情儀表設計</h1></div>', unsafe_allow_html=True)

tab_dashboard, tab_chat, tab_editor = st.tabs([
    "📊 戰情儀表板",
    "🤖 AI 戰情助理",
    "📝 資料管理",
])

if st.session_state.force_tab_index is not None:
    components.html(f"<script>window.parent.document.querySelectorAll('button[data-baseweb=\"tab\"]')[{st.session_state.force_tab_index}].click();</script>", height=0)
    st.session_state.force_tab_index = None

# ============================================================
# TAB 1: 戰情儀表板
# ============================================================
with tab_dashboard:
    try:
        source_df = st.session_state.df_data
        STATION_DATA = [{
            "name": str(int(row['Station'])), "id": int(row['Station']),
            "capacities": parse_list_from_string(row['capacities']),
            "probs": parse_list_from_string(row['probs']), "p": row['p'],
            "power": row['power'], "k": row.get('k', 1.0)
        } for _, row in source_df.iterrows()]
        FIXED_N = len(STATION_DATA)
    except:
        STATION_DATA, FIXED_N = [], 0

    if FIXED_N == 0:
        st.warning("無有效工作站資料")
    else:
        with st.sidebar:
            st.markdown('<div style="padding:12px 10px; background-color: rgba(255, 255, 255, 0.08); border-radius: 8px; margin-bottom: 15px;"><h3 style="margin:0; color:#ffffff">系統參數面板</h3></div>', unsafe_allow_html=True)
            demand = st.number_input("輸出量 (𝑑)", min_value=1, step=100, key="sim_d")
            carbon_factor = st.number_input("CO₂ 係數 (kg/kWh)", min_value=0.001, step=0.001, format="%.3f", key="sim_cf")
            tb_val = st.number_input("厚度參數 ($t_b$)", step=0.01, format="%.2f", key="sim_tb")
            st.divider()
            res = calculate_metrics(demand, carbon_factor, STATION_DATA, tb_val)
            rel_val = res.get('reliability', 0)
            sys_status_sidebar = "green" if rel_val > 0.95 else "yellow" if rel_val >= 0.9 else "red"
            status_colors = {"green": "#4cd37a", "yellow": "#ffd86b", "red": "#ff6b6b"}
            status_bgs = {"green": "rgba(76, 211, 122, 0.05)", "yellow": "rgba(255, 216, 107, 0.05)", "red": "rgba(255, 107, 107, 0.05)"}
            status_texts = {"green": "可靠度正常", "yellow": "可靠度警告", "red": "可靠度過低"}
            st.markdown(f'<div style="background-color: {status_bgs[sys_status_sidebar]}; padding: 12px; border-radius: 8px; text-align: center;"><span style="color: {status_colors[sys_status_sidebar]}; font-weight: 700; font-size: 16px;">{status_texts[sys_status_sidebar]} : {rel_val:.4f}</span></div>', unsafe_allow_html=True)
            st.markdown('<div style="padding:15px; background-color: rgba(255,255,255,0.05); border-radius: 8px; margin-top: 25px; border: 1px solid rgba(255,255,255,0.1);"><h4 style="margin-top:0; color:#e6eef6; font-size: 16px; border-bottom: 1px solid rgba(255,255,255,0.2); padding-bottom: 8px;">🚦 狀態燈號閾值說明</h4><div style="font-size: 0.9rem; color: #ddd; margin-top: 10px;"><div style="margin-bottom: 8px;"><b>系統可靠度</b></div><div style="display: flex; justify-content: space-between; margin-bottom: 4px;"><span style="color:#4cd37a;">🟢 正常</span> <span>＞ 0.95</span></div><div style="display: flex; justify-content: space-between; margin-bottom: 4px;"><span style="color:#ffd86b;">🟡 警告</span> <span>0.90 ~ 0.95</span></div><div style="display: flex; justify-content: space-between; margin-bottom: 16px;"><span style="color:#ff6b6b;">🔴 危險</span> <span>＜ 0.90</span></div><div style="margin-bottom: 8px;"><b>總碳排放 (kg)</b></div><div style="display: flex; justify-content: space-between; margin-bottom: 4px;"><span style="color:#4cd37a;">🟢 正常</span> <span>0 ~ 70</span></div><div style="display: flex; justify-content: space-between; margin-bottom: 4px;"><span style="color:#ffd86b;">🟡 警告</span> <span>71 ~ 100</span></div><div style="display: flex; justify-content: space-between;"><span style="color:#ff6b6b;">🔴 危險</span> <span>＞ 100</span></div></div></div>', unsafe_allow_html=True)

        sys_reliability = res.get('reliability', 0)
        sys_carbon = res.get('carbon_emission', 0)
        sys_status = "green" if sys_reliability > 0.95 else "yellow" if sys_reliability >= 0.9 else "red"
        sys_anim = "kpi-pulse" if sys_status == "yellow" else "kpi-shake" if sys_status == "red" else ""
        sys_anim_line = "line-pulse" if sys_status == "yellow" else "line-blink" if sys_status == "red" else ""
        line_states = []
        for i, station in enumerate(STATION_DATA):
            is_failed = res["rounded_inputs"][i] > (max(station["capacities"]) if station["capacities"] else 0)
            line_states.append("line-fail line-blink" if is_failed else f"line-{sys_status} {sys_anim_line}")

        st.markdown("### 🕸️ 生產線即時拓樸監控")
        if "selected_node_idx" not in st.session_state: st.session_state.selected_node_idx = None
        station_labels = ["🔽 吹瓶站", "🔽 充填站", "🔽 套標站", "🔽 包裝站", "🔽 疊棧站"]

        btn_cols = st.columns(FIXED_N)
        for i, col in enumerate(btn_cols):
            with col:
                is_first = (i == 0)
                is_last = (i == FIXED_N - 1)
                if FIXED_N == 1:
                    line_left, line_width, node_left_pos = "0", "100%", "0"
                else:
                    if is_first: line_left, line_width, node_left_pos = "0", "calc(100% + 0.5rem)", "0"
                    elif is_last: line_left, line_width, node_left_pos = "-0.5rem", "calc(100% + 0.5rem)", "-0.5rem"
                    else: line_left, line_width, node_left_pos = "-0.5rem", "calc(100% + 1rem)", "-0.5rem"
                node_id = STATION_DATA[i]["id"]
                prev_node_id = "0" if is_first else STATION_DATA[i-1]["id"]
                l_class = line_states[i]
                html = '<div style="position: relative; width: 100%; height: 100px; display: flex; justify-content: center; align-items: center; z-index: 0;">'
                html += f'<div class="{l_class}" style="position: absolute; left: {line_left}; width: {line_width}; height: 3px; background: #ccc; top: 50%; transform: translateY(-50%); z-index: 1;"><div class="arrow-head" style="position: absolute; right: 28px; top: -4.5px; border-top: 6px solid transparent; border-bottom: 6px solid transparent; border-left: 10px solid #ccc;"></div></div>'
                html += f'<div class="arc-label" style="position: absolute; top: 15px; left: calc(50% + 0.5rem); transform: translateX(-50%); z-index: 3;"><i>a</i><sub>{node_id}</sub></div>'
                html += f'<div style="position: absolute; left: {node_left_pos}; top: 50%; transform: translate(-50%, -50%); z-index: 4; display: flex; align-items: center;">'
                if is_first:
                    html += '<div style="position: absolute; right: 100%; display: flex; align-items: center;"><span style="margin-right: 8px; color: #fff; font-weight: 700; font-size: 16px; text-shadow: 0 2px 4px rgba(0,0,0,0.8);">Input</span><div style="width: 30px; height: 2px; background: #ccc; position: relative; margin-right: 5px;"><div style="position: absolute; right: 0; top: -4px; border-top: 5px solid transparent; border-bottom: 5px solid transparent; border-left: 8px solid #ccc;"></div></div></div>'
                    html += '<div class="topo-node" style="background: #111111; border-color: #555;"><div class="topo-node-content"></div></div></div>'
                else:
                    html += f'<div class="topo-node" style="background: #23395B; border-color: rgba(255,255,255,0.4);"><div class="topo-node-content"><i>n</i><sub>{prev_node_id}</sub></div></div></div>'
                if is_last:
                    html += f'<div style="position: absolute; left: calc(100% + 0.5rem); top: 50%; transform: translate(-50%, -50%); z-index: 4; display: flex; align-items: center;">'
                    html += '<div class="topo-node" style="background: #111111; border-color: #555;"><div class="topo-node-content"></div></div>'
                    html += '<div style="position: absolute; left: 100%; display: flex; align-items: center;"><div style="width: 30px; height: 2px; background: #ccc; position: relative;"><div style="position: absolute; right: 0; top: -4px; border-top: 5px solid transparent; border-bottom: 5px solid transparent; border-left: 8px solid #ccc;"></div></div><span style="margin-left: 10px; color: #fff; font-weight: 700; font-size: 16px; text-shadow: 0 2px 4px rgba(0,0,0,0.8);">Output</span></div>'
                    html += '</div>'
                html += '</div>'
                st.markdown(html, unsafe_allow_html=True)
                label = station_labels[i] if i < len(station_labels) else f"🔽 工作站 {STATION_DATA[i]['id']}"
                if st.button(label, key=f"n_btn_{i}", type="primary" if st.session_state.selected_node_idx == i else "secondary", use_container_width=True):
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
                st.markdown(f'<div class="detail-card-highlight"><h5 style="margin-bottom: 15px; color: #fff;">🔍 {st_detail_name} 詳細數據</h5><div style="display: flex; justify-content: space-between; text-align: center; gap: 10px;"><div style="flex: 1;"><div style="font-size: 0.9rem; color: rgba(255,255,255,0.7); margin-bottom: 4px;">投入量</div><div style="font-size: 1.5rem; font-weight: 700; color: #fff;">{res["rounded_inputs"][idx]}</div></div><div style="flex: 1;"><div style="font-size: 0.9rem; color: rgba(255,255,255,0.7); margin-bottom: 4px;">動態功率 (kW)</div><div style="font-size: 1.5rem; font-weight: 700; color: #fff;">{st_power:.2f}</div></div><div style="flex: 1;"><div style="font-size: 0.9rem; color: rgba(255,255,255,0.7); margin-bottom: 4px;">參數 (𝑘)</div><div style="font-size: 1.5rem; font-weight: 700; color: #fff;">{d_st.get("k", 1.0)}</div></div><div style="flex: 1;"><div style="font-size: 0.9rem; color: rgba(255,255,255,0.7); margin-bottom: 4px;">品質調整後成功率</div><div style="font-size: 1.5rem; font-weight: 700; color: #ffffff;">{res["pi_list"][idx]:.4f}</div></div><div style="flex: 1;"><div style="font-size: 0.9rem; color: rgba(255,255,255,0.7); margin-bottom: 4px;">碳排放 (kg)</div><div style="font-size: 1.5rem; font-weight: 700; color: #fff;">{st_carbon:.3f}</div></div><div style="flex: 1;"><div style="font-size: 0.9rem; color: rgba(255,255,255,0.7); margin-bottom: 4px;">耗損 (qty)</div><div style="font-size: 1.5rem; font-weight: 700; color: #ff6b6b;">{st_loss:.3f}</div></div></div></div>', unsafe_allow_html=True)

        k1, k2, k3, k4, k5 = st.columns([1,1,1,1,1], gap="large")
        with k1: st.markdown(f'<div class="kpi-box kpi-border-{sys_status} {sys_anim}"><div class="kpi-label">系統可靠度 (<span style="font-family: \'Times New Roman\', serif; font-style: italic;">R<sub>d</sub></span>)</div><div class="kpi-value">{res.get("reliability",0):.4f}</div></div>', unsafe_allow_html=True)
        with k2: st.markdown(f'<div class="kpi-box"><div class="kpi-label">輸出量 (𝑑)</div><div class="kpi-value">{demand}</div></div>', unsafe_allow_html=True)
        with k3: st.markdown(f'<div class="kpi-box"><div class="kpi-label">動態總功率 (kW)</div><div class="kpi-value">{res.get("total_energy",0):.3f}</div></div>', unsafe_allow_html=True)
        c_color = "green" if sys_carbon <= 70 else "yellow" if sys_carbon <= 100 else "red"
        with k4: st.markdown(f'<div class="kpi-box kpi-border-{c_color}"><div class="kpi-label">總碳排放 (kg)</div><div class="kpi-value">{sys_carbon:.3f}</div></div>', unsafe_allow_html=True)
        with k5: st.markdown(f'<div class="kpi-box kpi-border-red"><div class="kpi-label">總耗損 (qty)</div><div class="kpi-value">{res.get("total_loss",0):.3f}</div></div>', unsafe_allow_html=True)

        st.divider()
        st.markdown("### 📈 數據視覺化分析")
        plot_stations = [get_a_subscript(d['id']) for d in STATION_DATA]
        c1, c2 = st.columns(2)
        with c1:
            fig1 = go.Figure(go.Bar(x=plot_stations, y=res["losses"], marker_color='#60d3ff', name="耗損量"))
            fig1.update_layout(title=dict(text="各工作站耗損量", font=dict(size=22, color='black', weight='bold')), paper_bgcolor='white', plot_bgcolor='white', height=350, margin=dict(b=0), xaxis=dict(title=dict(text='工作站', font=dict(size=18, color='black')), type='category', color='#000000', showline=False, ticks='', ticklen=0, tickfont=dict(size=18, color='#000000', family='Times New Roman')), yaxis=dict(title=dict(text='耗損量', font=dict(size=18, color='black')), color='#000000', showline=True, linecolor='#000000', gridcolor='#000000', tickfont=dict(size=16, color='#000000'), range=[0, max(res["losses"])*1.15 if res["losses"] else 1], autorange=False, rangemode='tozero', zeroline=True, zerolinecolor='#000000'))
            st.plotly_chart(fig1, use_container_width=True)
        with c2:
            fig2 = go.Figure(go.Bar(x=plot_stations, y=res["energies"], marker_color='#ffcf60', name="動態功率"))
            fig2.update_layout(title=dict(text="各工作站動態功率 (kW)", font=dict(size=22, color='black', weight='bold')), paper_bgcolor='white', plot_bgcolor='white', height=350, margin=dict(b=0), xaxis=dict(title=dict(text='工作站', font=dict(size=18, color='black')), type='category', color='#000000', showline=False, ticks='', ticklen=0, tickfont=dict(size=18, color='#000000', family='Times New Roman')), yaxis=dict(title=dict(text='動態功率 (kW)', font=dict(size=18, color='black')), color='#000000', showline=True, linecolor='#000000', gridcolor='#000000', tickfont=dict(size=16, color='#000000'), range=[0, max(res["energies"])*1.15 if res["energies"] else 1], autorange=False, rangemode='tozero', zeroline=True, zerolinecolor='#000000'))
            st.plotly_chart(fig2, use_container_width=True)

        st.markdown("### 📉 系統可靠度敏感度分析")
        def get_dynamic_crit_d(_station_data, _tb_val):
            pi_list_local = [d.get('p', 0.96) * math.exp(-d.get('k', 1.0) * (_tb_val - 1.0)**2) for d in _station_data]
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
        fig3.add_trace(go.Scatter(x=[crit_d], y=[calculate_metrics(crit_d, carbon_factor, STATION_DATA, tb_val)['reliability']], mode='markers', name=f'臨界點 (𝑑={crit_d})', marker=dict(symbol='star', size=22, color='#ffd86b', line=dict(width=2, color='#ff0000')), cliponaxis=False))
        fig3.add_trace(go.Scatter(x=[demand], y=[res.get('reliability', 0)], mode='markers', name=f'當前輸出量 (𝑑={demand})', marker=dict(symbol='circle', size=14, color='#4cd37a', line=dict(width=2, color='#ffffff')), cliponaxis=False))
        fig3.update_layout(title=dict(text="系統可靠度敏感度分析", font=dict(size=22, color='black', weight='bold')), xaxis_title=dict(text="輸出量 (𝑑)", font=dict(size=18, color='black')), yaxis_title=dict(text="系統可靠度", font=dict(size=18, color='black')), paper_bgcolor='white', plot_bgcolor='white', height=400, margin=dict(l=20, r=20, t=40, b=20), legend=dict(yanchor="top", y=0.99, xanchor="right", x=0.99, font=dict(color="black", size=14)), xaxis=dict(color='#000000', showline=False, ticks='', ticklen=0, gridcolor='#000000', zeroline=False, tickfont=dict(size=16, color='#000000'), range=[10000, max(d_range_vals + [demand, crit_d]) + max(step, 1000)]), yaxis=dict(color='#000000', showline=True, linecolor='#000000', ticks='', ticklen=0, gridcolor='#000000', zeroline=True, zerolinecolor='#000000', tickmode='array', tickvals=[0, 0.2, 0.4, 0.6, 0.8, 1.0], tickfont=dict(size=16, color='#000000'), range=[0, 1.05], rangemode='tozero'))
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

# ============================================================
# TAB 2: AI 戰情助理
# ============================================================
with tab_chat:
    st.markdown("## 🤖 AI 戰情助理")
    st.markdown('<p style="color: #aac4d8; margin-top: -10px; margin-bottom: 20px; font-size: 1.05rem;">支援自然語言情境模擬、參數自動提取、多輪對話</p>', unsafe_allow_html=True)

    try:
        source_df_chat = st.session_state.df_data
        STATION_DATA_CHAT = [{
            "name": str(int(row['Station'])), "id": int(row['Station']),
            "capacities": parse_list_from_string(row['capacities']),
            "probs": parse_list_from_string(row['probs']), "p": row['p'],
            "power": row['power'], "k": row.get('k', 1.0)
        } for _, row in source_df_chat.iterrows()]
    except:
        STATION_DATA_CHAT = []

    if not st.session_state.chat_history:
        st.markdown("""
        <div style="
            background: linear-gradient(135deg, rgba(63,230,255,0.06), rgba(11,22,38,0.8));
            border: 1.5px solid rgba(63,230,255,0.25);
            border-radius: 16px;
            padding: 28px 32px;
            margin-bottom: 20px;
        ">
            <div style="color: #3fe6ff; font-weight: 700; font-size: 1.3rem; margin-bottom: 14px;">
                👋 哈囉！我是您的 AI 戰情助理
            </div>
            <p style="color: #c8e6f5; margin: 0; line-height: 2; font-size: 1.05rem;">
                我可以幫您：<br>
                🔍 <b>模擬分析</b>：輸入「如果產量 (<i>d</i>) 改成 15000，系統可靠嗎？」<br>
                📊 <b>狀態診斷</b>：輸入「目前碳排 (CO₂ 係數) 情況如何？」<br>
                ⚙️ <b>參數調整</b>：輸入「厚度參數 (<i>t</i><sub>b</sub>) 改 0.85，CO₂ 係數 0.5，結果如何？」<br>
                📉 <b>可靠度查詢</b>：輸入「目前系統可靠度 (<i>R</i><sub><i>d</i></sub>) 是否達標？」
            </p>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <style>
        .chat-label-user { text-align:right; font-size:0.88rem; color:#7abadb; margin-bottom:4px; }
        .chat-label-ai   { text-align:left;  font-size:0.88rem; color:#3fe6ff; margin-bottom:4px; }
        .chat-bubble-user {
            background: #ffffff;
            border: 1.5px solid #d0d0d0;
            border-radius:16px 16px 4px 16px; padding:13px 20px;
            margin:0 0 8px 80px; color:#1a1a1a;
            font-size:1.05rem; line-height:1.75; }
        .chat-bubble-ai {
            background: #1a1440;
            border: 1.5px solid #7f77dd;
            border-radius:16px 16px 16px 4px; padding:13px 20px;
            margin:0 80px 6px 0; color:#cecbf6;
            font-size:1.05rem; line-height:1.85;
            box-shadow:0 2px 8px rgba(0,0,0,0.3); }
        .sim-summary {
            background:rgba(255,255,255,0.04);
            border:1px solid rgba(255,255,255,0.12);
            border-radius:10px; padding:10px 18px;
            margin:2px 80px 18px 0;
            font-size:0.92rem; color:#aac4d8; line-height:1.8; }
        </style>
        """, unsafe_allow_html=True)

        for turn in st.session_state.chat_history:
            safe_user = html_lib.escape(turn["user"])
            safe_ai = html_lib.escape(turn["ai"]).replace("\n", "<br>")
            safe_ai = safe_ai.replace("【現況分析】", "<br><b style='color:#3fe6ff'>【現況分析】</b>")
            safe_ai = safe_ai.replace("【風險評估】", "<br><b style='color:#ffd86b'>【風險評估】</b>")
            safe_ai = safe_ai.replace("【建議措施】", "<br><b style='color:#4cd37a'>【建議措施】</b>")

            st.markdown(
                f'<div class="chat-label-user">您</div>'
                f'<div class="chat-bubble-user">{safe_user}</div>',
                unsafe_allow_html=True
            )
            st.markdown(
                f'<div class="chat-label-ai">🤖 AI 戰情助理</div>'
                f'<div class="chat-bubble-ai">{safe_ai}</div>',
                unsafe_allow_html=True
            )
            
            # 圖表渲染
            if turn.get("chart"):
                chart_type = turn["chart"]
                s = turn.get("sim_summary") or {
                    "d": st.session_state.sim_d,
                    "tb": st.session_state.sim_tb,
                    "cf": st.session_state.sim_cf,
                }
                chart_metrics = calculate_metrics(
                    s["d"], s["cf"], STATION_DATA_CHAT, s["tb"]
                )
                station_labels = ["吹瓶站", "充填站", "套標站", "包裝站", "疊棧站"]

                if chart_type == "loss":
                    fig = go.Figure(go.Bar(
                        x=station_labels, y=chart_metrics["losses"],
                        marker_color='#ff6b6b', name="耗損量"
                    ))
                    fig.update_layout(
                        title=dict(text="各工作站耗損量", font=dict(size=16, color='black')),
                        paper_bgcolor='white', plot_bgcolor='white',
                        height=280, margin=dict(l=10, r=10, t=40, b=10),
                        xaxis=dict(tickfont=dict(size=13, color='black')),
                        yaxis=dict(tickfont=dict(size=13, color='black'))
                    )
                    st.plotly_chart(fig, use_container_width=True, key=f"chat_chart_{turn['time']}_loss")

                elif chart_type == "energy":
                    fig = go.Figure(go.Bar(
                        x=station_labels, y=chart_metrics["energies"],
                        marker_color='#ffcf60', name="動態功率"
                    ))
                    fig.update_layout(
                        title=dict(text="各工作站動態功率 (kW)", font=dict(size=16, color='black')),
                        paper_bgcolor='white', plot_bgcolor='white',
                        height=280, margin=dict(l=10, r=10, t=40, b=10),
                        xaxis=dict(tickfont=dict(size=13, color='black')),
                        yaxis=dict(tickfont=dict(size=13, color='black'))
                    )
                    st.plotly_chart(fig, use_container_width=True, key=f"chat_chart_{turn['time']}_energy")

                elif chart_type == "carbon":
                    fig = go.Figure(go.Bar(
                        x=station_labels, y=chart_metrics["carbons"],
                        marker_color='#4cd37a', name="碳排放"
                    ))
                    fig.update_layout(
                        title=dict(text="各工作站碳排放 (kg)", font=dict(size=16, color='black')),
                        paper_bgcolor='white', plot_bgcolor='white',
                        height=280, margin=dict(l=10, r=10, t=40, b=10),
                        xaxis=dict(tickfont=dict(size=13, color='black')),
                        yaxis=dict(tickfont=dict(size=13, color='black'))
                    )
                    st.plotly_chart(fig, use_container_width=True, key=f"chat_chart_{turn['time']}_carbon")

                elif chart_type == "reliability":
                    d_range = list(range(5000, s["d"] + 5000, 500))
                    y_vals = [calculate_metrics(dv, s["cf"], STATION_DATA_CHAT, s["tb"])["reliability"] for dv in d_range]
                    fig = go.Figure(go.Scatter(
                        x=d_range, y=y_vals, mode='lines+markers',
                        line=dict(color='#3fe6ff', width=2)
                    ))
                    fig.add_trace(go.Scatter(
                        x=[s["d"]], y=[chart_metrics["reliability"]],
                        mode='markers', marker=dict(size=12, color='#ffd86b'),
                        name=f'當前 d={s["d"]}'
                    ))
                    fig.update_layout(
                        title=dict(text="可靠度敏感度分析", font=dict(size=16, color='black')),
                        paper_bgcolor='white', plot_bgcolor='white',
                        height=280, margin=dict(l=10, r=10, t=40, b=10),
                        xaxis=dict(tickfont=dict(size=13, color='black')),
                        yaxis=dict(tickfont=dict(size=13, color='black'), range=[0, 1.05])
                    )
                    st.plotly_chart(fig, use_container_width=True, key=f"chat_chart_{turn['time']}_rel")

            if turn.get("sim_summary"):
                s = turn["sim_summary"]
                rd_color = "#4cd37a" if s["rd"] > 0.95 else "#ffd86b" if s["rd"] >= 0.9 else "#ff6b6b"
                cb_color = "#4cd37a" if s["carbon"] <= 70 else "#ffd86b" if s["carbon"] <= 100 else "#ff6b6b"
                st.markdown(f"""
                <div class="sim-summary">
                    📊 <b style="color:#f3a21a">模擬摘要</b>&nbsp;&nbsp;
                    <i>d</i>={s["d"]}&nbsp;&nbsp;
                    <i>t</i><sub>b</sub>={s["tb"]}&nbsp;&nbsp;
                    CO₂ 係數={s["cf"]}&nbsp;&nbsp;
                    <span style="color:{rd_color};font-weight:700;">
                        <i>R</i><sub><i>d</i></sub>={s["rd"]:.4f}
                    </span>&nbsp;&nbsp;
                    <span style="color:{cb_color};font-weight:700;">
                        總碳排放={s["carbon"]:.1f} kg
                    </span>
                </div>
                """, unsafe_allow_html=True)

    col_clear_btn, _ = st.columns([1, 4])
    with col_clear_btn:
        if st.button("🗑️ 清除對話", use_container_width=True):
            st.session_state.chat_history = []
            st.rerun()

    st.markdown('<p style="color: #f3a21a; font-weight: 700; font-size: 1.05rem; margin-bottom: 10px;">💡 快速提問</p>', unsafe_allow_html=True)
    q_cols = st.columns(4)
    quick_questions = [
        "目前系統狀態如何？",
        "瓶胚厚度改成0.9，可靠度還達標嗎？",
        "哪個工作站耗損最高？",
        "產量增加到18000，碳排會超標嗎？"
    ]
    for i, (col, qq) in enumerate(zip(q_cols, quick_questions)):
        with col:
            if st.button(qq, key=f"qq_{i}", use_container_width=True):
                st.session_state["pending_quick_q"] = qq
                st.rerun()

    st.markdown("<div style='margin-bottom: 6px;'></div>", unsafe_allow_html=True)

    pending_q = st.session_state.pop("pending_quick_q", None)
    user_input = st.chat_input("💬 請輸入您的問題（例：如果產量改成15000，系統可靠度還安全嗎？）")
    final_query = pending_q or user_input

    if final_query:
        if not GOOGLE_API_KEYS or all("你的" in k for k in GOOGLE_API_KEYS):
            st.error("⚠️ 請先在程式碼頂部填入有效的 Google API Key（GOOGLE_API_KEYS 列表）！")
        elif not STATION_DATA_CHAT:
            st.error("⚠️ 無有效工作站資料，請先在「資料管理」頁面設定。")
        else:
            with st.spinner("🤖 AI 戰情助理分析中..."):
                try:
                    gemini_client = get_gemini_client()

                    current_params = {
                        "d":  st.session_state.sim_d,
                        "tb": st.session_state.sim_tb,
                        "cf": st.session_state.sim_cf,
                    }
                    current_metrics = calculate_metrics(
                        current_params["d"], current_params["cf"],
                        STATION_DATA_CHAT, current_params["tb"]
                    )
                    extracted, ai_reply, ai_chart = call_ai_single(
                        gemini_client, final_query,
                        current_params, current_metrics,
                        st.session_state.chat_history
                    )
                    sim_changed = (
                        extracted["d"] != current_params["d"] or
                        abs(extracted["tb"] - current_params["tb"]) > 0.001 or
                        abs(extracted["cf"] - current_params["cf"]) > 0.001
                    )
                    sim_res = calculate_metrics(
                        extracted["d"], extracted["cf"],
                        STATION_DATA_CHAT, extracted["tb"]
                    ) if sim_changed else current_metrics

                    st.session_state.chat_history.append({
                        "user": final_query,
                        "ai": ai_reply,
                        "chart": ai_chart,
                        "time": datetime.now().strftime("%H:%M"),
                        "sim_summary": {
                            "d": extracted["d"], "tb": extracted["tb"], "cf": extracted["cf"],
                            "rd": sim_res.get("reliability", 0),
                            "carbon": sim_res.get("carbon_emission", 0)
                        } if sim_changed else None
                    })
                    if sim_changed:
                        st.session_state.pending_ai_updates = extracted
                        st.toast("✅ 參數已由 AI 戰情助理自動更新，儀表板同步刷新", icon="🔄")

                    st.rerun()

                except Exception as e:
                    st.error(f"❌ AI 呼叫失敗：{e}")

# ============================================================
# TAB 3: 資料管理（Excel 編輯）
# ============================================================
with tab_editor:
    st.subheader("Excel 資料編輯器 (支援動態長度)")
    col_upload, _ = st.columns([2, 1])
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
                st.session_state.force_tab_index = 2
                st.rerun()
            except Exception as e:
                st.error(f"讀取失敗: {e}")

    def maintain_editor_tab(): st.session_state.force_tab_index = 2

    edited_df = st.data_editor(
        st.session_state.df_data[['Station', 'p', 'power', 'k', 'capacities', 'probs']],
        num_rows="dynamic", use_container_width=True, key="editor_table",
        on_change=maintain_editor_tab, disabled=["k"],
        column_config={
            "Station": st.column_config.NumberColumn("站號 (𝑎ₙ)", min_value=1, step=1, required=True),
            "p": None,
            "power": st.column_config.NumberColumn("基礎功率 (kW)", help="最大功率，儀表板依比例動態計算"),
            "k": st.column_config.NumberColumn("參數 (𝑘)", format="%.2f"),
            "capacities": st.column_config.TextColumn("產能列表 (List)", help="例如 [0, 100, 200]"),
            "probs": st.column_config.TextColumn("機率列表 (List)", help="例如 [0.1, 0.4, 0.5]")
        }
    )

    col_reset, col_save = st.columns([1, 1])
    with col_reset:
        if st.button("🔄 重置為預設資料", use_container_width=True):
            st.session_state.df_data = get_default_data()
            st.session_state.force_tab_index = 2
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
                        long_rows.append({"Station": int(row['Station']), "Machine": 1, "Success rate": row['p'], "Power(kW)加工功率": row['power'], "capacity": caps[i], "Capacity_Prob": probs[i], "k": row.get('k', 1.0), "pi(deg)": pi_calc})
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