"""
App 2｜職業棋士每日極簡打卡系統
Mobile-First Streamlit App with Google Sheets Integration
"""

import streamlit as st
import gspread
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ─────────────────────────────────────────
# 0. 頁面設定（必須是第一個 st 指令）
# ─────────────────────────────────────────
st.set_page_config(
    page_title="每日訓練打卡",
    page_icon="♟️",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────
# Mobile-First CSS 注入
# ─────────────────────────────────────────
st.markdown("""
<style>
    /* ── 字型：思源黑體（聶永真御用）── */
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;700;900&display=swap');

    html, body, [class*="css"], .stMarkdown, .stRadio, .stSelectbox,
    .stButton, div, span, p, label {
        font-family: 'Noto Sans TC', 'Source Han Sans TC', sans-serif !important;
    }

    /* ── 全局手機最佳化 ── */
    .main .block-container {
        padding: 1rem 1rem 2rem 1rem;
        max-width: 480px;
        margin: 0 auto;
    }

    /* ── 教練語錄 ── */
    .coach-quote {
        border-left: 5px solid #f0a500;
        background: #111111;
        color: #dddddd;
        padding: 0.85rem 1.1rem;
        margin: 0.75rem 0 1rem 0;
        border-radius: 0 6px 6px 0;
        font-size: 1.05rem;
        font-weight: 700;
        line-height: 1.75;
        letter-spacing: 0.04em;
    }

    /* ── 星期大標題 ── */
    .day-title {
        font-size: 2.4rem;
        font-weight: 900;
        letter-spacing: -0.01em;
        margin: 0.4rem 0 0.2rem 0;
        color: #ffffff;
        line-height: 1.1;
    }

    /* ── 任務卡片 ── */
    .task-card {
        background: #111827;
        border: 1px solid #1f2d40;
        border-radius: 8px;
        padding: 0.75rem 1rem;
        margin-bottom: 0.4rem;
    }
    .task-name {
        font-size: 1.25rem;       /* 放大：撐滿卡片 */
        font-weight: 900;
        color: #f0f0f0;
        margin-bottom: 0.2rem;
        letter-spacing: 0.02em;
        line-height: 1.3;
    }
    .task-time {
        font-size: 0.78rem;
        font-weight: 400;
        color: #6b7280;
        letter-spacing: 0.05em;
    }

    /* ── 任務數量小標 ── */
    .task-count {
        font-size: 0.85rem;
        font-weight: 700;
        color: #9ca3af;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        margin-bottom: 0.6rem;
    }

    /* ── 送出按鈕 ── */
    .stButton > button {
        width: 100%;
        height: 3.4rem;
        font-size: 1.1rem;
        font-weight: 900;
        letter-spacing: 0.06em;
        border-radius: 8px;
        border: none;
        background: linear-gradient(135deg, #00b09b, #96c93d);
        color: white;
        cursor: pointer;
        transition: opacity 0.2s;
    }
    .stButton > button:disabled {
        background: #1f2937 !important;
        color: #4b5563 !important;
        cursor: not-allowed !important;
    }

    /* ── Radio 元件 ── */
    div[role="radiogroup"] label {
        font-size: 0.9rem !important;
        font-weight: 700 !important;
        letter-spacing: 0.02em !important;
        padding: 0.25rem 0 !important;
    }

    /* ── 隱藏 Streamlit 預設 header/footer ── */
    header[data-testid="stHeader"] { display: none; }
    footer { display: none; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# 1. Google Sheets 連線
# ─────────────────────────────────────────
def get_gspread_client():
    """Session 內快取 client，避免每次 radio 點擊都重建連線。"""
    if "gspread_client" not in st.session_state:
        creds_dict = {k: v for k, v in st.secrets["gcp_service_account"].items()}
        st.session_state.gspread_client = gspread.service_account_from_dict(creds_dict)
    return st.session_state.gspread_client

def fetch_quote() -> str:
    """從 Quote_DB 的 A1 格讀取本週金句，結果存在 session_state 避免重複請求。"""
    if "cached_quote" not in st.session_state:
        try:
            client = get_gspread_client()
            sheet = client.open("Quote_DB").sheet1
            val = sheet.acell("A1").value
            st.session_state.cached_quote = str(val).strip() if val else "每一天都是新的開始。"
        except Exception as e:
            st.session_state.cached_quote = f"[DEBUG] 讀取失敗：{type(e).__name__}: {e}"
    return st.session_state.cached_quote

def _sheet_to_dicts(rows: list[list]) -> list[dict]:
    """
    將無標題列的原始資料轉為 dict 清單。
    欄位順序（依截圖）：A=姓名 B=星期 C=時段 D=時間 E=任務名稱
    """
    COL = ["姓名", "星期", "時段", "時間", "任務名稱"]
    result = []
    for row in rows:
        # 跳過完全空白的列
        if not any(str(c).strip() for c in row):
            continue
        # 補齊欄位長度
        padded = list(row) + [""] * max(0, len(COL) - len(row))
        result.append({COL[i]: str(padded[i]).strip() for i in range(len(COL))})
    return result

@st.cache_data(ttl=300)
def fetch_schedule(name: str, weekday_str: str) -> list[dict]:
    """
    從 Schedule_DB 抓取指定學員、指定星期的訓練任務。
    工作表無需標題列，欄位順序：A=姓名 B=星期 C=時段 D=時間 E=任務名稱
    """
    try:
        client = get_gspread_client()
        sheet = client.open("Schedule_DB").sheet1
        rows = sheet.get_all_values()
        records = _sheet_to_dicts(rows)
        tasks = [
            r for r in records
            if r.get("姓名") == name and r.get("星期") == weekday_str
        ]
        # 依「時段」→「時間」排序
        def sort_key(r):
            order = {"早上": "0", "下午": "1", "晚上": "2"}
            return order.get(r.get("時段", ""), "9") + r.get("時間", "")
        tasks.sort(key=sort_key)
        return tasks
    except Exception as e:
        st.error(f"讀取 Schedule_DB 失敗：{e}")
        return []

def get_student_names() -> list[str]:
    """從 Schedule_DB 讀取不重複的學員姓名清單。"""
    try:
        client = get_gspread_client()
        sheet = client.open("Schedule_DB").sheet1
        rows = sheet.get_all_values()
        records = _sheet_to_dicts(rows)
        names = sorted({r["姓名"] for r in records if r.get("姓名")})
        return names
    except Exception as e:
        st.error(f"讀取學員名單失敗：{e}")
        return []

def append_logs(rows: list[list]):
    """批次寫入紀錄到 Log_DB 工作表。"""
    import traceback
    try:
        client = get_gspread_client()
        sh = client.open("Log_DB")
        sheet = sh.sheet1
        sheet.append_rows(rows, value_input_option="USER_ENTERED")
    except Exception as e:
        raise RuntimeError(f"{type(e).__name__}: {e}\n{traceback.format_exc()}")

def send_email_notification(name: str, weekday: str, rows: list[list]):
    """學員送出後，寄 Email 通知教練。"""
    try:
        cfg = st.secrets["email"]
        sender   = cfg["sender"]
        receiver = cfg["receiver"]
        password = cfg["password"]

        # 組合任務清單 HTML
        rows_html = "".join(
            f"<tr><td style='padding:4px 12px;'>{r[3]}</td>"
            f"<td style='padding:4px 12px; font-size:1.2em;'>{r[4]}</td></tr>"
            for r in rows
        )

        html = f"""
        <div style="font-family:'Noto Sans TC',sans-serif; max-width:480px;">
          <h2 style="margin-bottom:4px;">♟️ 訓練日誌已送出</h2>
          <p style="color:#666; margin-top:0;">{rows[0][0]}</p>
          <p><strong>{name}</strong> 完成了 <strong>{weekday}</strong> 的訓練打卡：</p>
          <table style="border-collapse:collapse; width:100%;">
            <thead>
              <tr style="background:#f0f0f0;">
                <th style="padding:4px 12px; text-align:left;">任務</th>
                <th style="padding:4px 12px; text-align:left;">狀態</th>
              </tr>
            </thead>
            <tbody>{rows_html}</tbody>
          </table>
          <p style="margin-top:16px; color:#999; font-size:0.85em;">
            ✅ 完美燃燒　🔺 浮動妥協
          </p>
        </div>
        """

        msg = MIMEMultipart("alternative")
        msg["Subject"] = f"♟️ [{name}] {weekday} 訓練日誌"
        msg["From"]    = sender
        msg["To"]      = receiver
        msg.attach(MIMEText(html, "html", "utf-8"))

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, password)
            server.sendmail(sender, receiver, msg.as_string())
    except Exception as e:
        # Email 失敗不阻斷主流程，只顯示警告
        st.warning(f"⚠️ Email 通知發送失敗：{e}")

# ─────────────────────────────────────────
# 2. 日期 / 星期計算
# ─────────────────────────────────────────
WEEKDAY_MAP = {
    0: "星期一",
    1: "星期二",
    2: "星期三",
    3: "星期四",
    4: "星期五",
    5: "星期六",
    6: "星期日",
}

today = datetime.now()
today_str = today.strftime("%Y-%m-%d %H:%M")
weekday_str = WEEKDAY_MAP[today.weekday()]

# ─────────────────────────────────────────
# 3. Session State 初始化
# ─────────────────────────────────────────
if "submitted" not in st.session_state:
    st.session_state.submitted = False
if "submit_summary" not in st.session_state:
    st.session_state.submit_summary = []
if "choices" not in st.session_state:
    st.session_state.choices = {}

# ─────────────────────────────────────────
# 4. 成功畫面（送出後顯示，擋住其餘內容）
# ─────────────────────────────────────────
if st.session_state.submitted:
    perfect = sum(1 for r in st.session_state.submit_summary if r[4] == "✅")
    compromise = sum(1 for r in st.session_state.submit_summary if r[4] == "🔺")
    total = len(st.session_state.submit_summary)

    st.markdown(f"""
    <div style="text-align:center; padding: 3rem 1rem;">
        <div style="font-size:4rem; margin-bottom:0.5rem;">✅</div>
        <div style="font-size:2rem; font-weight:900; color:#ffffff; letter-spacing:0.03em;">
            日誌已送出
        </div>
        <div style="font-size:0.9rem; color:#6b7280; margin-top:0.5rem; letter-spacing:0.06em;">
            {st.session_state.submit_summary[0][0] if st.session_state.submit_summary else ""}
        </div>
        <div style="margin-top:2rem; background:#111827; border-radius:12px; padding:1.2rem 1.5rem; text-align:left;">
            <div style="font-size:0.75rem; color:#6b7280; letter-spacing:0.1em; margin-bottom:0.8rem;">本日戰況</div>
            <div style="display:flex; gap:2rem; justify-content:center;">
                <div style="text-align:center;">
                    <div style="font-size:2.5rem; font-weight:900; color:#34d399;">{perfect}</div>
                    <div style="font-size:0.75rem; color:#6b7280;">完美燃燒</div>
                </div>
                <div style="text-align:center;">
                    <div style="font-size:2.5rem; font-weight:900; color:#f59e0b;">{compromise}</div>
                    <div style="font-size:0.75rem; color:#6b7280;">浮動妥協</div>
                </div>
                <div style="text-align:center;">
                    <div style="font-size:2.5rem; font-weight:900; color:#9ca3af;">{total}</div>
                    <div style="font-size:0.75rem; color:#6b7280;">總任務</div>
                </div>
            </div>
        </div>
        <div style="font-size:0.85rem; color:#4b5563; margin-top:1.5rem;">
            📧 日誌已寄送到教練信箱
        </div>
    </div>
    """, unsafe_allow_html=True)

    if st.button("↩ 返回（再次打卡）", use_container_width=True):
        st.session_state.submitted = False
        st.session_state.submit_summary = []
        st.session_state.choices = {}
        st.rerun()
    st.stop()

# ─────────────────────────────────────────
# 5. 介面：頂部區塊
# ─────────────────────────────────────────
st.markdown("♟️ **每日訓練打卡**", unsafe_allow_html=False)

# 學員選單
names = get_student_names()
if not names:
    st.warning("⚠️ 尚未讀取到學員名單，請確認 Schedule_DB 設定。")
    st.stop()

selected_name = st.selectbox("選擇學員", names, label_visibility="collapsed",
                              placeholder="── 請選擇學員姓名 ──")

# 星期大標題
st.markdown(f'<div class="day-title">📅 {weekday_str}</div>', unsafe_allow_html=True)

# 教練語錄（從 Quote_DB A1 動態讀取）
quote = fetch_quote()
st.markdown(
    f'<div class="coach-quote">「{quote}」</div>',
    unsafe_allow_html=True,
)

st.divider()

# ─────────────────────────────────────────
# 5. 讀取今日任務
# ─────────────────────────────────────────
if selected_name:
    tasks = fetch_schedule(selected_name, weekday_str)
else:
    tasks = []

if not tasks:
    st.info("今天沒有排定的訓練任務，或尚未選擇學員。")
    st.stop()

# ─────────────────────────────────────────
# 6. 任務列表 + 單選元件
# ─────────────────────────────────────────
OPTIONS = ["未選擇", "✅ 完美燃燒", "🔺 浮動妥協"]

st.markdown(f'<div class="task-count">共 {len(tasks)} 項任務｜請逐一如實填寫</div>', unsafe_allow_html=True)

for i, task in enumerate(tasks):
    task_name = str(task.get("任務名稱", f"任務 {i+1}")).strip()
    task_period = str(task.get("時段", "")).strip()   # 早上/下午/晚上
    task_time   = str(task.get("時間", "")).strip()   # 07:00-08:00

    time_label = " ".join(filter(None, [task_period, task_time]))

    st.markdown(
        f'<div class="task-card">'
        f'<div class="task-name">{task_name}</div>'
        + (f'<div class="task-time">🕐 {time_label}</div>' if time_label else "")
        + '</div>',
        unsafe_allow_html=True,
    )

    key = f"task_{i}_{task_name}"
    current_val = st.session_state.choices.get(key, OPTIONS[0])

    choice = st.radio(
        label=task_name,
        options=OPTIONS,
        index=OPTIONS.index(current_val),
        key=key,
        horizontal=True,
        label_visibility="collapsed",
    )
    st.session_state.choices[key] = choice

st.divider()

# ─────────────────────────────────────────
# 7. 防呆鎖定 + 送出按鈕
# ─────────────────────────────────────────
all_filled = all(
    v != "未選擇" for v in st.session_state.choices.values()
) and len(st.session_state.choices) == len(tasks)

unanswered = sum(1 for v in st.session_state.choices.values() if v == "未選擇")

if not all_filled:
    st.caption(f"⚠️ 還有 **{unanswered}** 項任務尚未填寫，填完才能送出。")

submit_clicked = st.button(
    "📤 送出今日訓練日誌",
    disabled=not all_filled,
    use_container_width=True,
)

# ─────────────────────────────────────────
# 8. 資料寫入 Log_DB
# ─────────────────────────────────────────
if submit_clicked and all_filled:
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    rows_to_write = []
    for i, task in enumerate(tasks):
        task_name = str(task.get("任務名稱", f"任務 {i+1}")).strip()
        key = f"task_{i}_{task_name}"
        status = st.session_state.choices.get(key, "未選擇")
        # 將 emoji 標籤轉為純狀態符號
        status_clean = "✅" if "完美燃燒" in status else "🔺"
        rows_to_write.append([now_str, selected_name, weekday_str, task_name, status_clean])

    try:
        append_logs(rows_to_write)
        send_email_notification(selected_name, weekday_str, rows_to_write)
        st.session_state.submitted = True
        st.session_state.submit_summary = rows_to_write
        st.session_state.choices = {}
        fetch_schedule.clear()
        st.rerun()
    except Exception as e:
        st.error(f"❌ 寫入失敗，請稍後再試：{e}")
