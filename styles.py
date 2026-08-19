"""
Enterprise-Grade Training System UI
Premium design inspired by Linear, Vercel, Stripe
"""
import streamlit.components.v1 as components

def _render_html(html: str, height: int = None):
    """Render HTML via iframe with CSS injection"""
    # Inject brand CSS into every iframe
    css_inline = """
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=Noto+Sans+SC:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: 'Inter', 'Noto Sans SC', -apple-system, BlinkMacSystemFont, sans-serif;
      -webkit-font-smoothing: antialiased;
      -moz-osx-font-smoothing: grayscale;
    }
    :root {
      --navy-900: #0a1628;
      --navy-800: #0f2340;
      --navy-700: #152d50;
      --navy-600: #1a3c6e;
      --navy-500: #2a5ca8;
      --gold-500: #c9a84c;
      --gold-400: #d4b86a;
      --gold-300: #e0c888;
      --slate-50: #f8fafc;
      --slate-100: #f1f5f9;
      --slate-200: #e2e8f0;
      --slate-300: #cbd5e1;
      --slate-400: #94a3b8;
      --slate-500: #64748b;
      --slate-600: #475569;
      --slate-700: #334155;
      --slate-800: #1e293b;
      --slate-900: #0f172a;
      --green-500: #22c55e;
      --green-400: #4ade80;
      --red-500: #ef4444;
      --orange-500: #f97316;
      --shadow-sm: 0 1px 2px rgba(0,0,0,.05);
      --shadow-md: 0 4px 6px -1px rgba(0,0,0,.1), 0 2px 4px -2px rgba(0,0,0,.1);
      --shadow-lg: 0 10px 15px -3px rgba(0,0,0,.1), 0 4px 6px -4px rgba(0,0,0,.1);
      --shadow-xl: 0 20px 25px -5px rgba(0,0,0,.1), 0 8px 10px -6px rgba(0,0,0,.1);
      --shadow-glow: 0 0 40px rgba(201,168,76,.15);
      --radius-sm: 8px;
      --radius-md: 12px;
      --radius-lg: 16px;
      --radius-xl: 24px;
    }
    </style>
    """
    full_html = f"{css_inline}<body>{html}</body>"
    components.html(full_html, height=height or 100, scrolling=True)


# ============================================================================
# GLOBAL CSS INJECTION
# ============================================================================
GLOBAL_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=Noto+Sans+SC:wght@300;400;500;600;700&display=swap');

:root {
  --navy-900: #0a1628;
  --navy-800: #0f2340;
  --navy-700: #152d50;
  --navy-600: #1a3c6e;
  --navy-500: #2a5ca8;
  --gold-500: #c9a84c;
  --gold-400: #d4b86a;
  --slate-50: #f8fafc;
  --slate-100: #f1f5f9;
  --slate-200: #e2e8f0;
  --slate-400: #94a3b8;
  --slate-500: #64748b;
  --slate-600: #475569;
  --slate-800: #1e293b;
  --green-500: #22c55e;
  --red-500: #ef4444;
}

/* Hide Streamlit defaults */
#MainMenu, footer, header { visibility: hidden !important; }
.stDeployButton { display: none !important; }

/* Global app styling */
.stApp {
  font-family: 'Inter', 'Noto Sans SC', -apple-system, sans-serif !important;
  background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%) !important;
}

/* Sidebar - Premium dark theme */
section[data-testid="stSidebar"] {
  background: linear-gradient(180deg, #0a1628 0%, #0f2340 50%, #152d50 100%) !important;
  border-right: 1px solid rgba(255,255,255,.06) !important;
}
section[data-testid="stSidebar"]::before {
  content: '';
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  height: 200px;
  background: linear-gradient(135deg, rgba(201,168,76,.08) 0%, transparent 100%);
  pointer-events: none;
}
section[data-testid="stSidebar"] .stMarkdown p,
section[data-testid="stSidebar"] .stMarkdown h1,
section[data-testid="stSidebar"] .stMarkdown h2,
section[data-testid="stSidebar"] .stMarkdown h3 {
  color: #fff !important;
}
section[data-testid="stSidebar"] .stMarkdown small {
  color: rgba(255,255,255,.5) !important;
}

/* Sidebar buttons - Glass effect */
section[data-testid="stSidebar"] .stButton > button {
  background: rgba(255,255,255,.04) !important;
  color: rgba(255,255,255,.85) !important;
  border: 1px solid rgba(255,255,255,.08) !important;
  border-radius: 10px !important;
  transition: all .2s cubic-bezier(.4,0,.2,1) !important;
  backdrop-filter: blur(10px) !important;
  font-weight: 500 !important;
}
section[data-testid="stSidebar"] .stButton > button:hover {
  background: rgba(255,255,255,.1) !important;
  border-color: rgba(201,168,76,.3) !important;
  transform: translateX(4px) !important;
  box-shadow: 0 0 20px rgba(201,168,76,.1) !important;
}
section[data-testid="stSidebar"] .stButton > button:active {
  transform: translateX(2px) !important;
  background: rgba(201,168,76,.15) !important;
}

/* Primary buttons - Premium gradient */
.stButton > button[kind="primary"],
.stButton > button[data-testid="stBaseButton-primary"] {
  background: linear-gradient(135deg, #1a3c6e 0%, #2a5ca8 100%) !important;
  border: none !important;
  color: #fff !important;
  font-weight: 600 !important;
  letter-spacing: .3px !important;
  box-shadow: 0 4px 15px rgba(26,60,110,.3) !important;
  transition: all .3s cubic-bezier(.4,0,.2,1) !important;
}
.stButton > button[kind="primary"]:hover {
  transform: translateY(-2px) !important;
  box-shadow: 0 8px 25px rgba(26,60,110,.4) !important;
}

/* Secondary buttons */
.stButton > button:not([kind="primary"]) {
  background: #fff !important;
  border: 1px solid #e2e8f0 !important;
  color: #334155 !important;
  font-weight: 500 !important;
  transition: all .2s ease !important;
}
.stButton > button:not([kind="primary"]):hover {
  border-color: #cbd5e1 !important;
  box-shadow: 0 4px 12px rgba(0,0,0,.08) !important;
}

/* Input fields - Modern design */
.stTextInput > div > div > input,
.stTextArea > div > div > textarea,
.stSelectbox > div > div {
  border: 2px solid #e2e8f0 !important;
  border-radius: 10px !important;
  background: #fff !important;
  transition: all .2s ease !important;
  font-size: .95rem !important;
}
.stTextInput > div > div > input:focus,
.stTextArea > div > div > textarea:focus {
  border-color: #2a5ca8 !important;
  box-shadow: 0 0 0 4px rgba(42,92,168,.1) !important;
}

/* Tabs - Premium style */
.stTabs [data-baseweb="tab-list"] {
  gap: 8px !important;
  background: transparent !important;
}
.stTabs [data-baseweb="tab"] {
  border-radius: 10px !important;
  font-weight: 500 !important;
  padding: 10px 20px !important;
}
.stTabs [aria-selected="true"] {
  background: linear-gradient(135deg, #1a3c6e, #2a5ca8) !important;
  color: #fff !important;
}

/* Expanders */
.streamlit-expanderHeader {
  border-radius: 12px !important;
  font-weight: 600 !important;
  background: #fff !important;
  border: 1px solid #e2e8f0 !important;
}
.streamlit-expanderHeader:hover {
  border-color: #cbd5e1 !important;
}

/* Progress bar */
.stProgress > div > div > div {
  background: linear-gradient(90deg, #1a3c6e, #c9a84c) !important;
  border-radius: 100px !important;
}

/* Dividers */
hr {
  border: none !important;
  border-top: 1px solid #e2e8f0 !important;
  margin: 1rem 0 !important;
}

/* Alerts */
.stAlert {
  border-radius: 12px !important;
  border-left-width: 4px !important;
}
</style>
"""


def inject_global_css():
    """Inject global CSS into Streamlit"""
    from streamlit.components.v1 import html
    css = GLOBAL_CSS.replace('<style>', '').replace('</style>', '').strip()
    html(f"<style>{css}</style>", height=0)


# ============================================================================
# LOGIN PAGE - Premium glass morphism design
# ============================================================================
def render_login():
    _render_html("""
<style>
.login-page {
  min-height: 100vh;
  display: flex;
  align-items: center;
  justify-content: center;
  background: linear-gradient(135deg, #0a1628 0%, #0f2340 30%, #152d50 60%, #1a3c6e 100%);
  position: relative;
  overflow: hidden;
  padding: 20px;
}
.login-page::before {
  content: '';
  position: absolute;
  width: 800px;
  height: 800px;
  border-radius: 50%;
  background: radial-gradient(circle, rgba(201,168,76,.12) 0%, transparent 70%);
  top: -300px;
  right: -200px;
  animation: float 20s ease-in-out infinite;
}
.login-page::after {
  content: '';
  position: absolute;
  width: 600px;
  height: 600px;
  border-radius: 50%;
  background: radial-gradient(circle, rgba(42,92,168,.15) 0%, transparent 70%);
  bottom: -200px;
  left: -150px;
  animation: float 15s ease-in-out infinite reverse;
}
@keyframes float {
  0%, 100% { transform: translate(0, 0) scale(1); }
  50% { transform: translate(30px, -30px) scale(1.05); }
}

/* Grid pattern overlay */
.grid-pattern {
  position: absolute;
  inset: 0;
  background-image:
    linear-gradient(rgba(255,255,255,.03) 1px, transparent 1px),
    linear-gradient(90deg, rgba(255,255,255,.03) 1px, transparent 1px);
  background-size: 60px 60px;
  pointer-events: none;
}

.login-card {
  position: relative;
  z-index: 10;
  width: 440px;
  max-width: 95vw;
  background: rgba(255,255,255,.03);
  backdrop-filter: blur(40px);
  -webkit-backdrop-filter: blur(40px);
  border: 1px solid rgba(255,255,255,.08);
  border-radius: 24px;
  padding: 48px 40px;
  box-shadow:
    0 32px 64px rgba(0,0,0,.3),
    0 0 0 1px rgba(255,255,255,.05) inset;
}

.login-logo {
  text-align: center;
  margin-bottom: 40px;
}
.login-logo .icon {
  width: 80px;
  height: 80px;
  margin: 0 auto 20px;
  background: linear-gradient(135deg, rgba(201,168,76,.2), rgba(201,168,76,.05));
  border: 1px solid rgba(201,168,76,.3);
  border-radius: 20px;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 2.5rem;
  box-shadow: 0 8px 32px rgba(201,168,76,.2);
}
.login-logo h1 {
  font-size: 1.5rem;
  font-weight: 700;
  color: #fff;
  letter-spacing: 1px;
  margin-bottom: 8px;
}
.login-logo p {
  font-size: .9rem;
  color: rgba(255,255,255,.5);
  font-weight: 400;
}

.login-divider {
  height: 1px;
  background: linear-gradient(90deg, transparent, rgba(255,255,255,.1), transparent);
  margin: 24px 0;
}

.login-footer {
  text-align: center;
  margin-top: 24px;
  font-size: .8rem;
  color: rgba(255,255,255,.3);
}
</style>
<div class="login-page">
  <div class="grid-pattern"></div>
  <div class="login-card">
    <div class="login-logo">
      <div class="icon">📦</div>
      <h1>供应链培训系统</h1>
      <p>Enterprise Training Platform</p>
    </div>
  </div>
</div>
""", height=700)


# ============================================================================
# DASHBOARD - Premium card-based design
# ============================================================================
def render_dashboard_header(name: str, dept: str, emp_id: str = ""):
    _render_html(f"""
<style>
.dash-header {{
  padding: 32px 0 24px;
}}
.dash-header h1 {{
  font-size: 1.75rem;
  font-weight: 800;
  color: #0f172a;
  margin-bottom: 8px;
  letter-spacing: -.5px;
}}
.dash-header h1 span {{
  background: linear-gradient(135deg, #1a3c6e, #c9a84c);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  background-clip: text;
}}
.dash-header p {{
  color: #64748b;
  font-size: .95rem;
}}
.dash-header .user-info {{
  display: inline-flex;
  align-items: center;
  gap: 12px;
  margin-top: 16px;
  padding: 10px 20px;
  background: #fff;
  border: 1px solid #e2e8f0;
  border-radius: 12px;
  font-size: .85rem;
  color: #475569;
  box-shadow: 0 1px 3px rgba(0,0,0,.05);
}}
.dash-header .user-info .badge {{
  background: linear-gradient(135deg, #1a3c6e, #2a5ca8);
  color: #fff;
  padding: 4px 12px;
  border-radius: 20px;
  font-size: .75rem;
  font-weight: 600;
}}
</style>
<div class="dash-header">
  <h1>欢迎回来, <span>{name}</span></h1>
  <p>跟踪您的培训进度，完成所有模块</p>
  <div class="user-info">
    <span>👤 {name}</span>
    <span>|</span>
    <span>{dept}</span>
    {"<span>|</span><span>ID: " + emp_id + "</span>" if emp_id else ""}
    <span class="badge">Employee</span>
  </div>
</div>
""", height=160)


def render_stats_card(icon: str, label: str, value: str, color: str = "#1a3c6e"):
    _render_html(f"""
<style>
.stat-card {{
  background: #fff;
  border-radius: 16px;
  padding: 24px;
  border: 1px solid #e2e8f0;
  box-shadow: 0 1px 3px rgba(0,0,0,.04);
  transition: all .3s cubic-bezier(.4,0,.2,1);
}}
.stat-card:hover {{
  transform: translateY(-4px);
  box-shadow: 0 12px 24px rgba(0,0,0,.08);
}}
.stat-card .icon {{
  width: 48px;
  height: 48px;
  border-radius: 12px;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 1.5rem;
  margin-bottom: 16px;
  background: {color}15;
}}
.stat-card .label {{
  font-size: .8rem;
  color: #94a3b8;
  text-transform: uppercase;
  letter-spacing: .5px;
  font-weight: 600;
  margin-bottom: 4px;
}}
.stat-card .value {{
  font-size: 1.75rem;
  font-weight: 700;
  color: #0f172a;
}}
</style>
<div class="stat-card">
  <div class="icon">{icon}</div>
  <div class="label">{label}</div>
  <div class="value">{value}</div>
</div>
""", height=150)


def render_progress_card(pct: int):
    _render_html(f"""
<style>
.progress-card {{
  background: linear-gradient(135deg, #0a1628 0%, #152d50 100%);
  border-radius: 20px;
  padding: 32px;
  color: #fff;
  position: relative;
  overflow: hidden;
}}
.progress-card::before {{
  content: '';
  position: absolute;
  width: 200px;
  height: 200px;
  border-radius: 50%;
  background: rgba(201,168,76,.1);
  top: -50px;
  right: -50px;
}}
.progress-card .header {{
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 20px;
  position: relative;
}}
.progress-card .title {{
  font-size: 1.1rem;
  font-weight: 600;
}}
.progress-card .pct {{
  font-size: 2rem;
  font-weight: 800;
  background: linear-gradient(135deg, #c9a84c, #d4b86a);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
}}
.progress-bar-bg {{
  height: 12px;
  background: rgba(255,255,255,.1);
  border-radius: 100px;
  overflow: hidden;
  position: relative;
}}
.progress-bar-fill {{
  height: 100%;
  background: linear-gradient(90deg, #c9a84c, #d4b86a);
  border-radius: 100px;
  width: {pct}%;
  transition: width 1s cubic-bezier(.4,0,.2,1);
  position: relative;
}}
.progress-bar-fill::after {{
  content: '';
  position: absolute;
  inset: 0;
  background: linear-gradient(90deg, transparent, rgba(255,255,255,.3), transparent);
  animation: shimmer 2s infinite;
}}
@keyframes shimmer {{
  0% {{ transform: translateX(-100%); }}
  100% {{ transform: translateX(100%); }}
}}
.progress-hint {{
  margin-top: 12px;
  font-size: .85rem;
  color: rgba(255,255,255,.5);
}}
</style>
<div class="progress-card">
  <div class="header">
    <div class="title">总体进度</div>
    <div class="pct">{pct}%</div>
  </div>
  <div class="progress-bar-bg">
    <div class="progress-bar-fill"></div>
  </div>
  <div class="progress-hint">完成所有模块和考试即可获得认证</div>
</div>
""", height=180)


def render_module_card_v2(title: str, status: str, chapters: str, exam: str,
                          duration: str, module_id: int, is_completed: bool = False):
    status_config = {
        "未开始": {"color": "#94a3b8", "bg": "#f1f5f9", "icon": "○"},
        "进行中": {"color": "#f97316", "bg": "#fff7ed", "icon": "◐"},
        "已完成": {"color": "#22c55e", "bg": "#f0fdf4", "icon": "●"},
    }
    cfg = status_config.get(status, status_config["未开始"])

    _render_html(f"""
<style>
.module-v2 {{
  background: #fff;
  border-radius: 20px;
  padding: 28px;
  border: 1px solid #e2e8f0;
  box-shadow: 0 1px 3px rgba(0,0,0,.04);
  transition: all .3s cubic-bezier(.4,0,.2,1);
  position: relative;
  overflow: hidden;
}}
.module-v2::before {{
  content: '';
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  height: 4px;
  background: {"linear-gradient(90deg, #22c55e, #4ade80)" if is_completed else "linear-gradient(90deg, #1a3c6e, #c9a84c)"};
}}
.module-v2:hover {{
  transform: translateY(-4px);
  box-shadow: 0 16px 32px rgba(0,0,0,.1);
}}
.module-v2 .top {{
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
  margin-bottom: 16px;
}}
.module-v2 .title {{
  font-size: 1.1rem;
  font-weight: 700;
  color: #0f172a;
  line-height: 1.4;
}}
.module-v2 .status {{
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 6px 14px;
  border-radius: 20px;
  font-size: .75rem;
  font-weight: 600;
  background: {cfg["bg"]};
  color: {cfg["color"]};
}}
.module-v2 .stats {{
  display: flex;
  gap: 24px;
  padding: 16px 0;
  border-top: 1px solid #f1f5f9;
}}
.module-v2 .stat {{
  display: flex;
  flex-direction: column;
  gap: 4px;
}}
.module-v2 .stat-label {{
  font-size: .7rem;
  color: #94a3b8;
  text-transform: uppercase;
  letter-spacing: .5px;
}}
.module-v2 .stat-value {{
  font-size: .95rem;
  font-weight: 600;
  color: #334155;
}}
.module-v2 .actions {{
  display: flex;
  gap: 12px;
  margin-top: 16px;
}}
.module-v2 .btn {{
  flex: 1;
  padding: 12px 20px;
  border-radius: 10px;
  font-weight: 600;
  font-size: .85rem;
  text-align: center;
  cursor: pointer;
  transition: all .2s;
  border: none;
}}
.module-v2 .btn-primary {{
  background: linear-gradient(135deg, #1a3c6e, #2a5ca8);
  color: #fff;
}}
.module-v2 .btn-secondary {{
  background: #f1f5f9;
  color: #475569;
}}
</style>
<div class="module-v2">
  <div class="top">
    <div class="title">{title}</div>
    <div class="status">{cfg["icon"]} {status}</div>
  </div>
  <div class="stats">
    <div class="stat">
      <div class="stat-label">课时</div>
      <div class="stat-value">{chapters}</div>
    </div>
    <div class="stat">
      <div class="stat-label">考试</div>
      <div class="stat-value">{exam}</div>
    </div>
    <div class="stat">
      <div class="stat-label">时长</div>
      <div class="stat-value">{duration}</div>
    </div>
  </div>
</div>
""", height=200)


# ============================================================================
# EXAM PAGE - Clean, focused design
# ============================================================================
def render_exam_timer_v2(minutes: int, seconds: int, is_warning: bool = False):
    bg = "linear-gradient(135deg, #ef4444, #dc2626)" if is_warning else "linear-gradient(135deg, #0f172a, #1e293b)"
    _render_html(f"""
<style>
.timer-v2 {{
  display: inline-flex;
  align-items: center;
  gap: 12px;
  padding: 14px 28px;
  background: {bg};
  border-radius: 100px;
  color: #fff;
  font-weight: 700;
  font-size: 1.25rem;
  font-variant-numeric: tabular-nums;
  box-shadow: 0 4px 20px rgba(0,0,0,.2);
  {"animation: pulse-v2 1s infinite;" if is_warning else ""}
}}
@keyframes pulse-v2 {{
  0%, 100% {{ opacity: 1; transform: scale(1); }}
  50% {{ opacity: .8; transform: scale(1.02); }}
}}
.timer-v2 .icon {{
  font-size: 1.1rem;
}}
</style>
<div class="timer-v2">
  <span class="icon">⏱</span>
  <span>{minutes:02d}:{seconds:02d}</span>
</div>
""", height=60)


def render_question_v2(number: int, question: str, q_type: str, per_q: int, total: int):
    type_config = {
        "single": {"label": "Single Choice", "color": "#3b82f6", "bg": "#eff6ff"},
        "multi": {"label": "Multiple Choice", "color": "#8b5cf6", "bg": "#f5f3ff"},
        "short": {"label": "Short Answer", "color": "#10b981", "bg": "#ecfdf5"},
    }
    cfg = type_config.get(q_type, type_config["single"])

    _render_html(f"""
<style>
.question-v2 {{
  background: #fff;
  border-radius: 16px;
  padding: 28px;
  border: 1px solid #e2e8f0;
  margin-bottom: 20px;
  position: relative;
}}
.question-v2 .header {{
  display: flex;
  align-items: center;
  gap: 12px;
  margin-bottom: 16px;
}}
.question-v2 .num {{
  width: 36px;
  height: 36px;
  border-radius: 10px;
  background: linear-gradient(135deg, #0f172a, #1e293b);
  color: #fff;
  display: flex;
  align-items: center;
  justify-content: center;
  font-weight: 700;
  font-size: .85rem;
}}
.question-v2 .type {{
  padding: 4px 12px;
  border-radius: 6px;
  font-size: .7rem;
  font-weight: 600;
  background: {cfg["bg"]};
  color: {cfg["color"]};
}}
.question-v2 .score {{
  margin-left: auto;
  font-size: .8rem;
  color: #94a3b8;
  font-weight: 500;
}}
.question-v2 .text {{
  font-size: 1.05rem;
  font-weight: 500;
  color: #1e293b;
  line-height: 1.7;
}}
</style>
<div class="question-v2">
  <div class="header">
    <div class="num">{number}</div>
    <div class="type">{cfg["label"]}</div>
    <div class="score">{per_q} pts</div>
  </div>
  <div class="text">{question}</div>
</div>
""", height=120)


# ============================================================================
# RESULT PAGE - Celebration design
# ============================================================================
def render_result_v2(score: int, passed: bool, module_name: str):
    if passed:
        bg = "linear-gradient(135deg, #059669, #10b981)"
        icon = "🎉"
        msg = "Congratulations!"
        sub = "You've passed the exam"
    else:
        bg = "linear-gradient(135deg, #dc2626, #ef4444)"
        icon = "😔"
        msg = "Keep Learning"
        sub = "Review the material and try again"

    _render_html(f"""
<style>
.result-v2 {{
  text-align: center;
  padding: 48px 32px;
}}
.score-circle {{
  width: 180px;
  height: 180px;
  border-radius: 50%;
  background: {bg};
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  margin: 0 auto 32px;
  color: #fff;
  box-shadow: 0 16px 48px rgba(0,0,0,.2);
  position: relative;
}}
.score-circle::before {{
  content: '';
  position: absolute;
  inset: -4px;
  border-radius: 50%;
  background: {bg};
  z-index: -1;
  opacity: .3;
  filter: blur(20px);
}}
.score-circle .number {{
  font-size: 3.5rem;
  font-weight: 800;
  line-height: 1;
}}
.score-circle .label {{
  font-size: 1rem;
  font-weight: 500;
  opacity: .9;
  margin-top: 4px;
}}
.result-v2 .icon {{
  font-size: 3rem;
  margin-bottom: 16px;
}}
.result-v2 .message {{
  font-size: 1.75rem;
  font-weight: 700;
  color: #0f172a;
  margin-bottom: 8px;
}}
.result-v2 .sub {{
  font-size: 1rem;
  color: #64748b;
  margin-bottom: 32px;
}}
.result-v2 .module-name {{
  display: inline-block;
  padding: 8px 20px;
  background: #f1f5f9;
  border-radius: 8px;
  font-size: .85rem;
  color: #475569;
  font-weight: 500;
}}
</style>
<div class="result-v2">
  <div class="score-circle">
    <div class="number">{score}</div>
    <div class="label">/ 100</div>
  </div>
  <div class="icon">{icon}</div>
  <div class="message">{msg}</div>
  <div class="sub">{sub}</div>
  <div class="module-name">{module_name}</div>
</div>
""", height=380)


# ============================================================================
# CERTIFICATE - Premium elegant design
# ============================================================================
def render_certificate_v2(name: str, emp_id: str, dept: str,
                          modules: list, date: str):
    rows = ""
    for mod_name, score, mins in modules:
        status = "PASSED" if score >= 80 else "FAILED"
        status_color = "#22c55e" if score >= 80 else "#ef4444"
        rows += f"""
        <tr>
          <td style="padding:14px 20px;border-bottom:1px solid #e2e8f0;font-weight:500">{mod_name}</td>
          <td style="padding:14px 20px;border-bottom:1px solid #e2e8f0;text-align:center;font-weight:600;color:#1a3c6e">{score}</td>
          <td style="padding:14px 20px;border-bottom:1px solid #e2e8f0;text-align:center;color:#64748b">{mins} min</td>
          <td style="padding:14px 20px;border-bottom:1px solid #e2e8f0;text-align:center"><span style="color:{status_color};font-weight:600">{status}</span></td>
        </tr>"""

    _render_html(f"""
<style>
.cert-v2 {{
  max-width: 800px;
  margin: 0 auto;
  background: #fff;
  border-radius: 24px;
  padding: 56px;
  border: 1px solid #e2e8f0;
  box-shadow: 0 24px 48px rgba(0,0,0,.08);
  position: relative;
  overflow: hidden;
}}
.cert-v2::before {{
  content: '';
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  height: 6px;
  background: linear-gradient(90deg, #0f2340, #1a3c6e, #c9a84c, #1a3c6e, #0f2340);
}}
.cert-v2::after {{
  content: '';
  position: absolute;
  inset: 16px;
  border: 2px solid rgba(201,168,76,.2);
  border-radius: 16px;
  pointer-events: none;
}}
.cert-v2 .header {{
  text-align: center;
  margin-bottom: 40px;
  position: relative;
}}
.cert-v2 .badge {{
  display: inline-flex;
  align-items: center;
  gap: 8px;
  padding: 8px 20px;
  background: linear-gradient(135deg, rgba(201,168,76,.1), rgba(201,168,76,.05));
  border: 1px solid rgba(201,168,76,.3);
  border-radius: 100px;
  font-size: .75rem;
  font-weight: 600;
  color: #c9a84c;
  text-transform: uppercase;
  letter-spacing: 1px;
  margin-bottom: 20px;
}}
.cert-v2 .title {{
  font-size: 2rem;
  font-weight: 800;
  color: #0f172a;
  letter-spacing: 2px;
  margin-bottom: 8px;
}}
.cert-v2 .subtitle {{
  font-size: 1rem;
  color: #94a3b8;
  font-weight: 400;
}}
.cert-v2 .body {{
  text-align: center;
  margin-bottom: 32px;
  font-size: 1rem;
  color: #475569;
  line-height: 1.8;
}}
.cert-v2 .body strong {{
  color: #1a3c6e;
  font-weight: 600;
}}
.cert-v2 table {{
  width: 100%;
  border-collapse: collapse;
  margin: 24px 0;
}}
.cert-v2 th {{
  background: #0f172a;
  color: #fff;
  padding: 14px 20px;
  font-weight: 600;
  font-size: .8rem;
  text-transform: uppercase;
  letter-spacing: .5px;
}}
.cert-v2 th:first-child {{
  border-radius: 10px 0 0 0;
}}
.cert-v2 th:last-child {{
  border-radius: 0 10px 0 0;
}}
.cert-v2 .seal {{
  display: flex;
  justify-content: center;
  margin: 32px 0;
}}
.cert-v2 .seal-icon {{
  width: 100px;
  height: 100px;
  border: 3px solid #dc2626;
  border-radius: 50%;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  color: #dc2626;
  font-weight: 800;
  font-size: .7rem;
  transform: rotate(-12deg);
  line-height: 1.3;
}}
.cert-v2 .signatures {{
  display: flex;
  justify-content: space-around;
  margin-top: 40px;
}}
.cert-v2 .sig {{
  text-align: center;
  color: #64748b;
  font-size: .85rem;
}}
.cert-v2 .sig .line {{
  width: 160px;
  border-bottom: 2px solid #e2e8f0;
  margin: 12px auto 8px;
}}
.cert-v2 .footer {{
  text-align: center;
  margin-top: 32px;
  padding-top: 20px;
  border-top: 1px solid #e2e8f0;
  font-size: .8rem;
  color: #94a3b8;
}}
</style>
<div class="cert-v2">
  <div class="header">
    <div class="badge">🏆 Official Certificate</div>
    <div class="title">CERTIFICATE OF COMPLETION</div>
    <div class="subtitle">Training Program Certification</div>
  </div>

  <div class="body">
    This is to certify that <strong>{name}</strong><br>
    Employee ID: <strong>{emp_id}</strong> | Department: <strong>{dept}</strong><br>
    has successfully completed all training modules and passed the examinations.
  </div>

  <table>
    <thead>
      <tr>
        <th style="text-align:left">模块</th>
        <th>分数</th>
        <th>时长</th>
        <th>结果</th>
      </tr>
    </thead>
    <tbody>
      {rows}
    </tbody>
  </table>

  <div class="seal">
    <div class="seal-icon">
      YCF<br>SUPPLY<br>CHAIN
    </div>
  </div>

  <div class="signatures">
    <div class="sig">
      <div class="line"></div>
      Training Manager
    </div>
    <div class="sig">
      <div class="line"></div>
      HR Department
    </div>
  </div>

  <div class="footer">
    Certificate No: YCF-{date}-{emp_id} | Issued: {date}
  </div>
</div>
""", height=750)
