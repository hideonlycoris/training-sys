"""
Training System Brand CSS Styles
"""
import streamlit.components.v1 as components

def _render_html(html_content: str, height: int = None):
    """Helper to render HTML via components.html for Streamlit Cloud compatibility"""
    components.html(html_content, height=height or 100, scrolling=True)

BRAND_CSS = """
<style>
:root {
  --primary: #1a3c6e;
  --primary-light: #2a5ca8;
  --accent: #c9a84c;
  --bg: #f4f6f9;
  --card: #fff;
  --text: #222;
  --text-light: #666;
  --success: #27ae60;
  --danger: #e74c3c;
  --warning: #f39c12;
  --radius: 12px;
  --shadow: 0 4px 20px rgba(0,0,0,.08);
}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
.stApp {
  font-family: 'Noto Sans SC', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
  background: var(--bg);
}
.brand-bar {
  background: linear-gradient(135deg, #0f2647 0%, #1a3c6e 60%, #2a5ca8 100%);
  color: #fff;
  padding: 16px 32px;
  margin: -1rem -1rem 1rem -1rem;
  display: flex;
  align-items: center;
  justify-content: space-between;
  box-shadow: 0 4px 20px rgba(0,0,0,.2);
  position: relative;
  overflow: hidden;
}
.brand-bar::before {
  content: '';
  position: absolute;
  width: 300px;
  height: 300px;
  border-radius: 50%;
  background: rgba(201,168,76,.06);
  top: -150px;
  right: -80px;
}
.brand-bar .brand-left {
  display: flex;
  align-items: center;
  gap: 14px;
  position: relative;
  z-index: 1;
}
.brand-bar .brand-icon {
  width: 42px;
  height: 42px;
  background: rgba(201,168,76,.2);
  border: 2px solid rgba(201,168,76,.4);
  border-radius: 10px;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 1.3rem;
}
.brand-bar .brand-text h1 {
  font-size: 1.1rem;
  font-weight: 700;
  margin: 0;
  letter-spacing: 1px;
}
.brand-bar .brand-text p {
  font-size: .75rem;
  margin: 2px 0 0;
  opacity: .75;
  letter-spacing: .5px;
}
.brand-bar .brand-right {
  display: flex;
  align-items: center;
  gap: 16px;
  font-size: .85rem;
  position: relative;
  z-index: 1;
}
.brand-bar .user-chip {
  background: rgba(255,255,255,.12);
  padding: 6px 16px;
  border-radius: 20px;
  backdrop-filter: blur(4px);
}
.brand-bar .dept-badge {
  background: rgba(201,168,76,.2);
  color: #c9a84c;
  padding: 4px 12px;
  border-radius: 12px;
  font-size: .75rem;
  font-weight: 600;
}
.login-wrapper {
  min-height: 100vh;
  display: flex;
  align-items: center;
  justify-content: center;
  background: linear-gradient(135deg, #0f2647 0%, #1a3c6e 40%, #2a5ca8 100%);
  position: relative;
  overflow: hidden;
  margin: -60px -20px -20px -20px;
  padding: 40px 20px;
}
.login-wrapper::before {
  content: '';
  position: absolute;
  width: 600px;
  height: 600px;
  border-radius: 50%;
  background: rgba(201,168,76,.08);
  top: -200px;
  right: -150px;
}
.login-wrapper::after {
  content: '';
  position: absolute;
  width: 400px;
  height: 400px;
  border-radius: 50%;
  background: rgba(255,255,255,.03);
  bottom: -100px;
  left: -100px;
}
.login-card {
  background: #fff;
  border-radius: 20px;
  padding: 48px 44px;
  width: 420px;
  max-width: 92vw;
  box-shadow: 0 24px 64px rgba(0,0,0,.25);
  position: relative;
  z-index: 1;
}
.login-card .logo-area {
  text-align: center;
  margin-bottom: 32px;
}
.login-card .logo-area .company-icon {
  width: 72px;
  height: 72px;
  background: linear-gradient(135deg, #1a3c6e, #2a5ca8);
  border-radius: 18px;
  display: flex;
  align-items: center;
  justify-content: center;
  margin: 0 auto 16px;
  font-size: 2rem;
  box-shadow: 0 8px 24px rgba(26,60,110,.3);
}
.login-card .logo-area h1 {
  font-size: 1.4rem;
  color: var(--primary);
  font-weight: 700;
  margin-bottom: 4px;
}
.login-card .logo-area p {
  font-size: .85rem;
  color: var(--text-light);
}
.dashboard-header {
  margin-bottom: 24px;
}
.dashboard-header h2 {
  font-size: 1.5rem;
  color: var(--primary);
  font-weight: 700;
  margin-bottom: 4px;
}
.dashboard-header p {
  color: var(--text-light);
  font-size: .9rem;
}
.progress-section {
  background: var(--card);
  border-radius: var(--radius);
  padding: 24px 28px;
  margin-bottom: 24px;
  box-shadow: var(--shadow);
}
.progress-outer {
  background: #e8ebf0;
  border-radius: 20px;
  height: 22px;
  overflow: hidden;
  margin: 12px 0 6px;
}
.progress-inner {
  height: 100%;
  background: linear-gradient(90deg, var(--primary), var(--accent));
  border-radius: 20px;
  transition: width .6s ease;
  min-width: 0;
}
.progress-label {
  font-size: .82rem;
  color: var(--text-light);
  text-align: right;
  margin-top: 4px;
}
.module-card {
  background: var(--card);
  border-radius: var(--radius);
  padding: 24px;
  box-shadow: var(--shadow);
  border-left: 5px solid var(--primary);
  transition: transform .2s, box-shadow .2s;
  margin-bottom: 16px;
  height: 100%;
}
.module-card:hover {
  transform: translateY(-3px);
  box-shadow: 0 8px 30px rgba(0,0,0,.12);
}
.module-card.completed {
  border-left-color: var(--success);
}
.module-card.in-progress {
  border-left-color: var(--warning);
}
.module-card h3 {
  font-size: 1.05rem;
  margin-bottom: 8px;
  color: var(--primary);
  font-weight: 600;
}
.module-card .meta {
  font-size: .82rem;
  color: var(--text-light);
  margin-bottom: 12px;
}
.module-card .meta span {
  margin-right: 16px;
}
.status-badge {
  display: inline-block;
  padding: 4px 14px;
  border-radius: 20px;
  font-size: .75rem;
  font-weight: 600;
}
.status-badge.not-started {
  background: #f0f0f0;
  color: #999;
}
.status-badge.in-progress {
  background: #fff3cd;
  color: #856404;
}
.status-badge.done {
  background: #d4edda;
  color: #155724;
}
.module-content-header {
  margin-bottom: 24px;
}
.module-content-header h2 {
  font-size: 1.3rem;
  color: var(--primary);
  font-weight: 700;
}
.exam-header-bar {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 24px;
  flex-wrap: wrap;
  gap: 12px;
}
.exam-header-bar h2 {
  font-size: 1.25rem;
  color: var(--primary);
  font-weight: 700;
}
.exam-timer {
  background: var(--danger);
  color: #fff;
  padding: 8px 24px;
  border-radius: 24px;
  font-size: 1.1rem;
  font-weight: 700;
  font-variant-numeric: tabular-nums;
  box-shadow: 0 4px 12px rgba(231,76,60,.3);
  text-align: center;
}
.exam-timer.warning {
  animation: pulse 1s infinite;
}
@keyframes pulse {
  0%, 100% { opacity: 1; }
  50% { opacity: .6; }
}
.question-card {
  background: var(--card);
  border-radius: var(--radius);
  padding: 24px;
  margin-bottom: 16px;
  box-shadow: var(--shadow);
  border-left: 4px solid var(--accent);
}
.question-card .q-number {
  font-size: .8rem;
  color: var(--accent);
  font-weight: 700;
  margin-bottom: 6px;
  text-transform: uppercase;
}
.question-card .q-text {
  font-size: .95rem;
  font-weight: 600;
  margin-bottom: 16px;
  line-height: 1.6;
  color: var(--text);
}
.question-card .q-type-badge {
  display: inline-block;
  padding: 2px 10px;
  border-radius: 12px;
  font-size: .7rem;
  font-weight: 600;
  margin-left: 8px;
}
.q-type-badge.single { background: #e3f2fd; color: #1565c0; }
.q-type-badge.multi { background: #f3e5f5; color: #7b1fa2; }
.q-type-badge.short { background: #e8f5e9; color: #2e7d32; }
.result-container {
  text-align: center;
  padding: 40px 20px;
}
.result-circle {
  width: 160px;
  height: 160px;
  border-radius: 50%;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  margin: 0 auto 24px;
  font-size: 2.4rem;
  font-weight: 900;
  box-shadow: 0 8px 30px rgba(0,0,0,.1);
}
.result-circle.pass {
  background: linear-gradient(135deg, #d4edda, #b7e4c7);
  color: var(--success);
}
.result-circle.fail {
  background: linear-gradient(135deg, #fde8e8, #f5c6c6);
  color: var(--danger);
}
.result-circle small {
  font-size: .8rem;
  font-weight: 400;
}
.cert-preview {
  background: #fffef7;
  border: 3px solid var(--accent);
  border-radius: 8px;
  padding: 48px;
  text-align: center;
  position: relative;
  margin-bottom: 24px;
  box-shadow: 0 8px 32px rgba(201,168,76,.15);
  max-width: 700px;
  margin-left: auto;
  margin-right: auto;
}
.cert-preview::before {
  content: '';
  position: absolute;
  inset: 8px;
  border: 1.5px solid var(--accent);
  border-radius: 4px;
  pointer-events: none;
}
.cert-preview .cert-title {
  font-size: 1.8rem;
  font-weight: 700;
  color: var(--primary);
  margin-bottom: 4px;
  letter-spacing: 4px;
}
.cert-preview .cert-subtitle {
  font-size: .9rem;
  color: var(--text-light);
  margin-bottom: 28px;
}
.cert-preview .cert-body {
  font-size: 1rem;
  line-height: 2.2;
  margin-bottom: 24px;
  text-align: left;
  padding: 0 20px;
}
.cert-preview .cert-body strong {
  color: var(--primary);
}
.cert-detail-table {
  width: 100%;
  border-collapse: collapse;
  margin: 20px 0;
  font-size: .85rem;
  text-align: left;
}
.cert-detail-table th {
  background: var(--primary);
  color: #fff;
  padding: 10px 14px;
  font-weight: 500;
}
.cert-detail-table td {
  padding: 10px 14px;
  border-bottom: 1px solid #e8e0c8;
}
.cert-seal {
  width: 100px;
  height: 100px;
  border: 3px solid #c0392b;
  border-radius: 50%;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  color: #c0392b;
  font-weight: 900;
  font-size: .75rem;
  margin: 20px 0;
  transform: rotate(-15deg);
  line-height: 1.3;
  text-align: center;
}
section[data-testid="stSidebar"] {
  background: linear-gradient(180deg, #0f2647 0%, #1a3c6e 100%);
}
section[data-testid="stSidebar"] .stMarkdown p,
section[data-testid="stSidebar"] .stMarkdown h1,
section[data-testid="stSidebar"] .stMarkdown h2,
section[data-testid="stSidebar"] .stMarkdown h3 {
  color: #fff !important;
}
section[data-testid="stSidebar"] .stButton > button {
  background: rgba(255,255,255,.1);
  color: #fff;
  border: 1px solid rgba(255,255,255,.15);
  border-radius: 10px;
  transition: all .2s;
}
section[data-testid="stSidebar"] .stButton > button:hover {
  background: rgba(255,255,255,.2);
  border-color: rgba(201,168,76,.4);
}
.stButton > button {
  border-radius: 10px;
  font-weight: 600;
  transition: all .2s;
}
.stButton > button:hover {
  transform: translateY(-1px);
  box-shadow: 0 4px 12px rgba(0,0,0,.15);
}
.stButton > button[kind="primary"] {
  background: linear-gradient(135deg, var(--primary), var(--primary-light));
  border: none;
}
.stButton > button[kind="primary"]:hover {
  background: linear-gradient(135deg, var(--primary-light), #3a6cb8);
  box-shadow: 0 6px 20px rgba(26,60,110,.3);
}
.stTextInput > div > div > input,
.stTextArea > div > div > textarea {
  border-radius: 10px;
  border: 2px solid #e0e4ea;
  transition: border .2s, box-shadow .2s;
}
.stTextInput > div > div > input:focus,
.stTextArea > div > div > textarea:focus {
  border-color: var(--primary-light);
  box-shadow: 0 0 0 3px rgba(42,92,168,.1);
}
@keyframes fadeIn {
  from { opacity: 0; transform: translateY(12px); }
  to { opacity: 1; transform: translateY(0); }
}
</style>
"""


def inject_brand_css():
    """Inject brand CSS into Streamlit page"""
    css_content = BRAND_CSS.replace('<style>', '').replace('</style>', '').strip()
    _render_html(f"<style>{css_content}</style>", height=0)


def render_brand_bar(display_name: str, department: str, emp_id: str = ""):
    """Render top brand bar"""
    dept_html = f'<span class="dept-badge">{department}</span>' if department else ""
    emp_html = f' | ID: {emp_id}' if emp_id else ""
    _render_html(f"""
<div class="brand-bar">
  <div class="brand-left">
    <div class="brand-icon">📦</div>
    <div class="brand-text">
      <h1>驭长风供应链</h1>
      <p>新员工培训系统</p>
    </div>
  </div>
  <div class="brand-right">
    {dept_html}
    <span class="user-chip">👤 {display_name}{emp_html}</span>
  </div>
</div>
""", height=80)


def render_login_page():
    """Render login page layout"""
    _render_html("""
<div class="login-wrapper">
  <div class="login-card">
    <div class="logo-area">
      <div class="company-icon">📦</div>
      <h1>驭长风供应链</h1>
      <p>新员工入职培训系统</p>
    </div>
  </div>
</div>
""", height=600)


def render_progress_bar(pct: int, label: str = "总进度"):
    """Render branded progress bar"""
    _render_html(f"""
<div class="progress-section">
  <div style="display:flex;justify-content:space-between;align-items:baseline">
    <strong>{label}</strong>
    <span class="progress-label">{pct}%</span>
  </div>
  <div class="progress-outer">
    <div class="progress-inner" style="width:{pct}%"></div>
  </div>
</div>
""", height=100)


def render_module_card(title: str, status: str, chapters_info: str, exam_info: str,
                       time_info: str, module_id: int):
    """Render module card"""
    status_class = {
        "已完成": "done",
        "进行中": "in-progress",
        "未开始": "not-started"
    }.get(status, "not-started")

    card_class = {
        "已完成": "completed",
        "进行中": "in-progress",
        "未开始": ""
    }.get(status, "")

    _render_html(f"""
<div class="module-card {card_class}">
  <h3>{title}</h3>
  <span class="status-badge {status_class}">{status}</span>
  <div class="meta" style="margin-top:12px">
    <span>📖 {chapters_info}</span>
    <span>📝 {exam_info}</span>
    <span>⏱️ {time_info}</span>
  </div>
</div>
""", height=120)


def render_exam_timer(minutes: int, seconds: int, is_warning: bool = False):
    """Render exam countdown timer"""
    warning_class = " warning" if is_warning else ""
    _render_html(f"""
<div class="exam-timer{warning_class}">
  ⏱️ {minutes:02d}:{seconds:02d}
</div>
""", height=60)


def render_question_card(number: int, question: str, q_type: str, per_q: int):
    """Render question card"""
    type_labels = {"single": "单选", "multi": "多选", "short": "简答"}
    type_classes = {"single": "single", "multi": "multi", "short": "short"}
    label = type_labels.get(q_type, "")
    css_class = type_classes.get(q_type, "single")

    _render_html(f"""
<div class="question-card">
  <div class="q-number">第 {number} 题 <span class="q-type-badge {css_class}">{label}</span> · {per_q}分</div>
  <div class="q-text">{question}</div>
</div>
""", height=100)


def render_result_circle(score: int, passed: bool):
    """Render score circle"""
    circle_class = "pass" if passed else "fail"
    label = "通过" if passed else "未通过"
    _render_html(f"""
<div class="result-container">
  <div class="result-circle {circle_class}">
    {score}
    <small>{label}</small>
  </div>
</div>
""", height=220)


def render_certificate_preview(name: str, emp_id: str, dept: str,
                                module_scores: list, date_str: str):
    """Render certificate preview"""
    rows_html = ""
    for mod_name, score, mins in module_scores:
        result = '✅ 通过' if score >= 80 else '❌ 未通过'
        rows_html += f"""
        <tr>
          <td>{mod_name}</td>
          <td style="text-align:center">{score}分</td>
          <td style="text-align:center">{mins}分钟</td>
          <td style="text-align:center">{result}</td>
        </tr>"""

    _render_html(f"""
<div class="cert-preview">
  <div class="cert-title">培 训 合 格 证 书</div>
  <div class="cert-subtitle">TRAINING COMPLETION CERTIFICATE</div>
  <div class="cert-body">
    兹证明 <strong>{name}</strong>（工号：{emp_id}，部门：{dept}）<br>
    已完成驭长风供应链新员工入职培训全部课程，并通过各项考核。
  </div>
  <table class="cert-detail-table">
    <thead>
      <tr>
        <th>培训模块</th>
        <th>考核成绩</th>
        <th>学习时长</th>
        <th>考核结果</th>
      </tr>
    </thead>
    <tbody>
      {rows_html}
    </tbody>
  </table>
  <div class="cert-seal">驭长风<br>供应链<br>OFFICIAL</div>
  <div style="display:flex;justify-content:space-around;margin-top:28px;font-size:.85rem;color:#666">
    <div style="text-align:center">
      <div style="width:140px;border-top:1px solid #999;margin:8px auto 4px"></div>
      培训主管
    </div>
    <div style="text-align:center">
      <div style="width:140px;border-top:1px solid #999;margin:8px auto 4px"></div>
      人力资源部
    </div>
  </div>
  <p style="margin-top:20px;font-size:.8rem;color:#999">证书编号: YCF-{date_str}-{emp_id} | 签发日期: {date_str}</p>
</div>
""", height=600)
