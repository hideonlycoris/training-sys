"""
驭长风供应链 — 新员工培训系统 (Streamlit)
Features: account management, training modules, exams, PDF certificates, course upload, analytics
"""


import re
import google.generativeai as genai
import streamlit as st
import sqlite3, hashlib, json, os, time, io, random
from datetime import datetime
from pathlib import Path
import mammoth
import openpyxl
from pptx import Presentation
from styles import inject_brand_css, render_brand_bar, render_login_page, render_progress_bar, render_module_card, render_exam_timer, render_question_card, render_result_circle, render_certificate_preview

# --- 云端适配：移除 Windows COM 依赖，云端不支持 Office 自动转 PDF ---

# --------------- FILE STORAGE (云端适配) ---------------
import urllib.parse

UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# 云端适配：移除本地服务器和 Office 转 PDF，改为直接下载/内嵌显示


# --------------- FILE STORAGE ---------------
UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# --------------- DB ---------------
DB_PATH = os.path.join(os.path.dirname(__file__), "training.db")

def get_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        display_name TEXT NOT NULL,
        emp_id TEXT DEFAULT '',
        department TEXT DEFAULT '',
        role TEXT DEFAULT 'user',
        created_at TEXT DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS progress (
        user_id INTEGER,
        module_id INTEGER,
        chapter_idx INTEGER,
        read_done INTEGER DEFAULT 0,
        PRIMARY KEY (user_id, module_id, chapter_idx)
    );
    CREATE TABLE IF NOT EXISTS exam_results (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        module_id INTEGER,
        score INTEGER,
        answers TEXT,
        taken_at TEXT DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS training_time (
        user_id INTEGER,
        module_id INTEGER,
        seconds INTEGER DEFAULT 0,
        PRIMARY KEY (user_id, module_id)
    );
    CREATE TABLE IF NOT EXISTS modules (
        id INTEGER PRIMARY KEY,
        title TEXT NOT NULL,
        chapters TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS exam_questions (
        module_id INTEGER PRIMARY KEY,
        questions TEXT NOT NULL,
        exam_count INTEGER DEFAULT 0
    );
    """)
    try:
        c.execute("ALTER TABLE exam_questions ADD COLUMN exam_count INTEGER DEFAULT 0")
    except sqlite3.OperationalError:
        pass # 如果列已经存在会报错，我们直接忽略即可
    # ====================================================

    # Seed admin account if not exists
    admin_hash = hashlib.sha256("admin123".encode()).hexdigest()
    # Seed admin account if not exists
    admin_hash = hashlib.sha256("admin123".encode()).hexdigest()
    c.execute("INSERT OR IGNORE INTO users (username, password_hash, display_name, role) VALUES (?,?,?,?)",
              ("admin", admin_hash, "管理员", "admin"))
    # Seed default modules and exams if empty
    if c.execute("SELECT COUNT(*) FROM modules").fetchone()[0] == 0:
        defaults = get_default_modules()
        for mid, mod in defaults.items():
            c.execute("INSERT INTO modules (id, title, chapters) VALUES (?,?,?)",
                      (mid, mod["title"], json.dumps(mod["chapters"], ensure_ascii=False)))
        default_exams = get_default_exams()
        for mid, qs in default_exams.items():
            c.execute("INSERT INTO exam_questions (module_id, questions) VALUES (?,?)",
                      (mid, json.dumps(qs, ensure_ascii=False)))
    conn.commit()
    conn.close()


# --------------- AUTH HELPERS ---------------
def hash_pw(pw: str) -> str:
    return hashlib.sha256(pw.encode()).hexdigest()

def authenticate(username: str, password: str):
    conn = get_db()
    row = conn.execute("SELECT * FROM users WHERE username=? AND password_hash=?",
                       (username, hash_pw(password))).fetchone()
    conn.close()
    return dict(row) if row else None

def create_user(username, password, display_name, emp_id, department, role="user"):
    conn = get_db()
    try:
        conn.execute("INSERT INTO users (username,password_hash,display_name,emp_id,department,role) VALUES (?,?,?,?,?,?)",
                     (username, hash_pw(password), display_name, emp_id, department, role))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()

def get_all_users():
    conn = get_db()
    rows = conn.execute("SELECT * FROM users ORDER BY id").fetchall()
    conn.close()
    return [dict(r) for r in rows]

def delete_user(uid):
    conn = get_db()
    conn.execute("DELETE FROM users WHERE id=?", (uid,))
    conn.execute("DELETE FROM progress WHERE user_id=?", (uid,))
    conn.execute("DELETE FROM exam_results WHERE user_id=?", (uid,))
    conn.execute("DELETE FROM training_time WHERE user_id=?", (uid,))
    conn.commit()
    conn.close()

def update_user(uid, **kwargs):
    conn = get_db()
    for k, v in kwargs.items():
        conn.execute(f"UPDATE users SET {k}=? WHERE id=?", (v, uid))
    conn.commit()
    conn.close()


# --------------- PROGRESS DB ---------------
def get_read_checks(user_id, module_id, chapter_count):
    conn = get_db()
    rows = conn.execute("SELECT chapter_idx, read_done FROM progress WHERE user_id=? AND module_id=?",
                        (user_id, module_id)).fetchall()
    conn.close()
    checks = [False] * chapter_count
    for r in rows:
        if r["chapter_idx"] < chapter_count:
            checks[r["chapter_idx"]] = bool(r["read_done"])
    return checks

def mark_chapter_read(user_id, module_id, chapter_idx):
    conn = get_db()
    conn.execute("INSERT OR REPLACE INTO progress (user_id,module_id,chapter_idx,read_done) VALUES (?,?,?,1)",
                 (user_id, module_id, chapter_idx))
    conn.commit()
    conn.close()

def save_exam_result(user_id, module_id, score, answers):
    conn = get_db()
    conn.execute("INSERT INTO exam_results (user_id,module_id,score,answers) VALUES (?,?,?,?)",
                 (user_id, module_id, score, json.dumps(answers)))
    conn.commit()
    conn.close()

def get_best_exam(user_id, module_id):
    conn = get_db()
    row = conn.execute("SELECT MAX(score) as best FROM exam_results WHERE user_id=? AND module_id=?",
                       (user_id, module_id)).fetchone()
    conn.close()
    return row["best"] if row and row["best"] is not None else None

def get_training_time(user_id, module_id):
    conn = get_db()
    row = conn.execute("SELECT seconds FROM training_time WHERE user_id=? AND module_id=?",
                       (user_id, module_id)).fetchone()
    conn.close()
    return row["seconds"] if row else 0

def add_training_time(user_id, module_id, seconds):
    conn = get_db()
    conn.execute("""INSERT INTO training_time (user_id,module_id,seconds) VALUES (?,?,?)
                    ON CONFLICT(user_id,module_id) DO UPDATE SET seconds=seconds+?""",
                 (user_id, module_id, seconds, seconds))
    conn.commit()
    conn.close()


# --------------- DEFAULT TRAINING DATA ---------------
DEPARTMENTS = [
    "运营一部", "运营二部", "运营三部", "运营四部",
    "开发部", "采购部", "财务部", "人力行政部",
    "物控组", "海运物流组", "仓储组", "品质组", "供应链综合组",
]

def get_default_modules():
    return {
        1: {
            "title": "模块一：供应链管理与滚动计划",
            "chapters": [
                {
                    "title": "1. 计划的定义与三大特征",
                    "html": "<h4>一、什么是计划？</h4><p>计划是在了解了市场与竞争对手后，结合自身供应链情况所做出的销售策略 [cite: 1]。</p><h4>二、计划的三大特征</h4><ul><li><b>所有预测都是错的，但有预测比没有强</b>：错多错少很关键，比竞争对手错的少，综合成本更低 [cite: 1]。</li><li><b>多职能参与</b>：需要开发、采购/生产、物控、财务等协同参与 [cite: 1]。</li><li><b>循环预测，逐渐逼近</b>：执行中遇到失准问题，需要立刻进行计划修正，用执行弥补不足 [cite: 1]。</li></ul>"
                },
                {
                    "title": "2. 物控的核心职能与三道防线",
                    "html": "<h4>一、物控的五大核心职能</h4><ul><li>生成需求计划并达成机会与风险平衡 [cite: 1]</li><li>根据海运和生产周期制定发货计划 [cite: 1]</li><li>与工厂沟通排产（未来2-3个月） [cite: 1]</li><li>追踪监控销售（每日/每周完成率） [cite: 1]</li><li>确立产品退出机制 [cite: 1]</li></ul><h4>二、供应链的三道防线</h4><p>1. <b>准确的计划</b>（第一道防线） [cite: 1]<br>2. <b>安全库存</b>（第二道防线）：针对重点产品采用 <b>3+2模式</b>（海外3个月，工厂2个月） [cite: 1]<br>3. <b>供应链执行</b>（第三道防线）：如果前两道防线失守，压力全转移到执行，全员陷入救火模式 [cite: 1]。</p>"
                },
                {
                    "title": "3. 滚动计划与异常管控",
                    "html": "<h4>一、滚动计划的核心逻辑</h4><p>生产周期越长的产品，供应链链条越长，反应速度越慢 [cite: 1]。例如，7月份的出货向前对应9月的销售，向后却对应5月份的生产和4月的订单 [cite: 1]。</p><h4>二、异常销售管控原则</h4><ul><li><b>无明显波动不变更</b>：哪怕能卖更多，也要通过涨价、控广告将销量控制在20%上下浮动 [cite: 1]。</li><li><b>超预期断货</b>：如果失控，要及时预估断货月份，并将该月的销售计划调整为0 [cite: 1]。</li><li><b>环环相扣</b>：每月做新计划时，必须对照上月提供的计划记录 [cite: 1]。</li></ul>"
                },
                {
                    "title": "4. 产品品级定义",
                    "html": "<h4>一、四大品级定义</h4><ul><li><b>爆品</b>：类目Top，超额准备库存，投入最好资源。变体数量严格控制在 <b>4个以内</b> [cite: 1]。</li><li><b>利润品</b>：毛利率15%以上（美国市场标准） [cite: 1]。</li><li><b>新品</b>：上线三个月内，随后定级，决定是返单还是清尾 [cite: 1]。</li><li><b>清尾品</b>：月销售低于 30件 或 毛利率小于 5% [cite: 1]。</li></ul>"
                }
            ]
        },
        2: {
            "title": "模块二：海外仓物流与发货规范",
            "chapters": [
                {
                    "title": "1. FBA直发与AGL规范",
                    "html": "<h4>一、提报计划与发货标准</h4><ul><li>每月 <b>20号</b> 提交未来五个月销售计划及三个月发货计划 [cite: 3]。</li><li>FBA国内直发：<b>2CBM起收</b>，不满按2CBM收费 [cite: 3]。</li><li>AGL（亚马逊物流）：<b>1CBM起收</b>，优点是不分仓免锁仓费、上架快 [cite: 2, 3]。</li></ul><h4>二、操作规范</h4><ul><li>严格按审核后的计划数建货件，超出将被物流部<b>驳回不发</b> [cite: 3]。</li><li>AGL后台“货好时间”最少多预留 <b>3个工作日</b> 给物流部贴标和预约，否则逾期需重新排船 [cite: 3]。</li><li>文件命名：`BOX LABEL-SKU-FBA ID-箱数-地址代称`，至少<b>提前一周</b>给物流部 [cite: 3]。</li></ul>"
                },
                {
                    "title": "2. 海外仓费用与附加费机制",
                    "html": "<h4>一、基础费用</h4><p>主要包括入库费、仓储费（按体积）、出库费 [cite: 4]。</p><h4>二、尾程快递附加费 (重点)</h4><ul><li><b>AHS额外处理费</b>：最长边超48in / 次长边超30in / 围长>105in / 单件超50lb（符合其一即收） [cite: 4]。</li><li><b>Oversize超尺寸附加费</b>：最长边超96in 或 围长>130in [cite: 4]。</li></ul><h4>三、Wayfair仓省钱策略</h4><p>2026年1月起，110lb以上将加收oversize费（约40-50美金），而Wayfair仓暂不加收。建议如劈材器、吹雪机试发Wayfair仓节约尾程成本 [cite: 4]。</p>"
                },
                {
                    "title": "3. 时效与索赔协同",
                    "html": "<h4>一、时效标准</h4><ul><li>散货到港后（涉拆柜打托），送仓时效 <b>7-12天</b> [cite: 3]。</li><li>海外仓出库时效 <b>1-2个工作日</b> [cite: 4]。</li><li>尾程超 <b>7个工作日</b> 无更新视为丢件，可发起索赔 [cite: 4]。</li></ul><h4>二、退件与索赔协同</h4><ul><li>退件处理需在销毁时间前填写意见，逾期 <b>默认销毁</b> [cite: 4]。</li><li>索赔登记必须提供规定的必填照片，否则降低成功率。索赔周期为1-3个月，超3个月默认失败 [cite: 4]。</li></ul>"
                }
            ]
        },
        3: {
            "title": "模块三：易仓系统操作与异常处理",
            "chapters": [
                {
                    "title": "1. 订单自动化与推单",
                    "html": "<h4>一、自动与人工推单</h4><p>开启自动推单后，系统会根据收货人距离<b>自动分仓</b>，匹配运费最低的仓库和渠道，切勿手动选择 [cite: 5]。</p><p>注意：金额为0的<b>换货订单</b>不在自动推单范围内，仍需手动检查 [cite: 5]。</p><h4>二、订单拦截（作废）</h4><p>在“待发货”中勾选点击截单，转至问题件，<b>绝对不要勾选</b>“确认服务商已删除订单”。转为问题件后再作废即可 [cite: 5]。</p>"
                },
                {
                    "title": "2. 团队库存与缺货排查",
                    "html": "<h4>一、如何查看团队库存？</h4><p>顶部导航栏点击“仓储” -> “库存查询（团队）”，输入SKU即可查看本部门的可用库存 [cite: 5]。</p><h4>二、有库存却显示“缺货订单”的原因</h4><ul><li><b>Listing未绑定负责人</b>：导致系统无法识别团队库存。绑好负责人后，需在待发货审核处点击“批量更新团队” [cite: 5]。</li><li><b>系统拆单Bug</b>：当仓库只剩1件，而订单要2件时，拆单会锁定这1件导致发走一单缺货一单 [cite: 5]。</li></ul>"
                },
                {
                    "title": "3. 常见异常件处理",
                    "html": "<h4>一、推送失败异常</h4><ul><li>提示 <b>duplication</b>：在原单号后加上“-1”重新新建订单 [cite: 5]。</li><li>分仓异常（如EfurdenUS）：可能是SKU未建立映射，导致海外仓无法识别 [cite: 5]。</li><li>地址异常：收件人名字太短（少于2个字符）或太长（超20个字符），以及地址中包含 PO BOX 字眼 [cite: 5]。</li></ul><h4>二、试算物流费</h4><p>运费主要由<b>邮编</b>决定，试算时随意输入收件人姓名和地址，只要替换为目标邮编即可出结果 [cite: 5]。</p>"
                }
            ]
        }
    }

def get_default_exams():
    return {
        1: [
            {"q": "计划的三个特征中，第一个是什么？", "opts": ["所有预测都是错的但有预测比没有强", "预测只需销售部门参与", "计划制定后不需要修改", "预测越乐观越好"], "ans": 0},
            {"q": "供应链第二道防线是？", "opts": ["准确计划", "安全库存", "供应链执行", "价格调整"], "ans": 1},
            {"q": "重点产品备货模式是？", "opts": ["海外1+工厂1", "海外2+工厂3", "海外3+工厂2", "海外5个月"], "ans": 2},
            {"q": "物控核心职能不包括？", "opts": ["生成需求计划", "制定发货计划", "产品广告投放", "产品退出机制"], "ans": 2},
            {"q": "计划达成率低于多少属于库存滞销？", "opts": ["90%", "80%", "70%", "60%"], "ans": 2},
            {"q": "滚动计划中完成率正负多少以内视为合理？", "opts": ["10%", "15%", "20%", "30%"], "ans": 2},
            {"q": "以60天生产+45天海运，1月销售的货最晚何时发走？", "opts": ["前一年10月", "前一年11月初", "前一年12月", "当年1月初"], "ans": 1},
            {"q": "清尾品标准是？", "opts": ["月销<100或毛利<10%", "月销<50或毛利<8%", "月销<30或毛利<5%", "月销<20或毛利<3%"], "ans": 2},
            {"q": "爆品变体数量最多不超过？", "opts": ["2个", "3个", "4个", "6个"], "ans": 2},
            {"q": "采购订单从提交到工厂排产一般需要？", "opts": ["3-5天", "7-14天", "15-20天", "21-30天"], "ans": 1},
        ],
        2: [
            {"q": "每月几号提交滚动销售计划？", "opts": ["1号", "10号", "15号", "20号"], "ans": 3},
            {"q": "FBA国内直发最低按多少CBM收费？", "opts": ["1CBM", "2CBM", "3CBM", "5CBM"], "ans": 1},
            {"q": "AGL最低按多少CBM收费？", "opts": ["0.5CBM", "1CBM", "2CBM", "3CBM"], "ans": 1},
            {"q": "AGL优点不包括？", "opts": ["官方物流有保障", "上架时效快", "不分仓免锁仓费", "旺季运力充足"], "ans": 3},
            {"q": "直发条码文件至少提前多久给物流部？", "opts": ["3天", "5天", "一周", "两周"], "ans": 2},
            {"q": "箱唛打印格式要求？", "opts": ["A4纸", "热敏4×6英寸", "热敏6×8英寸", "A5纸"], "ans": 1},
            {"q": "整柜到港送仓正常时效？", "opts": ["3天", "5天", "7天内", "14天"], "ans": 2},
            {"q": "带电产品发货不需要提供的资料？", "opts": ["MSDS", "UN38.3报告", "海运鉴定报告", "使用说明书"], "ans": 3},
            {"q": "AGL贸易术语应选择？", "opts": ["到岸价", "到门价", "工厂交货", "离岸价"], "ans": 2},
            {"q": "超出审核数量的货件物流部如何处理？", "opts": ["正常发货", "延迟发货", "驳回不发", "加收费用"], "ans": 2},
        ],
        3: [
            {"q": "开启自动推单后，系统会根据什么自动分仓？", "opts": ["商品重量", "收货人距离", "仓库库存", "订单金额"], "ans": 1},
            {"q": "金额为0的换货订单如何处理？", "opts": ["自动推单", "需要手动检查", "直接删除", "忽略不处理"], "ans": 1},
            {"q": "订单拦截（作废）的正确操作是？", "opts": ["直接勾选确认删除", "转为问题件后再作废", "联系客服处理", "等待系统自动处理"], "ans": 1},
            {"q": "查看团队库存的路径是？", "opts": ["首页→库存查询", "仓储→库存查询（团队）", "订单→库存管理", "设置→库存查看"], "ans": 1},
            {"q": "有库存却显示缺货订单，可能的原因是？", "opts": ["库存数据错误", "Listing未绑定负责人", "系统服务器故障", "网络连接问题"], "ans": 1},
            {"q": "系统拆单Bug是指什么情况？", "opts": ["订单被重复拆分", "仓库只剩1件但订单要2件", "拆单后无法追踪", "拆单导致运费增加"], "ans": 1},
            {"q": "推送失败提示duplication时如何处理？", "opts": ["重新创建订单", "在原单号后加\"-1\"重新新建", "联系技术支持", "删除原订单重下"], "ans": 1},
            {"q": "地址异常不包括以下哪种情况？", "opts": ["收件人名字太短", "地址包含PO BOX", "邮编格式错误", "收件人名字太长"], "ans": 2},
            {"q": "试算物流费时主要由什么决定？", "opts": ["商品重量", "邮编", "订单金额", "配送时效"], "ans": 1},
            {"q": "海外仓出库时效标准是？", "opts": ["当天", "1-2个工作日", "3-5个工作日", "一周内"], "ans": 1},
        ],
    }

# --------------- COURSE PARSER ---------------
def parse_docx_to_chapters(file_bytes: bytes) -> list:
    """智能 Word 解析器：支持基于编号特征（一、/1./第一部分）的语义切分"""
    
    # 转换 Word 为 HTML
    result = mammoth.convert_to_html(io.BytesIO(file_bytes))
    html = result.value

    # 定义章节标题的识别特征（匹配：一、 1. 第一章 第一部分 等）
    # 同时保留原有的 h1-h3 标签识别
    chapter_regex = re.compile(
        r'(<h[1-3][^>]*>.*?</h[1-3]>|' # 标准标题标签
        r'<p><strong>\s*[一二三四五六七八九十]+[、.].*?</strong></p>|' # 粗体中文编号
        r'<p>\s*[一二三四五六七八九十]+[、.].*?</p>|' # 普通中文编号
        r'<p><strong>\s*\d+[、.].*?</strong></p>|' # 粗体数字编号
        r'<p>\s*\d+[、.].*?</p>|' # 普通数字编号
        r'<p><strong>\s*第[一二三四五六七八九十]+\s*[章节部].*?</strong></p>|' # 第X章
        r'<p>\s*第[一二三四五六七八九十]+\s*[章节部].*?</p>)', # 第X章
        re.IGNORECASE
    )

    parts = chapter_regex.split(html)
    chapters = []
    current_title = "前言/说明"
    current_content = ""

    for part in parts:
        if not part.strip(): continue
        
        # 判断这一段是否匹配标题特征
        if chapter_regex.match(part):
            # 如果之前的章节有实质内容，先保存
            # 过滤掉只有几个字且无后续内容的“假章节”
            if current_content.strip() and len(re.sub(r'<[^>]+>', '', current_content).strip()) > 5:
                chapters.append({"title": current_title, "html": current_content})
            
            # 提取新标题纯文本作为章节名
            new_title = re.sub(r'<[^>]+>', '', part).strip()
            # 清洗掉标题末尾可能的标点
            current_title = re.sub(r'[、.]$', '', new_title)
            current_content = f"<h4>{new_title}</h4>" # 章节内容以标题开始
        else:
            current_content += part

    # 收尾
    if current_content.strip():
        chapters.append({"title": current_title, "html": current_content})

    # 如果解析出来的章节太少且内容很多，尝试按强行平分逻辑（兜底方案）
    if len(chapters) <= 1 and len(html) > 2000:
        # 这里可以加入进一步细分的逻辑，目前先返回整体
        pass

    return chapters

def parse_pptx_to_chapters(file_bytes: bytes) -> list:
    """智能 PPT 解析器：支持过渡页自动合并，确保每个章节都有其实际内容"""
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    prs = Presentation(io.BytesIO(file_bytes))
    temp_slides = []

    def extract_html(shape):
        res = ""
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                t = para.text.strip()
                if t: res += f"<p>{t}</p>"
        elif shape.has_table:
            res += "<table border='1' style='border-collapse:collapse; width:100%; font-size:14px;'>"
            for row in shape.table.rows:
                res += "<tr>"
                for cell in row.cells:
                    txt = "<br>".join([p.text.strip() for p in cell.text_frame.paragraphs if p.text.strip()])
                    res += f"<td style='padding:4px; border:1px solid #ccc;'>{txt}</td>"
                res += "</tr>"
            res += "</table>"
        elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for s in shape.shapes: res += extract_html(s)
        return res

    # 1. 提取所有页面的内容
    for slide in prs.slides:
        title = slide.shapes.title.text.strip() if slide.shapes.title else ""
        content_html = ""
        for shape in slide.shapes:
            if shape == slide.shapes.title: continue
            content_html += extract_html(shape)
        temp_slides.append({"title": title, "content": content_html})

    # 2. 智能逻辑合并
    final_chapters = []
    buffer_title = ""
    buffer_content = ""

    for i, slide in enumerate(temp_slides):
        # 如果本页有标题
        if slide["title"]:
            # 如果之前 buffer 里已经攒了内容，先结算
            if buffer_content.strip():
                final_chapters.append({"title": buffer_title or "模块开始", "html": f"<h4>{buffer_title}</h4>{buffer_content}"})
                buffer_content = ""
            
            # 更新当前的标题 buffer
            buffer_title = slide["title"]
            buffer_content += slide["content"]
        else:
            # 如果本页没标题，说明是上一页的延续内容，直接叠加内容
            if not buffer_title and i == 0: buffer_title = "课程介绍"
            buffer_content += slide["content"]

        # 特殊处理：如果这页结束后 buffer 里的字数太少（说明可能是纯过渡页），
        # 不结算，继续往后攒，直到遇到下一个有内容的页面
    
    # 最后的结算
    if buffer_title or buffer_content:
        # 修正标题，防止只有“PART ONE”这种没意义的标题
        display_title = buffer_title if len(buffer_title) > 2 else "课程章节"
        final_chapters.append({"title": display_title, "html": f"<h4>{buffer_title}</h4>{buffer_content}"})

    # 3. 二次清洗：剔除只有“目录”、“谢谢观看”等无意义章节
    cleaned_chapters = []
    useless_keywords = ["目录", "CONTENTS", "CONTENT", "谢谢", "THANK", "Q&A", "封面"]
    for c in final_chapters:
        text_content = re.sub(r'<[^>]+>', '', c["html"]).strip()
        if any(kw in c["title"].upper() for kw in useless_keywords) and len(text_content) < 50:
            continue
        cleaned_chapters.append(c)

    return cleaned_chapters

def parse_exam_xlsx(file_bytes: bytes) -> list:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
    ws = wb.active
    questions = []
    ans_map = {"A": 0, "B": 1, "C": 2, "D": 3}
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
        if not row or not row[0]:
            continue
        q = str(row[0]).strip()
        opts = [str(row[j] or "").strip() for j in range(1, 5)]
        ans_raw = str(row[5] or "").strip() if len(row) > 5 else ""
        qtype_raw = str(row[6] or "").strip() if len(row) > 6 else ""

        qtype = "single"
        if qtype_raw in ("多选", "multi", "多选题"):
            qtype = "multi"
        elif qtype_raw in ("简答", "short", "简答题"):
            qtype = "short"
        elif len(ans_raw) > 1 and all(c in "ABCDabcd" for c in ans_raw.replace(",", "").replace("，", "")):
            qtype = "multi"

        if qtype == "short":
            keywords = [k.strip() for k in ans_raw.replace("，", "|").replace(",", "|").split("|") if k.strip()]
            questions.append({"q": q, "type": "short", "keywords": keywords, "answer": ans_raw})
        elif qtype == "multi":
            letters = [c.upper() for c in ans_raw.replace(",", "").replace("，", "") if c.upper() in ans_map]
            ans_indices = sorted(set(ans_map[c] for c in letters))
            valid_opts = [o for o in opts if o]
            if not valid_opts or not ans_indices:
                continue
            questions.append({"q": q, "opts": valid_opts, "ans": ans_indices, "type": "multi"})
        else:
            ans = ans_map.get(ans_raw.upper())
            if any(not o for o in opts) or ans is None:
                continue
            questions.append({"q": q, "opts": opts, "ans": ans, "type": "single"})
    return questions

def parse_exam_docx(file_bytes: bytes) -> list:
    result = mammoth.extract_raw_text(io.BytesIO(file_bytes))
    lines = [l.strip() for l in result.value.split("\n") if l.strip()]
    questions = []
    i = 0
    opt_pat = re.compile(r'^[A-Da-d][.、．:：\s]')
    ans_line_pat = re.compile(r'(?:答案|answer|正确答案)\s*[：:]\s*(.*)', re.I)
    ans_map = {"A": 0, "B": 1, "C": 2, "D": 3}

    while i < len(lines):
        q_line = lines[i]
        if i + 4 < len(lines):
            cands = lines[i+1:i+5]
            if all(opt_pat.match(c) for c in cands):
                strip_opt = lambda s: re.sub(r'^[A-Da-d][.、．:：\s]\s*', '', s)
                opts = [strip_opt(c) for c in cands]
                if i + 5 < len(lines):
                    ans_match = ans_line_pat.search(lines[i+5])
                    if ans_match:
                        ans_text = ans_match.group(1).strip().upper()
                        letters = [c for c in ans_text.replace(",", "").replace("，", "") if c in ans_map]
                        clean_q = re.sub(r'^[\d]+[.、．:：\s]\s*', '', q_line)
                        if len(letters) > 1:
                            questions.append({"q": clean_q or q_line, "opts": opts,
                                              "ans": sorted(set(ans_map[c] for c in letters)), "type": "multi"})
                        elif len(letters) == 1:
                            questions.append({"q": clean_q or q_line, "opts": opts,
                                              "ans": ans_map[letters[0]], "type": "single"})
                        i += 6
                        continue
                i += 1
                continue

        if i + 1 < len(lines):
            ans_match = ans_line_pat.search(lines[i+1])
            if ans_match:
                ans_text = ans_match.group(1).strip()
                clean_q = re.sub(r'^[\d]+[.、．:：\s]\s*', '', q_line)
                keywords = [k.strip() for k in ans_text.replace("，", "|").replace(",", "|").split("|") if k.strip()]
                questions.append({"q": clean_q or q_line, "type": "short",
                                  "keywords": keywords, "answer": ans_text})
                i += 2
                continue
        i += 1
    return questions

# --------------- PDF CERTIFICATE ---------------
def generate_certificate_pdf(name, emp_id, dept, module_scores, module_times):
    from fpdf import FPDF

    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()
    W, H = 297, 210

    pdf.set_draw_color(201, 168, 76)
    pdf.set_line_width(1.5)
    pdf.rect(8, 8, W - 16, H - 16)
    pdf.set_line_width(0.5)
    pdf.rect(12, 12, W - 24, H - 24)

    pdf.set_font("Helvetica", "B", 26)
    pdf.set_text_color(26, 60, 110)
    pdf.set_xy(0, 30)
    pdf.cell(W, 10, "TRAINING COMPLETION CERTIFICATE", align="C")

    pdf.set_font("Helvetica", "", 14)
    pdf.set_text_color(100, 100, 100)
    pdf.set_xy(0, 44)
    pdf.cell(W, 8, "YuChangFeng Supply Chain Training", align="C")

    pdf.set_font("Helvetica", "", 11)
    pdf.set_text_color(50, 50, 50)
    pdf.set_xy(0, 62)
    pdf.cell(W, 8, f"Employee: {name}    ID: {emp_id}    Department: {dept}", align="C")
    pdf.set_xy(0, 74)
    pdf.cell(W, 8, "Has successfully completed the New Employee Onboarding Training Program", align="C")
    pdf.set_xy(0, 82)
    pdf.cell(W, 8, "and passed all required examinations.", align="C")

    y = 98
    pdf.set_fill_color(26, 60, 110)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 9)
    col_w = [80, 40, 40, 37]
    headers = ["Module", "Score", "Duration", "Result"]
    x_start = (W - sum(col_w)) / 2
    x = x_start
    for i, h in enumerate(headers):
        pdf.set_xy(x, y)
        pdf.cell(col_w[i], 8, h, border=1, fill=True, align="C")
        x += col_w[i]

    pdf.set_text_color(50, 50, 50)
    pdf.set_font("Helvetica", "", 9)
    y += 8
    for mod_name, score, mins in module_scores:
        x = x_start
        vals = [mod_name, str(score), f"{mins} min", "PASS"]
        for i, v in enumerate(vals):
            pdf.set_xy(x, y)
            pdf.cell(col_w[i], 8, v, border=1, align="C")
            x += col_w[i]
        y += 8

    pdf.set_draw_color(192, 57, 43)
    pdf.set_line_width(0.8)
    cx, cy = W / 2, 155
    pdf.circle(cx, cy, 14, style="D")
    pdf.set_font("Helvetica", "B", 7)
    pdf.set_text_color(192, 57, 43)
    pdf.set_xy(cx - 20, cy - 4)
    pdf.cell(40, 5, "YUCHANGFENG", align="C")
    pdf.set_xy(cx - 20, cy + 1)
    pdf.cell(40, 5, "OFFICIAL SEAL", align="C")

    pdf.set_draw_color(100, 100, 100)
    pdf.set_line_width(0.3)
    pdf.line(50, 182, 110, 182)
    pdf.line(187, 182, 247, 182)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(100, 100, 100)
    pdf.set_xy(50, 184)
    pdf.cell(60, 6, "Training Manager", align="C")
    pdf.set_xy(187, 184)
    pdf.cell(60, 6, "HR Department", align="C")

    now = datetime.now()
    cert_no = f"YCF-{now.strftime('%Y-%m%d')}-{emp_id}"
    pdf.set_font("Helvetica", "", 7)
    pdf.set_text_color(150, 150, 150)
    pdf.set_xy(16, H - 14)
    pdf.cell(100, 5, f"Certificate No: {cert_no}")
    pdf.set_xy(W - 116, H - 14)
    pdf.cell(100, 5, f"Date: {now.strftime('%Y-%m-%d')}", align="R")

    return pdf.output()

# --------------- UI HELPERS ---------------
def fmt_time(seconds):
    m, s = divmod(int(seconds), 60)
    h, m = divmod(m, 60)
    return f"{h}小时{m}分{s}秒" if h else f"{m}分{s}秒"

def get_modules():
    conn = get_db()
    rows = conn.execute("SELECT id, title, chapters FROM modules ORDER BY id").fetchall()
    conn.close()
    return {r["id"]: {"title": r["title"], "chapters": json.loads(r["chapters"])} for r in rows}

def save_module(mid, title, chapters):
    conn = get_db()
    conn.execute("INSERT OR REPLACE INTO modules (id, title, chapters) VALUES (?,?,?)",
                 (mid, title, json.dumps(chapters, ensure_ascii=False)))
    conn.commit()
    conn.close()

def delete_module_db(mid):
    conn = get_db()
    conn.execute("DELETE FROM modules WHERE id=?", (mid,))
    conn.execute("DELETE FROM exam_questions WHERE module_id=?", (mid,))
    conn.execute("DELETE FROM progress WHERE module_id=?", (mid,))
    conn.execute("DELETE FROM exam_results WHERE module_id=?", (mid,))
    conn.execute("DELETE FROM training_time WHERE module_id=?", (mid,))
    conn.commit()
    conn.close()

def get_next_module_id():
    conn = get_db()
    row = conn.execute("SELECT MAX(id) as m FROM modules").fetchone()
    conn.close()
    return (row["m"] or 0) + 1

def get_exams():
    conn = get_db()
    rows = conn.execute("SELECT module_id, questions, exam_count FROM exam_questions").fetchall()
    conn.close()
    result = {}
    for r in rows:
        result[r["module_id"]] = {
            "questions": json.loads(r["questions"]),
            "exam_count": r["exam_count"] or 0,
        }
    return result

def save_exam_questions(mid, questions, exam_count=0):
    conn = get_db()
    conn.execute("INSERT OR REPLACE INTO exam_questions (module_id, questions, exam_count) VALUES (?,?,?)",
                 (mid, json.dumps(questions, ensure_ascii=False), exam_count))
    conn.commit()
    conn.close()

def delete_exam_questions_db(mid):
    conn = get_db()
    conn.execute("DELETE FROM exam_questions WHERE module_id=?", (mid,))
    conn.commit()
    conn.close()

# --------------- PAGE: LOGIN ---------------
def page_login():
    # 渐变背景容器
    st.markdown("""
<div class="login-wrapper">
  <div class="login-card">
    <div class="logo-area">
      <div class="company-icon">📦</div>
      <h1>驭长风供应链</h1>
      <p>新员工入职培训系统</p>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

    # 将表单放在卡片区域内（通过 Streamlit 布局）
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<div style='margin-top:-180px'></div>", unsafe_allow_html=True)
        tab_login, tab_reg = st.tabs(["🔑 登录", "📝 注册"])
        with tab_login:
            username = st.text_input("用户名", key="login_user", placeholder="请输入用户名")
            password = st.text_input("密码", type="password", key="login_pw", placeholder="请输入密码")
            if st.button("登 录", type="primary", use_container_width=True):
                user = authenticate(username, password)
                if user:
                    st.session_state.user = user
                    st.session_state.page = "dashboard"
                    st.rerun()
                else:
                    st.error("用户名或密码错误")
        with tab_reg:
            new_user = st.text_input("用户名", key="reg_user", placeholder="设置用户名")
            new_pw = st.text_input("密码", type="password", key="reg_pw", placeholder="设置密码")
            new_name = st.text_input("姓名", key="reg_name", placeholder="真实姓名")
            new_emp = st.text_input("工号", key="reg_emp", placeholder="员工工号")
            new_dept = st.selectbox("部门", DEPARTMENTS, key="reg_dept")
            if st.button("注 册", use_container_width=True):
                if not new_user or not new_pw or not new_name:
                    st.warning("请填写必填项")
                elif create_user(new_user, new_pw, new_name, new_emp, new_dept):
                    st.success("注册成功，请登录")
                else:
                    st.error("用户名已存在")

# --------------- PAGE: ADMIN ---------------
def page_admin():
    st.markdown("## 账号管理")
    users = get_all_users()

    with st.expander("➕ 新建账号"):
        c1, c2 = st.columns(2)
        nu = c1.text_input("用户名", key="adm_nu")
        np = c2.text_input("密码", type="password", key="adm_np")
        c3, c4, c5 = st.columns(3)
        nn = c3.text_input("姓名", key="adm_nn")
        ne = c4.text_input("工号", key="adm_ne")
        nd = c5.selectbox("部门", DEPARTMENTS, key="adm_nd")
        nr = st.selectbox("角色", ["user", "admin"], key="adm_nr")
        if st.button("创建"):
            if nu and np and nn:
                if create_user(nu, np, nn, ne, nd, nr):
                    st.success("创建成功")
                    st.rerun()
                else:
                    st.error("用户名已存在")

    st.markdown("### 现有账号")
    for u in users:
        cols = st.columns([2, 2, 2, 1, 1, 1])
        cols[0].write(u["username"])
        cols[1].write(u["display_name"])
        cols[2].write(u["department"])
        cols[3].write(u["role"])
        if u["role"] != "admin" or u["id"] != 1:
            if cols[4].button("重置密码", key=f"rst_{u['id']}"):
                update_user(u["id"], password_hash=hash_pw("123456"))
                st.success(f"{u['username']} 密码已重置为 123456")
            if cols[5].button("删除", key=f"del_{u['id']}"):
                delete_user(u["id"])
                st.rerun()

# --------------- PAGE: DASHBOARD ---------------
def page_dashboard():
    user = st.session_state.user
    modules = get_modules()
    exams = get_exams()

    # 品牌化仪表盘头部
    st.markdown(f"""
<div class="dashboard-header">
  <h2>📊 培训仪表盘</h2>
  <p>欢迎回来，<strong>{user['display_name']}</strong>！请依次完成以下培训模块。</p>
</div>
""", unsafe_allow_html=True)

    # 计算总进度
    exam_modules = sum(1 for e in exams.values() if e.get("questions"))
    total_items = sum(len(m["chapters"]) for m in modules.values()) + exam_modules
    completed = 0
    for mid, mod in modules.items():
        checks = get_read_checks(user["id"], mid, len(mod["chapters"]))
        completed += sum(checks)
        best = get_best_exam(user["id"], mid)
        if best is not None and best >= 80:
            completed += 1
    pct = int(completed / total_items * 100) if total_items else 0

    # 品牌化进度条
    render_progress_bar(pct, "总体培训进度")

    # 模块卡片网格
    num_cols = min(len(modules), 2)
    cols = st.columns(num_cols)
    for i, (mid, mod) in enumerate(modules.items()):
        col = cols[i % num_cols]
        with col:
            checks = get_read_checks(user["id"], mid, len(mod["chapters"]))
            read_done = sum(checks)
            total_ch = len(mod["chapters"])
            best = get_best_exam(user["id"], mid)
            t = get_training_time(user["id"], mid)

            status = "已完成" if (read_done == total_ch and best is not None and best >= 80) else ("进行中" if read_done > 0 else "未开始")
            score_str = f"{best}分" if best is not None else "--"

            # 渲染品牌化卡片
            render_module_card(
                title=mod["title"],
                status=status,
                chapters_info=f"章节 {read_done}/{total_ch}",
                exam_info=f"考试 {score_str}",
                time_info=f"时长 {fmt_time(t)}",
                module_id=mid
            )

            # 操作按钮
            c1, c2 = st.columns(2)
            if c1.button("📖 进入学习", key=f"learn_{mid}", use_container_width=True):
                st.session_state.page = "module"
                st.session_state.current_module = mid
                st.rerun()
            if c2.button("📝 参加考试", key=f"exam_{mid}", use_container_width=True):
                st.session_state.page = "exam"
                st.session_state.current_module = mid
                st.session_state.exam_started = False
                st.rerun()

    # 证书按钮
    all_pass = all(
        (get_best_exam(user["id"], mid) or 0) >= 80
        for mid, e in exams.items() if e.get("questions")
    )
    all_read = all(
        all(get_read_checks(user["id"], mid, len(modules[mid]["chapters"])))
        for mid in modules
    )
    if all_pass and all_read:
        st.markdown("---")
        st.markdown("""
<div style="text-align:center;padding:20px 0">
  <p style="color:#27ae60;font-size:1.1rem;font-weight:600;margin-bottom:12px">🎉 恭喜！您已完成全部培训课程</p>
</div>
""", unsafe_allow_html=True)
        if st.button("🎓 查看培训合格证书", type="primary", use_container_width=True):
            st.session_state.page = "certificate"
            st.rerun()

# --------------- SCORING HELPERS ---------------
def score_short_answer(user_ans: str, item: dict) -> tuple[float, str]:
    """基于 Google Gemini 的语义化智能打分，并返回详细判分理由"""
    
    user_ans = str(user_ans or "").strip()
    if not user_ans:
        return 0.0, "学员未作答，计0分。"
    
    ref_answer = item.get("answer", "")
    keywords = item.get("keywords", [])

    # 读取 Gemini 配置开关
    USE_AI_SCORING = st.secrets.get("gemini", {}).get("use_ai_scoring", False)
    
    if USE_AI_SCORING and ref_answer:
        try:
            api_key = st.secrets["gemini"]["api_key"]
            model_name = st.secrets["gemini"].get("model_name", "gemini-2.5-flash")
            
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(model_name)
            
            prompt = f"""你是一个严谨的供应链系统新员工培训考试阅卷助手。
            请对比【标准参考答案】与【学员回答】，判断学员是否答到了核心语义。
            
            【标准参考答案】：{ref_answer}
            【学员回答】：{user_ans}
            
            评分规则（满分1.0）：
            - 核心意思完全一致，即使措辞和描述方式不同，给 1.0
            - 表达了部分核心意思，给 0.5
            - 意思完全无关或错误，给 0.0
            
            请严格输出纯 JSON 字符串，不要带 markdown 标记(如 ```json )。格式必须如下：
            {{"score": 0.5, "reason": "这里写30字以内的评分理由，指出答对了什么或欠缺了什么"}}"""

            response = model.generate_content(
                prompt,
                generation_config=genai.types.GenerationConfig(
                    temperature=0.1,
                )
            )
            
            llm_reply = response.text.strip()
            clean_json = llm_reply.replace("```json", "").replace("```", "").strip()
            
            result = json.loads(clean_json)
            score = float(result.get("score", 0.0))
            reason = result.get("reason", "AI打分成功")
            
            return min(max(score, 0.0), 1.0), f"🤖 阅卷助手: {reason}"
                
        except Exception as e:
            print(f"Gemini 阅卷异常，降级为常规匹配: {e}")
            pass 
            
    # ================= 常规匹配算法 (降级方案) =================
    score = 0.0
    if keywords:
        hits = sum(1 for kw in keywords if kw.lower() in user_ans.lower())
        score = hits / len(keywords) if keywords else 0.0
    
    if ref_answer:
        ref_words = set(ref_answer.lower())
        ans_words = set(user_ans.lower())
        overlap = len(ref_words.intersection(ans_words))
        if len(ref_words) > 0:
            sim = overlap / len(ref_words)
            score = max(score, sim)
            
    if score >= 0.6: 
        return 1.0, "⚙️ 系统常规判定：命中多数核心关键词"
    elif score >= 0.3: 
        return 0.5, "⚙️ 系统常规判定：命中部分核心关键词"
    else: 
        return 0.0, "⚙️ 系统常规判定：未命中核心关键词或偏题"

# --------------- PAGE: MODULE CONTENT ---------------
def page_module():
    import urllib.parse
    user = st.session_state.user
    mid = st.session_state.current_module
    modules = get_modules()
    mod = modules[mid]

    # 品牌化模块头部
    st.markdown(f"""
<div class="module-content-header">
  <h2>📖 {mod['title']}</h2>
</div>
""", unsafe_allow_html=True)

    if st.button("← 返回仪表盘"):
        st.session_state.page = "dashboard"
        st.rerun()

    checks = get_read_checks(user["id"], mid, len(mod["chapters"]))

    for idx, ch in enumerate(mod["chapters"]):
        title_text = ch.get("title", "未命名章节")
        status_icon = "✅" if checks[idx] else "📖"

        with st.expander(f"{status_icon} 课件：{title_text}", expanded=not checks[idx]):
            file_path = ch.get("file_path", "")
            file_type = ch.get("file_type", "pdf")
            old_html = ch.get("html", "")

            if file_path:
                if not os.path.exists(file_path):
                    st.error("⚠️ 文件路径失效，请管理员重新上传此模块。")
                else:
                    if os.path.getsize(file_path) == 0:
                        st.error("⚠️ 该文件转换失败（大小为0字节），请检查本地 Office 是否卡死。")
                    else:
                        # 云端适配：直接提供下载（无法使用本地服务器预览）
                        with open(file_path, "rb") as f:
                            st.download_button(label=f"📥 下载课件 ({file_type.upper()})", data=f, file_name=f"{title_text}.{file_type}", key=f"dl_{mid}_{idx}")
            elif old_html:
                st.markdown(old_html, unsafe_allow_html=True)
            else:
                st.info("该章节暂无内容。")

            st.markdown("<br>", unsafe_allow_html=True)
            if not checks[idx]:
                if st.button(f"✅ 确认完成学习", key=f"read_{mid}_{idx}", type="primary"):
                    mark_chapter_read(user["id"], mid, idx)
                    st.rerun()
            else:
                st.success("✅ 学习已完成")

# --------------- PAGE: EXAM ---------------
def page_exam():
    user = st.session_state.user
    mid = st.session_state.current_module
    modules = get_modules()
    exams = get_exams()
    mod = modules[mid]
    exam_data = exams.get(mid, {})
    all_questions = exam_data.get("questions", []) if isinstance(exam_data, dict) else exam_data
    exam_count = exam_data.get("exam_count", 0) if isinstance(exam_data, dict) else 0

    # 品牌化考试头部
    st.markdown(f"""
<div class="exam-header-bar">
  <div>
    <h2>📝 {mod['title']} — 在线考试</h2>
  </div>
  <div>
    <a href="#" onclick="window.history.back(); return false;" style="color:#666;text-decoration:none;font-size:.85rem">← 返回仪表盘</a>
  </div>
</div>
""", unsafe_allow_html=True)

    if st.button("← 返回仪表盘", key="exam_back"):
        st.session_state.page = "dashboard"
        st.rerun()

    if not all_questions:
        st.warning("本模块暂无考试题目")
        return

    checks = get_read_checks(user["id"], mid, len(mod["chapters"]))
    if not all(checks):
        st.warning("⚠️ 请先完成所有章节的阅读确认后再参加考试")
        return

    best = get_best_exam(user["id"], mid)
    if best is not None:
        st.info(f"📊 历史最高分: {best}分 {'(已通过 ✅)' if best >= 80 else ''}")

    EXAM_DURATION = 30 * 60
    actual_count = exam_count if (exam_count > 0 and exam_count < len(all_questions)) else len(all_questions)

    if not st.session_state.get("exam_started"):
        # 考试说明卡片
        st.markdown(f"""
<div class="question-card" style="border-left-color:var(--primary)">
  <div class="q-number">考试说明</div>
  <div class="q-text">
    题库共 <strong>{len(all_questions)}</strong> 题，本次考试随机抽取 <strong>{actual_count}</strong> 题<br>
    每题 <strong>{round(100/actual_count)}</strong> 分，<strong>80分</strong>及格，限时 <strong>30分钟</strong>
  </div>
</div>
""", unsafe_allow_html=True)
        if st.button("🚀 开始考试", type="primary", use_container_width=True):
            if actual_count < len(all_questions):
                selected = random.sample(all_questions, actual_count)
            else:
                selected = list(all_questions)
                random.shuffle(selected)
            st.session_state.exam_questions = selected
            st.session_state.exam_started = True
            st.session_state.exam_start_time = time.time()
            st.session_state.exam_answers = {}
            st.rerun()
        return

    questions = st.session_state.get("exam_questions", all_questions)
    elapsed = time.time() - st.session_state.exam_start_time
    remaining = max(0, EXAM_DURATION - elapsed)
    rm, rs = divmod(int(remaining), 60)

    # 渲染品牌化倒计时
    is_warning = remaining <= 300  # 最后5分钟警告
    render_exam_timer(rm, rs, is_warning)

    if remaining <= 0:
        st.error("⏰ 考试时间已到，自动提交")
        _submit_exam(user, mid, questions)
        return

    per_q = round(100 / len(questions))
    answers = st.session_state.get("exam_answers", {})

    for qi, item in enumerate(questions):
        qtype = item.get("type", "single")

        # 渲染品牌化题目卡片
        render_question_card(qi + 1, item["q"], qtype, per_q)

        if qtype == "single":
            sel = st.radio(
                f"q_{mid}_{qi}",
                options=item["opts"],
                index=answers.get(qi),
                key=f"eq_{mid}_{qi}",
                label_visibility="collapsed",
            )
            if sel is not None:
                answers[qi] = item["opts"].index(sel)
        elif qtype == "multi":
            prev = answers.get(qi, [])
            selected = []
            for oi, opt in enumerate(item["opts"]):
                checked = st.checkbox(f"{chr(65+oi)}. {opt}", value=(oi in prev), key=f"eq_{mid}_{qi}_{oi}")
                if checked:
                    selected.append(oi)
            answers[qi] = sorted(selected)
        elif qtype == "short":
            prev = answers.get(qi, "")
            txt = st.text_area("请输入答案", value=prev, key=f"eq_{mid}_{qi}", height=100,
                              placeholder="请输入您的答案...")
            answers[qi] = txt

    st.session_state.exam_answers = answers

    # 提交按钮
    st.markdown("<div style='text-align:center;margin-top:24px'>", unsafe_allow_html=True)
    if st.button("📋 提交考试", type="primary", use_container_width=True):
        unanswered = 0
        for qi, item in enumerate(questions):
            qtype = item.get("type", "single")
            a = answers.get(qi)
            if a is None or (qtype == "multi" and len(a) == 0) or (qtype == "short" and not str(a).strip()):
                unanswered += 1
        if unanswered > 0:
            st.warning(f"⚠️ 还有 {unanswered} 题未作答，请检查后再提交")
        else:
            _submit_exam(user, mid, questions)
    st.markdown("</div>", unsafe_allow_html=True)

def _submit_exam(user, mid, questions):
    answers = st.session_state.get("exam_answers", {})
    per_q = round(100 / len(questions))
    score = 0
    details = []
    
    for qi, item in enumerate(questions):
        qtype = item.get("type", "single")
        a = answers.get(qi)
        q_score = 0
        ai_reason = "" # 初始化判分理由
        
        if qtype == "single":
            if a == item.get("ans"):
                q_score = per_q
        elif qtype == "multi":
            correct = sorted(item.get("ans", []))
            if sorted(a or []) == correct:
                q_score = per_q
        elif qtype == "short":
            # 接收分数比例和判卷理由
            ratio, ai_reason = score_short_answer(a, item)
            q_score = round(per_q * ratio)
            
        score += q_score
        details.append({
            "q": item["q"], 
            "type": qtype, 
            "earned": q_score, 
            "max": per_q,
            "user_ans": a, 
            "correct": item.get("ans") or item.get("answer", ""),
            "ai_reason": ai_reason # 将理由存入详情
        })
        
    score = min(score, 100)
    elapsed = time.time() - st.session_state.exam_start_time
    add_training_time(user["id"], mid, int(elapsed))
    save_exam_result(user["id"], mid, score, {"answers": answers, "details": details})
    
    st.session_state.exam_started = False
    st.session_state.exam_questions = None
    st.session_state.last_score = score
    st.session_state.last_details = details
    st.session_state.page = "result"
    st.rerun()

# --------------- PAGE: RESULT ---------------
def page_result():
    score = st.session_state.get("last_score", 0)
    mid = st.session_state.get("current_module", 1)
    modules = get_modules()
    mod = modules[mid]
    passed = score >= 80
    details = st.session_state.get("last_details", [])

    # 品牌化成绩展示
    render_result_circle(score, passed)

    if passed:
        st.markdown(f"""
<div style="text-align:center;margin-bottom:24px">
  <h3 style="color:#27ae60">🎉 恭喜通过考试！</h3>
  <p style="color:#666">{mod['title']} — 得分 {score}/100</p>
</div>
""", unsafe_allow_html=True)
    else:
        st.markdown(f"""
<div style="text-align:center;margin-bottom:24px">
  <h3 style="color:#e74c3c">😔 未通过考试</h3>
  <p style="color:#666">{mod['title']} — 得分 {score}/100（及格线: 80分）</p>
</div>
""", unsafe_allow_html=True)

    if details:
        with st.expander("📝 查看逐题解析与判分详情", expanded=True):
            for i, d in enumerate(details):
                icon = "✅" if d["earned"] == d["max"] else ("⚠️" if d["earned"] > 0 else "❌")
                type_label = {"single": "单选", "multi": "多选", "short": "简答"}.get(d["type"], "")

                st.markdown(f"**{icon} {i+1}. [{type_label}] {d['q']}**")

                if d['type'] == "short":
                    st.markdown(f"> **你的回答**: {d.get('user_ans', '未作答')}")
                    st.markdown(f"> **标准答案**: {d['correct']}")
                    st.markdown(f"**得分**: {d['earned']}/{d['max']} 分")
                    if d.get("ai_reason"):
                        st.info(d["ai_reason"])
                else:
                    st.markdown(f"**得分**: {d['earned']}/{d['max']} 分")

                st.divider()

    # 操作按钮
    c1, c2 = st.columns(2)
    if c1.button("📊 返回仪表盘", use_container_width=True):
        st.session_state.page = "dashboard"
        st.rerun()
    if not passed:
        if c2.button("🔄 重新考试", type="primary", use_container_width=True):
            st.session_state.page = "exam"
            st.session_state.exam_started = False
            st.rerun()

# --------------- PAGE: CERTIFICATE ---------------
def page_certificate():
    user = st.session_state.user
    modules = get_modules()

    module_scores = []
    for mid, mod in modules.items():
        best = get_best_exam(user["id"], mid) or 0
        t = get_training_time(user["id"], mid)
        module_scores.append((mod["title"], best, t // 60))

    date_str = datetime.now().strftime('%Y-%m-%d')

    # 渲染品牌化证书预览
    render_certificate_preview(
        name=user["display_name"],
        emp_id=user.get("emp_id", ""),
        dept=user.get("department", ""),
        module_scores=module_scores,
        date_str=date_str
    )

    # 下载按钮
    st.markdown("<div style='text-align:center;margin-top:24px'>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        if st.button("📥 下载PDF证书", type="primary", use_container_width=True):
            pdf_bytes = generate_certificate_pdf(
                user["display_name"],
                user.get("emp_id", ""),
                user.get("department", ""),
                module_scores,
                {},
            )
            st.download_button(
                "点击下载",
                data=pdf_bytes,
                file_name=f"certificate_{user.get('emp_id','')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
    with c2:
        if st.button("📊 返回仪表盘", use_container_width=True):
            st.session_state.page = "dashboard"
            st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

# --------------- PAGE: UPLOAD COURSE ---------------
def page_upload_course():
    if st.session_state.user.get("role") != "admin":
        st.warning("仅管理员可使用此功能")
        return
    st.markdown("## 上传培训课件 (支持原格式 1:1 在线预览)")
    st.info("系统会自动将上传的 Word/PPT 在后台转换为网页可读格式，保证排版原汁原味。")

    modules = get_modules()
    mode = st.radio("操作", ["新建模块", "替换现有模块课件"], horizontal=True, key="course_mode")

    if mode == "替换现有模块课件" and modules:
        mod_opts = {mid: mod["title"] for mid, mod in modules.items()}
        replace_mid = st.selectbox("选择要替换的模块", options=list(mod_opts.keys()),
                                   format_func=lambda x: mod_opts[x], key="replace_mid")
        mod_title = mod_opts[replace_mid]
    else:
        replace_mid = None
        mod_title = st.text_input("模块名称", placeholder="例: 模块三：仓储管理")

    uploaded_files = st.file_uploader("选择课件文件 (支持多选 PDF/DOCX/PPTX)", type=["pdf", "docx", "pptx", "doc", "ppt"], accept_multiple_files=True)

    btn_label = "上传并替换课件" if replace_mid else "上传并创建模块"
    if uploaded_files and mod_title and st.button(btn_label, type="primary"):
        chapters = []
        try:
            for f in uploaded_files:
                safe_filename = f"{int(time.time())}_{f.name}"
                file_path = os.path.join(UPLOAD_DIR, safe_filename)

                # 保存原始文件
                with open(file_path, "wb") as out:
                    out.write(f.read())

                ext = f.name.split('.')[-1].lower()
                # 云端适配：直接存储原始文件，提供下载
                chapters.append({"title": f.name, "file_path": file_path, "file_type": ext})

            if replace_mid:
                save_module(replace_mid, mod_title, chapters)
                conn = get_db()
                conn.execute("DELETE FROM progress WHERE module_id=?", (replace_mid,))
                conn.commit()
                conn.close()
                st.success(f"已替换模块 [{mod_title}] 的课件。")
            else:
                new_id = get_next_module_id()
                save_module(new_id, mod_title, chapters)
                save_exam_questions(new_id, [])
                st.success(f"已创建模块 [{mod_title}]。")

        except Exception as e:
            st.error(f"处理失败: {e}")

# --------------- PAGE: UPLOAD EXAM ---------------
def page_upload_exam():
    if st.session_state.user.get("role") != "admin":
        st.warning("仅管理员可使用此功能")
        return
    st.markdown("## 题库管理")

    modules = get_modules()
    if not modules:
        st.warning("请先创建培训模块")
        return

    exams = get_exams()
    mod_opts = {mid: mod["title"] for mid, mod in modules.items()}
    selected_mid = st.selectbox("选择模块", options=list(mod_opts.keys()), format_func=lambda x: mod_opts[x])

    exam_data = exams.get(selected_mid, {})
    current_qs = exam_data.get("questions", []) if isinstance(exam_data, dict) else []
    current_count = exam_data.get("exam_count", 0) if isinstance(exam_data, dict) else 0

    st.markdown(f"当前题库: **{len(current_qs)}** 题 | 每次考试抽取: **{current_count if current_count > 0 else '全部'}** 题")

    st.markdown("---")
    st.markdown("### 考试抽题设置")
    new_count = st.number_input("每次考试随机抽取题数 (0=全部出题)", min_value=0,
                                max_value=max(len(current_qs), 100), value=current_count, key="exam_count_input")
    if new_count != current_count:
        if st.button("保存抽题设置"):
            save_exam_questions(selected_mid, current_qs, new_count)
            st.success(f"已设置每次考试抽取 {new_count if new_count > 0 else '全部'} 题")
            st.rerun()

    # --- Batch entry ---
    st.markdown("---")
    st.markdown("### ⚡ 批量快速录入")
    st.markdown("支持直接粘贴文本，每行一题。格式：`题目 | 选项A | 选项B | 选项C | 选项D | 答案(A/B/C/D)`")
    batch_text = st.text_area("在此粘贴题库文本", height=150, key="batch_upload_text")

    if st.button("批量解析入库", type="primary"):
        if batch_text.strip():
            new_items = []
            ans_map = {"A": 0, "B": 1, "C": 2, "D": 3}
            for line in batch_text.split("\n"):
                line = line.strip()
                if not line: continue
                parts = [p.strip() for p in line.split("|")]
                if len(parts) >= 6:
                    q_text, a, b, c, d, ans_str = parts[:6]
                    # 去除干扰字符，提取纯大写字母答案
                    ans_clean = ans_str.upper().replace('答案', '').replace(':', '').replace('：', '').replace(',', '').replace('，', '').strip()
                    
                    # 识别答案中包含了哪些有效选项字母
                    letters = [char for char in ans_clean if char in ans_map]
                    
                    if len(letters) > 1:
                        # 识别为多选题
                        ans_indices = sorted(list(set([ans_map[char] for char in letters])))
                        new_items.append({
                            "q": q_text, 
                            "opts": [a, b, c, d], 
                            "ans": ans_indices, 
                            "type": "multi"
                        })
                    elif len(letters) == 1:
                        # 识别为单选题
                        new_items.append({
                            "q": q_text, 
                            "opts": [a, b, c, d], 
                            "ans": ans_map[letters[0]], 
                            "type": "single"
                        })
            
            if new_items:
                updated_qs = current_qs + new_items
                save_exam_questions(selected_mid, updated_qs, exam_count=10) # 自动设置为抽取10道题
                
                # 统计一下导入了多少单选，多少多选
                singles = sum(1 for q in new_items if q["type"] == "single")
                multis = sum(1 for q in new_items if q["type"] == "multi")
                
                st.success(f"成功批量录入 {len(new_items)} 道题 (单选 {singles} 题, 多选 {multis} 题)！当前考试已自动设置为随机抽取 10 题。")
                st.rerun()
            else:
                st.error("未能解析出有效题目，请检查格式分隔符是否为 '|' 且答案是否包含A/B/C/D")

    st.markdown("---")
    st.markdown("### 从文件导入")
    st.markdown("""
支持: **Excel (.xlsx)** 列: 题目、选项A-D、答案、题型(可选) | **Word (.docx)** 题目+选项+答案行
- 单选: 答案=A  多选: 答案=AB,题型=多选  简答: 选项留空,答案=关键词(|分隔),题型=简答
""")
    uploaded = st.file_uploader("选择题库文件", type=["xlsx", "docx"], key="exam_upload")

    if uploaded and st.button("解析并导入题库"):
        file_bytes = uploaded.read()
        try:
            if uploaded.name.endswith(".xlsx"):
                questions = parse_exam_xlsx(file_bytes)
            else:
                questions = parse_exam_docx(file_bytes)

            if not questions:
                st.error("未能从文件中提取到有效题目")
                return

            save_exam_questions(selected_mid, current_qs + questions, new_count)

            singles = sum(1 for q in questions if q.get("type", "single") == "single")
            multis = sum(1 for q in questions if q.get("type") == "multi")
            shorts = sum(1 for q in questions if q.get("type") == "short")
            st.success(f"已导入 {len(questions)} 题 (单选 {singles} / 多选 {multis} / 简答 {shorts})")
            st.rerun()
        except Exception as e:
            st.error(f"解析失败: {e}")

    # --- Manual entry ---
    st.markdown("---")
    st.markdown("### 手工录入单道题目")
    
    # 【关键修改】：将题型选择器移动到 form 的外面，使其可以实时触发界面更新
    qtype = st.selectbox("题型", ["单选", "多选", "简答"], key="mq_type")
    
    with st.form("manual_q", clear_on_submit=True):
        q_text = st.text_input("题目")
        if qtype in ("单选", "多选"):
            c1, c2 = st.columns(2)
            opt_a = c1.text_input("选项A")
            opt_b = c2.text_input("选项B")
            c3, c4 = st.columns(2)
            opt_c = c3.text_input("选项C")
            opt_d = c4.text_input("选项D")
            if qtype == "单选":
                ans_choice = st.selectbox("正确答案", ["A", "B", "C", "D"])
            else:
                ans_multi = st.multiselect("正确答案(可多选)", ["A", "B", "C", "D"])
        else:
            answer_text = st.text_input("参考答案 (关键词用|分隔)")

        submitted = st.form_submit_button("添加到题库")
        if submitted and q_text:
            ans_map = {"A": 0, "B": 1, "C": 2, "D": 3}
            if qtype == "单选":
                opts = [opt_a, opt_b, opt_c, opt_d]
                if all(opts):
                    new_item = {"q": q_text, "opts": opts, "ans": ans_map[ans_choice], "type": "single"}
                else:
                    st.warning("请填写所有选项")
                    new_item = None
            elif qtype == "多选":
                opts = [opt_a, opt_b, opt_c, opt_d]
                valid_opts = [o for o in opts if o]
                if len(valid_opts) >= 2 and ans_multi:
                    new_item = {"q": q_text, "opts": valid_opts,
                                "ans": sorted(ans_map[c] for c in ans_multi), "type": "multi"}
                else:
                    st.warning("至少2个选项和1个答案")
                    new_item = None
            else:
                keywords = [k.strip() for k in answer_text.replace("，", "|").replace(",", "|").split("|") if k.strip()]
                new_item = {"q": q_text, "type": "short", "keywords": keywords, "answer": answer_text}

            if new_item:
                updated_qs = current_qs + [new_item]
                save_exam_questions(selected_mid, updated_qs, new_count)
                st.success(f"已添加，当前共 {len(updated_qs)} 题")
                st.rerun()

    if current_qs:
        st.markdown("---")
        st.markdown(f"### 当前题库 ({len(current_qs)} 题)")
        for i, item in enumerate(current_qs):
            qtype = item.get("type", "single")
            label = {"single": "单选", "multi": "多选", "short": "简答"}.get(qtype, "")
            cols = st.columns([10, 1])
            cols[0].markdown(f"**{i+1}. [{label}] {item['q']}**")
            if cols[1].button("删除", key=f"delq_{selected_mid}_{i}"):
                updated = current_qs[:i] + current_qs[i+1:]
                save_exam_questions(selected_mid, updated, new_count)
                st.rerun()

# --------------- PAGE: MANAGE COURSES ---------------
def page_manage_courses():
    if st.session_state.user.get("role") != "admin":
        st.warning("仅管理员可使用此功能")
        return
    st.markdown("## 课件管理")

    modules = get_modules()
    exams = get_exams()

    if not modules:
        st.info("暂无培训模块")
        return

    for mid, mod in list(modules.items()):
        exam_data = exams.get(mid, {})
        qs = exam_data.get("questions", []) if isinstance(exam_data, dict) else []
        with st.expander(f"{mod['title']}  ({len(mod['chapters'])} 章 | {len(qs)} 题)"):
            new_name = st.text_input("模块名称", value=mod["title"], key=f"rename_{mid}")
            if new_name != mod["title"]:
                if st.button("保存名称", key=f"savename_{mid}"):
                    save_module(mid, new_name, mod["chapters"])
                    st.success("已更新")
                    st.rerun()

            st.markdown("**章节列表:**")
            for idx, ch in enumerate(mod["chapters"]):
                st.markdown(f"{idx+1}. {ch['title']}")

            st.markdown("---")
            st.markdown("**危险操作**")
            if st.button(f"删除整个模块", key=f"delmod_{mid}", type="secondary"):
                st.session_state[f"confirm_del_{mid}"] = True

            if st.session_state.get(f"confirm_del_{mid}"):
                st.warning(f"确认删除 [{mod['title']}]？此操作将清除该模块的所有学员进度和考试记录。")
                c1, c2 = st.columns(2)
                if c1.button("确认删除", key=f"yes_del_{mid}", type="primary"):
                    delete_module_db(mid)
                    st.session_state.pop(f"confirm_del_{mid}", None)
                    st.success("已删除")
                    st.rerun()
                if c2.button("取消", key=f"no_del_{mid}"):
                    st.session_state.pop(f"confirm_del_{mid}", None)
                    st.rerun()

            if qs:
                if st.button(f"清空题库 ({len(qs)} 题)", key=f"delexam_{mid}"):
                    save_exam_questions(mid, [])
                    st.success("题库已清空")
                    st.rerun()

# --------------- PAGE: STUDENT ANALYTICS ---------------
def page_analytics():
    if st.session_state.user.get("role") != "admin":
        st.warning("仅管理员可使用此功能")
        return
    st.markdown("## 考生分析")

    conn = get_db()
    users = [dict(r) for r in conn.execute(
        "SELECT id, username, display_name, department, emp_id FROM users WHERE role='user' ORDER BY id"
    ).fetchall()]

    if not users:
        st.info("暂无普通用户")
        conn.close()
        return

    modules = get_modules()
    exams = get_exams()

    st.markdown("### 全员概览")
    header_cols = ["姓名", "部门", "工号"] + [mod["title"] for mod in modules.values()] + ["总进度"]
    rows = []
    for u in users:
        row = [u["display_name"], u["department"], u["emp_id"]]
        total_ch = 0
        done_ch = 0
        total_exams = 0
        passed_exams = 0
        for mid, mod in modules.items():
            ch_count = len(mod["chapters"])
            checks = get_read_checks(u["id"], mid, ch_count)
            read_n = sum(checks)
            best = get_best_exam(u["id"], mid)
            total_ch += ch_count
            done_ch += read_n
            exam_data = exams.get(mid, {})
            has_exam = bool(exam_data.get("questions") if isinstance(exam_data, dict) else exam_data)
            if has_exam:
                total_exams += 1
                if best is not None and best >= 80:
                    passed_exams += 1
            score_str = f"{best}分" if best is not None else "--"
            status = f"{read_n}/{ch_count}章 | {score_str}"
            row.append(status)
        total_items = total_ch + total_exams
        done_items = done_ch + passed_exams
        pct = round(done_items / total_items * 100) if total_items else 0
        row.append(f"{pct}%")
        rows.append(row)

    import pandas as pd
    df = pd.DataFrame(rows, columns=header_cols)
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### 个人详细分析")
    user_opts = {u["id"]: f"{u['display_name']} ({u['emp_id']})" for u in users}
    sel_uid = st.selectbox("选择考生", options=list(user_opts.keys()), format_func=lambda x: user_opts[x])

    sel_user = next(u for u in users if u["id"] == sel_uid)
    st.markdown(f"**{sel_user['display_name']}** | 部门: {sel_user['department']} | 工号: {sel_user['emp_id']}")

    for mid, mod in modules.items():
        st.markdown(f"#### {mod['title']}")
        ch_count = len(mod["chapters"])
        checks = get_read_checks(sel_uid, mid, ch_count)
        read_n = sum(checks)
        st.progress(read_n / ch_count if ch_count else 0, text=f"阅读进度: {read_n}/{ch_count}")

        exam_rows = conn.execute(
            "SELECT score, answers, taken_at FROM exam_results WHERE user_id=? AND module_id=? ORDER BY taken_at DESC",
            (sel_uid, mid)
        ).fetchall()

        if exam_rows:
            best = max(r["score"] for r in exam_rows)
            st.markdown(f"考试次数: **{len(exam_rows)}** | 最高分: **{best}** | 最近一次: **{exam_rows[0]['score']}**")

            scores = [r["score"] for r in reversed(exam_rows)]
            if len(scores) > 1:
                st.line_chart({"得分": scores}, height=150)

            latest_ans = exam_rows[0]["answers"]
            try:
                ans_data = json.loads(latest_ans)
                details = ans_data.get("details", []) if isinstance(ans_data, dict) else []
            except (json.JSONDecodeError, TypeError):
                details = []

            # 需求3：考生个性化能力分析报告
            if details:
                wrong = [d for d in details if d["earned"] < d["max"]]
                if wrong:
                    st.markdown("#### 🔍 个性化能力诊断")
                    
                    error_types = {"single": 0, "multi": 0, "short": 0}
                    for d in wrong:
                        error_types[d.get("type", "single")] += 1
                        
                    col1, col2 = st.columns(2)
                    with col1:
                        st.warning(f"**核心薄弱点**：共 {len(wrong)} 处知识盲区")
                        st.markdown(f"- 概念记忆偏差（单/多选错误）: **{error_types['single'] + error_types['multi']}** 题")
                        st.markdown(f"- 流程与实操理解不足（简答题失分）: **{error_types['short']}** 题")
                        
                    with col2:
                        st.info("**💡 系统建议复习方向**")
                        if error_types['short'] > error_types['single'] + error_types['multi']:
                            st.markdown("- 建议重新阅读**系统流程与规范**章节，加强对底层逻辑的理解。")
                        else:
                            st.markdown("- 建议重点巩固**硬性指标与基础概念**，可利用碎片时间回顾模块大纲。")

                    with st.expander(f"薄弱环节明细与判分追踪 ({len(wrong)} 题)"):
                        for d in wrong:
                            type_label = {"single": "单选", "multi": "多选", "short": "简答"}.get(d.get("type", "single"), "")
                            st.markdown(f"**[{type_label}]** {d['q']} (得分 {d['earned']}/{d['max']})")
                            
                            if d.get("type") in ["single", "multi"]:
                                st.markdown(f"- *系统标准答案*：{d['correct']}")
                            else:
                                st.markdown(f"- *考生回答*：{d.get('user_ans', '未作答')}")
                                st.markdown(f"- *参考答案/关键词*：{d['correct']}")
                                if d.get("ai_reason"):
                                    # 管理员视角高亮展示 AI 判卷理由
                                    st.caption(f"**{d['ai_reason']}**")
                            st.divider()
                else:
                    st.success("🎉 最近一次考试表现完美，无知识盲区。")
        else:
            st.caption("尚未参加考试")

        t = get_training_time(sel_uid, mid)
        st.caption(f"学习时长: {fmt_time(t)}")

    conn.close()

# --------------- SIDEBAR & MAIN ---------------
def main():
    st.set_page_config(page_title="驭长风供应链培训系统", page_icon="📦", layout="wide")
    inject_brand_css()
    init_db()

    if "page" not in st.session_state:
        st.session_state.page = "login"

    if "user" not in st.session_state:
        page_login()
        return

    user = st.session_state.user

    with st.sidebar:
        st.markdown(f"### 📦 驭长风培训")
        st.markdown(f"**{user['display_name']}**")
        st.caption(f"👤 {user.get('department','')} | {user.get('emp_id','')}")
        st.markdown("---")

        if st.button("📊 培训仪表盘", use_container_width=True):
            st.session_state.page = "dashboard"
            st.rerun()

        if user["role"] == "admin":
            st.markdown("**管理员功能**")
            if st.button("👥 账号管理", use_container_width=True):
                st.session_state.page = "admin"
                st.rerun()

            if st.button("📁 上传课件", use_container_width=True):
                st.session_state.page = "upload_course"
                st.rerun()

            if st.button("📝 题库管理", use_container_width=True):
                st.session_state.page = "upload_exam"
                st.rerun()

            if st.button("⚙️ 课件管理", use_container_width=True):
                st.session_state.page = "manage_courses"
                st.rerun()

            if st.button("📈 考生分析", use_container_width=True):
                st.session_state.page = "analytics"
                st.rerun()

        st.markdown("---")
        if st.button("🚪 退出登录", use_container_width=True):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()

    page = st.session_state.page
    if page == "dashboard":
        page_dashboard()
    elif page == "admin" and user["role"] == "admin":
        page_admin()
    elif page == "module":
        page_module()
    elif page == "exam":
        page_exam()
    elif page == "result":
        page_result()
    elif page == "certificate":
        page_certificate()
    elif page == "upload_course":
        page_upload_course()
    elif page == "upload_exam":
        page_upload_exam()
    elif page == "manage_courses":
        page_manage_courses()
    elif page == "analytics":
        page_analytics()
    else:
        page_dashboard()

if __name__ == "__main__":
    main()