"""
Training-SYS Database Layer (Supabase)
所有数据库操作统一在此文件中，app.py 只需 import 使用
表名使用 training_ 前缀避免与其他项目冲突
"""

import json
import hashlib
from supabase import create_client, Client

# 表名常量
TABLE_USERS = "training_users"
TABLE_PROGRESS = "training_progress"
TABLE_EXAM_RESULTS = "training_exam_results"
TABLE_TRAINING_TIME = "training_training_time"
TABLE_MODULES = "training_modules"
TABLE_EXAM_QUESTIONS = "training_exam_questions"

# Supabase 连接（从 Streamlit Secrets 读取）
def get_supabase() -> Client:
    import streamlit as st
    url = st.secrets.get("supabase_url", "")
    key = st.secrets.get("supabase_key", "")
    if not url or not key:
        st.error("❌ 未配置 Supabase 数据库。请在 Settings → Secrets 中添加 supabase_url 和 supabase_key")
        st.stop()
    return create_client(url, key)


# ============ INIT ============

def init_db():
    """检查 Supabase 数据库连接"""
    try:
        sb = get_supabase()
        sb.table(TABLE_USERS).select("id").limit(1).execute()
    except Exception as e:
        import streamlit as st
        st.error(f"❌ 数据库连接失败: {e}")
        st.info("请确保已在 Supabase Dashboard 中执行 supabase_migration.sql 创建数据表")
        st.stop()


# ============ USERS ============

def hash_pw(pw: str) -> str:
    return hashlib.sha256(pw.encode()).hexdigest()


def authenticate(username: str, password: str):
    sb = get_supabase()
    result = sb.table(TABLE_USERS).select("*").eq("username", username).eq("password_hash", hash_pw(password)).execute()
    return result.data[0] if result.data else None


def create_user(username, password, display_name, emp_id, department, role="user"):
    sb = get_supabase()
    try:
        sb.table(TABLE_USERS).insert({
            "username": username,
            "password_hash": hash_pw(password),
            "display_name": display_name,
            "emp_id": emp_id,
            "department": department,
            "role": role,
        }).execute()
        return True
    except Exception:
        return False


def get_all_users():
    sb = get_supabase()
    result = sb.table(TABLE_USERS).select("*").order("id").execute()
    return result.data


def delete_user(uid):
    sb = get_supabase()
    sb.table(TABLE_USERS).delete().eq("id", uid).execute()
    sb.table(TABLE_PROGRESS).delete().eq("user_id", uid).execute()
    sb.table(TABLE_EXAM_RESULTS).delete().eq("user_id", uid).execute()
    sb.table(TABLE_TRAINING_TIME).delete().eq("user_id", uid).execute()


def update_user(uid, **kwargs):
    ALLOWED_COLUMNS = {"password_hash", "display_name", "role", "department", "emp_id"}
    sb = get_supabase()
    for k, v in kwargs.items():
        if k not in ALLOWED_COLUMNS:
            raise ValueError(f"不允许更新列: {k}")
    update_data = {k: v for k, v in kwargs.items() if k in ALLOWED_COLUMNS}
    if update_data:
        sb.table(TABLE_USERS).update(update_data).eq("id", uid).execute()


def seed_admin(admin_pw: str):
    """种子管理员账号"""
    sb = get_supabase()
    existing = sb.table(TABLE_USERS).select("id").eq("username", "admin").execute()
    if not existing.data:
        sb.table(TABLE_USERS).insert({
            "username": "admin",
            "password_hash": hash_pw(admin_pw),
            "display_name": "管理员",
            "role": "admin",
        }).execute()


# ============ PROGRESS ============

def get_read_checks(user_id, module_id, chapter_count):
    sb = get_supabase()
    result = sb.table(TABLE_PROGRESS).select("chapter_idx, read_done").eq("user_id", user_id).eq("module_id", module_id).execute()
    checks = [False] * chapter_count
    for r in result.data:
        if r["chapter_idx"] < chapter_count:
            checks[r["chapter_idx"]] = bool(r["read_done"])
    return checks


def mark_chapter_read(user_id, module_id, chapter_idx):
    sb = get_supabase()
    existing = sb.table(TABLE_PROGRESS).select("user_id").eq("user_id", user_id).eq("module_id", module_id).eq("chapter_idx", chapter_idx).execute()
    if existing.data:
        sb.table(TABLE_PROGRESS).update({"read_done": 1}).eq("user_id", user_id).eq("module_id", module_id).eq("chapter_idx", chapter_idx).execute()
    else:
        sb.table(TABLE_PROGRESS).insert({
            "user_id": user_id,
            "module_id": module_id,
            "chapter_idx": chapter_idx,
            "read_done": 1,
        }).execute()


# ============ EXAM RESULTS ============

def save_exam_result(user_id, module_id, score, answers):
    sb = get_supabase()
    sb.table(TABLE_EXAM_RESULTS).insert({
        "user_id": user_id,
        "module_id": module_id,
        "score": score,
        "answers": json.dumps(answers),
    }).execute()


def get_best_exam(user_id, module_id):
    sb = get_supabase()
    result = sb.table(TABLE_EXAM_RESULTS).select("score").eq("user_id", user_id).eq("module_id", module_id).order("score", desc=True).limit(1).execute()
    return result.data[0]["score"] if result.data else None


# ============ TRAINING TIME ============

def get_training_time(user_id, module_id):
    sb = get_supabase()
    result = sb.table(TABLE_TRAINING_TIME).select("seconds").eq("user_id", user_id).eq("module_id", module_id).execute()
    return result.data[0]["seconds"] if result.data else 0


def add_training_time(user_id, module_id, seconds):
    sb = get_supabase()
    existing = sb.table(TABLE_TRAINING_TIME).select("user_id").eq("user_id", user_id).eq("module_id", module_id).execute()
    if existing.data:
        current = sb.table(TABLE_TRAINING_TIME).select("seconds").eq("user_id", user_id).eq("module_id", module_id).execute()
        new_seconds = (current.data[0]["seconds"] if current.data else 0) + seconds
        sb.table(TABLE_TRAINING_TIME).update({"seconds": new_seconds}).eq("user_id", user_id).eq("module_id", module_id).execute()
    else:
        sb.table(TABLE_TRAINING_TIME).insert({
            "user_id": user_id,
            "module_id": module_id,
            "seconds": seconds,
        }).execute()


# ============ MODULES ============

def get_modules():
    sb = get_supabase()
    result = sb.table(TABLE_MODULES).select("id, title, chapters").order("id").execute()
    return {r["id"]: {"title": r["title"], "chapters": json.loads(r["chapters"])} for r in result.data}


def save_module(mid, title, chapters):
    sb = get_supabase()
    existing = sb.table(TABLE_MODULES).select("id").eq("id", mid).execute()
    if existing.data:
        sb.table(TABLE_MODULES).update({
            "title": title,
            "chapters": json.dumps(chapters, ensure_ascii=False),
        }).eq("id", mid).execute()
    else:
        sb.table(TABLE_MODULES).insert({
            "id": mid,
            "title": title,
            "chapters": json.dumps(chapters, ensure_ascii=False),
        }).execute()


def delete_module_db(mid):
    sb = get_supabase()
    sb.table(TABLE_MODULES).delete().eq("id", mid).execute()
    sb.table(TABLE_EXAM_QUESTIONS).delete().eq("module_id", mid).execute()
    sb.table(TABLE_PROGRESS).delete().eq("module_id", mid).execute()
    sb.table(TABLE_EXAM_RESULTS).delete().eq("module_id", mid).execute()
    sb.table(TABLE_TRAINING_TIME).delete().eq("module_id", mid).execute()


def get_next_module_id():
    sb = get_supabase()
    result = sb.table(TABLE_MODULES).select("id").order("id", desc=True).limit(1).execute()
    if result.data:
        return result.data[0]["id"] + 1
    return 1


# ============ EXAM QUESTIONS ============

def get_exams():
    sb = get_supabase()
    result = sb.table(TABLE_EXAM_QUESTIONS).select("module_id, questions, exam_count").execute()
    exam_data = {}
    for r in result.data:
        exam_data[r["module_id"]] = {
            "questions": json.loads(r["questions"]),
            "exam_count": r["exam_count"] or 0,
        }
    return exam_data


def save_exam_questions(mid, questions, exam_count=0):
    sb = get_supabase()
    existing = sb.table(TABLE_EXAM_QUESTIONS).select("module_id").eq("module_id", mid).execute()
    if existing.data:
        sb.table(TABLE_EXAM_QUESTIONS).update({
            "questions": json.dumps(questions, ensure_ascii=False),
            "exam_count": exam_count,
        }).eq("module_id", mid).execute()
    else:
        sb.table(TABLE_EXAM_QUESTIONS).insert({
            "module_id": mid,
            "questions": json.dumps(questions, ensure_ascii=False),
            "exam_count": exam_count,
        }).execute()


def delete_exam_questions_db(mid):
    sb = get_supabase()
    sb.table(TABLE_EXAM_QUESTIONS).delete().eq("module_id", mid).execute()


# ============ DEFAULT DATA ============

def seed_default_modules(modules_dict, exams_dict):
    """种子默认模块和考试（如果表为空）"""
    sb = get_supabase()
    existing = sb.table(TABLE_MODULES).select("id").limit(1).execute()
    if not existing.data:
        for mid, mod in modules_dict.items():
            sb.table(TABLE_MODULES).insert({
                "id": mid,
                "title": mod["title"],
                "chapters": json.dumps(mod["chapters"], ensure_ascii=False),
            }).execute()
        for mid, qs in exams_dict.items():
            sb.table(TABLE_EXAM_QUESTIONS).insert({
                "module_id": mid,
                "questions": json.dumps(qs, ensure_ascii=False),
            }).execute()
