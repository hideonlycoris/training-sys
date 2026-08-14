"""
SQLite → Supabase 数据迁移脚本
运行方式: streamlit run migrate_to_supabase.py
或: python migrate_to_supabase.py（需先设置环境变量）

使用前请确保:
1. Supabase 中已执行 supabase_migration.sql 创建表
2. .streamlit/secrets.toml 中配置了 supabase_url 和 supabase_key
3. training.db 文件存在于当前目录
"""

import sqlite3
import json
import os
from supabase import create_client

# 配置
DB_PATH = os.path.join(os.path.dirname(__file__), "training.db")

def get_supabase():
    """从环境变量或 secrets.toml 读取配置"""
    try:
        import streamlit as st
        url = st.secrets["supabase_url"]
        key = st.secrets["supabase_key"]
    except:
        url = os.environ.get("SUPABASE_URL", "")
        key = os.environ.get("SUPABASE_KEY", "")

    if not url or not key:
        print("错误: 未配置 Supabase 连接信息")
        print("请在 .streamlit/secrets.toml 或环境变量中设置 SUPABASE_URL 和 SUPABASE_KEY")
        return None

    return create_client(url, key)


def migrate_users(sb):
    """迁移用户表"""
    print("迁移 users 表...")
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT * FROM users").fetchall()
    conn.close()

    count = 0
    for r in rows:
        try:
            sb.table("users").insert({
                "id": r["id"],
                "username": r["username"],
                "password_hash": r["password_hash"],
                "display_name": r["display_name"],
                "emp_id": r["emp_id"] or "",
                "department": r["department"] or "",
                "role": r["role"] or "user",
                "created_at": r["created_at"] or None,
            }).execute()
            count += 1
        except Exception as e:
            print(f"  跳过用户 {r['username']}: {e}")

    print(f"  完成: {count}/{len(rows)} 条用户记录")


def migrate_progress(sb):
    """迁移学习进度表"""
    print("迁移 progress 表...")
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute("SELECT * FROM progress").fetchall()
    conn.close()

    count = 0
    for r in rows:
        try:
            sb.table("progress").insert({
                "user_id": r[0],
                "module_id": r[1],
                "chapter_idx": r[2],
                "read_done": r[3],
            }).execute()
            count += 1
        except Exception as e:
            print(f"  跳过进度记录: {e}")

    print(f"  完成: {count}/{len(rows)} 条进度记录")


def migrate_exam_results(sb):
    """迁移考试结果表"""
    print("迁移 exam_results 表...")
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute("SELECT * FROM exam_results").fetchall()
    conn.close()

    count = 0
    for r in rows:
        try:
            sb.table("exam_results").insert({
                "id": r[0],
                "user_id": r[1],
                "module_id": r[2],
                "score": r[3],
                "answers": r[4],
                "taken_at": r[5] or None,
            }).execute()
            count += 1
        except Exception as e:
            print(f"  跳过考试记录: {e}")

    print(f"  完成: {count}/{len(rows)} 条考试记录")


def migrate_training_time(sb):
    """迁移培训时间表"""
    print("迁移 training_time 表...")
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute("SELECT * FROM training_time").fetchall()
    conn.close()

    count = 0
    for r in rows:
        try:
            sb.table("training_time").insert({
                "user_id": r[0],
                "module_id": r[1],
                "seconds": r[2],
            }).execute()
            count += 1
        except Exception as e:
            print(f"  跳过时间记录: {e}")

    print(f"  完成: {count}/{len(rows)} 条时间记录")


def migrate_modules(sb):
    """迁移模块表"""
    print("迁移 modules 表...")
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute("SELECT * FROM modules").fetchall()
    conn.close()

    count = 0
    for r in rows:
        try:
            sb.table("modules").insert({
                "id": r[0],
                "title": r[1],
                "chapters": r[2],
            }).execute()
            count += 1
        except Exception as e:
            print(f"  跳过模块: {e}")

    print(f"  完成: {count}/{len(rows)} 条模块记录")


def migrate_exam_questions(sb):
    """迁移考题表"""
    print("迁移 exam_questions 表...")
    conn = sqlite3.connect(DB_PATH)
    # 检查是否有 exam_count 列
    cursor = conn.execute("PRAGMA table_info(exam_questions)")
    columns = [col[1] for col in cursor.fetchall()]
    has_exam_count = "exam_count" in columns

    if has_exam_count:
        rows = conn.execute("SELECT * FROM exam_questions").fetchall()
    else:
        rows = conn.execute("SELECT module_id, questions, 0 FROM exam_questions").fetchall()
    conn.close()

    count = 0
    for r in rows:
        try:
            sb.table("exam_questions").insert({
                "module_id": r[0],
                "questions": r[1],
                "exam_count": r[2] if has_exam_count else 0,
            }).execute()
            count += 1
        except Exception as e:
            print(f"  跳过考题: {e}")

    print(f"  完成: {count}/{len(rows)} 条考题记录")


def main():
    print("=" * 50)
    print("Training-SYS: SQLite → Supabase 数据迁移")
    print("=" * 50)

    # 检查 SQLite 文件
    if not os.path.exists(DB_PATH):
        print(f"错误: 找不到 {DB_PATH}")
        print("请确保 training.db 文件在当前目录")
        return

    # 连接 Supabase
    sb = get_supabase()
    if not sb:
        return

    print(f"数据库文件: {DB_PATH}")
    print()

    # 执行迁移
    migrate_users(sb)
    migrate_modules(sb)
    migrate_exam_questions(sb)
    migrate_progress(sb)
    migrate_exam_results(sb)
    migrate_training_time(sb)

    print()
    print("=" * 50)
    print("迁移完成！")
    print("请在 Supabase Dashboard 中检查数据是否正确")
    print("=" * 50)


if __name__ == "__main__":
    main()
