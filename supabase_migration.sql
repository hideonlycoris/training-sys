-- Training-SYS Supabase 建表语句
-- 在 Supabase Dashboard → SQL Editor 中执行此脚本

-- 1. users 表
CREATE TABLE IF NOT EXISTS users (
    id BIGSERIAL PRIMARY KEY,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    display_name TEXT NOT NULL,
    emp_id TEXT DEFAULT '',
    department TEXT DEFAULT '',
    role TEXT DEFAULT 'user',
    created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 2. progress 表
CREATE TABLE IF NOT EXISTS progress (
    user_id BIGINT,
    module_id BIGINT,
    chapter_idx BIGINT,
    read_done INTEGER DEFAULT 0,
    PRIMARY KEY (user_id, module_id, chapter_idx)
);

-- 3. exam_results 表
CREATE TABLE IF NOT EXISTS exam_results (
    id BIGSERIAL PRIMARY KEY,
    user_id BIGINT,
    module_id BIGINT,
    score INTEGER,
    answers TEXT,
    taken_at TIMESTAMPTZ DEFAULT NOW()
);

-- 4. training_time 表
CREATE TABLE IF NOT EXISTS training_time (
    user_id BIGINT,
    module_id BIGINT,
    seconds INTEGER DEFAULT 0,
    PRIMARY KEY (user_id, module_id)
);

-- 5. modules 表
CREATE TABLE IF NOT EXISTS modules (
    id BIGINT PRIMARY KEY,
    title TEXT NOT NULL,
    chapters TEXT NOT NULL
);

-- 6. exam_questions 表
CREATE TABLE IF NOT EXISTS exam_questions (
    module_id BIGINT PRIMARY KEY,
    questions TEXT NOT NULL,
    exam_count INTEGER DEFAULT 0
);

-- 启用 Row Level Security（可选，增强安全性）
-- ALTER TABLE users ENABLE ROW LEVEL SECURITY;
-- ALTER TABLE progress ENABLE ROW LEVEL SECURITY;
-- ALTER TABLE exam_results ENABLE ROW LEVEL SECURITY;
-- ALTER TABLE training_time ENABLE ROW LEVEL SECURITY;
-- ALTER TABLE modules ENABLE ROW LEVEL SECURITY;
-- ALTER TABLE exam_questions ENABLE ROW LEVEL SECURITY;
