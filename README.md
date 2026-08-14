# 📚 新员工培训系统（云端版）

基于 Streamlit 的在线培训系统，支持账号管理、课程学习、在线考试、证书生成。

## 🚀 部署到 Streamlit Cloud（免费）

### 第一步：上传到 GitHub
1. 在 GitHub 创建新仓库（如 `training-sys`）
2. 把本文件夹所有文件上传到仓库
3. 注意：不要上传 `training.db`、`uploads/`、`app_local.py`、`.streamlit/secrets.toml`

### 第二步：连接 Streamlit Cloud
1. 访问 https://share.streamlit.io
2. 用 GitHub 账号登录
3. 点 "New app"
4. 选择你的仓库、选 `app.py`、点 "Deploy"

### 第三步：配置 Supabase 数据库
1. 访问 https://supabase.com 注册账号（免费）
2. 创建新项目，记录项目 URL 和 API Key
3. 在 Supabase Dashboard → SQL Editor 中执行 `supabase_migration.sql` 创建数据表
4. 如果有旧数据，在本地运行迁移脚本：`python migrate_to_supabase.py`

### 第四步：配置密钥
1. 在 Streamlit Cloud 点你的 app → "Settings"
2. 找到 "Secrets" 部分
3. 粘贴以下内容（替换为你的真实值）：
```toml
# 管理员登录密码（重要：改成你自己的强密码）
admin_password = "你的管理员密码"

# 密码重置默认值
default_reset_password = "你的重置密码"

# Supabase 数据库配置
supabase_url = "https://xxx.supabase.co"
supabase_key = "sb_publishable_xxx"

# Gemini AI 配置
[gemini]
api_key = "你的Gemini API Key"
model_name = "gemini-2.5-flash"
use_ai_scoring = true
```

### 第五步：完成
等待部署完成，获得一个 `https://xxx.streamlit.app` 的链接

---

## ⚠️ 云端版与本地版的区别

| 功能 | 本地版 | 云端版 |
|-----|-------|-------|
| 账号管理 | ✅ | ✅ |
| 课程学习 | ✅ | ✅（仅支持下载预览） |
| 在线考试 | ✅ | ✅ |
| PDF 证书 | ✅ | ✅ |
| 课件上传 | ✅（自动转PDF预览） | ✅（直接下载） |
| 数据持久性 | ✅ 永久 | ✅ 永久（Supabase 云数据库） |
| 简答AI打分 | ✅ | ✅（需配置 Gemini Key） |

## 📝 本地运行
```bash
pip install -r requirements.txt
streamlit run app.py
```

## 🔑 管理员账号
- 用户名：admin
- 密码：在 Streamlit Secrets 中配置的 `admin_password` 值
