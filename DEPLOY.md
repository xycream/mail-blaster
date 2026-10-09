# Mail Blaster — Railway 部署指南

> 把你的邮件群发工具部署到 Railway，让它 24/7 在线、有正式 HTTPS 域名。

---

## 部署架构

```
GitHub repo (你的代码)
        ↓ push
Railway (自动 build)
        ↓ 启动
Docker container (uvicorn:8000)
        ↓ 读取
Env Vars (SMTP / 密钥)
        ↓ 持久化
Volume /data (上传文件 / 任务状态)
```

---

## 步骤 1: 准备 GitHub 仓库

把 `mail-app/` 目录下的代码（不含 `uploads/` `data/` `.smtp_config.json`）推到一个空的 GitHub 仓库。

---

## 步骤 2: Railway 创建项目

1. 登录 https://railway.com
2. **New Project** → **Deploy from GitHub repo**
3. 选你的仓库，Railway 自动检测 Dockerfile 并开始构建

---

## 步骤 3: 绑定持久化 Volume（重要）

**否则 Railway 重启后上传的图片、任务状态会丢失！**

1. 项目页面 → **+ New** → **Volume**
2. 挂载路径填 `/app/data`
3. 在 Mail Blaster service 里点 **Variables** → 暂时不用设，但确认 Volume 已 attach 到此 service

或者：把 Volume 挂载到 `/app/uploads` 也行（图片存储）—— **建议挂两个**：
- Volume A → `/app/uploads`（图片）
- Volume B → `/app/data`（任务状态）

---

## 步骤 4: 配置环境变量

Mail Blaster service → **Variables** → 逐个填入：

| 变量名 | 值 | 说明 |
|---|---|---|
| `SMTP_HOST` | `smtp.gmail.com` | Gmail SMTP |
| `SMTP_PORT` | `465` | SSL 端口 |
| `SMTP_USERNAME` | `liveology.ny@gmail.com` | 发件邮箱 |
| `SMTP_PASSWORD` | `你的16位App Password` | Gmail App Password（不是登录密码）|
| `SENDER_NAME` | `Luna Hei` | 收件人看到的发件人 |
| `APP_PASSWORD` | `你自己设的访问密码` | **🔒 链接防泄漏门禁**（必填，见下方说明）|
| `PORT` | Railway 自动设置 | 不要手动改 |

⚠️ **绝对不要**把 `SMTP_PASSWORD` 写到代码里或者 commit 到 Git。

### 怎么生成 Gmail App Password

1. 登录 Gmail 账号 → https://myaccount.google.com/security
2. 开启 **两步验证**（必须）
3. 进入 **应用专用密码** → 选择「其他(自定义)」→ 命名「Mail Blaster」
4. 复制生成的 16 位字符（带空格也行，会自动去除）

---

## 步骤 5: 部署

1. 填好变量后，Railway 自动重新部署
2. 在 **Settings** → **Networking** → **Generate Domain** 拿到正式域名
   - 例如 `mail-blaster-production.up.railway.app`
3. 打开域名，看到 Mail Blaster 界面 = 部署成功 ✅
4. 访问 `https://你的域名/api/health` 检查健康状态

---

## 验证清单

- [ ] `https://你的域名/api/health` 返回 `{"status":"ok",...}`
- [ ] 打开首页看到 SMTP 已配置（前端不显示密码）
- [ ] 测发一封 TEST 邮件到自己邮箱确认 SMTP 通
- [ ] 上传一张图片，刷新后图片还在（说明 Volume 生效）
- [ ] 关掉 Railway service 再开 → 任务状态还在

---

## 本地开发（可选）

```bash
cd mail-app
pip install -r requirements.txt
python3 main.py
# 访问 http://localhost:8000
```

本地模式会自动读 `.smtp_config.json`（已 gitignore）。
