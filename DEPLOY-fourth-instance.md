# 第四个发信实例部署指南（luna03@myliveology.org · Google Workspace）

与 luna04 完全相同的模式：同一套代码（`xycream/mail-blaster`，分支
`enterprise-two-product-links`）在 Railway 起**第四个独立服务**。四个实例互不影响，
代码零改动——双链接、实例标识、草稿通道全部由环境变量控制。

| 实例 | 发信账号 | 域名 | 界面标识 | 顶栏颜色 |
|---|---|---|---|---|
| 个人邮箱 | `liveology.ny@gmail.com` | `liveologyus.up.railway.app` | Mail Blaster | 蓝 #1a73e8 |
| 企业 A | `allyinfo@myliveology.org` | `mail-blaster-copy-production.up.railway.app` | Liveology 企业邮箱 | 紫 #7c3aed |
| 企业 B | `luna04@myliveology.org` | `mail-blaster-copy-copy-production.up.railway.app` | Liveology 企业邮箱 · Luna04 | 青绿 #0d9488 |
| **企业 C（本指南）** | `luna03@myliveology.org` | 新建服务自动生成 | Liveology 企业邮箱 · Luna03 | 橙 #d97706 |

工作方式与其它实例相同：**读 Excel → IMAP 写草稿 → 人工在 Gmail 点发送**。
程序不发信、不需要 Google OAuth。

---

## 0. 前置条件（账号本人）

同域策略已放行（allyinfo / luna04 已走通），只需 `luna03@myliveology.org` 本人：

1. 开两步验证：👉 https://myaccount.google.com/signinoptions/two-step-verification
2. 生成应用专用密码（App 选 Mail，设备名建议 `Railway Mail Blaster luna03`）：
   👉 https://myaccount.google.com/apppasswords
3. 拿到 16 位 App Password（填变量时去掉空格）。

## 1. Railway：新建第四个服务

👉 https://railway.com/dashboard/

1. 对 luna04 服务（或任一企业服务）点 `⋯` → **Duplicate service**
2. **Settings → Source** 分支确认仍是：
   ```
   enterprise-two-product-links
   ```
3. **Settings → Networking → Generate Domain**，记下新域名
4. **Settings → Volumes → New Volume**，挂载路径：
   ```
   /app/data
   ```

## 2. Variables（环境选 Production，逐条 Update，最后 Redeploy）

```env
SMTP_HOST=smtp.gmail.com
SMTP_PORT=465
SMTP_USERNAME=luna03@myliveology.org
SMTP_PASSWORD=<luna03 的 16 位 App Password，去掉空格>

DRAFT_BACKEND=imap
IMAP_HOST=imap.gmail.com
IMAP_PORT=993

INSTANCE_LABEL=Liveology 企业邮箱 · Luna03
INSTANCE_COLOR=#d97706

PRODUCT_LINKS=2

UPLOAD_DIR=/app/data/uploads
JOB_STATE_FILE=/app/data/job_state.json
UPLOAD_HISTORY_FILE=/app/data/upload_history.json
GOOGLE_TOKEN_FILE=/app/data/google_tokens.json
GOOGLE_DRAFT_LEDGER=/app/data/draft_ledger.json

APP_PASSWORD=<自定义网页访问密码，只发给实际操作的人>
```

⚠️ 三个不要（从 luna04 复制服务时最容易带过来的脏变量）：

- **删掉 `GOOGLE_CLIENT_ID` / `GOOGLE_CLIENT_SECRET`**（个人实例的 Gmail API 备用通道
  凭据，本实例走 IMAP 用不到，留着只会让 health 里 OAuth 回调指向错误域名）
- **删掉或改掉 `PUBLIC_BASE_URL`**（若保留，值应是本实例新域名；不建议保留）
- **核对 `INSTANCE_LABEL` / `INSTANCE_COLOR` / `SMTP_*` 全部换成本实例的值**——
  Duplicate 会原样带过来，不核对就直接用旧邮箱发信了

## 3. 部署后核验（发新域名给我）

```
GET /api/health
  → build_sha: 1370852（或更新）
  → instance_label: Liveology 企业邮箱 · Luna03
  → sender: luna03@myliveology.org
  → smtp_configured: true
  → upload_dir: /app/data/uploads
```

然后用 1 行测试名单跑 1 封草稿（收件人写 luna03 自己），在 Gmail 草稿箱核对
（👉 https://mail.google.com/mail/u/0/#drafts）双按钮与变量替换无误，再上真实名单。

## 4. 安全与纪律（同其它实例）

- App Password = 整个邮箱的 IMAP 读写权限，只发给配置的人；跑通后建议重置轮换
- 新邮箱预热：第 1 天 1 封 → 第 2 天 20 封 → 观察退信/垃圾箱再放量
- 一个服务同一时间只允许一人操作网页
- 与其它三个实例可**同时**跑建草稿任务，互不影响（各自的容器、名单、草稿箱）
