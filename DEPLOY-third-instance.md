# 第三个发信实例部署指南（luna04@myliveology.org · Google Workspace）

同一套代码（`xycream/mail-blaster`，分支 `enterprise-two-product-links`）在 Railway 上跑
**第三个独立服务**，与现有两个实例互不影响：

| 实例 | 发信账号 | 域名 | 界面标识 | 产品链接 |
|---|---|---|---|---|
| 个人邮箱 | `liveology.ny@gmail.com` | `liveologyus.up.railway.app` | Mail Blaster | 单链接 |
| 企业邮箱 A | `allyinfo@myliveology.org` | `mail-blaster-copy-production.up.railway.app` | Liveology 企业邮箱 | 双链接 |
| **企业邮箱 B（本指南）** | `luna04@myliveology.org` | 新建服务自动生成 | Liveology 企业邮箱 · Luna04 | 双链接（可切单） |

工作方式与另两个企业实例相同：**读 Excel → IMAP 写草稿 → 人工在 Gmail 点发送**。
程序不发信、不需要 Google OAuth。

---

## 0. 前置条件（账号本人/管理员）

1. `luna04@myliveology.org` 与 `allyinfo` 同域同策略：若 allyinfo 已走通（已验证），
   IMAP 与 App Password 对该域已放行，**只需账号本人**完成：
   - 开两步验证：👉 https://myaccount.google.com/signinoptions/two-step-verification
   - 生成应用专用密码（App 选 Mail，设备名建议 `Railway Mail Blaster luna04`）：
     👉 https://myaccount.google.com/apppasswords
2. 拿到 16 位 App Password（形如 `abcd efgh ijkl mnop`，填变量时去空格）。

## 1. Railway：新建第三个服务

👉 https://railway.com/dashboard/

1. 进入现有 Project → 对任一企业服务点 `⋯` → **Duplicate service**
   （或在画布点 **+ New Service → GitHub Repo → xycream/mail-blaster**）
2. **关键**：新服务的 **Settings → Source**（或 Repository）里，分支选：
   ```
   enterprise-two-product-links
   ```
   （不要选 main —— main 是单链接版，也没有本次的 PRODUCT_LINKS 开关）
3. 生成域名：**Settings → Networking → Generate Domain**，记下 `xxxx.up.railway.app`
4. **Settings → Volumes → New Volume**，挂载路径：
   ```
   /app/data
   ```

## 2. Variables（环境选 Production，逐条 Update，最后 Redeploy）

```env
SMTP_HOST=smtp.gmail.com
SMTP_PORT=465
SMTP_USERNAME=luna04@myliveology.org
SMTP_PASSWORD=<luna04 的 16 位 App Password，去掉空格>

DRAFT_BACKEND=imap
IMAP_HOST=imap.gmail.com
IMAP_PORT=993

INSTANCE_LABEL=Liveology 企业邮箱 · Luna04
INSTANCE_COLOR=#0d9488

PRODUCT_LINKS=2

UPLOAD_DIR=/app/data/uploads
JOB_STATE_FILE=/app/data/job_state.json
UPLOAD_HISTORY_FILE=/app/data/upload_history.json
GOOGLE_TOKEN_FILE=/app/data/google_tokens.json
GOOGLE_DRAFT_LEDGER=/app/data/draft_ledger.json

APP_PASSWORD=<自定义网页访问密码，只发给实际操作的人>
```

说明：

- `PRODUCT_LINKS=2` → 双产品链接（`{product_url}` + `{product_url_2}` 双按钮）；
  以后想让这个实例回到单链接，把值改成 `1` 并 Redeploy 即可，不用改代码。
- **不要**复制 allyinfo 服务的 `GOOGLE_CLIENT_ID` / `GOOGLE_CLIENT_SECRET` /
  `PUBLIC_BASE_URL`（那是个人实例的 Gmail API 备用通道配置，本实例走 IMAP 用不到，
  复制过来还会让 health 里的 OAuth 回调指向错误域名）。
- `INSTANCE_LABEL` 用 `· Luna04` 后缀区分两个企业实例；顶栏颜色建议用青绿色
  `#0d9488`（个人蓝、allyinfo 紫、luna04 青绿，三色区分）。

## 3. 部署后我来核验

给我新域名后我会检查：

```
GET /api/health
  → build_sha: d2574cf          ← 确认拉的是 enterprise-two-product-links 最新提交
  → instance_label: Liveology 企业邮箱 · Luna04
  → sender: luna04@myliveology.org
  → smtp_configured: true
  → upload_dir: /app/data/uploads   ← Volume 生效
```

然后用 1 行测试名单跑 1 封草稿（收件人写 luna04 自己），你在 Gmail 草稿箱核对
（👉 https://mail.google.com/mail/u/0/#drafts），确认双按钮、变量替换无误后再上真实名单。

## 4. 安全提醒

- App Password 权限 = 整个邮箱的 IMAP 读写，只发给配置的人；跑通后建议重置轮换
- 新邮箱有信誉预热期：第 1 天 1 封 → 第 2 天 20 封 → 观察退信/垃圾箱再逐步放量
  （Workspace 额度约 2000 收件人/天，但新账号别第一天就冲满）
- 与其它实例一样：**一个服务同时只允许一人操作**（服务端名单是单份的）
