# 第二个发信实例部署指南（Google Workspace 企业邮箱）

同一套代码（`xycream/mail-blaster`）在 Railway 上跑**两个独立服务**，各自一套凭据、
各自的名单与日志，互不影响。不需要 fork 仓库、不需要新代码。

| 实例 | 发信账号 | Railway 域名 | 界面顶部标识 |
|---|---|---|---|
| 个人邮箱（已有） | `liveology.ny@gmail.com` | `liveologyus.up.railway.app` | `Mail Blaster` |
| **企业邮箱（本指南）** | `allyinfo@myliveology.org` | 新建服务自动生成 | `Liveology 企业邮箱` |

两个实例都做同一件事：**读 Excel → 写 Gmail 草稿 → 人工在 Gmail 点发送**（IMAP 通道，
程序不发信、不申请任何发信权限）。

---

## 0. 前置条件（必须 Workspace 管理员配合，这步不通后面全白做）

在 **Google Admin Console**（https://admin.google.com）确认三项：

1. **IMAP 已开启**
   `Apps → Google Workspace → Gmail → User settings → 选择该账号所在部门`
   → 勾选 **"Let users access Gmail through IMAP"**
2. **允许应用专用密码 / 不受限的应用**
   `Security → Authentication → Settings for app passwords and less secure apps`
   → 不能被组织策略关闭（Workspace 若禁用 App Password，就必须走 OAuth，本指南不适用）
3. **该账号已开启两步验证**（App Password 的前置条件）
   由 `allyinfo@myliveology.org` 自己登录 Google 账号安全页开启 2SV 后，
   才能在 https://myaccount.google.com/apppasswords 生成 16 位应用专用密码

拿到 **16 位 App Password**（形如 `abcd efgh ijkl mnop`）才能继续。

### Workspace 发信额度（比个人 Gmail 宽松）

* 约 **2000 收件人/天**（个人 Gmail 是 500）
* 单封外部收件人 ≤ 100
* 新邮箱/新域名有预热期：批量发之前先小量试跑，别第一天就上千

---

## 1. Railway：复制出第二个服务

1. 打开 https://railway.com/dashboard/ ，进入现有 Project
2. 点 **"+" / New Service → Deploy from GitHub repo**，仓库选 `xycream/mail-blaster`
   （也可以对现有服务点 `⋯ → Duplicate service`，更快）
3. 新服务会自动生成域名，记下它（形如 `xxxx-yyyy.up.railway.app`）

> 两个服务共用仓库没问题：改动会同时部署到两边，各自读自己的环境变量。

## 2. 给新服务配 Volume（重要）

**Settings → Volumes → New Volume**，挂载路径填：

```
/app/data
```

不挂 Volume 的话，重新部署会丢掉：任务状态、名单上传历史、草稿去重台账。

## 3. 新服务的 Variables

**Variables 标签，环境选 Production**，逐条填写并**每条点 Update**，全部填完点 **Deploy/Redeploy**：

```
# ── 企业邮箱凭据 ──
SMTP_HOST        smtp.gmail.com
SMTP_PORT        465
SMTP_USERNAME    allyinfo@myliveology.org
SMTP_PASSWORD    <刚生成的 16 位 App Password，去掉空格>

# ── 草稿模式 ──
DRAFT_BACKEND    imap

# ── 界面标识（防混淆，必须与个人实例不同）──
INSTANCE_LABEL   Liveology 企业邮箱
INSTANCE_COLOR   #7c3aed

# ── 持久化路径（Volume 挂载在 /app/data）──
UPLOAD_DIR           /app/data/uploads
JOB_STATE_FILE       /app/data/job_state.json
UPLOAD_HISTORY_FILE  /app/data/upload_history.json
GOOGLE_TOKEN_FILE    /app/data/google_tokens.json
GOOGLE_DRAFT_LEDGER  /app/data/draft_ledger.json

# ── 访问密码：强烈建议设置，否则任何拿到域名的人都能进后台 ──
APP_PASSWORD     <自定义密码，只发给实际操作的人>
```

> `INSTANCE_LABEL` / `INSTANCE_COLOR` 是 commit `f023cac` 新加的：页面标题、顶栏颜色、
> 以及每次「发送 / 建草稿」的确认框都会显示**当前真实发信账号**，避免把企业名单传到
> 个人实例（历史教训：曾误发 59 封旧名单）。

## 4. 验证（我来做，你只需告诉我新域名）

我会依次检查：

```
GET https://<新域名>/api/health
  → "instance_label": "Liveology 企业邮箱"
  → "sender": "allyinfo@myliveology.org"        ← 必须是企业邮箱
  → "smtp_configured": true
  → "upload_dir": "/app/data/uploads"           ← 说明 Volume 生效

GET https://<新域名>/api/draft_backend
  → "imap": { "ready": true, "account": "allyinfo@myliveology.org" }
```

## 5. 首跑 SOP（和现有流程一致，别跳）

1. 打开新实例页面，**先确认顶栏写着「Liveology 企业邮箱」**
2. 上传 Excel → 确认行数正确（页面会显示读的是哪个工作表）
3. 填主题 / 模板 / 本轮统一参数
4. 「本次创建数量」先填 **1** → 点 **📝 创建 Gmail 草稿**
5. 用 `allyinfo@myliveology.org` 登录 Gmail → 打开
   👉 https://mail.google.com/mail/u/0/#drafts
   核对：收件人、变量替换、图片、链接、发件人显示名
6. 单封没问题 → 再填 20 / 50 逐步放大
7. **发送永远是人工在 Gmail 里点**，程序只写草稿

## 6. 常见问题

| 现象 | 原因 | 处理 |
|---|---|---|
| `IMAP APPEND` 报 `LOGIN FAILED` | App Password 错 / 2SV 未开 / 管理员禁用 | 重新生成 App Password；确认第 0 节三项 |
| 报 `该邮箱没有 Drafts 文件夹` | IMAP 未对该账号开启 | 管理员开 IMAP 后等几分钟 |
| 登录 Gmail 时被要求验证/账号受限 | 新邮箱批量行为触发风控 | 降低每日量、加预热；不要短时间大量发 |
| 草稿箱里看不到 | 登错账号（个人 Gmail 也有一个实例） | 看顶栏标识 + Gmail 右上角头像确认账号 |

## 7. 已实测结论（2026-10-19，`allyinfo@myliveology.org`）

用线上同一份代码直接对该企业邮箱跑过，结果：

- IMAP 登录成功：`allyinfo@myliveology.org authenticated (Success)`
- **注意：该邮箱界面语言是中文，IMAP 里草稿箱名是 `[Gmail]/&g0l6Pw-`（"草稿" 的
  modified-UTF7 编码），不是 `[Gmail]/Drafts`。** 代码按 LIST 的 `\Drafts` 属性动态
  解析文件夹名，因此自动兼容，无需配置；遇到其它语言界面同样适用。
- APPEND 写入成功，读回确认带 `\Draft` 标志、From/To/Subject 与 HTML 正文完整

---

## 附：个人实例要不要也补 `INSTANCE_LABEL`

建议给现有 `liveologyus` 服务加两个变量，让两边一眼可辨：

```
INSTANCE_LABEL   Liveology 个人邮箱
INSTANCE_COLOR   #1a73e8
```

（不填则默认显示 `Mail Blaster`，蓝色，功能不受影响。）
