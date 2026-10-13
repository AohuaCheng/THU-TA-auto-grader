# THU-TA-auto-grader

清华大学（THU）**助教/教师**用的半自动批改工具：从网络学堂下载作业、雨课堂 AI 批改、分数汇总、生成可上传网络学堂的反馈文档，并批量回传成绩与评语。

> 非清华官方软件，仅供本人课程、本人账号使用。使用前请阅读文末「安全说明」。

基于 [learn2018-autodown](https://github.com/Trinkle23897/learn2018-autodown) 的浏览器登录与助教端 API 改造。

## 功能一览

| 模块 | 说明 |
|------|------|
| 网络学堂 | 批量下载学生提交、查看名单、CSV 批量回传成绩与**评语附件** |
| 雨课堂 | Reflective 规则 AI 批改，自动下载带评分的 feedback docx |
| 分数汇总 | 从 feedback 导出 xlsx；**生成符合上传网络学堂格式的反馈 docx** |
| 字数检查 | 统计正文英文词数（不含题目与参考文献），在 xlsx 中标注明显超出范围 |
| 修复工具 | 批量批改后 feedback 错位时，按正文相似度重新对齐 |

## 部署方式（三选一）

| 方式 | 适合谁 | 怎么做 |
|------|--------|--------|
| **A. 一键脚本（推荐新手）** | 不想研究 uv / pip | 克隆仓库后执行 `./scripts/setup.sh`，日常用 `./scripts/menu.sh` 或 `./scripts/run.sh …` |
| **B. uv** | 已安装 [uv](https://docs.astral.sh/uv/) | `uv sync`，命令形如 `uv run python main.py …` |
| **C. pip + venv** | 只有系统自带的 Python | `python3 -m venv .venv` → `pip install -e .`，用 `.venv/bin/python main.py …` |

**Mac / Linux 一键安装与菜单：**

```bash
cd THU-TA-auto-grader
chmod +x scripts/*.sh
./scripts/setup.sh      # 自动选 uv 或 venv；创建 ../downloads、../templates；复制 .env 示例
./scripts/menu.sh       # 数字菜单，按 Essay 流程操作
# 或直接: ./scripts/run.sh login
```

**Windows（PowerShell）：**

```powershell
cd THU-TA-auto-grader
.\scripts\setup.ps1
.\scripts\run.ps1 login
```

下文命令行示例同时给出 **`./scripts/run.sh`** 与 **`uv run python main.py`** 两种写法，效果相同。

## 推荐给其他助教：第一次怎么用

下面按**真实操作顺序**写，不需要会写代码，只要在终端里复制粘贴命令即可（Mac 打开「终端」，Windows 可用 PowerShell）。

### 0. 准备（只需做一次）

1. **安装 Google Chrome**（登录清华 SSO、雨课堂时会自动弹出）。
2. **安装 Python 环境**（任选其一）：
   - **推荐**：运行 `./scripts/setup.sh`（有 uv 会用 uv，没有则用 venv + pip）；
   - 或自行安装 [uv](https://docs.astral.sh/uv/getting-started/installation/) 后 `uv sync`；
   - 或 Python 3.10+ 下 `python3 -m venv .venv && .venv/bin/pip install -e .`。
3. **建好工作区文件夹**（名字可自定，下面用 `Essays` 举例）：

```
Essays/
├── THU-TA-auto-grader/    ← git clone 本仓库到这里
├── downloads/             ← 程序会自动创建子文件夹，也可先手动建好
└── templates/             ← 放课程提供的 xlsx/docx 模板（见下）
```

4. **向课程负责人或上一届助教索取**（本仓库**不包含**这些文件）：
   - 分数汇总 xlsx 模板（如 `MINC 2025 Essay 分数汇总 样例.xlsx`）
   - 网络学堂上传用反馈 docx 模板（如 `仅反馈样例 这样上传至网络学堂.docx`）
5. 克隆仓库并安装依赖：

```bash
cd Essays
git clone https://github.com/AohuaCheng/THU-TA-auto-grader.git
cd THU-TA-auto-grader
./scripts/setup.sh
```

（若不用脚本：`uv sync` 或见上文「部署方式 C」。）

6. **配置账号与 URL**（`setup.sh` 已复制示例文件，若未运行脚本则手动执行）：

```bash
cp .env.example .env
cp credentials.local.env.example credentials.local.env
```

- `credentials.local.env`：填 `THU_USERNAME`、`THU_PASSWORD`（网络学堂登录）。
- `.env`：
  - `THU_ASSIGNMENT_URL`：打开网络学堂 → 某次作业 → **助教批改列表页**，从浏览器地址栏**整段复制**（须含 `wlkcid` 和 `zyid`）。
  - `YUKETANG_*`：在雨课堂「智能批改」页面 URL 或页面元素里找 workspace / rule id（第一次可请会用工具的助教帮填一次）。

### 1. 登录（每次 session 过期后重做）

```bash
./scripts/run.sh login                   # 网络学堂（会弹 Chrome，完成 SSO）
./scripts/run.sh yuketang-login          # 雨课堂（常需微信扫码）
# 等价: uv run python main.py login
```

登录成功后会生成本地的 `session.json` / `yuketang_session.json`（已 git 忽略，勿外传）。

### 2. 一次 Essay 批改的完整流程

| 步骤 | 命令 | 说明 |
|------|------|------|
| ① 下载 | `./scripts/run.sh download` | 作业保存到 `../downloads/<作业名>/学号_姓名/` |
| ② 雨课堂批改 | `./scripts/run.sh grade --dir ../downloads/Essay1` | 已有 feedback 的会自动跳过；中断可重跑 |
| ③ 导出分数 | `./scripts/run.sh export-scores --dir ../downloads/Essay1` | 生成 `../downloads/Minc2026Essays分数汇总.xlsx` |
| ④ 生成上传用反馈 | `./scripts/run.sh format-feedback --dir ../downloads/Essay1` | 生成符合网络学堂上传格式的 `Firstround_feedback_学号.docx` |
| ⑤ 字数列（可选） | `./scripts/run.sh annotate-word-count` | 在 xlsx 增加「字数」列 |
| ⑥ 准备上传 CSV | `./scripts/run.sh prepare-upload` | 生成 `../downloads/Essay1/learn_upload.csv` |
| ⑦ 回传网络学堂 | 先 dry-run，再 `--confirm` | 见下文；菜单里选 9 / 10 |

**回传网络学堂（务必先 dry-run）：**

```bash
./scripts/run.sh upload --grades ../downloads/Essay1/learn_upload.csv
# 确认终端无报错后：
./scripts/run.sh upload --grades ../downloads/Essay1/learn_upload.csv --confirm
```

`dry-run` = **只检查、不提交**；`--confirm` = **真正写入网络学堂**。

### 3. 容易踩坑的地方（给其他助教）

| 困惑点 | 说明 |
|--------|------|
| 命令要在哪个目录跑？ | 在 **`THU-TA-auto-grader/`** 下运行 `./scripts/…`；脚本会自动进入正确目录。 |
| 不想记命令？ | 用 **`./scripts/menu.sh`** 按数字选择。 |
| `../downloads/Essay1` 是什么？ | `Essay1` 是下载时根据作业标题生成的文件夹名，以你机器上实际文件夹为准（如 `Essay1`）。 |
| 模板找不到 | `export-scores` / `format-feedback` 会读 `../templates/`，模板需自己放进该目录。 |
| 雨课堂批改卡住 | 保持 Chrome 窗口**不要关**；可看 `grade --limit 1` 先测一人。 |
| feedback 和 student 对不上 | `uv run python fix_feedback_alignment.py --dir ../downloads/Essay1`（无 uv 时用 `.venv/bin/python fix_feedback_alignment.py …`）。 |
| PDF/txt 作业 | 下载后会自动转成 docx 批改，原文件仍保留；汇总表「备注」里可能提示格式问题。 |
| 上传后学生看不到附件 | 在网络学堂批阅页检查「评语附件」是否上传成功；本工具上传的是 `Firstround_feedback_学号.docx`。 |
| 优秀作业 | 在 xlsx「优秀作业」列填 `1`，`prepare-upload` 会带入 CSV，上传时勾选网络学堂「设为优秀作业」。 |

## 目录结构（与 git 的关系）

```
Essays/                      # 工作区根（一般不必是 git 仓库）
├── THU-TA-auto-grader/      # 本仓库
├── downloads/               # 学生作业、feedback、汇总表（不提交 git）
└── templates/               # 课程模板（不提交 git）
```

程序默认读写上级目录的 `downloads/` 与 `templates/`（可在 `.env` 修改）。

## 配置说明

**登录凭据** — `credentials.local.env`：

```env
THU_USERNAME=你的学工号
THU_PASSWORD=你的密码
```

**作业与雨课堂** — `.env` 示例见 `.env.example`。

## 命令参考（简）

### 雨课堂 AI 批改

```bash
./scripts/run.sh grade --dir ../downloads/Essay1
./scripts/run.sh grade --dir ../downloads/Essay1 --regrade-only
./scripts/run.sh grade --dir ../downloads/Essay1 --student 2023010785
./scripts/run.sh grade --dir ../downloads/Essay1 --dry-run --limit 3
```

feedback 保存为 `{原稿名}_feedback.docx`。文件名含 `#` 的 docx 上传雨课堂时会自动复制为 `No.` 版本。

### 分数汇总与网络学堂格式反馈

```bash
./scripts/run.sh export-scores --dir ../downloads/Essay1
./scripts/run.sh format-feedback --dir ../downloads/Essay1
./scripts/run.sh annotate-word-count
```

`format-feedback` 输出的 `Firstround_feedback_学号.docx` 即为**符合网络学堂「评语附件」要求**的 docx（由雨课堂 feedback 中的 Review Comments 生成）。

## 安全说明

- 请仅在自己的账号、自己的课程上使用。
- `credentials.local.env`、`session.json`、`yuketang_session.json` 含敏感信息，已加入 `.gitignore`，**勿提交或转发**。
- 不要将含学生姓名的 xlsx/csv 上传到公开仓库。

## 已知限制

1. 网络学堂上传依赖网页表单解析，平台改版后可能需要更新工具。
2. 首次上传务必 dry-run，并抽 1–2 人在网页上核对成绩、评语、附件。
3. 工具面向单次作业流程，不做整门课镜像备份。
