# Content Production Lab · 中文内容生产线

> 一个以 **Claude Code 技能包（Skills）** 为引擎、以**中文内容生产**为交付物的个人工作站。
> 覆盖「选题调研 → 深度研究 → 公众号文章 → 短视频脚本 → 专业 DOCX / PDF / PPT / 表格」的完整生产线，并把流程沉淀为**可复用、可验证**的技能包与分类归档目录。

所有文档类交付物均以 **`.docx`** 产出，并在交付前通过 **渲染 → 逐页视觉验收** 的质量门（见 [§ 质量门](#-质量门-renderandverify)），保证版面无误。

---

## 📌 项目定位

本仓库**不是软件工程**：无构建、无测试。它是由三部分构成的**内容生产线**：

1. **技能包**（`.claude/skills/`）：把可复用的创作流程封装为可被 Claude 触发的 Skill；
2. **分类工作区**（`dayouyuan/`、`geopolitical/`、`qiminyaosu/`、`reports/`、`series/` 等）：按生产日期/主题归档交付物；
3. **知识资产**（`research/`、`出行攻略/`、`课程教学/`、`stock-fund/`）：长期沉淀的研究报告与技术资料。

> 中文优先：正文、命名、交付物均为中文；文档类交付以 `.docx` 为准。

---

## ✨ 核心能力（技能包）

| 能力 | 作用 | 交付 |
|------|------|------|
| **wechat-high-energy-commentary** | 调研选题 → 高能中文时评公众号文章（配图建议） | DOCX |
| **geopolitical-deep-analysis-wechat** | 地缘政治深度分析：多学派解读、情景推演、战略判断 | DOCX |
| **article-to-short-video-script** | 已完成文章 → 可直接配音/剪辑的短视频执行稿（微信视频号/抖音/B站） | DOCX |
| **modern-qimin-jimeng-video** | 现代齐民要术六卷（衣食住行养用）生活科普 → 即梦 AI 动画分镜稿 | DOCX |
| **docx** | `.docx` 全流程：创建/编辑/批注/修订/合并/水印/审校 | DOCX |
| **pdfs** | PDF 处理：读取/OCR/表单/涂改/转换/差异对比（render→verify→operate→re-verify） | PDF |
| **slides** | 基于 pptxgenjs 的幻灯片制作与导出（含十余个专业模板） | PPTX |
| **spreadsheets** | 表格处理：读取/清洗/转换/生成 | XLSX |

> 技能位于 `.claude/skills/<name>/`，`SKILL.md` 携带 `name` + `description` 元数据，可被 Claude 自动匹配触发或 `/name` 手动调用；内部按 `SKILL.md` → `references/*.md` 渐进展开。

---

## 🗂 目录结构

```
content-prod-lab/
├── CLAUDE.md                # 权威项目指南（先读）
├── AGENTS.md                # 技能包/工作区检索速查表
├── README.md                # 本文件
├── docs/gotchas.md          # 踩坑日志（可追加）
│
├── .claude/
│   ├── skills/              # 8 个可触发技能包
│   └── settings.local.json  # 权限 allow-list
│
│   # —— 生产流水线工作区 ——
├── dayouyuan/               # 公众号文章 → 短视频流水线（prompts/ + articles/YYYYMMDD/）
├── geopolitical/            # 地缘政治深度分析（prompts/ + articles/YYYYMMDD/）
├── qiminyaosu/              # 现代齐民要术（选题圣经 + prompts/ + articles/）
├── series/                  # 系列长文（每系列一目录：钢斧之后/，命名 系列名_第X章_标题）
│
│   # —— 交付物归档库 ——
├── reports/                 # 城市/机构年度报告深度研究（110 项）
├── research/                # 深度研究报告库（按学科分子目录，见下）
├── stock-fund/              # 权益基金/股票研究 + 数据管线
├── opencode/                # OpenCode/Claude Code 技术交付（cookbook/ 教程/ 方案/ 汇报/ 深度研究）
├── proposals/               # 技术/产品方案
├── 出行攻略/                # 城市两日游/深度旅行攻略（南京/宁波/扬州/东京 等）
└── 课程教学/                # 经济学讲义与课件（宏观/微观/货币金融学）
```

**`research/` 学科分类**（按主题归档，便于检索）：

```
research/
├── AI与训练          # Agent 训练、Harness 研究
├── LLM底层技术       # 词向量/RoPE/Transformer 等底层原理
├── 宏观政策与经济    # 政策、经济、金融案例研究
├── 市场情报          # 媒体情报、市场与非标信息
├── 工程与源码        # 源码分析与工程实践
├── 教育升学          # 升学政策、高校研究
├── 历史与军事        # 历史/军事深入调查
└── 个人成长          # 健康、职业与个人效能研究
```

> 归档约定：凡新产出/外部文件，先按主题归入对应目录，避免顶层再积累散落文件。目录说明以 `AGENTS.md` 为准。

---

## 🔄 典型工作流

**公众号文章 → 短视频**

```
① wechat-high-energy-commentary     选题调研 → 公众号终稿（DOCX）
② article-to-short-video-script     同一文章 → 短视频执行稿（DOCX）
③ 归档 dayouyuan/articles/YYYYMMDD/
```

**地缘政治深度分析**

```
① geopolitical-deep-analysis-wechat  选题 → 多学派深度分析（DOCX）
② 归档 geopolitical/articles/YYYYMMDD/
```

**生活科普 → AI 动画**

```
① modern-qimin-jimeng-video          选题 → ≤10s 分镜稿（DOCX）
```

---

## ⚠️ 质量门：REQUIRED `render_docx.py`

**任何 `.docx` 交付物都必须渲染成逐页图片并人工核验**，才能交付：

```bash
cd .claude/skills/docx
python render_docx.py in.docx --output_dir out                # 逐页 PNG
python render_docx.py in.docx --output_dir out --emit_pdf     # 可选：追加 PDF
```

打开 `page-*.png` 逐页核对 **裁切 / 重叠 / 缺字符 / 表格碎裂**。文本抽取无法暴露出版面缺陷；渲染图片仅作 QA，非交付物。

> 依赖 `pdf2image` + Poppler + LibreOffice。渲染失败时先修复环境，**不可跳过验证**。若修复不了环境问题，把情况记入 `docs/gotchas.md`。

---

## 🧭 使用注意事项

- 目标语言：**中文**。请在同一会话按技能说明使用，交付物为 `.docx`/`.pdf`/`.pptx`/`.xlsx`。
- **内容红线**：
  - 不复用任何创作者的原句/口头禅，只从事实重建叙事机制；
  - 事实 / 推断 / 玩梗严格区分（事实直接陈述；推断加限定词；玩梗明显到不可误读），绝不把「前后发生」升格为「因果」；
  - 图片须「解释」而非「装饰」；AI 图片不得冒充真实新闻图；不嵌入无授权的新闻图片。
- **工程纪律**：
  - 不写「meme-first」正文，不用无证据的确定性语言（如「毫无疑问/必然/已彻底证明」）；
  - 避免 `git add -A`（历史上误提交过 137MB 技能 zip）；用精确 `git add <路径>`；
  - 不在 macOS `sleep`-poll（无 GNU `timeout`），用后台任务。

---

## 🔧 环境要求

| 依赖 | 用途 |
|------|------|
| Claude Code（最新版） | Skills / Artifacts 支持 |
| Python 3 + pdf2image + Poppler | `render_docx.py` 渲染 |
| LibreOffice | DOCX→PDF 渲染对比（可选但推荐） |
| Node.js | `slides` 技能（pptxgenjs） |

---

## 📎 附注

- 踩过的坑统一沉淀于 [`docs/gotchas.md`](docs/gotchas.md)，遇到新问题可随时追加。
- GitHub 大流量推送走 `ssh.github.com:443`（见 `docs/gotchas.md`）。
- 本仓库版权归作者所有；内容交付物仅供个人工作区使用。