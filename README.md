# Content Production Lab

一个以 **Claude Code 技能（Skills）** 为引擎、以 **中文内容生产** 为交付物的个人工作区。把"写公众号文章 → 转成短视频脚本 → 生成专业 DOCX/幻灯片/PDF"的完整生产线，沉淀成可复用的技能包与工作目录。

> 生产内容以中文为主，所有文档类交付均以 `.docx` 产出，并通过 **渲染验证门**（见下）保证版面无误。

---

## ✨ 核心能力

| 技能 | 用途 |
|------|------|
| **wechat-high-energy-commentary** | 调研选题 → 撰写高能中文时评公众号文章（配图、DOCX 交付） |
| **geopolitical-deep-analysis-wechat** | 地缘政治深度分析 → 多学派解读、情景推演、战略判断的中文公众号文章（DOCX 交付） |
| **article-to-short-video-script** | 把已完成文章重新工程化 → 可直接配音/剪辑的短视频制作执行稿（微信视频号/抖音/B站），非摘要式搬运 |
| **modern-qimin-jimeng-video** | 现代齐民要术六卷（衣食住行养用）生活科普选题 → 即梦 AI 动画短视频生产稿（≤10s 分镜、桥接帧、干净帧向量） |
| **docx** | `.docx` 全流程：创建/编辑/批注/修订/合并/水印/审校，含 `render_docx.py` 渲染→视觉验收 |
| **pdfs** | PDF 处理：读取/OCR/表单/涂改/转换/差异对比，`render → verify → operate → re-verify` |
| **slides** | 基于 pptxgenjs 的 PPT 制作与导出（含十多个专业模板） |
| **spreadsheets** | 表格处理（读取/清洗/转换/生成） |

> 以上技能均在 `.claude/skills/<name>/`，每个 SKILL.md 有 `name` + `description` 的元数据，可被 Claude 自动触发或 `/name` 调用。使用 `SKILL.md` → 按需加载 `references/*.md` 的方式渐进展开。

---

## 📂 目录结构

```
articles/
├── CLAUDE.md              # Claude Code 工作指南（本项目）
├── docs/
│   └── gotchas.md         # 踩坑记录（可追加）
├── .claude/
│   ├── skills/            # 8 个可被 Claude 触发的技能包
│   └── settings.local.json # 权限 allow-list
├── dayouyuan/             # 公众号文章 → 短视频 流水线工作区
│   ├── prompts/           # 写作/视频重构 Prompt 模板
│   └── articles/YYYYMMDD/ # 按生产日期归档的交付 DOCX
├── qiminyaosu/            # 现代齐民要术 工作区
│   ├── 现代齐民要术_六卷选题圣经/  # 六卷选题库（DOCX + XLSX）
│   ├── prompts/  articles/  docs/
├── geopolitical/          # 地缘政治深度分析 工作区（prompts/ + articles/YYYYMMDD/）
└── docs/
    └── gotchas.md         # 踩坑记录（可追加）
```

---

## 🚀 用法

### 调用技能

```bash
# 在 Claude Code 中直接触发（agent 自动匹配 description 或手动指定）
/wechat-high-energy-commentary
/geopolitical-deep-analysis-wechat
/article-to-short-video-script
/modern-qimin-jimeng-video
/docx
/pdfs
/slides
/spreadsheets
```

### DOCX 交付验证门框（讲卫生，不省）

任何产出 `.docx` 的任务都必须**渲染成逐页图片并人工检查**才能交付：

```bash
cd .claude/skills/docx
python render_docx.py in.docx --output_dir out        # 逐页 PNG
python render_docx.py in.docx --output_dir out --emit_pdf   # + PDF
```

打开 `page-*.png` 逐页核对裁切/重叠/缺字符/表格碎裂。文本抽取无法暴露出版面缺陷。渲染图片仅用于 QA，非交付物。

> 需要 `pdf2image` + Poppler + LibreOffice。若渲染失败，先修 LibreOffice 环境，别跳过验证。

### 工作区使用流程（dayouyuan 流水线）

1. `wechat-high-energy-commentary` → 出公众号终稿（DOCX，过验证门）
2. `article-to-short-video-script` → 同一文章 → 短视频执行稿（DOCX）
3. 归档至 `articles/YYYYMMDD/`

---

## ⚙️ 环境要求

- Claude Code（最新版，含 Skills、Artifacts 支持）
- Python 3 + `pdf2image`（`render_docx.py`），以及 LibreOffice（可选，用于 PDF 渲染对比）
- Node.js（`slides` 技能用 pptxgenjs）
- GitHub 推送右侧的环境：大流量走 `ssh.github.com:443`（见 `docs/gotchas.md`）

---

## 📌 备注

- **内容红线**：不复用任何创作者的原句/口头禅，仅重建叙事机制；事实/推断/玩梗严格区分。
- 生产产物为中文，请在 Claude 会话中按技能说明使用。
- 踩过的坑统一沉淀在 [`docs/gotchas.md`](docs/gotchas.md)，遇到新问题随时追加。