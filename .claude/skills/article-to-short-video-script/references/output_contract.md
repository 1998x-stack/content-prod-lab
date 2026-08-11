# Output and DOCX Contract

## Default package title

Use:

# 《短视频制作脚本》

## 1. Video titles

Provide 5 candidates and mark 1 recommended title.

Good titles are:
- short
- anomalous
- visual
- conflict-driven
- factually supportable

Do not invent facts for click-through.

## 2. Cover text

Provide 3 candidates, normally **6–14 Chinese characters**.

The cover line should be readable as large mobile-screen text and should not require a subtitle to make sense.

## 3. Video metadata

Include:
- 预计时长
- 预计字数
- 内容类型
- 核心钩子
- 核心结论
- 主隐喻
- 目标平台/version when relevant

## 4. Clean full voiceover

Provide one uninterrupted script ready for:
- teleprompter
- TTS
- human voice actor

Do not mix in:
- camera instructions
- URLs
- source citations
- bracketed editor notes

The voiceover should be independently readable from beginning to end.

## 5. Timed execution / shot script

Use time blocks based on actual information beats, not a fixed 5-second grid.

Example structure:

### 00:00–00:07
**口播：** …

**画面：** …

**屏幕大字：** …

**素材建议：** …

Continue through the whole video.

## 6. Key subtitle list

Extract the strongest:
- numbers
- names
- dates
- places
- rules
- reversals
- key concepts
- quotable conclusions

## 7. Asset checklist

Organize into:
- 必须素材
- 推荐素材
- 可替代素材

## 8. Source notes

Record core asset/fact sources. Prefer:

**official / primary → authoritative media → professional institutions → licensed/public-domain repositories.**

## 9. Publishing copy

Default: create a concise **WeChat Channels** publishing caption. Do not retell the entire video.

If user explicitly requests Douyin/Bilibili copy, adapt accordingly.

## DOCX deliverable

Unless user says `不要Word` or `只要口播`, create an editable `.docx` in this order:

1. 推荐视频标题
2. 封面文案
3. 视频基本信息
4. 完整纯口播稿
5. 分镜执行稿
6. 重点字幕
7. 素材清单
8. 素材来源
9. 视频号发布文案

### DOCX design

The document is a **production execution file**, not a decorative magazine layout.

Prioritize:
- clean hierarchy
- easy scanning
- visible timecodes
- clear distinction between `口播 / 画面 / 字幕 / 素材`
- short shot blocks
- editable text
- stable image placement if reference frames/assets are embedded

Avoid complicated Word effects.

### QA before delivery

Follow the DOCX artifact skill/instructions.

Render the DOCX to images/PDF and visually inspect every page for:
- clipped text
- broken tables
- oversized images
- bad page breaks
- unreadable timecodes
- inconsistent hierarchy
- missing source labels

Fix issues before sharing the final file.

## User overrides

- `只要口播` → do not generate full shot list or DOCX unless explicitly requested.
- `不要Word` → provide the requested script sections in chat only.
- `短一点` → 60–120 sec target.
- `正常视频` → 2–4 min target.
- `深度视频` → 4–7 min target.
