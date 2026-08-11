---
name: article-to-short-video-script
description: >-
  Convert a completed Chinese WeChat/current-affairs article into a high-retention short-video production script for WeChat Channels, Douyin, or Bilibili, with a rewritten voiceover, hook, pacing, subtitles, visual plan, shot list, asset checklist, source notes, and final editable Word document. Use when the user says “把这篇转视频/做成视频稿/生成短视频版/视频号版/抖音版/B站版/只要口播”, provides a finished article to transform, or refers to the current conversation's completed article and wants a video version. This skill performs video re-engineering rather than summarization.
---

# Article → Short Video Script

## Goal

Turn an already completed Chinese news/current-affairs article into a **production-ready short-video package**, not a shortened article.

Core workflow:

**read source article → extract the video thesis → rebuild information order → design first 3 seconds → create retention beats → rewrite voiceover → design visual/subtitle logic → build shot list → source/plan assets → adapt platform → deliver editable DOCX.**

The target experience is:

> The viewer is hooked immediately, keeps receiving meaningful new information or reversals, and finishes the video with a clearer understanding of the event's mechanism, history, incentives, technology, ecology, business logic, or geopolitics.

Use high-level mechanisms such as strong hooks, information density, conversational narration, defamiliarization, personification, historical/game/business analogies, register mismatch, long-short sentence rhythm, information gaps, small reversals, and knowledge-as-entertainment.

Do **not** copy distinctive signature phrases, famous passages, recurring verbal tics, or highly recognizable wording from any specific creator. Recreate the narrative mechanics from the source facts.

## Load supporting guidance

Read only the references needed for the task:

- Video rewrite structure, hooks, pacing, narration → `references/video_rewrite_style.md`
- Platform length and adaptation → `references/platform_adaptation.md`
- Visuals, subtitles, shot design, asset sourcing → `references/visual_and_assets.md`
- Fact preservation and research behavior → `references/fact_integrity.md`
- Final production package and DOCX → `references/output_contract.md`

If generating a `.docx`, follow the platform DOCX artifact instructions. Render the document to page images and visually inspect the final document before delivery.

## Input routing

### A — User provides a finished article

Read the article and start conversion directly.

### B — A finished article exists earlier in the current conversation

When the user says “转视频 / 做成视频稿 / 视频号版 / 短视频版”, reuse the existing article. Do not ask the user to paste it again.

### C — User provides a document/file

Read the relevant article from the provided file. Preserve its factual backbone and source context.

### D — User asks for platform variant

- `视频号版` → default balanced deep-news short video.
- `抖音版` → faster entry, tighter setup, stronger early payoff, more compression.
- `B站版` → allow more history, mechanism, context, and slower explanatory sections.
- `只要口播` → output only title options + clean voiceover unless the user asks for additional sections.

Do not ask follow-up questions if the source article and requested platform are sufficiently clear. Make editorial decisions and proceed.

## The conversion rule

A WeChat article depends on voluntary reading. A short video depends on **continued attention**.

Therefore never convert by preserving paragraph order and merely deleting material.

Instead, rebuild the story around this question:

> **Why would the viewer listen to the next sentence?**

Information may be reordered. Results may appear before causes. The strongest contrast, number, visual, contradiction, or outcome may move to the beginning.

## Extract the video engine

Before drafting, determine internally:

1. What happened?
2. What is the strongest anomaly or contradiction?
3. What is the biggest reversal?
4. Which 3 facts are indispensable?
5. Which mechanism must be explained?
6. Which one or two numbers create scale?
7. Is there a useful historical parallel?
8. What can be visualized especially well?
9. What should the viewer understand by the end?
10. What can be cut without damaging that understanding?

Then write one internal thesis:

> **这条视频真正要讲的是：______。**

Remove or compress branches that do not serve this thesis.

## Default duration routing

Use content complexity, not a fixed platform stereotype.

- **60–120 seconds** / ~350–700 Chinese characters: single event, oddity, simple policy, one main reversal.
- **2–4 minutes** / ~700–1400 Chinese characters: default; event + mechanism + context + one expansion.
- **4–7 minutes** / ~1400–2500 Chinese characters: geopolitics, technology, business, history, or complex social systems.

Do not force a complicated story into a few dozen seconds when accuracy and comprehension would collapse.

## Non-negotiable opening gate

The first sentence must already contain meaningful information.

Never open with:

- “大家好，今天我们来聊一下……”
- “近日，据媒体报道……”
- “最近发生了一件……”
- “这件事非常离谱，到底有多离谱呢……”
- empty calls to like/follow

Prefer one of:

- result-first contradiction
- abnormal behavior
- one shocking number
- historical mismatch
- a specific question whose answer matters

The first **3 seconds** should expose the conflict or anomaly.

The first **15 seconds** should usually do three things:

1. show the anomaly;
2. create an information gap (“为什么？”);
3. promise a worthwhile explanation without revealing every answer at once.

## Retention structure

Use this as a flexible rhythm, not visible section labels:

**hook → amplify anomaly → essential timeline → first explanation → first reversal → knowledge expansion → second escalation → real answer → callback.**

Every ~15–30 seconds, introduce at least one meaningful new stimulus:

- fact
- actor
- number
- question
- visual
- historical context
- mechanism
- analogy
- small reversal

If a section explains one abstraction for too long without new value, compress or visualize it.

## Voiceover contract

The narration must sound natural when spoken aloud.

- Prefer conversational syntax and clean pauses.
- Avoid nested parentheses and long written-language constructions.
- Avoid stacking multiple technical terms in one sentence.
- Avoid reading too many exact numbers aloud.
- Use long sentences to move information, then short lines as rhythmic punches.
- Use questions and turns to propel explanation, not as filler.

A joke or metaphor must do at least one useful job: explain, compress, pace, visualize, or improve memory.

Default information/entertainment balance: **about 70/30**.

## Main metaphor

Choose the observation system that best fits the story: game/rules, history, business, ecology, legal mechanics, geography, logistics, etc.

Use at most **1 main metaphor + 1 auxiliary metaphor**. Do not mechanically call everything a BUG, server, version, or battlefield.

## Fact preservation

The video may compress language but must not compress away uncertainty.

Use three levels:

- **Verified fact** → state directly.
- **Reasonable analysis/inference** → retain qualifiers such as “可能 / 一个重要原因 / 从目前资料看”.
- **Rhetorical analogy/hypothetical** → make it clearly figurative or speculative.

Never change “A happened before B” into “A caused B” simply to make the script cleaner.

If the source article contains a claim that appears materially outdated, disputed, unsupported, or internally inconsistent, verify it before preserving it in the final narration.

For severe violence, war casualties, disaster, mass death, or minors harmed: reduce jokes automatically. Do not make victims or suffering the punch line.

## Visual-first writing

Write narration and visuals together.

The video should not be a talking article laid over random screenshots. Each information unit should have a corresponding visual function.

Think in visual units of roughly **3–8 seconds** when useful, using:

- scene footage / news photos
- people
- maps
- documents / announcements / judgments
- products / factories / companies
- historical images and maps
- data cards / simple charts
- timelines
- motion/zoom/crop on stills
- clearly labeled diagrams or AI concept visuals when real imagery is inappropriate/unavailable

Do not make narration and on-screen text redundantly say the same thing when the screen can carry precise dates, names, or numbers.

## Completion / retention QA

Before finalizing, check:

- **3s:** Does the opening independently create interest with real information?
- **10–15s:** Is there a concrete information gap and a reason to continue?
- **30s:** Has the viewer already learned something valuable?
- **Middle:** Is there any >30s stretch with no new meaningful information, visual, or reversal?
- **Later:** Is there a second escalation beyond the surface event?
- **End:** Does the viewer leave with a deeper answer than the opening premise?

If not, reorder before polishing language.

## Default final deliverable

Unless the user narrows the request, produce a complete **《短视频制作脚本》** containing:

1. 5 video title candidates + 1 recommended title
2. 3 cover-text candidates (normally 6–14 Chinese characters)
3. video metadata: estimated duration, estimated character count, content type, core hook, core conclusion, main metaphor
4. clean full voiceover script, ready for teleprompter/TTS/voice actor
5. timed shot-by-shot execution script
6. key subtitle/highlight list
7. prioritized asset checklist
8. asset/source notes
9. concise WeChat Channels publishing copy
10. editable `.docx`

If user says `只要口播`, stop after the clean voiceover (plus minimal title options if useful) and do not force a full production package.

## Core principle

A WeChat article answers:

> **“这件事到底是怎么回事？”**

A short video must additionally answer:

> **“为什么观众下一秒还愿意继续听？”**

So the correct transformation is not “delete paragraphs”. It is:

**find strongest hook → move contrast earlier → create information gaps → deliver value quickly → keep adding meaningful information → let visuals carry part of the explanation → preserve essential knowledge → complete a final cognitive upgrade → deliver a production-ready DOCX.**
