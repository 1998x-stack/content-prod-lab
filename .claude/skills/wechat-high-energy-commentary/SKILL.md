---
name: wechat-high-energy-commentary
description: >-
  Research, select, write, illustrate, and deliver high-energy Chinese current-affairs commentary for WeChat Official Accounts. Use when the user gives a news/topic to investigate and turn into a WeChat article, asks “今天写什么/最近有什么选题/帮我找几个新闻”, selects a previously proposed topic by number, asks to “继续挖/梗多一点/严肃一点”, or wants a researched article with images and a final editable Word document. This skill emphasizes factual verification, narrative reframing, knowledge-rich explanations, mobile-first WeChat formatting, image sourcing, and DOCX delivery.
---

# WeChat High-Energy Commentary

## Goal

Produce a **publish-ready WeChat Official Account article**, not a generic news summary:

**research → find the core anomaly → reframe the story → explain mechanisms/history → write for mobile reading → add useful images → deliver an editable DOCX.**

The desired reading experience is: **the reader follows an unusually entertaining story while genuinely understanding the rules, history, technology, ecology, business incentives, or international relations behind it.**

Use high-level mechanisms such as strong openings, high information density, colloquial internet language, defamiliarization, personification, extended metaphors, register mismatch, long-sentence momentum, short punch lines, questions, reversals, and knowledge-as-entertainment.

Do **not** copy distinctive wording, signature phrases, famous passages, or fixed verbal tics from any specific writer or creator. Recreate the narrative mechanics from the facts of the current story.

## Load supporting guidance

Read the smallest relevant reference before execution:

- Topic discovery or breaking/current news → `references/research_and_fact_check.md`
- Writing/rewrite → `references/style_guide.md`
- WeChat layout, images, captions, copyright → `references/wechat_and_images.md`
- Final deliverable / Word → `references/output_contract.md`

If producing a `.docx`, also follow the platform's DOCX artifact skill/instructions. **Render the DOCX to page images and visually inspect every page before delivery.**

## Mode routing

### Mode A — User gives a topic

Examples: “写一下这个政策”, “写某公司最近这个事故”, “写一下X为什么会这样”.

1. Do **not** draft from memory when the topic is current or externally verifiable.
2. Search and verify the event first.
3. Build a timeline and evidence map.
4. Identify the single strongest anomaly/contradiction.
5. Choose the best narrative coordinate system.
6. Write the WeChat article.
7. Source or create useful images and place them near the relevant passages.
8. Produce the final DOCX unless the user explicitly says “不要Word”.

Default: do not ask follow-up questions when the topic is sufficiently clear. Make reasonable editorial decisions and proceed.

### Mode B — User does not know what to write

Triggers include: “今天写什么？”, “最近有什么选题？”, “帮我找几个新闻”, “不知道写啥”.

Search the **most recent 72 hours** of domestic and international news. Favor stories with:

- absurdity or strong contradiction
- policy/legal loopholes
- international reversals or unusual diplomacy
- technology accidents/failures
- corporate conflict
- social controversies with a non-obvious mechanism
- animal/ecology oddities
- dormant historical disputes returning
- ordinary-looking headlines with a surprisingly deep system underneath
- high discussion but low public understanding

Do not rank only by trending status. Rank by **story potential + explanatory depth + verifiability + WeChat shareability + image potential**.

Return **3–5 candidates only**. Do not write the full article yet.

For each candidate use:

- `选题 X｜标题`
- `发生了什么：` 2–4 sentences
- `最值得写的点：` the core contradiction/anomaly
- `推荐视角：` one strong narrative frame
- `可以深挖：` 3–5 directions
- `配图潜力：` likely useful visuals
- `文章潜力：★★★★★`

When the user replies with a number, immediately continue with research → article → images → DOCX. Do not ask them to restate the choice.

## Non-negotiable research gate

Before writing a current-affairs article:

1. Verify time, location, actors, sequence, current status, and latest development.
2. Prefer primary sources: government, court, police, company filings/releases, direct statements, official statistics, laws, papers, international organizations.
3. Cross-check load-bearing numbers and claims with at least two reliable sources when practical.
4. Investigate why now, why here, which rule changed, winners/losers, technical limits, geography, demography, economics, historical residue, and overlooked numbers.
5. Search opposing/alternative explanations for controversial claims.

Never treat “A happened before B” as proof that A caused B.

## Find the article engine

Before drafting, answer internally in one sentence:

> **这件事最离谱/最反常/最违反直觉的地方是什么？**

Do not settle for the event category (“发生了移民潮”). Find the contradiction (“拼命进去后发现规则/目的地与想象不同，又大量返回”).

Then select a more explanatory observation frame. Possibilities include:

- game/system rules: player, map, version, patch, loophole
- animal/ecology: humans as one species among others
- historical recurrence: similar mechanism, while stating differences
- corporate power struggle: board, capital, product, incentives
- legal/institutional mechanics
- geographic constraint
- formal political/military/academic language applied to an everyday/absurd object

Use only **1–2 main metaphor systems** per article. Do not force game language into every story.

## Drafting contract

Default article length: **2500–4500 Chinese characters**.

- “短一点” → about 1500–2500
- “长文” → about 4000–7000

Default information/entertainment balance: roughly **70/30**.

Structure flexibly around this rhythm:

1. **炸场** — first 100–200 Chinese characters: contradiction, scene, question, result-first reveal, or historical mismatch. Never begin with “近日，据媒体报道……”.
2. **放大荒诞** — concrete details, questions, defamiliarization, personification.
3. **回到时间线** — explain how the situation developed.
4. **第一次解释** — accurate professional fact + plain-language translation.
5. **第一次升级** — key actors, numbers, incentives, loopholes, history.
6. **知识扩展** — 3–8 genuine knowledge gains; use analogies to reduce complexity.
7. **第二次升级** — move from the episode to systems, incentives, technology, social psychology, historical patterns, or geopolitics.
8. **重新定义事件** — show why the real problem differs from the opening impression.
9. **结尾** — return to the opening, create a historical loop, or end on one memorable redefinition. Avoid generic “理性看待/总而言之”.

## Fact language levels

Keep three levels distinct:

- **Verified fact** → state directly.
- **Reasonable inference** → use qualifiers such as “可能 / 更可能是 / 一个重要原因是 / 从目前资料看”.
- **Rhetorical joke / hypothetical** → make it unmistakably figurative, satirical, or speculative.

**Accuracy > drama > jokes.**

For mass death, war casualties, severe disaster, minors harmed, or severe violence: automatically reduce jokes. Never make victims or suffering the punch line; critique decisions, institutions, technical failures, bureaucracy, or responsible actors instead.

## Images

Images must explain, not decorate. For a typical 2500–4500-character article, use roughly **4–8 images when useful**, not by quota.

Prioritize:

- cover candidate
- news scene
- key people/subjects/products/buildings
- maps for geopolitics, borders, routes, enclaves, war, transport
- historical photos/maps/documents
- simple data charts when one number/trend is crucial

Prefer official/public-domain/clearly licensed material, institutional materials, Wikimedia Commons where appropriate, corporate press materials, or self-generated charts/maps/diagrams. Keep source metadata. Never present an AI-generated image as real news photography; label it as a concept/illustration when used.

## WeChat mobile layout

- Short paragraphs, usually 1–4 mobile lines.
- Use whitespace around turns and punch lines.
- Bold only high-value conclusions, numbers, reversals, or key judgments.
- Use 3–6 natural subheads at most; avoid textbook headings like “一、事件背景”.
- Place each image beside the section it explains.
- Captions should explain **why the image matters**, not merely name what it shows.

## Delivery behavior

Unless the user asks otherwise, a selected/explicit topic should result in:

1. 3–5 title candidates
2. one recommended main title
3. optional WeChat summary/subtitle if useful
4. full WeChat article
5. inline images + captions + image sources
6. 5–10 key factual sources
7. editable `.docx`

If the user says `只给选题`, stop after candidate topics.
If the user says `不要Word`, deliver the complete article without generating DOCX.
If the user says `继续挖`, deepen research, seek contrary evidence/history/data, then revise rather than merely lengthen.
If the user says `梗多一点`, increase defamiliarization, personification, analogy, and register mismatch without weakening factual discipline.
If the user says `严肃一点`, reduce internet slang while preserving narrative momentum and explanatory depth.

## Final quality gate

Before delivery, verify:

- Can the anomaly be stated in one sentence?
- Is there a better observation frame than the original news framing?
- Does the article teach at least 3 meaningful things?
- Are mechanism/history/incentives explained?
- Are causality and inference labeled correctly?
- Would the article still stand if all jokes were removed?
- Is the mobile reading rhythm clean?
- Does each image serve a clear information function?
- Are title and cover ideas compelling without inventing facts?
- Is the final DOCX visually clean after render-and-inspect QA?

If the article collapses after removing the jokes, research more before publishing.

## Core principle

Never **invent jokes first and stuff news into them**.

Use this order:

**investigate facts → find the anomaly → reframe the observation point → explain the mechanism → add knowledge → design mobile rhythm → configure images → let humor grow naturally from the facts → verify → deliver DOCX.**
