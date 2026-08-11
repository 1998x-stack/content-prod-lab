---
name: geopolitical-deep-analysis-wechat
description: Researches and writes professional, source-grounded geopolitical analysis on countries, regions, conflicts, alliances, sanctions, trade, energy, technology, security, and strategic competition. Use when a user provides a geopolitical topic and wants current web research, multi-school analysis, scenario forecasting, strategic judgment, or a polished Chinese WeChat public-account article delivered as a visually verified Word .docx.
metadata:
  version: "1.0.0"
  language: "zh-CN"
---

# Geopolitical Deep Analysis for WeChat

## Mission

Turn a geopolitical topic into a rigorous, current, multi-framework analysis and a publication-ready Chinese WeChat long-form article in `.docx` format.

The skill must do more than summarize news. It must explain:

1. why the issue exists,
2. what structural forces constrain the actors,
3. what each actor wants and fears,
4. how relative capabilities and dependencies are changing,
5. what second-order reactions follow,
6. which future scenarios are plausible,
7. which observable indicators could confirm or falsify the analysis.

The final deliverable is a polished Word document suitable for a WeChat public-account editorial workflow.

## Trigger and inputs

Use this skill when the user supplies a topic such as:

- a country or bilateral relationship,
- a war, crisis, territorial dispute, alliance, sanctions regime, or diplomatic shift,
- a strategic region, sea lane, chokepoint, energy corridor, or technology supply chain,
- a question about power transition, deterrence, strategic competition, or geopolitical risk,
- a request for a deep geopolitical WeChat article or Word report.

Minimum required input: **topic**.

Optional inputs:

- time horizon,
- geographic scope,
- focal question,
- preferred depth,
- desired audience,
- whether scenario probabilities are wanted.

If optional inputs are missing, default to the latest available information, deep analysis, a general educated Chinese readership, scenario analysis, and cited sources.

## Non-negotiable workflow

### 1. Define the real strategic question

Convert a broad topic into a researchable question. Identify:

- time range,
- geographic scope,
- principal actors,
- disputed facts,
- decision or strategic outcome being explained.

Do not reduce the question to a news recap.

### 2. Research before judging

For current or time-sensitive topics, browse the web before analysis.

Follow the source hierarchy and verification rules in `references/research-and-sourcing.md`.

Build a fact base that separates:

- confirmed facts,
- contested claims,
- actor narratives,
- analytical inference.

### 3. Build the strategic map

Start with geography where relevant:

- location and distance,
- sea/land access,
- mountains, plains, rivers, islands,
- ports, bases, chokepoints, corridors,
- resource and population geography,
- strategic depth and logistics.

Ask: **What flows through this space, and who can interrupt, defend, or exploit those flows?**

### 4. Analyze actors and relative power

For each major actor, identify:

- maximum objective,
- minimum acceptable outcome,
- red lines,
- core fears,
- time preference,
- military/economic/industrial/demographic capabilities,
- alliance support,
- state capacity to convert resources into usable power.

Never equate raw resources with usable strategic power.

### 5. Apply multiple theories

Use only frameworks that add explanatory value. Usually test at least three of:

- Mahan / maritime power,
- Mackinder / continental geography,
- Spykman / rimland,
- defensive realism,
- offensive realism,
- neoclassical realism,
- security dilemma,
- liberal institutionalism / interdependence,
- geoeconomics / weaponized interdependence,
- constructivism / identity and norms,
- critical geopolitics / discourse,
- domestic politics and regime interests.

For each framework used, state:

- what it explains,
- what it misses,
- whether its assumptions fit the current case.

See `references/methodology.md`.

### 6. Model action and reaction

Do not stop at first-order effects.

Trace:

`Actor A action -> Actor B reaction -> alliance/market/domestic response -> second-order effect -> long-run strategic consequence`

Look specifically for:

- security dilemmas,
- balancing,
- abandonment/entrapment fears,
- escalation spirals,
- deterrence failure,
- strategic blowback,
- self-fulfilling threat perceptions.

### 7. Separate tactical from strategic outcomes

Explicitly distinguish:

- tactical,
- operational,
- strategic,
- grand-strategic outcomes.

Do not infer strategic success from a battlefield gain, territory gain, sanction announcement, summit, or headline alone.

Use a net-strategic-benefit lens:

`direct gains - direct costs - adversary response - alliance response - long-term opportunity cost`

### 8. Analyze time and scenarios

Evaluate short, medium, and long horizons where useful.

Generate at least three scenarios for open-ended questions:

- baseline,
- escalation/adverse shift,
- de-escalation or alternative structural shift.

For every scenario provide:

- trigger conditions,
- mechanism,
- winners/losers,
- early-warning indicators.

Only provide probability ranges when useful, label them as analytical estimates, and avoid false precision.

### 9. Red-team the conclusion

Before writing the final article, challenge the main thesis with at least three strong counterarguments.

Ask:

- What assumption could be wrong?
- What actor may be misread?
- What capability is hard to measure?
- What new technology or political change could invalidate the model?
- What evidence would falsify the conclusion?

Revise the analysis if the counterarguments are stronger.

### 10. Write for WeChat, not for an academic journal

Transform the research into a readable Chinese public-account article according to `references/wechat-editorial.md` and `templates/analysis-outline.md`.

The article must retain research rigor while using:

- a strong but non-clickbait title,
- a short subtitle,
- a 150-300 Chinese-character lead,
- a compact “核心判断” section,
- clear numbered sections,
- short mobile-friendly paragraphs,
- selective bold emphasis,
- narrow tables only when they improve comprehension,
- maps/figures only when they answer a real analytical question,
- a final section with key monitoring indicators and falsifiable judgments.

### 11. Generate the final Word document

The final formal deliverable must be `.docx` unless the user explicitly requests another format.

Follow `references/docx-production.md`.

When a DOCX creation capability is available, use it. The document must be rendered to page images and visually inspected before delivery. Fix layout defects and re-render until clean.

Return the final `.docx`; do not clutter delivery with internal QA files unless requested.

## Required analytical output contract

The final article should normally contain:

1. Title and subtitle
2. Lead
3. 核心判断
4. What the issue really is
5. Geographic/strategic value
6. Actor objectives and fears
7. Relative power and constraints
8. Multi-theory interpretation
9. Action-reaction chain and second-order effects
10. Who currently has which type of advantage
11. Structural weaknesses of each side
12. Future scenarios
13. Key monitoring indicators: signal vs noise
14. Red-team / strongest counterargument
15. Final judgment
16. References / further reading

Adapt structure to the topic rather than forcing empty sections.

## Analytical writing rules

- Explanation is not justification.
- Distinguish legal judgment, moral judgment, actor motivation, and strategic effectiveness.
- Do not treat official claims as verified facts.
- Do not use geography as destiny.
- Prefer relative, deployable power over headline aggregates.
- Separate intent, capability, perceived capability, and adversary perception.
- Use exact dates when chronology matters.
- State uncertainty plainly.
- Avoid partisan advocacy, propaganda vocabulary, civilizational essentialism, and deterministic claims.
- Avoid sensational words such as “必然”, “彻底崩溃”, “大结局”, or “稳赢” unless the evidence genuinely warrants them.
- Keep military discussion at strategic/analytical level; do not provide operational instructions for harming people, targeting, weapons use, sabotage, or evasion.

## Quality gate

Before delivery confirm:

### Evidence
- Current claims are sourced.
- Load-bearing facts use high-quality sources.
- Disputed claims are labeled.
- Facts, narratives, and inference are separated.

### Analysis
- Geography is considered where relevant.
- Relative power and conversion capacity are considered.
- At least three useful analytical lenses are tested on complex issues.
- Second-order effects are examined.
- Tactical and strategic outcomes are distinguished.
- Key uncertainties and falsifiers are stated.

### Editorial
- The article has a clear thesis.
- Sections progress logically rather than repeating news.
- Paragraphs are mobile-readable.
- Tables fit narrow screens.
- Every visual has analytical purpose.
- Headline is attractive but not sensational.

### DOCX
- Heading styles are semantic, not just manually bolded.
- Fonts render correctly in Chinese.
- Tables and images fit the page.
- Captions and sources are present.
- Hyperlinks are usable.
- Every rendered page has been visually checked.
- No clipping, overlap, broken glyphs, orphaned headings, or stray tool citation tokens remain.

## Progressive references

Load these only when needed:

- `references/research-and-sourcing.md` — web research, source hierarchy, citations, contested facts
- `references/methodology.md` — analytical frameworks and model-selection rules
- `references/wechat-editorial.md` — Chinese WeChat article structure and writing style
- `references/docx-production.md` — Word layout, render/verify workflow, final packaging
- `references/safety-and-neutrality.md` — strategic-level safety and neutral framing
- `templates/analysis-outline.md` — reusable article scaffold
