#!/usr/bin/env python3
"""
Recheck each production docx against:
  (A) modern-qimin-jimeng-video skill hard rules
  (B) TODO.json acceptance criteria

Reports per-story pass/fail. Flags generic-placeholder values that violate
skill rule #3 (every shot must have a concrete START_FRAME / END_FRAME describing
subject/position/material/light/composition — not boilerplate).
"""
import os, re, json, glob
from docx import Document

ROOT = "/Users/x/Desktop/content-prod-lab/qiminyaosu/ralph/生产稿_600"
VOL_DIR = {"衣事":"01_衣事","食事":"02_食事","居事":"03_居事","养身":"04_养身","行事":"05_行事","器用":"06_器用"}
VOL_ORDER = {"衣事":1,"食事":2,"居事":3,"养身":4,"行事":5,"器用":6}

# ---------- generic-placeholder detectors (skill rule #3, #4) ----------
GENERIC_START = ["稳定停在可桥接的尾帧", "主体与开头相扣", "（后期", "作下一镜首帧。"]
GENERIC_END = ["稳定停在可桥接的尾帧", "作下一镜首帧", "（后期"]

def _check_doc(fpath):
    doc = Document(fpath)
    paras = [p.text.strip() for p in doc.paragraphs]
    text_joined = "\n".join(paras)
    shots = []
    # collect Shot Cards
    shot_cards = []
    card_start = False
    cur = {}
    for t in paras:
        if t in ("S01","S02","S03","S04","S05","S06"):
            if cur: shot_cards.append(cur)
            cur = {"id": t}; continue
        if t.startswith("一、"):  # stop at sections after shots typically 七、 is before
            continue
        for k in ("DURATION","FUNCTION","START_FRAME","END_FRAME","BRIDGE_TO_NEXT","JIMENG_PROMPT","VISUAL_EVENT","CAMERA"):
            if t.startswith(k+"：") or t.startswith(k+"（"):
                # cv
                val = t.split("：",1)[1] if "：" in t else t.split("（",1)[-1]
                cur[k] = val
        if t.startswith("9:16"):
            cur["JIMENG_PROMPT"] = t
    if cur: shot_cards.append(cur)

    issues = []
    # A1: each shot ≤10s
    if shot_cards:
        for card in shot_cards:
            dur = card.get("DURATION","")
            m = re.search(r"(\d+)", dur)
            if m and int(m.group(1)) > 10:
                issues.append("shot '%s' DURATION>10s"%card.get("id","?"))
    else:
        issues.append("没有识别到 Shot Card")

    # A2: START/END/BRIDGE present & not generic
    for card in shot_cards:
        st = card.get("START_FRAME","")
        en = card.get("END_FRAME","")
        br = card.get("BRIDGE_TO_NEXT","")
        if not st: issues.append("缺 START_FRAME")
        elif any(g in st for g in GENERIC_START): issues.append("START_FRAME 为通用占位")
        if not en: issues.append("缺 END_FRAME")
        elif any(g in en for g in GENERIC_END): issues.append("END_FRAME 为通用占位")
        if not br: issues.append("缺 BRIDGE_TO_NEXT")
        # JIMENG prompt non-empty
        if not card.get("JIMENG_PROMPT","").startswith("9:16"):
            issues.append("缺 JIMENG_PROMPT")

    # A3: AI text / pseudo-evidence negative constraints present in prompt
    prompt_txt = " ".join(c.get("JIMENG_PROMPT","") for c in shot_cards)
    if "无中文文字" not in prompt_txt and "无文字" not in prompt_txt:
        issues.append("Prompt 缺'不生成文字'约束")
    if "伪" not in prompt_txt:
        issues.append("Prompt 缺'伪实验/伪专家'约束")

    # A4: unified 片尾
    if "现代齐民要术｜衣食住行，件件有方法。" not in text_joined:
        issues.append("片尾缺失/不一致")

    # A5: every shot ≤10s in 分镜总表 durations
    # (table check optional; covered by cards)

    # A6: health safety (养身 must not diagnose)
    if "04_养身" in fpath:
        if "一般健康教育" not in text_joined and "不做一个" not in text_joined:
            issues.append("养身缺健康边界注记")

    return shot_cards, issues


def recheck_all():
    files = sorted(glob.glob(os.path.join(ROOT,"*","*_生产稿.docx")))
    # map file base id (tid) 
    results = {}
    for f in files:
        base = os.path.basename(f)
        # regex extract id like 01-01-1 from 现代齐民要术_01-01-1_标题_生产稿.docx
        m = re.match(r"现代齐民要术_(\d{2}-\d{2}-\d)_", base)
        tid = m.group(1) if m else base
        cards, issues = _check_doc(f)
        results[tid] = {"file":f, "cards":len(cards), "issues":issues, "pass":len(issues)==0}
    return results


if __name__ == "__main__":
    r = recheck_all()
    total = len(r)
    passed = sum(1 for v in r.values() if v["pass"])
    failed = total - passed
    print(f"total={total} passed={passed} failed={failed}")
    # count issue types
    from collections import Counter
    c = Counter()
    for v in r.values():
        for i in v["issues"]:
            # give category
            cat = i.split(" ")[0] if " " in i else i
            c[cat]+=1
    print("Issue categories:")
    for k,n in c.most_common():
        print(f"  {k}: {n}")
    # sample 5 failing with issues
    shown=0
    for tid,v in r.items():
        if not v["pass"] and shown<6:
            print(f"\n{os.path.basename(v['file'])}")
            for i in v["issues"]: print("   -",i)
            shown+=1
    json.dump({k:{"cards":v["cards"],"issues":v["issues"],"pass":v["pass"]} for k,v in r.items()},
              open("/tmp/audit_result.json","w",encoding="utf-8"),ensure_ascii=False,indent=1)
    print("\nsaved /tmp/audit_result.json")