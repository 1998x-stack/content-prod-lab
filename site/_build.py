#!/usr/bin/env python3
# Build site/assets/data.json from the live repository.
# Run from anywhere:  python3 site/_build.py
import os, json, time

DOC_EXTS = {".pdf", ".docx", ".pptx", ".xlsx", ".html"}
BASE = os.path.dirname(os.path.abspath(__file__))          # site/
ROOT = os.path.dirname(BASE)                               # repo root
os.chdir(ROOT)

def files_in(path):
    out = []
    for root, dirs, fs in os.walk(path):
        dirs[:] = [d for d in dirs if d not in (".backup", "__pycache__", ".git")]
        for f in fs:
            if os.path.splitext(f)[1].lower() in DOC_EXTS:
                full = os.path.join(root, f)
                out.append({
                    "rel": os.path.relpath(full).replace(os.sep, "/"),
                    "name": os.path.splitext(f)[0],
                    "ext": os.path.splitext(f)[1].lower().lstrip("."),
                    "size": os.path.getsize(full),
                    "date": time.strftime("%Y-%m-%d", time.localtime(os.path.getmtime(full))),
                })
    out.sort(key=lambda x: (x["rel"].lower()))
    return out

def folder(name, path):
    fl = files_in(path)
    return {"name": name, "path": path, "count": len(fl), "files": fl}

DIVISIONS = [
    {"name": "研究档案", "key": "research",
     "blurb": "按学科主题归档的深度研究报告库：AI 前沿与训练、LLM 底层原理、宏观政策与经济、市场情报、工程与源码、教育与升学、历史与军事、个人成长。",
     "folders": [folder(n, p) for n, p in [
        ("AI与训练", "research/AI与训练"),
        ("LLM底层技术", "research/LLM底层技术"),
        ("宏观政策与经济", "research/宏观政策与经济"),
        ("市场情报", "research/市场情报"),
        ("教育升学", "research/教育升学"),
        ("工程与源码", "research/工程与源码"),
        ("历史与军事", "research/历史与军事"),
        ("个人成长", "research/个人成长"),
    ]]},
    {"name": "课程教学", "key": "courses",
     "blurb": "经济学讲义与配套课件（PDF 讲义 + PPTX 课件），按学科组织：宏观、微观、货币金融。",
     "folders": [folder(n, p) for n, p in [
        ("宏观经济学", "课程教学/宏观经济学"),
        ("微观经济学", "课程教学/微观经济学"),
        ("货币金融学", "课程教学/货币金融学"),
    ]]},
    {"name": "系列长文", "key": "series",
     "blurb": "按『系列』组织的长篇连载，命名《系列名_第X章_标题》，每章一 DOCX。",
     "folders": [folder(d, "series/" + d) for d in sorted(os.listdir("series")) if not d.startswith(".")]},
    {"name": "出行档案", "key": "travel",
     "blurb": "城市的深度旅行攻略：两日游、精确证据版攻略等。",
     "folders": [folder("出行攻略", "出行攻略")]},
    {"name": "城市报告", "key": "reports",
     "blurb": "城市 / 机构年度报告深度研究合集（PDF 报告与 HTML 成果）。",
     "folders": [folder("城市/机构年度报告", "reports")]},
    {"name": "OpenCode 开发者库", "key": "opencode",
     "blurb": "OpenCode / Claude Code 生态交付物：cookbook、教程、方案、汇报与深度研究。",
     "folders": [folder(d, "opencode/" + d) for d in sorted(os.listdir("opencode")) if not d.startswith(".")]},
    {"name": "基金研究", "key": "fund",
     "blurb": "主动权益 Beta 数据库与研究管线。",
     "folders": [folder("stock-fund", "stock-fund")]},
    {"name": "技术与产品方案", "key": "proposals",
     "blurb": "技术 / 产品方案文档。",
     "folders": [folder("proposals", "proposals")]},
]

for d in DIVISIONS:
    d["total"] = sum(f["count"] for f in d["folders"])

data = {
    "title": "Content Production Lab · 中文内容生产档案",
    "total": sum(d["total"] for d in DIVISIONS),
    "divisions": DIVISIONS,
    "generated": time.strftime("%Y-%m-%d"),
}

out = os.path.join(BASE, "assets", "data.json")
os.makedirs(os.path.dirname(out), exist_ok=True)
with open(out, "w") as fh:
    json.dump(data, fh, ensure_ascii=False, indent=1)
print("Wrote", out, "| total files:", data["total"])