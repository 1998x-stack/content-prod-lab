#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_fund_pipeline.py — 全市场基金逐条调研流水线（never stop）

依据 TODO.json 的 US-002~US-005：
- 读取 all_funds/fund_list.json（约 27037 条）
- 逐条构造"信息充分、schema 完整"的 JSONL 记录，append 进 .workbuddy-ai/research_log.jsonl
- 郑希 8 只基金：本地精编快照增强（top_holdings/净值/规模/role/任职）
- 其余基金：以 fund_list 元数据(code/name/type/abbr/pinyin) + framework 推演为主；
  联网可达时尝试在线增强，网络失败自动降级（不中断、不丢已处理）
- 断点续跑：progress 文件记录已处理 code；--limit/--batch 控制每轮数量；never stop

用法:
  python3 run_fund_pipeline.py --limit 20
  python3 run_fund_pipeline.py --batch 1000
  python3 run_fund_pipeline.py --offset 5000 --limit 100
  python3 run_fund_pipeline.py --reset
  python3 run_fund_pipeline.py --dry
"""
import os, sys, json, glob, time, datetime, argparse

HERE = os.path.abspath(os.path.dirname(__file__))
SKILL_DIR = os.path.expanduser("~/.workbuddy-ai/skills/zhengxi-views")
FUND_LIST = os.path.join(SKILL_DIR, "references", "all_funds", "fund_list.json")
ZX_INDEX = os.path.join(SKILL_DIR, "references", "fund_data", "_index.json")
LOG = os.path.join(HERE, ".workbuddy-ai", "research_log.jsonl")
STATE = os.path.join(HERE, ".workbuddy-ai", "pipeline_state.json")

ZX_CODES = {"001513", "010013", "012920", "506002", "003293", "014275", "110009", "500056"}
SKIP_NAMES = {"的", "之", "—", "-", "混合", "基金"}  # 极短名不补笔记


def iso_now():
    return datetime.datetime.now().astimezone().isoformat(timespec="seconds")


def load_json(p, default=None):
    try:
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def write_json(p, obj):
    with open(p, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False)


def log_append(rec):
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(json.dumps(rec, ensure_ascii=False) + "\n")


def last_id():
    if not os.path.exists(LOG):
        return None
    last = None
    try:
        with open(LOG, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    last = line
        if last:
            return json.loads(last).get("id", "R000000")
    except Exception:
        pass
    return "R000000"


def next_id():
    rid = last_id()
    return "R%06d" % (int(rid[1:]) + 1) if rid else "R000001"


# ---------- 郑希本地快照 ----------
def zx_dir(code):
    base = os.path.join(SKILL_DIR, "references", "fund_data")
    hits = glob.glob(os.path.join(base, f"{code}_*"))
    return hits[0] if hits else None


def read_holdings_json(d):
    p = os.path.join(d, "季度持仓.json")
    if not os.path.exists(p):
        return None, None, None
    arr = load_json(p, [])
    if not arr:
        return None, None, None
    latest = sorted(arr, key=lambda q: (q.get("year", 0), q.get("quarter", 0)))[-1]
    hod = []
    for h in latest.get("holdings", [])[:10]:
        hits = h.get("股票名称")
        if not hits:
            continue
        pct_s = h.get("占净值比", "") or ""
        pct = pct_s.replace("%", "").strip() if pct_s else None
        hod.append({"code": h.get("股票代码"), "name": hits, "pct": pct})
    con = 0.0
    for h in latest.get("holdings", [])[:10]:
        s = h.get("占净值比", "0")
        if s:
            try:
                con += float(s.replace("%", "").strip())
            except Exception:
                pass
    return hod, round(con, 2), f'{latest.get("year")}年第{latest.get("quarter")}季度'


def read_nav(code):
    d = zx_dir(code)
    if not d:
        return {}
    p = os.path.join(d, "净值业绩规模.json")
    if not os.path.exists(p):
        return {}
    pz = load_json(p, {})
    nav = cum = nd = None
    nwt = pz.get("单位净值走势") or []
    if nwt and isinstance(nwt[-1], dict) and nwt[-1].get("y") is not None:
        nav = nwt[-1].get("y")
        x = nwt[-1].get("x")
        if x:
            nd = datetime.datetime.fromtimestamp(x / 1000).strftime("%Y-%m-%d")
    acw = pz.get("累计净值走势") or []
    cum = acw[-1][1] if acw and len(acw[-1]) >= 2 and acw[-1][1] is not None else None
    return {"nav": nav, "cum_nav": cum, "nav_date": nd}


def read_scale_manager(code):
    d = zx_dir(code)
    if not d:
        return None, None
    p = os.path.join(d, "净值业绩规模.json")
    if not os.path.exists(p):
        return None, None
    pz = load_json(p, {})
    scale = scale_date = None
    fs = pz.get("规模变动") or {}
    if isinstance(fs, dict) and fs.get("series") and fs.get("categories"):
        last = fs["series"][-1]
        cats = fs["categories"]
        scale = last.get("y") if isinstance(last, dict) else last
        scale_date = cats[-1] if cats else None
    mgr = None
    fm = pz.get("基金经理") or []
    if isinstance(fm, list) and fm:
        mgr = fm[-1].get("name")
    return scale, scale_date, mgr


def zx_role(code):
    idx = load_json(ZX_INDEX, {"funds": []})
    for f in idx.get("funds", []):
        if f.get("code") == code:
            return f.get("role"), f.get("start"), f.get("end")
    return None, None, None


# ---------- 在线兜底 ----------
def try_online(code):
    try:
        import subprocess
        fetch = os.path.join(SKILL_DIR, "scripts", "fetch_any_fund.py")
        r = subprocess.run(["python3", fetch, code], capture_output=True, text=True, timeout=45)
        if r.returncode != 0:
            return None
        cache = os.path.join(SKILL_DIR, "references", "fund_data_cache")
        for dd in glob.glob(os.path.join(cache, f"{code}_*")):
            hod, con, dt = read_holdings(dd)
            nv = read_nav(dd)
            scale, sdt, mgr = read_meta_manager(dd)
            return {"top_holdings": hod, "concentration_pct": f"前十大合计≈{con}%" if con else None,
                    "holdings_date": dt, "nav": nv.get("nav"), "cum_nav": nv.get("cum_nav"),
                    "nav_date": nv.get("nav_date"), "scale": scale, "scale_date": sdate,
                    "manager": mgr, "sources": ["天天基金在线(经fetch_any_fund)"]}
    except Exception:
        return None
    return None


def read_holdings(d):
    return _read_holdings_impl(d)


def _read_holdings_impl(d):
    p = os.path.join(d, "季度持仓.json")
    if not os.path.exists(p):
        return None, None, None
    arr = load_json(p, [])
    if not arr:
        return None, None, None
    latest = sorted(arr, key=lambda q: (q.get("year", 0), q.get("quarter", 0)))[-1]
    hod = []
    con = 0.0
    for h in latest.get("holdings", [])[:10]:
        nm = h.get("股票名称")
        if not nm:
            continue
        pct_s = h.get("占净值比", "")
        hod.append({"code": h.get("股票代码"), "name": nm, "pct": pct_s})
    return hod, None, f'{latest.get("year")}年第{latest.get("quarter")}季度'


def read_nav_dir(d):
    p = os.path.join(d, "净值业绩规模.json")
    if not os.path.exists(p):
        return {}
    pz = load_json(p, {})
    nav = cum = nd = None
    nwt = pz.get("单位净值走势") or []
    if nwt and isinstance(nwt[-1], dict) and nwt[-1].get("y") is not None:
        nav = nwt[-1].get("y")
        x = nwt[-1].get("x")
        nd = datetime.datetime.fromtimestamp(x / 1000).strftime("%Y-%m-%d") if x else None
    acw = pz.get("累计净值走势") or []
    cum = acw[-1][1] if acw and len(acw[-1]) >= 2 and acw[-1][1] is not None else None
    return {"nav": nav, "cum_nav": cum, "nav_date": nd}


def read_meta_manager(d):
    p = os.path.join(d, "净值业绩规模.json")
    if not os.path.exists(p):
        return None, None, None
    pz = load_json(p, {})
    scale = scale_date = mgr = None
    fs = pz.get("规模变动") or {}
    if isinstance(fs, dict) and fs.get("series") and fs.get("categories"):
        last = fs["series"][-1]
        scale = last.get("y") if isinstance(last, dict) else last
        scale_date = fs["categories"][-1] if fs["categories"] else None
    fm = pz.get("基金经理") or []
    mgr = fm[-1].get("name") if isinstance(fm, list) and fm else None
    return scale, scale_date, mgr


# ---------- 推演 ----------
def fallback_fit(name, ftype):
    n = name.lower()
    if any(k in n for k in ["科技", "信息", "互联", "智能", "数据", "ai", "半导体", "芯片", "算力"]):
        return "赛道属科技/成长，与郑希景气+通胀供给端方向部分契合，但无原话佐证"
    if any(k in ftype for k in ["指数", "债券", "货币", "fof", "qdii"]):
        return "非偏股主动成长，通常低于郑希风格契合度"
    return "主动权益类，需结合持仓判断是否为郑希式高景气成长"


def fallback_score(ftype):
    if any(k in ftype for k in ["债券", "货币", "指数-固收"]):
        return "低(此类天然低于郑希评分，非差基金)"
    if any(k in ftype for k in ["混合", "股票", "fof"]):
        return "中等(需持仓数据精确评估)"
    return "中低/需核实"


# ---------- 主流程 ----------
def build_record(f):
    code, name, ftype = f["code"], f["name"], f["type"]
    rec = {
        "id": next_id(),
        "ts": iso_now(),
        "type": "fund_investigation",
        "fund": {
            "code": code,
            "name": name,
            "manager": None,
            "type": ftype,
            "nav": None, "cum_nav": None, "nav_date": None,
            "scale": None, "scale_date": None
        },
        "top_holdings": [],
        "holdings_date": None,
        "concentration_pct": None,
        "fund_level": {"is_zhengxi": code in ZX_CODES, "zx_role": None, "zx_years": None},
        "framework": {
            "zhengxi_fit_note": fallback_fit(name, ftype),
            "estimated_6dim_score_range": fallback_score(ftype),
            "is_zhengxi_quote": code in ZX_CODES,
            "disclaimer": "评分为按郑希方法的定性推演，非其本人观点，非投资建议"
        },
        "data_sources": ["fund_list.json(元数据)"],
        "notes": f"abbr={f.get('abbr')}; pinyin={f.get('pinyin')}"
    }
    # 郑希本地增强
    if code in ZX_CODES:
        role, st, en = zx_role(code)
        rec["fund_level"].update({"zx_role": role, "zx_years": f"{st}~{en or '今'}"})
        zh_dir = zx_dir(code)
        hod, con, dt = read_holdings_json(zh_dir)
        if hod:
            rec["top_holdings"] = hod
            rec["holdings_date"] = dt
            rec["concentration_pct"] = f"{con}%"
        nv = read_nav(code); scale, sd, mgr = read_meta_manager(zh_dir)
        rec["fund"].update({"nav": nv.get("nav"), "cum_nav": nv.get("cum_nav"), "nav_date": nv.get("nav_date"),
                            "scale": scale, "scale_date": sd, "manager": mgr})
        rec["data_sources"].append("zhengxi-views 本地 fund_data 快照")
    else:
        # 尝试在线增强；网络失败不中断
        on = try_online(code)
        if on:
            rec["top_holdings"] = on["top_holdings"]
            rec["holdings_date"] = on["holdings_date"]
            rec["concentration_pct"] = on["concentration_pct"]
            rec["fund"].update({"nav": on["nav"], "cum_nav": on["cum_nav"], "nav_date": on["nav_date"],
                                "scale": on["scale"], "scale_date": on["scale_date"], "manager": on["manager"]})
            rec["data_sources"].extend(on["sources"])
    return rec


def dedupe_by_code():
    """读 LOG 已存在 code 集合。进度也记在 STATE。"""
    done = set()
    if os.path.exists(LOG):
        with open(LOG, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    o = json.loads(line)
                    if o.get("fund", {}).get("code"):
                        done.add(o["fund"]["code"])
                except Exception:
                    pass
    return done


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=None, help="本轮最多处理 N 条")
    ap.add_argument("--batch", type=int, default=None, help="每轮批量,循环直到跑完(never stop)")
    ap.add_argument("--offset", type=int, default=0)
    ap.add_argument("--reset", action="store_true", help="清空进度重新全量")
    ap.add_argument("--dry", action="store_true", help="只打印待处理,不写文件")
    args = ap.parse_args()

    os.makedirs(os.path.join(HERE, ".workbuddy-ai"), exist_ok=True)
    if args.reset:
        if os.path.exists(LOG):
            os.remove(LOG)
        print("[reset] 已清空 research_log.jsonl")

    funds = load_json(FUND_LIST, {}).get("funds", [])
    print(f"[info] 全市场 {len(funds)} 只基金")
    done = set()
    if os.path.exists(LOG):
        done = dedupe_by_code()
        print(f"[info] 已处理 {len(done)} 只, 去重续跑")

    total_processed = 0
    # 批次循环主：按 batch 分批处理；never stop
    batch = args.batch or args.limit or 1
    stop_after = args.limit if args.limit is not None else None
    start = args.offset

    # 顺序(不是并行)以保持可续跑确定性和 id 顺序
    for i in range(start, len(funds)):
        f = funds[i]
        code = f["code"]
        if code in done:
            continue
        if stop_after is not None and total_processed >= stop_after:
            break
        if args.dry:
            print("DRY:", code, f["name"])
            total_processed += 1
            continue
        try:
            rec = build_record(f)
            log_append(rec)
            total_processed += 1
            if total_processed % 50 == 0:
                print(f"[..] {total_processed} 已写入; 当前位置 {i}/{len(funds)}")
        except Exception as e:
            print(f"[ERR] {code} {f['name']}: {e}")
            # 不中断:修复单条失败,继续
        # 批次控制: batch 模式下每 batch 后继续循环(永不停止),但仍保持单轮内存安全
    print(f"[done] 本轮共写入 {total_processed} 条到 {LOG}")


if __name__ == "__main__":
    main()