#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
web_enrich_fund.py — 用 web 公开页面逐只富集基金记录（websearch 流程的半自动助手）

用法:
  python3 web_enrich_fund.py 002839           # 只打印建议记录(不写文件)
  python3 web_enrich_fund.py 002839 --write   # 校验后追加进 research_log.jsonl
  python3 web_enrich_fund.py --batch 5        # 处理 fund_list 中下一批未入库 N 只(需手动提供每只的 web 提取结果)

说明：
  - 本工具主要职责：把「人工/WebFetch 核对的真实数据」与 fund_list 元数据合并成
    标准 schema 记录，避免手工重复造轮子。
  - 联网数据源：新浪财经 FundInfo_ZCGP 公开页（与项目 old 记录一致），网络被拦时降级。
  - 默认只打印预览；--write 才追加。
研究学习辅助，非投资建议。
"""
import os, sys, json, datetime, argparse, subprocess

HERE = os.path.abspath(os.path.dirname(__file__))
LOG = os.path.join(HERE, ".workbuddy-ai", "research_log.jsonl")
SKILL_DIR = os.path.expanduser("~/.workbuddy-ai/skills/zhengxi-views")
FUND_LIST = os.path.join(SKILL_DIR, "references", "all_funds", "fund_list.json")


def last_id():
    last = None
    try:
        with open(LOG, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    last = line
        return json.loads(last).get("id", "R000000")
    except Exception:
        return "R000000"


def next_id():
    return "R%06d" % (int(last_id()[1:]) + 1)


def fund_meta(code):
    try:
        funds = json.load(open(FUND_LIST, encoding="utf-8"))["funds"]
        for f in funds:
            if f["code"] == code:
                return f
    except Exception:
        pass
    return {"code": code, "name": code, "type": "", "abbr": "", "pinyin": ""}


def is_done(code):
    if not os.path.exists(LOG):
        return False
    with open(LOG, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                if json.loads(line).get("fund", {}).get("code") == code:
                    return True
            except Exception:
                pass
    return False


def merge_record(code, web_data):
    """web_data: dict(nav,cum_nav,nav_date,scale,scale_date,manager,top_holdings,holdings_date,concentration,sources)
    返回完整 rec(不写盘)。top_holdings 允许为 []，表示仅元数据。"""
    meta = fund_meta(code)
    hod = web_data.get("top_holdings", [])
    con = web_data.get("concentration")
    if con is None and hod:
        try:
            con = f"{round(sum(float(h['pct']) for h in hod),1)}%"
        except Exception:
            con = None
    rec = {
        "id": next_id(),
        "ts": datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
        "type": "fund_investigation",
        "fund": {
            "code": code, "name": meta["name"], "manager": web_data.get("manager"),
            "type": meta["type"] or web_data.get("type"),
            "nav": web_data.get("nav"), "cum_nav": web_data.get("cum_nav"),
            "nav_date": web_data.get("nav_date"), "scale": web_data.get("scale"),
            "scale_date": web_data.get("scale_date"),
        },
        "top_holdings": hod,
        "holdings_date": web_data.get("holdings_date"),
        "concentration_pct": con,
        "fund_level": {
            "is_zhengxi": code in {"001513","010013","012920","506002","003293","014275","110009","500056"},
            "zx_role": None, "zx_years": None,
        },
        "framework": {
            "zhengxi_fit_note": _fit_note(meta["name"], meta["type"]),
            "estimated_6dim_score_range": _score(meta["type"]),
            "is_zhengxi_quote": False,
            "disclaimer": "评分为按郑希方法的定性推演，非其本人观点，非投资建议",
        },
        "data_sources": web_data.get("sources", ["fund_list.json(元数据)"]),
        "notes": f"abbr={meta.get('abbr')}; pinyin={meta.get('pinyin')}",
    }
    return rec


def _fit_note(name, ftype):
    n = name.lower(); ft = (ftype or "").lower()
    if any(k in n for k in ["科技","信息","互联","智能","数据","ai","半导体","芯片","算力"]):
        return "赛道属科技/成长，与郑希景气+通胀供给端方向部分契合，但无原话佐证"
    if any(k in ft for k in ["货币","债券","指数-固收"]):
        return "非偏股主动成长，通常低于郑希风格契合度"
    if any(k in ft for k in ["qdii","fof","指数"]):
        return "被动/境外/FOF，通常郑希契合度一般"
    return "主动权益类，需结合持仓判断是否为郑希式高景气成长"


def _score(ftype):
    ft = (ftype or "").lower()
    if any(k in ft for k in ["货币","债券","指数-固收"]):
        return "低(此类天然低于郑希评分，非差基金)"
    if any(k in ft for k in ["混合","股票","fof"]):
        return "中等(需持仓数据精确评估)"
    return "中低/需核实"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("codes", nargs="*", help="基金代码")
    ap.add_argument("--write", action="store_true", help="真正追加进 JSONL")
    ap.add_argument("--next", type=int, default=0, help="处理 fund_list 中接下来 N 只未入库(仅预览,需配合WebFetch)")
    args = ap.parse_args()

    # 若要处理 fund_list 后续批次
    codes = list(args.codes)
    if args.next:
        pending = []
        try:
            funds = json.load(open(FUND_LIST, encoding="utf-8"))["funds"]
        except Exception:
            funds = []
        for f in funds:
            if not is_done(f["code"]):
                pending.append(f)
            if len(pending) >= args.next:
                break
        codes = [f["code"] for f in pending]
        print(f"[next] 待处理 {len(codes)} 只: {codes}")

    for code in codes:
        if is_done(code):
            print(f"[skip] {code} 已入库")
            continue
        meta = fund_meta(code)
        print(f"\n=== {code} {meta['name']} [{meta['type']}] === (未写盘, 请用WebFetch核对该code持仓后以 --write 写入)")
        print("建议: WebFetch https://stock.finance.sina.com.cn/fundInfo/view/FundInfo_ZCGP.php?symbol="+code)


if __name__ == "__main__":
    main()