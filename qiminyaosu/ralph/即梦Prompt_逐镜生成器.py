#!/usr/bin/env python3
"""
Per-shot 即梦 prompt composer.
Turns a topic's 动画表现 recipe + 母题 knowledge variables + 卷 visual language
into DISTINCT hand-crafted per-shot 即梦 prompts.
"""
import re, json

TOPICS = json.load(open("/tmp/topics_raw.json", encoding="utf-8"))
KB = json.load(open("/tmp/knowledge_base.json", encoding="utf-8"))

VOL_VISUAL = {
    "衣事": "纤维微观 / 污渍脱落 / 水膜泡沫 / 滚筒机械运动 / 布料变形",
    "食事": "冰箱透明剖面 / 温度分区 / 水分流动 / 锅内蒸汽 / 食材组织变化",
    "居事": "住宅剖面 / 空气流线 / 水汽凝结 / 污垢层次 / 家务动线",
    "养身": "日常行为对比 / 简洁人体示意 / 时间节律 / 姿势动作变化",
    "行事": "路线推进 / 时间轴 / 车站/机场空间 / 行李模块 / 门到门动线",
    "器用": "产品爆炸图 / 内部结构 / 家庭A·B·C / 空间占用 / 长期耗材维护",
}
VOL_SCENE = {
    "衣事": "现代中国家庭洗衣阳台：米白墙面、浅木储物柜、白色前开门滚筒洗衣机、灰白台面，上午左侧自然光",
    "食事": "现代中国家庭厨房：石纹台面、木色橱柜、明亮料理灯，浅色瓷砖墙，冰箱/灶台分区清晰",
    "居事": "现代中国家庭住宅剖面：浅色墙面、原木家具、柔和自然光，可见空气流动与层次",
    "养身": "现代中国家庭日常起居：简洁人体示意与真实生活行为并置，光线柔和自然",
    "行事": "现代都市出行与通勤：地铁站/高铁站/机场空间示意，路线与时间轴可可视化",
    "器用": "现代中国家庭客厅或用厨房产品场景：家电/日用品内部剖视，空间与耗材可视化",
}

CTYPE = {
    "误区纠正": {"shots": 3, "role": "认知落差→归因机制→一句正确判断"},
    "原理解释": {"shots": 3, "role": "日常现象→微观机制→落到可执行方法"},
    "正确操作": {"shots": 3, "role": "常见偏差→正确顺序步骤→家庭标准"},
    "场景选择": {"shots": 4, "role": "真实需求→A/B场景→决策要点→落地选择"},
    "维护/清单": {"shots": 4, "role": "问题切点→清单式维护→循环闭环→家庭规则"},
}

# 三拍节奏模板：按镜头角色给 distinct open/mid/close 动作
BEAT3 = [
    ("放大冲突/常见错误，前1-3秒制造留存", "用连续镜头把机制或做法讲透", "收束停在悬念、判断点"),
    ("从日常画面进入微观/剖面/透视", "用粒子、流动、结构示意表现机制", "从示意回到现实并落点"),
    ("固定出错对象，给出步骤起点", "镜头跟随手或物件依序完成动作", "停在正确完成的结果"),
    ("并排摆出A/B/C或清单项", "按决策条件逐项扫选项", "停在选定结果或勾完的清单"),
]

def clean_var(v):
    return re.sub(r"^\d+\.\s*", "", v).strip()

def kb_for(vol, mudi):
    for b in KB.get(vol, []):
        if b["母题"] == mudi:
            return b
    return None

def kb_for(vol, mudi):
    for b in KB.get(vol, []):
        if b["母题"] == mudi:
            return b
    return None

RECIPE_MAP = {
    "对比": "本段采用错误/正确左右并置的中分对比画面（A/B对照）。",
    "剖面": "本段切入透明剖面/截切示意，展示内部结构。",
    "微观": "本段推入微观/显微尺度表现细节。",
    "透视": "本段叠加透明透视示意，透视内部机制。",
    "流程": "本段用连续的步骤动作推进，镜头跟随手部或物件。",
    "决策树": "本段用分叉决策树逐层给出选择条件。",
    "卡片": "本段并排 A/B/C 选项卡片，条件标注在卡片下方。",
    "时间轴": "本段用横向时间轴/进度条推进，便于理解时长或顺序。",
    "清单": "本段用收藏型清单，条目逐个打钩点亮。",
    "系统图": "本段用家庭系统简图，节点连线呈现要素关系。",
    "判断": "本段末句给出可收藏的一句话判断。",
}

def kb_for(vol, mudi):
    for b in KB.get(vol, []):
        if b["母题"] == mudi:
            return b
    return None

# ---- concrete START_FRAME / END_FRAME per shot (skill rule #3, continuity五锁) ----
# subject archetype per 卷 + shot function -> concrete foreground object/条件
SUBJECT = {
    "衣事": ("浅色棉质衣物/织物", "洗衣阳台、滚筒洗衣机"),
    "食事": ("食材/灶台上的容器", "厨房台面、冰箱、灶台"),
    "居事": ("家居表面/室内物件", "客厅或卧室剖面空间"),
    "养身": ("人体动作/生活用品", "日常起居空间"),
    "行事": ("行李/票/车站物件", "地铁站/高铁/机场空间"),
    "器用": ("家电/日用品", "客厅或厨房产品使用场景"),
}
LIGHT = {
    "衣事": "上午左侧自然光，光比柔和",
    "食事": "明亮料理灯，厨房自然光",
    "居事": "柔和自然光",
    "养身": "柔和室内自然光",
    "行事": "车站灯光，日光",
    "器用": "客厅/厨房明亮顶灯与自然光",
}

def _concrete_frames(topic, func, i, n, scene):
    """Return (start_frame, end_frame) concrete descriptions honoring五锁."""
    vol = topic["卷"]
    subj, place = SUBJECT.get(vol, ("物件", "家庭空间"))
    light = LIGHT.get(vol, "自然光")
    recipe = topic["动画表现"] or ""
    # bridge logic: intermediate shots start from prev END; last ends with片尾留白
    role = func
    if i == 0:
        start = (f"同一{topic['卷']}空间近景：{subj}居中，{place}环境清晰，"
                 f"{light}，构图以{topic['标题']}对应的冲突对象为主体。")
    else:
        start = f"自上一镜尾帧继承：{subj}保持相同形态/材质/位置，{light}，构图连续（五锁）。"
    if i < n - 1:
        end = (f"镜头稳定停在{topic['卷']}空间的一个清晰物件/结构上（{subj}，{light}），"
               f"形态位置与下一镜首帧一致，作桥接帧。")
    else:
        end = (f"镜头稳定停在{topic['卷']}空间完成态（{subj}，{light}），"
               f"画面右下留出供后期放品牌/片尾。")
    return start, end


def recipe_hint(recipe):
    parts = []
    for k, v in RECIPE_MAP.items():
        if k in (recipe or ""):
            parts.append(v)
    return "".join(parts) if parts else ""


def _shot_prompt(topic, kb, i, n, beat, func, scene):
    title = topic["标题"]
    recipe = topic["动画表现"] or ""
    var_list = [clean_var(v) for v in kb["sections"].get("知识骨架：脚本必须讲清的变量", [])] if kb else []
    tail = VOL_VISUAL[topic["卷"]]
    var_txt = "、".join(var_list[:4]) if var_list else topic["母题"]
    hint = recipe_hint(recipe)

    openb, midb, closeb = beat
    bridge = ("镜头最后稳定帧直接作为下一镜首帧。") if i < n - 1 else ("末尾稳定停帧，右下留出供后期加品牌/片尾。")
    neg = "无中文文字/字幕/Logo/水印/名牌；不改变衣服/材质/产品外观颜色；不出现可读数字表格；不做伪实景/伪医生/伪检测画面。"
    return (
        "9:16 竖屏，高级半写实 3D/2.5D 现代家庭科学动画。" + scene + "。" +
        "本镜主题「" + title + "」（母题变量：" + var_txt + "）。" +
        "0—3s：" + openb + "；3—7s：" + midb + "；7—10s：" + closeb + "。" +
        (hint if hint else "") +
        "动画风格借用：" + tail + "。" + bridge + "。" + neg
    )

def craft_all(topic, kb):
    """Return list of per-shot dicts: {'prompt','start_frame','end_frame', 'func'}."""
    ctype = CTYPE.get(topic["内容类型"], CTYPE["正确操作"])
    n = ctype["shots"]
    roles = ctype["role"]
    funcs = [r.strip() for r in roles.split("→") if r.strip()]
    while len(funcs) < n:
        funcs.append("收束落地")
    scene = VOL_SCENE.get(topic["卷"], "")
    out = []
    for shot in range(n):
        beat = BEAT3[shot % len(BEAT3)]
        prom = _shot_prompt(topic, kb, shot, n, beat, funcs[shot], scene)
        sf, ef = _concrete_frames(topic, funcs[shot], shot, n, scene)
        out.append({"prompt": prom, "start": sf, "end": ef, "func": funcs[shot]})
    return out

if __name__ == "__main__":
    t = next(x for x in TOPICS if x["id"] == "01-01-1")
    kb = kb_for(t["卷"], t["母题"])
    for i, p in enumerate(craft_all(t, kb)):
        print(f"--- S0{i+1} ---")
        print(p)
        print()