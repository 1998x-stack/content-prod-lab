"""matplotlib 中文字体配置（macOS 适用）。

优先使用系统自带 Arial Unicode（完整 CJK 覆盖），
以便图内中文标题正常渲染。
"""
import os

import matplotlib
import matplotlib.font_manager as fm
from matplotlib import pyplot as plt

CANDIDATES = [
    "/System/Library/Fonts/Arial Unicode.ttf",
    "/System/Library/Fonts/PingFang.ttc",
    "/System/Library/Fonts/Hiragino Sans GB.ttc",
    "/System/Library/Fonts/STHeiti Medium.ttc",
]


def setup_matplotlib_chinese():
    installed = []
    for path in CANDIDATES:
        if os.path.exists(path):
            try:
                fm.fontManager.addfont(path)
                installed.append(path)
            except Exception:
                pass
    if installed:
        plt.rcParams["font.sans-serif"] = [
            "Arial Unicode MS", "PingFang SC", "Hiragino Sans GB",
            "STHeiti", "sans-serif",
        ]
        plt.rcParams["axes.unicode_minus"] = False
    return installed