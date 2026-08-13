"""Iris 数据集可视化脚本。

图 1  鸢尾花按品种分布（sepal & petal）的散点矩阵    -> output/figures/1_pairplot.png
图 2  四种特征在各品种下的箱线图                    -> output/figures/2_boxplot.png
图 3  特征分布直方图（按品种着色）                   -> output/figures/3_histograms.png
图 4  相关性热力图                                  -> output/figures/4_correlation_heatmap.png
图 5  PCA 二维投影（4 维数据降维到平面）             -> output/figures/5_pca_scatter.png
"""
import os
import sys

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns

from config_font import setup_matplotlib_chinese

setup_matplotlib_chinese()

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_PATH = os.path.join(ROOT, "data", "iris.csv")
FIG_DIR = os.path.join(ROOT, "output", "figures")

sns.set_theme(style="whitegrid")
PALETTE = {"setosa": "#1f77b4", "versicolor": "#ff7f0e", "virginica": "#2ca02c"}


def main():
    os.makedirs(FIG_DIR, exist_ok=True)
    df = pd.read_csv(DATA_PATH)
    df["species"] = df["species"].astype("category")

    # 图 1: 散点矩阵（plotting="scatter" 保持经典风格）
    g = sns.PairGrid(df, hue="species", palette=PALETTE, corner=True)
    g.map_lower(sns.scatterplot, s=18)
    g.map_diag(sns.histplot, edgecolor="w", alpha=0.7)
    g.map_upper(sns.kdeplot, fill=True, alpha=0.35, levels=4)
    g.add_legend(title="Species")
    g.figure.suptitle("图 1   Iris 各特征两两分布（颜色=品种）", y=1.02, fontsize=12)
    g.savefig(os.path.join(FIG_DIR, "1_pairplot.png"), dpi=150, bbox_inches="tight")
    plt.close("all")
    print("[1/5] 1_pairplot.png 已生成")

    # 2: 箱线图
    fig, axes = plt.subplots(2, 2, figsize=(12, 8))
    for ax, col in zip(axes.ravel(), df.columns[:-1]):
        sns.boxplot(data=df, x="species", y=col, hue="species",
                    palette=PALETTE, ax=ax, legend=False)
        ax.set_title(f"{col} 按品种分布")
        ax.set_xlabel("")
    fig.suptitle("图 2   Iris 四个特征在各品种下的箱线图", fontsize=13)
    fig.tight_layout()
    fig.savefig(os.path.join(FIG_DIR, "2_boxplot.png"), dpi=150, bbox_inches="tight")
    plt.close("all")
    print("[2/5] 2_boxplot.png 已生成")

    # 3: 直方图
    fig, axes = plt.subplots(2, 2, figsize=(12, 8))
    for ax, col in zip(axes.ravel(), df.columns[:-1]):
        for sp in df["species"].cat.categories:
            sns.histplot(df.loc[df["species"] == sp, col], kde=True,
                         label=sp, ax=ax, alpha=0.55, palette=PALETTE)
        ax.set_title(f"{col}")
        ax.set_ylabel("")
    axes[0, 0].legend(title="品种", bbox_to_anchor=(1.05, 1))
    fig.suptitle("图 3   Iris 特征分布直方图", fontsize=13)
    fig.tight_layout()
    fig.savefig(os.path.join(FIG_DIR, "3_histograms.png"), dpi=150, bbox_inches="tight")
    plt.close("all")
    print("[3/5] 3_histograms.png 已生成")

    # 4: 相关性热力图
    fig, ax = plt.subplots(figsize=(8, 6.5))
    corr = df.drop(columns="species").corr()
    sns.heatmap(corr, annot=True, fmt=".2f", cmap="coolwarm", vmin=-1, vmax=1,
                center=0, ax=ax, linewidths=0.5)
    ax.set_title("图 4   Iris 特征相关性矩阵")
    fig.tight_layout()
    fig.savefig(os.path.join(FIG_DIR, "4_correlation_heatmap.png"), dpi=150, bbox_inches="tight")
    plt.close("all")
    print("[4/5] 4_correlation_heatmap.png 已生成")

    # 5: PCA 可视化
    from sklearn.decomposition import PCA

    X = df.drop(columns="species")
    y = df["species"]
    pca = PCA(n_components=2)
    X_pca = pca.fit_transform(X)
    explained = pca.explained_variance_ratio_

    fig, ax = plt.subplots(figsize=(9, 7))
    for i, sp in enumerate(y.cat.categories):
        mask = y == sp
        ax.scatter(X_pca[mask, 0], X_pca[mask, 1], s=45, alpha=0.85,
                   label=sp, color=PALETTE[sp])
    ax.set_xlabel(f"PCA 第一主成分（解释方差 {explained[0]*100:.1f}%）")
    ax.set_ylabel(f"PCA 第二主成分（解释方差 {explained[1]*100:.1f}%）")
    ax.set_title(f"图 5   Iris PCA 二维投影（累计解释方差 {explained.sum()*100:.1f}%）")
    ax.legend()
    fig.tight_layout()
    fig.savefig(os.path.join(FIG_DIR, "5_pca_scatter.png"), dpi=150, bbox_inches="tight")
    plt.close("all")
    print(f"[5/7] 5_pca_scatter.png 已生成（PCA 累计解释方差 {explained.sum()*100:.1f}%）")

    print("\n全部 5 张图已输出到:", FIG_DIR)


if __name__ == "__main__":
    sys.exit(main())