"""Iris 数据集下载脚本。

数据来源（二选一，优先尝试在线 UCI 镜像，失败则回退到 sklearn 内置数据）：
  1. UCI ML Repository: https://archive.ics.uci.edu/ml/datasets/Iris
  2. scikit-learn 内置 `datasets.load_iris()`（自 0.24 起数据被打包在安装目录内，无需联网）

输出：
  - data/iris.csv           原始数据（无表头，含目标列 "species"）
  - data/iris_summary.txt   基本统计信息
"""
import os
import sys

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(ROOT, "data")
CSV_PATH = os.path.join(DATA_DIR, "iris.csv")
SUMMARY_PATH = os.path.join(DATA_DIR, "iris_summary.txt")

COLUMNS = [
    "sepal_length_cm",
    "sepal_width_cm",
    "petal_length_cm",
    "petal_width_cm",
    "species",
]

SPPECIES_NAMES = ["setosa", "versicolor", "virginica"]


def download_from_uci():
    """从 UCI 镜像下载原始 iris.data（需联网）。"""
    url = "https://archive.ics.uci.edu/ml/machine-learning-databases/iris/iris.data"
    df = pd.read_csv(url, header=None, names=COLUMNS)
    df["species"] = df["species"].astype("category")
    return df


def load_from_sklearn():
    """从 sklearn 内置数据加载（离线可用）。"""
    from sklearn.datasets import load_iris

    data = load_iris()
    df = pd.DataFrame(data.data, columns=COLUMNS[:-1])
    df["species"] = [data.target_names[i] for i in data.target]
    df["species"] = df["species"].astype("category")
    return df


def main():
    os.makedirs(DATA_DIR, exist_ok=True)

    try:
        df = download_from_uci()
        source = "UCI (https://archive.ics.uci.edu/ml/datasets/Iris) via network download"
    except Exception:
        df = load_from_sklearn()
        source = "scikit-learn bundled data (offline)"

    df["species"] = df["species"].str.replace("Iris-", "", regex=False)
    df.to_csv(CSV_PATH, index=False)
    print(f"[1/2] 数据集下载完成 -> {CSV_PATH}")
    print(f"      来源: {source}")
    print(f"      形状: {df.shape[0]} 行 x {df.shape[1]} 列")
    print(f"      类别: {dict(df['species'].value_counts())}")

    with open(SUMMARY_PATH, "w", encoding="utf-8") as f:
        f.write("=" * 60 + "\n")
        f.write("Iris 数据集统计摘要\n")
        f.write("=" * 60 + "\n\n")
        f.write(f"来源: {source}\n")
        f.write(f"行数: {len(df)}, 列数: {len(df.columns)}\n\n")
        f.write("描述统计（按数值列）:\n")
        f.write(df.drop(columns="species").describe().to_string())
        f.write("\n\n各类别样本数:\n")
        f.write(df["species"].value_counts().to_string())
        f.write("\n\n缺失值检查:\n")
        f.write(df.isnull().sum().to_string())
        f.write("\n\n重复行数: {}\n".format(int(df.duplicated().sum())))
    print(f"[2/5] 统计摘要已保存 -> {SUMMARY_PATH}")


if __name__ == "__main__":
    sys.exit(main())