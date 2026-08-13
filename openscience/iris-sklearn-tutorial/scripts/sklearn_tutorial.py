"""Iris 数据集 + scikit-learn 教程脚本（分类全流程）。

流程：
  1) 加载数据（data/iris.csv）
  2) 划分训练集 / 测试集（70 / 30，分层抽样，固定随机种子保证可复现）
  3) 构建 Pipeline: 标准化(StandardScaler) + 分类器
  4) 五折交叉验证评估 5 种模型：KNN、Logistic 回归、SVM、
     CART 决策树、随机森林
  5) 在测试集上输出最优模型的分类报告与混淆矩阵
  6) 输出可解释性图：混淆矩阵 + CART 决策树结构

输出：
  - output/models_comparison.txt  5 折交叉验证准确性对比
  - output/classification_report.txt  测试集分类报告
  - output/figures/6_confusion_matrix.png  混淆矩阵
  - output/figures/7_decision_tree.png  决策树结构
"""
import os
import sys

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from sklearn import svm
from sklearn.ensemble import RandomForestClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (
    ConfusionMatrixDisplay,
    accuracy_score,
    classification_report,
    confusion_matrix,
)
from sklearn.model_selection import StratifiedKFold, cross_val_score, train_test_split
from sklearn.neighbors import KNeighborsClassifier
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler
from sklearn.tree import DecisionTreeClassifier, plot_tree

from config_font import setup_matplotlib_chinese

setup_matplotlib_chinese()

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_PATH = os.path.join(ROOT, "data", "iris.csv")
OUT_DIR = os.path.join(ROOT, "output")
FIG_DIR = os.path.join(OUT_DIR, "figures")

RANDOM_STATE = 42

# 建议：也可直接从 sklearn 加载（与下载脚本输出等价）
# from sklearn.datasets import load_iris; data = load_iris()


def load_data():
    df = pd.read_csv(DATA_PATH)
    X = df.drop(columns="species")
    y = df["species"].astype(str)
    return X, y


def build_models():
    return {
        "KNN (k=5)": KNeighborsClassifier(n_neighbors=5),
        "Logistic Regression": LogisticRegression(max_iter=500),
        "SVM (RBF)": svm.SVC(kernel="rbf", C=1.0, gamma="scale"),
        "Decision Tree": DecisionTreeClassifier(max_depth=3, random_state=RANDOM_STATE),
        "Random Forest": RandomForestClassifier(
            n_estimators=100, max_depth=4, random_state=RANDOM_STATE
        ),
    }


def compare_with_cv(X, y, models, n_splits=5):
    """对每个模型做分层 K 折交叉验证。"""
    results = {}
    cv = StratifiedKFold(n_splits=n_splits, shuffle=True, random_state=RANDOM_STATE)
    for name, clf in models.items():
        pipe = Pipeline([("scaler", StandardScaler()), ("clf", clf)])
        scores = cross_val_score(pipe, X, y, cv=cv, scoring="accuracy", n_jobs=-1)
        results[name] = (scores.mean(), scores.std(), scores)
    return results


def main():
    os.makedirs(FIG_DIR, exist_ok=True)
    X, y = load_data()
    print(f"[1/5] 数据已加载: {X.shape[0]} 样本 x {X.shape[1]} 特征，类别={sorted(set(y))}")

    # 2) train/test split（分层）
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.3, stratify=y, random_state=RANDOM_STATE
    )
    print(f"[2/5] 划分完成: 训练 {X_train.shape[0]} / 测试 {X_test.shape[0]}（30%）")

    # 3+4) 交叉验证模型比较
    models = build_models()
    results = compare_with_cv(X, y, models)

    lines = ["=" * 66, "5 折交叉验证：各模型准确率（mean±std）", "=" * 66]
    ranking = sorted(results.items(), key=lambda kv: -kv[1][0])
    for i, (name, (mean, std, _)) in enumerate(ranking, 1):
        lines.append(f"{i:>2}. {name:<20} acc = {mean:.4f} ± {std:.4f}")
        print(f"[3/5] 交叉验证: {name:<20} acc = {mean:.4f} ± {std:.4f}")
    with open(os.path.join(OUT_DIR, "models_comparison.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    # 5) 用最优模型在保留测试集上评估
    best_name, (_, _, _) = ranking[0]  # 枚举有序，best 为第一
    best_clf = models[best_name]
    pipe = Pipeline([("scaler", StandardScaler()), ("clf", best_clf)])
    pipe.fit(X_train, y_train)
    y_pred = pipe.predict(X_test)
    acc = accuracy_score(y_test, y_pred)

    report = "\n".join(
        [
            "=" * 66,
            f"测试集评估（最好的模型: {best_name}）",
            "=" * 66,
            f"测试集准确率: {acc:.4f}",
            "",
            classification_report(y_test, y_pred, digits=4),
        ]
    )
    print(f"[3/5] 测试集评估: {best_name} accuracy = {acc:.4f}")
    with open(os.path.join(OUT_DIR, "classification_report.txt"), "w", encoding="utf-8") as f:
        f.write(report + "\n")

    # 6) 混淆矩阵图
    cm = confusion_matrix(y_test, y_pred, labels=sorted(set(y)))
    fig, ax = plt.subplots(figsize=(7, 6))
    ConfusionMatrixDisplay(cm, display_labels=sorted(set(y))).plot(ax=ax, cmap="Blues")
    ax.set_title(f"图 6   混淆矩阵（{best_name}，测试集）")
    fig.tight_layout()
    fig.savefig(os.path.join(FIG_DIR, "6_confusion_matrix.png"), dpi=150,
                bbox_inches="tight")
    plt.close("all")
    print("[4/6] 6_confusion_matrix.png 已生成")

    # 7) 决策树可视化（logistic 回归亦可绘制，但树形图最直观）
    dt = DecisionTreeClassifier(max_depth=3, random_state=RANDOM_STATE)
    dt.fit(X_train, y_train)
    dt_acc = accuracy_score(y_test, dt.predict(X_test))
    fig, ax = plt.subplots(figsize=(16, 10))
    class_names = sorted(set(y))
    plot_tree(dt, feature_names=X.columns, class_names=class_names,
              filled=True, rounded=True, ax=ax, fontsize=9)
    ax.set_title(f"图 7   CART 决策树（测试集准确率 {dt_acc:.4f}）")
    fig.savefig(os.path.join(FIG_DIR, "7_decision_tree.png"), dpi=150,
                bbox_inches="tight")
    plt.close("all")
    print(f"[5/6] 7_decision_tree.png 已生成（决策树测试准确率 {dt_acc:.4f}）")

    print("\n全部完成。输出:")
    print("  output/models_comparison.txt")
    print("  output/classification_report.txt")
    print("  output/figures/6_confusion_matrix.png")
    print("  output/figures/7_decision_tree.png")


if __name__ == "__main__":
    sys.exit(main())