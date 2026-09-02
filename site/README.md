# 档案站点 · GitHub Pages

`content-prod-lab` 的静态档案站点：**总览页**（按专题聚合）+ **明细页**（每专题文件清单，可点击打开对应的 PDF / DOCX / PPTX / XLSX）。

## 结构与原理

```
site/
├── index.html        # 总览：8 大专题 + 内层分类，按专题进入
├── detail.html?p=<路径>  # 明细：某专题的文件清单，逐条「打开」链接
├── assets/
│   ├── site.css      # 设计系统（绿墨档案 × 朱红印章）
│   ├── app.js        # 渲染器（读取 data.json）
│   └── data.json     # 【自动生成】全库文件清单
└── _build.py         # 从仓库扫描重建 data.json
```

- 明细页通过 `detail.html?p=<文件夹相对路径>` 定位专题；文件夹路径即该目录在仓库中的位置。
- 每条文献的「打开」按钮指向仓库内真实文件（相对路径、UTF-8 编码），点击在新标签页打开 PDF 等。
- `data.json` 由 `python3 site/_build.py` 从仓库扫描生成（PDF/DOCX/PPTX/XLSX/HTML）。仓库结构变化后重新运行一次即可刷新站点。

## 部署到 GitHub Pages

> **关键**：本站链接指向仓库内真实文件（`../research/…` 等），因此 GitHub Pages 必须**从仓库根发布**，`site/` 之外的文件才会一起被发布、链接才能打开。仓库根已放一个 `index.html`，会自动重定向到 `site/index.html`。

1. 把 `site/` 与根 `index.html` 提交推送到仓库（`main` 分支）。
2. GitHub 仓库 **Settings → Pages → Source**：
   - 选择 **Deploy from a branch**；
   - **Branch** 选 `main`，**Folder** 选 `/（root）`；
   - 保存，等待构建完成。
3. 访问 `https://<用户名>.github.io/<仓库名>/` —— 会自动进入 `site/index.html`。

> 站内相对路径（`assets/…`、`../research/…`）基于根目录发布设计：`/site/assets/data.json` 为站点资源，`/research/…` 等为可下载文献。若从 `/site` 单独发布，正文中的 PDF 链接将无法访问。

## 重新生成数据

```bash
cd /Users/x/Desktop/content-prod-lab
python3 site/_build.py     # 覆盖写入 site/assets/data.json
```

## 设计要点

- **定位**：不是「暖米色 + 衬线 + 陶土」的通用模板，而是一套**绿墨纸档案 + 朱红印章**的文献编目空间。
- **印章（Signature）**：每个专题与文件前盖一枚朱红方印，章面取条目首字；悬停时「落章」。
- **结构即信息**：分类不套用 01/02/03 假编号——目录按下层建档；明细页按真实子目录自动分组；页首记录真实编目（X 专题 · Y 文献 · 编目日期）。
- 响应式到移动端、可见键盘焦点、尊重 `prefers-reduced-motion`。