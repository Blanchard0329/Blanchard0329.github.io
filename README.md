# 2026 June Exam Calendar Generator

一个可直接部署到 GitHub Pages 的静态网页工具，用于从你提供的考试总表中勾选考试并生成个性化日历。

## 已实现功能

- 按 `IG / AS / A2` 分类展示考试，支持搜索、多选、按科目/按层级批量全选。
- 冲突自动处理：当同一时间段考试冲突时，自动尝试切换到 `S2 / S3` 变体时段。
- 自动生成 `Supervised Break`（隔离休息）时间块。
- 导出 `PNG`（清爽高对比度，适合保存或打印）
- 导出 `ICS`（支持 iPhone / Android / Mac / Outlook），可选提醒：
  - 考前 1 天
  - 考前 2 小时

## 目录结构

- `index.html`：主页面
- `styles.css`：样式
- `app.js`：前端逻辑（筛选、冲突处理、导出）
- `data/exams-2026-june.json`：当前网页使用的 2026 June 考试数据
- `data/exams-2025-mj.json`：上一版 2025 May/June 考试数据，保留作备份和对照
- `scripts/extract_exam_data.py`：Excel 转 JSON 脚本（无第三方依赖）

## 本地运行

```bash
cd /Users/macbook/Documents/Playground
python3 -m http.server 8080
```

浏览器打开：`http://localhost:8080`

## 若 Excel 更新，重新生成数据

```bash
python3 /Users/macbook/Documents/Playground/scripts/extract_exam_data.py \
  --input "/Users/macbook/Downloads/20260416 To Candidate Provisional June 2026 Exam Timetable.xlsx" \
  --output /Users/macbook/Documents/Playground/data/exams-2026-june.json
```

## 部署到 GitHub Pages

1. 把当前目录推到 GitHub 仓库（例如 `main` 分支）。
2. 在 GitHub 仓库设置中打开 `Settings -> Pages`。
3. `Build and deployment` 选择：
   - Source: `Deploy from a branch`
   - Branch: `main`（或你的分支）
   - Folder: `/ (root)`
4. 保存后等待发布，GitHub 会给出访问链接。

## 说明

- 当前数据来自你提供的 Excel，覆盖时间范围：`2026-04-24` 到 `2026-06-12`。
- `AS/A2` 分类基于考试代码与试卷编号进行自动推断，适用于当前数据结构。
