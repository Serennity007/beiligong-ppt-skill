---
name: ppt制作
display_name_zh: PPT制作
display_name_en: PPT Production
description: "党校/政务/企业授课课件的批量改造与可读性优化：白底深字、标题分层（统摄标题+页内小标题）、目录重排、小标题色带、右上角完整院标、全篇字号分级放大与逐框实测回退、颜色加深、越框与配色体检、PDF/PNG 导出核验。适用于已有 pptx 的成体系改版，尤其是教师或领导反馈'字太小看不清''底太花''标题每页都不一样''院标缺半截'之后的整改。触发词：PPT改造、课件整改、白底深字、标题分层、字号放大、字体调大、院标、越框检查、批量改PPT。"
description_en: "Batch rework and readability optimization for lecture decks: white background with dark text, two-tier titles, TOC restyle, subtitle color bands, complete institution logo, graded font enlargement with per-box measured rollback, color deepening, overflow and contrast audit, PDF/PNG export verification."
version: v20261021
---

# PPT制作 Skill

面向**已有 .pptx 的成体系改版**，核心目标是"投影看得清、讲述有主线"。
沉淀自一次真实交付：186 页智能体安全课件，按教师录音与现场反馈连改四轮，最终 181 页交付。

## 什么时候用

- 教师/领导反馈：字太小看不清、屏幕不清楚、绿色底太花、标题每页都不一样、院标少半截子
- 需要对上百页课件做**统一版式、字号、配色、标题层级**的批量改造
- 需要"改完还能放心投"——必须有越框实测，而不是凭感觉放大字号

## 快速导航

| 任务 | 读这个 |
|---|---|
| **教师要求与验收标准（硬约束，先读）** | [references/教师要求清单.md](references/教师要求清单.md) |
| 完整改造流程（先做什么后做什么） | [workflow/工作流.md](workflow/工作流.md) |
| 版式与配色硬性规范 | [references/规范_版式与配色.md](references/规范_版式与配色.md) |
| 字号与可读性规范（放大多少、怎么回退） | [references/规范_字号与可读性.md](references/规范_字号与可读性.md) |
| 踩过的坑（文件被占用写坏、COM 误判等） | [references/常见坑.md](references/常见坑.md) |
| 真实案例一：教师录音改造指导书 | [references/案例_01_改造指导书_教师录音纠偏.md](references/案例_01_改造指导书_教师录音纠偏.md) |
| 真实案例二：逐轮改造记录 | [references/案例_02_改造记录_逐轮.md](references/案例_02_改造记录_逐轮.md) |

## 教师要求（硬约束，改造前先读）

教师录音里点名的问题，全部落在 [references/教师要求清单.md](references/教师要求清单.md)，
每条都有编号（A/B/C/D/E/F/G）和判定方式。**这些不是建议，是验收标准**：

- 停止旧做法：目录不再叠导读、不再按"标题唯一"验收、不再给页数图数配额
- 目录只呈现结构；**总图必须画清组成与关系**（不是把标题列表塞进方框）
- 顶部统摄标题 + 页内小标题两层；同一主题连续页共用同一标题
- 案例直接讲：事件 → 问题 → 防控方案（实际处置与建议分开）
- 白底深字、院标完整不缺半截、删"正文位置 / P.6—9"这类编辑定位信息
- 放不下时删冗句，不靠持续缩小字号

改造完用 `ppt_audit.py` 逐条过一遍；脚本判不了的（总图画对没、讲述顺序、口语化标题、连线压字）
会作为人工核对项列在报告末尾。

## 七步工作流（一句话版）

0. **读要求**：翻《教师要求清单》，确认哪些是硬约束、哪些本轮要解决
1. **扫描**：`ppt_scan.py` / `ppt_chrome.py` 摸清版式骨架与节点标签
2. **标题分层**：`ppt_titles.py` 生成统摄标题方案，**人工核对后再执行**
3. **白底深字**：`ppt_white.py`（`--titles` 写统摄标题；配置里的 `purge` 顺手删导读与页段）
4. **装饰件**：`ppt_deco.py` 目录重排 + 小标题色带 + 右上角完整院标
5. **放大字号**：`ppt_enlarge.py` → `ppt_fit.py`（PowerPoint 实测）→ `ppt_apply.py`（写回）
6. **验收交付**：`ppt_audit.py` 按教师要求体检 → `ppt_verify.py` → `ppt_export.py` / `ppt_shots.py`

## 四条不可违反的规则

1. **交付前必须过《教师要求清单》**：目录与总图、标题分层、案例三件套、白底深字、完整院标、无编辑定位信息——`ppt_audit.py` 能自动判的必须全 PASS，判不了的逐条人工勾。
2. **放大字号必须实测回退**：先放大，再调 PowerPoint 逐框测墨迹，放不下的按档回退，且**任何字号不得低于放大前原值**。禁止凭感觉一把放大。
3. **改底色必须同步改字色**：去掉绿底的同时必须把白字改成深色，否则出现"看不见的字"。
4. **写文件前确认未被占用**：pptx 被预览器/微信/Office 打开时写入会整体损坏（页序错乱、丢页）。每轮改前先备份底稿，改后立刻核验页数与页序。

## 脚本清单

全部脚本在 `scripts/`，用管理版 Python（已装 python-pptx / pywin32 / pymupdf）运行：

```
C:/Users/ztl/.workbuddy/binaries/python/envs/default/Scripts/python.exe scripts/xxx.py --help
```

| 脚本 | 作用 | 依赖 |
|---|---|---|
| `pptkit.py` | 公共工具（遍历/取色/亮度/中文字体），被其它脚本 import | python-pptx |
| `ppt_scan.py` | 全量形状扫描，输出 JSON（位置/填充/文字/字号） | python-pptx |
| `ppt_chrome.py` | 识别横幅/标题/小标题/底栏/页码/节点标签，出清单 | python-pptx |
| `ppt_titles.py` | 按节点分组生成统摄标题方案（供人工核对） | python-pptx |
| `ppt_white.py` | 白底深字 + 色号加深 + 写统摄标题 + 删导读页段（`purge`） | python-pptx |
| `ppt_deco.py` | 目录重排 + 小标题色带 + 右上角院标 | python-pptx |
| `ppt_enlarge.py` | 字号分级放大 + 颜色加深 + 收紧行距，出映射表 | python-pptx |
| `ppt_fit.py` | PowerPoint 实测越框并按档回退，出实测字号 | pywin32 + PowerPoint |
| `ppt_apply.py` | 把实测结果写回 pptx | python-pptx |
| `ppt_audit.py` | **按《教师要求清单》逐条验收**（11 项自动判 + 人工核对清单） | python-pptx |
| `ppt_verify.py` | 技术体检：浅底白字/残留浅底/页数/越框 | python-pptx |
| `ppt_export.py` | 导出 PDF 与 PNG | pywin32 + PowerPoint |
| `ppt_shots.py` | PDF 转 PNG 缩略图（跨平台） | pymupdf |

## 配置与案例

- `case/white_config.json` —— 白底深字的色号映射与去底清单
- `case/deco_config.json` —— 色带/院标/目录的位置尺寸与条目
- `case/enlarge_config.json` —— 字号放大曲线与颜色映射
- `case/title_override.json` —— 统摄标题改写表（教师口径）
- `case/*.py` —— 本次交付实际跑过的原始脚本（路径写死版），作对照参考
- `assets/README.md` —— **换院标**：把学院 LOGO 放进 `assets/`，改 `deco_config.json` 的 `logo.image` 一行即可
