# ppt制作

一个面向**已有 PPTX 成体系改版**的 Skill / 工作流：把"字太小看不清、底太花、标题每页都不一样、院标缺半截"这类反馈，变成一条可复跑的流水线。

沉淀自一次真实交付：某党校 186 页《智能体安全：技术趋势与风险防控》课件，按教师录音与现场反馈连改四轮，最终 181 页交付。真实指导书、逐轮改造记录、当初跑过的原始脚本全部保留在 `case/` 与 `references/`。

## 目录

```
ppt制作/
├── SKILL.md                 # Skill 入口（触发词、六步流程、三条红线）
├── README.md                # 本文件
├── workflow/工作流.md        # 六步工作流，含每条命令
├── references/
│   ├── 规范_版式与配色.md    # 白底深字的硬性规范、色号映射、院标规格
│   ├── 规范_字号与可读性.md  # 放大曲线、实测回退算法、越框修复阶梯
│   ├── 常见坑.md            # 8 个踩过的坑
│   ├── 案例_01_改造指导书_教师录音纠偏.md
│   └── 案例_02_改造记录_逐轮.md
├── scripts/
│   ├── pptkit.py            # 公共工具（遍历 / 取色 / 亮度 / 中文字体）
│   ├── ppt_scan.py          # 全量形状扫描
│   ├── ppt_chrome.py        # 版式骨架识别（横幅/标题/小标题/底栏/页码/节点标签）
│   ├── ppt_titles.py        # 统摄标题方案
│   ├── ppt_white.py         # 白底深字 + 写统摄标题
│   ├── ppt_deco.py          # 目录重排 + 小标题色带 + 右上角院标
│   ├── ppt_enlarge.py       # 字号分级放大（出映射表）
│   ├── ppt_fit.py           # PowerPoint 实测越框并回退（Windows）
│   ├── ppt_apply.py         # 实测字号写回
│   ├── ppt_verify.py        # 交付前体检
│   ├── ppt_export.py        # 导出 PDF / PNG（Windows）
│   └── ppt_shots.py         # PDF 转缩略图（跨平台）
└── case/
    ├── *_config.json        # 白底 / 装饰 / 放大的配置
    ├── title_override.json  # 统摄标题改写表
    └── *.py                 # 本次交付实际跑过的原始脚本（路径写死版，作对照）
```

## 环境

```bash
pip install python-pptx          # 必需
pip install pywin32 pymupdf      # 实测回退与导出需要（Windows + PowerPoint）
```

## 最简跑法

```bash
PY="C:/Users/ztl/.workbuddy/binaries/python/envs/default/Scripts/python.exe"

$PY scripts/ppt_chrome.py 底稿.pptx _build/chrome.json --txt _build/chrome.txt
$PY scripts/ppt_titles.py _build/chrome.json _build/plan.json --override case/title_override.json --txt _build/plan.txt
$PY scripts/ppt_white.py 底稿.pptx _build/v白底.pptx --config case/white_config.json --titles _build/plan.json
$PY scripts/ppt_deco.py  _build/v白底.pptx _build/v装饰.pptx --config case/deco_config.json
$PY scripts/ppt_enlarge.py _build/v装饰.pptx _build/v放大.pptx --map _build/size_map.json --config case/enlarge_config.json
$PY scripts/ppt_fit.py     _build/v放大.pptx --map _build/size_map.json --out _build/fit.json
$PY scripts/ppt_apply.py   _build/v放大.pptx --fit _build/fit.json --out 交付.pptx
$PY scripts/ppt_verify.py  交付.pptx --fit _build/fit.json
```

完整流程与每一步为什么这么做，见 [workflow/工作流.md](workflow/工作流.md)。

## 三条红线

1. **放大字号必须实测回退** —— 先放大，再调 PowerPoint 逐框测墨迹，放不下按档回退；任何字号不得低于放大前原值。
2. **改底色必须同步改字色** —— 去绿底不改白字，等于造出"看不见的字"。
3. **写文件前确认未被占用** —— PPTX 被预览器/微信/Office 打开时写入会整体损坏；改前备份、改后立刻核验页数页序。
