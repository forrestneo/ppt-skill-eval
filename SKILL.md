---
name: ppt-skill-benchmark
description: 对多款 PPT 生成 skill / AI PPT 工具做横向评测的完整工作流：双赛道（纯文字大纲 + 原版图片复刻）生成 18 页原生可编辑 PPTX、渲染逐页核验、打散盲评打分工作台、排行榜/象限图/打分矩阵图表、楷体 DOCX+PDF 评测报告与全部成品截图附录。只要用户提到「评测 PPT skill」「对比几个 PPT 生成工具哪个好」「给这些 PPT skill 打分排名」「搭一个盲评打分工作台」「把评测报告出成 DOCX/PDF」，即使没说"评测"二字（如"这几个 ppt skill 哪个最能打"），也使用本 skill。
---

# PPT Skill 横向评测工作流

把"对比若干款 PPT 生成 skill 谁强"做成一场可复现的评测：同一基准、双赛道、强制原生可编辑、渲染逐页核验、打散盲评、终榜报告。

## 硬规则（先读，违反即评测作废）

1. **页数铁律**：所有产物与基准页数 1:1（默认 18 页），页序对应主题。初版不足页数的，补齐重写，不许合并/拆分/缩水。
2. **原生可编辑铁律**：产物必须是文本框级可编辑的 PPTX（python-pptx / pptxgenjs / 官方转换器直出）。整页截图贴入 PPTX = 该产物不合格。交付后用 `scripts/verify_pptx.py` 核验：页数、形状数、文本框数、**图片数必须为 0**。
3. **多模态铁律**：图片复刻赛道每页生成前必须重新查看对应原图，按所见复刻，不许凭记忆。
4. **评测标准排序**：美观是主赛道；可编辑性是二筛（全员达标后不再计分）；相似度仅是参考项。核心标准一句话：「可编辑但丑 = 废物」。
5. **写法纪律**：报告结论先行（先排行榜后细节）；全篇统一评测对象数量口径（不要一会 6 个一会 7 个）；不写"应用户要求/用户指定"这类叙述腔，直接陈述做了什么；skill 名用其官方全名（如 GordenPPTSkill 不写 gorden-ppt-skill）。

## 九步流程

### Step 1 基准采集
从来源（微信文章/网页/PPTX）提取全部页面图存为 PNG（1080p），逐页多模态识别文字与版式，合并成文字大纲（约 1 万字量级）。微信图片直链用正则从 HTML 提取（mmbiz.qpic.cn 域名）。

### Step 2 Skill 就位
逐款安装并实测可用性。下载受阻用镜像：GitHub 走 `gh-proxy.com`，pip 走清华源，npm 走 npmmirror。各 skill 的安装要点与已知坑见 [references/skill-setup-notes.md](references/skill-setup-notes.md)；ppt-master 的完整跑法见 [references/ppt-master-workflow.md](references/ppt-master-workflow.md)。安装不等于可用——每款先跑通最小产物再进入正式生成。

### Step 3 文字赛道
唯一输入 = 文字大纲。逐款驱动生成完整 deck。选型覆盖尽量多的技术路线：模板换字 / HTML 设计系统 / pptxgenjs 原生直出 / AI Studio 应用 / SVG 工程化管线。

### Step 4 复刻赛道
唯一输入 = 基准页面图。逐页 1:1 对照复刻，页序、标题、数据与原图一致。

### Step 5 渲染核验
所有产物统一转 PNG 后逐页亲眼核对：
```bash
python scripts/render_pptx_pages.py <deck.pptx> --outdir <render_dir>
python scripts/verify_pptx.py <deck.pptx>
python scripts/make_collages.py <render_dir> --out <collage.png>
```
HTML deck 截图必须直接驱动 `deck.style.transform` 翻页（CDN 断网时键盘绑定失效）；视口 ≥1920×1080，否则裁切错位——这是历史误判的根因。

### Step 6 执行工具评审（过程分）
逐页审阅渲染图打 10 分制过程分（美观 60% + 相似度/覆盖 25% + 缺陷 15%；文字赛道美观 55% + 覆盖 25% + 缺陷 15% + 结构 5%）。**过程分只做交叉参考，永远不进排名**——实践证明它系统性偏宽。

### Step 7 盲评工作台
每套产物 18 张渲染图加入浏览器打分台：全部打散混排、随机五字符盲码匿名（如 BYQGD-07）、进度 localStorage 存档、CSV/JSON 导出、一键揭盲。用 `scripts/build_workbench.py` 增量加图（保留已有打分与顺序，新图稳定洗牌后追加）：
```bash
python scripts/build_workbench.py --workbench <工作台目录> --render-dir <render_dir> --track 文字赛道 --skill "skill 显示名"
```
工作台本体（index.html）支持增量合并：manifest 新增图片后，老图顺序与已打分数保留，新图追加到赛道末尾。

### Step 8 人工盲评
用户逐张 1–5 分，全部打完无遗漏后导出 CSV 存档。打分期间不看盲码对应关系。

### Step 9 揭盲统计与报告
揭盲计算均分：套分 = 18 页平均；综合 = 两赛道平均。用 `scripts/make_charts.py` 出图（分组柱状图 + 2.5 推荐线 + 四象限图，STKaiti）。报告按 [references/report-spec.md](references/report-spec.md) 的章节规范生成 DOCX（楷体 STKaiti、手写目录、图表内嵌、附录放全部成品 3×6 拼接图），再经 LibreOffice 转 PDF。报告模板见 `scripts/report_template.py`（复制后替换数据区）。

## 常见坑（历史教训）

- 控制台中文/引号地狱：Windows 下不用 `python -c "..."` 内联多行代码，一律写脚本文件再执行。
- LibreOffice 文本框宽度估算比浏览器渲染窄，ASCII 密集统计文本（"≥95%" "+5" "AUM"）会折行——写 SVG/HTML 版式时给这类文本加窄空格（U+2009）或去空格。
- 拼接图、图表、打分台数据都属评测证据，随报告一起交付，别只交正文。
- 转换/补全类产物（HTML→pptxgenjs、管线阻塞后补全）必须在报告中标注，与原生直出产物不可比的部分要说明。

## 交付清单

评测结束时应同时交付：各款 deck18.pptx、render18/ 渲染图、每套拼接图、盲评工作台（含揭盲 answers）、盲评 CSV 存档、柱状图/象限图、DOCX+PDF 报告。
