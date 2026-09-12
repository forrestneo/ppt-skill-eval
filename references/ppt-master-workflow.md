# ppt-master 跑法（Windows 实测）

ppt-master（github.com/hugohe3/ppt-master）以 SVG 为中间格式：逐页手写 1280×720 SVG → 官方质检 → 官方转换器出原生 DrawingML PPTX。规范文档约 30 万字，不必全读，按下面最小闭环跑。

## 双赛道对应
- 文字赛道 → Quick Generate 工作流（`workflows/profiles/quick-generate.md`）
- 复刻赛道 → Image to PPTX 工作流（`workflows/profiles/image-to-pptx.md`，Quick-only；其照片级管线依赖 Codex 宿主，ZCode 宿主按"确定性矢量重绘"执行并在 inventory 里注明）

## 步骤
1. 完整性门禁：`python skills/ppt-master/scripts/attribution_guard.py`，非 0 退出即停。
2. 建项目：`python .../scripts/project_manager.py init <name> --quick-generate`（项目生成在框架的 projects/ 下）。
3. 导入源：`python .../scripts/project_manager.py import-sources <project> <大纲.md 或 图片目录>`。
4. 手写 SVG 到 `<project>/svg_output/`，命名 `01_cover.svg`…（零填充，首页定画布）。核心契约：
   - 根：`viewBox="0 0 1280 720"`、`lang="zh-CN"`、`font-family="Microsoft YaHei"`、`data-pptx-page-role="cover|toc|section|content|ending"`；
   - 首个视觉元素 = 全幅 rect 背景，`id="bg" data-pptx-role="background"`；
   - 每个语义模块一个根级 `<g id="..." data-pptx-bounds="x y w h">`，**bounds 互不重叠（>1px 即 fail）**，组内元素用局部坐标 + 转换为绝对坐标写入（自研脚本生成时务必做平移，混用绝对/局部是最大翻车点）；
   - 段落 = `<text>` + 同 x 的 `<tspan dy=行高>`；文本不得溢出 bounds（>5% fail）；
   - 颜色大写 `#RRGGBB`；阴影只允许 `feDropShadow` 标准形态；禁 mask/class/style 元素/textPath/animate/script。
5. 质检（early + final）：
   `python .../scripts/svg_quality_checker.py <project> --quick-generate --canonical-authoring --stage final --json`
6. 导出：`python .../scripts/svg_to_pptx.py <project> --quick-generate --no-notes`
   （原生 DrawingML 形状/文本；产物默认落 `exports/`，附 postflight 报告。）

## 复刻赛道附加件
- 导入图片后写 `<project>/analysis/reconstruction_inventory.json`：每页 SHA-256、像素尺寸、可见区域清单（id/bbox/family/confidence/sufficiency）。文字以识别文件为准，图形确定性矢量重绘，原页图本身绝不作为幻灯片媒体打包。

## 实测数据（供预期管理）
- 18 页 × 2 赛道各约 2500 形状 / 920+ 文本框 / 0 图片；质检 0 error。
- 坑：ASCII 密集统计文本（≥95%、+5、AUM）导出后在 LibreOffice 渲染折行——SVG 里加窄空格 U+2009 加固。
