# 参赛 skill 安装与已知坑（实战笔记）

镜像：GitHub 直连不可用走 `https://gh-proxy.com/<zip-url>`；pip 走 `pypi.tuna.tsinghua.edu.cn`；npm 走 `npmmirror.com`。
Windows 纪律：多行 Python 一律写脚本文件执行（`python -c` 内联多行在 cmd 下会静默失败/GBK 乱码）；PowerShell 同理。

## GordenPPTSkill（模板换字，原生 PPTX）
- 仓库：github.com/GordenSun/GordenPPTSkill。目录放入 skills 目录即可，依赖 python-pptx。
- 工作流：模板 INDEX 选模板 → detail.json 读每槽容量（chars_per_line/max_lines/max_chars）→ 写 edits.json → build_pptx.py。
- 坑：构建自带出框检测但不阻断；超密度文案会重叠/截断，须按告警逐条打磨（一轮 18 页实测修 13 处）。

## ppt-master（SVG→原生 DrawingML）
- 仓库：github.com/hugohe3/ppt-master（53.4k star）。整仓 gh-proxy 下载，免 pip 安装。
- 详见 [ppt-master-workflow.md](ppt-master-workflow.md)。

## claude-pptx（Anthropic 官方 pptx skill）
- ZCode 插件缓存自带（document-skills/pptx）。pptxgenjs 宽幅 13.33×7.5 英寸。
- 结论式标题 + 原生 addChart 图表 + 禁 AI 味排版；密度主动降低是官方设计取舍。

## elite-ppt-pro（咨询风，原生 PPTX）
- vercel-labs/skills PR #813。pptxgenjs 直出。纪律：标题即结论、每页 ≥4 内容区、至多 2 强调色、来源标注。
- 坑：默认产出卡片下半空置、封面无设计元素——需按满版思路重写布局。

## deck-dna（HTML 设计 DNA）
- 仓库：github.com/dakjdakd/PPT-Design-DNA。gh-proxy 下载，Node 跑守卫脚本 ppt-layout-guard.js（48 项检查）。
- 坑：无 PPTX 导出路径（exports/deck.pptx 只是目录约定）；要 PPTX 须按其 Page Specs 用 pptxgenjs 规格化重写。

## guizang-ppt-skill（瑞士风 HTML）
- 仓库：github.com/op7418/guizang-ppt-skill。22 个登记版式 S01–S22，Motion One 动效。
- 坑 1：种子模板缺 21 个版式 CSS 类（ledger-row、brief-grid、matrix-fill、three-forces 等），不补则版式散架。
- 坑 2：HTML 截图必须直接驱动 `deck.style.transform` 翻页（CDN 断网键盘绑定失效）；视口 ≥1920×1080。

## codex-slides（AI Studio 应用）
- 仓库：github.com/nexu-io/codex-slides。npm install（设 ELECTRON_SKIP_BINARY_DOWNLOAD=1）+ next build + next start。
- 阻塞：AI 出图硬绑 chatgpt.com + ChatGPT OAuth；codex CLI spawn 在 Windows 新版 Node 报 EINVAL；空页导出 500。留证据（SSE 日志）后按规则以 pptxgenjs 原生补全。
