# -*- coding: utf-8 -*-
# 终版 DOCX v6：一次评测 · 7 款 skill · 双赛道 · 252 张盲评（全篇统一口径重写）
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

RED = RGBColor(0xB0, 0x1F, 0x24)
DARK = RGBColor(0x1C, 0x19, 0x17)
MUT = RGBColor(0x5A, 0x52, 0x4C)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
F = "STKaiti"

doc = Document()
for section in doc.sections:
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)

style = doc.styles["Normal"]
style.font.name = F
style.font.size = Pt(12)
style.element.rPr.rFonts.set(qn("w:eastAsia"), F)
style.paragraph_format.line_spacing = 1.4

def set_cn(run, size=None, bold=None, color=None, italic=None):
    run.font.name = F
    run.element.rPr.rFonts.set(qn("w:eastAsia"), F)
    if size: run.font.size = Pt(size)
    if bold is not None: run.font.bold = bold
    if color: run.font.color.rgb = color
    if italic is not None: run.font.italic = italic

def para(text="", size=12, bold=False, color=None, align=None, after=8, spacing=1.4, indent=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_cn(r, size, bold, color)
    if align: p.alignment = align
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = spacing
    if indent: p.paragraph_format.left_indent = Cm(indent)
    return p

def heading(text, level=1):
    h = doc.add_heading(text, level=level)
    for r in h.runs:
        set_cn(r, 22 if level == 1 else 16, True, RED if level == 1 else DARK)
    h.paragraph_format.space_before = Pt(24)
    h.paragraph_format.space_after = Pt(12)
    return h

def subheading(text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_cn(r, 14, True, DARK)
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(8)
    return p

def body(text, size=12, bold=False, color=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_cn(r, size, bold, color)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    return p

def note(text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_cn(r, 10, False, MUT)
    p.paragraph_format.space_after = Pt(4)
    return p

def bullets(items, size=12, indent=0.5):
    for it in items:
        p = doc.add_paragraph()
        r = p.add_run("• " + it)
        set_cn(r, size, False, DARK)
        p.paragraph_format.left_indent = Cm(indent + 0.5)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.4

def table(headers, rows, widths=None, fontsize=10):
    t = doc.add_table(rows=1, cols=len(headers)); t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]; c.text = ""
        r = c.paragraphs[0].add_run(h); set_cn(r, fontsize, True, WHITE)
        shd = OxmlElement("w:shd"); shd.set(qn("w:fill"), "B01F24")
        c._tc.get_or_add_tcPr().append(shd)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            r = cells[i].paragraphs[0].add_run(str(v)); set_cn(r, fontsize - 0.5)
    if widths:
        for i, w in enumerate(widths):
            for row in t.rows: row.cells[i].width = Cm(w)
    return t

def img(path, wcm=15.0, cap=None):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(path, width=Cm(wcm))
    if cap:
        cp = doc.add_paragraph(); cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = cp.add_run(cap); set_cn(r, 10, False, MUT)
        cp.paragraph_format.space_after = Pt(8)

def pagebreak(): doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

# ================= 封面 =================
para("", after=60)
para("7 款 PPT 生成 Skill 横向评测报告", 32, True, RED, WD_ALIGN_PARAGRAPH.CENTER, 16)
para("基于 18 页银行零售运营方案基准 · 双赛道 · 252 张盲评全部完成", 16, False, DARK, WD_ALIGN_PARAGRAPH.CENTER, 28)
img(os.path.join("sample", "images", "slide01.png"), 13.0)
para("评测基准样例：原版 PPT 封面页（全部 18 页样例见第一章拼接总览）", 10, False, MUT, WD_ALIGN_PARAGRAPH.CENTER, 24)
para("被评测对象（按最终名次）", 11, True, DARK, WD_ALIGN_PARAGRAPH.CENTER, 6)
para("GordenPPTSkill ｜ ppt-master ｜ claude-pptx（Anthropic 官方）", 11, False, DARK, WD_ALIGN_PARAGRAPH.CENTER, 2)
para("elite-ppt-pro ｜ deck-dna ｜ guizang-ppt-skill ｜ codex-slides", 11, False, DARK, WD_ALIGN_PARAGRAPH.CENTER, 20)
para("评测执行：达仔 × ZCode（GLM-5.3-Flash）", 12, True, DARK, WD_ALIGN_PARAGRAPH.CENTER, 4)
para("2026 年 9 月 12 日", 11, False, MUT, WD_ALIGN_PARAGRAPH.CENTER)
pagebreak()

# ================= 目录 =================
para("目 录", 20, True, DARK, WD_ALIGN_PARAGRAPH.CENTER, 20)
for item in [
    "摘要",
    "第一章 评测背景与目标样例",
    "第二章 参赛 Skill 详述与安装方法",
    "  2.0 为什么是这 7 款（选型理由）",
    "  2.1 GordenPPTSkill（A1）",
    "  2.2 ppt-master（A7）",
    "  2.3 claude-pptx（A6，Anthropic 官方）",
    "  2.4 elite-ppt-pro（A5）",
    "  2.5 deck-dna（A3）",
    "  2.6 guizang-ppt-skill（A2）",
    "  2.7 codex-slides（A4）",
    "  2.8 参赛资格裁定汇总",
    "第三章 评测设计与方法",
    "第四章 评测流程",
    "第五章 素材准备",
    "第六章 测试内容",
    "第七章 测试结果",
    "  7.1 排行榜柱状图",
    "  7.2 四象限定位图",
    "  7.3 打分矩阵总表",
    "  7.4 图片复刻赛道逐套详解",
    "  7.5 纯文字赛道逐套详解",
    "  7.6 执行工具过程分（参考）",
    "第八章 评测工具缺陷与产物缺陷台账",
    "第九章 结论与推荐",
    "第十章 证据与复现索引",
    "附录 全部 14 套成品 PPT 截图拼接总览",
]:
    is_chapter = not item.startswith("  ")
    p = doc.add_paragraph()
    r = p.add_run(item)
    set_cn(r, 13 if is_chapter else 11, is_chapter, DARK if is_chapter else MUT)
    p.paragraph_format.space_after = Pt(6 if is_chapter else 3)
    if not is_chapter:
        p.paragraph_format.left_indent = Cm(1.0)
pagebreak()

# ================= 摘要 =================
heading("摘要", 1)
body("本评测以微信公众号文章《某股份制银行零售部2026年二季度重点工作运营方案》的 18 页 PPT 为基准，由达仔使用 ZCode（内置模型 GLM-5.3-Flash）从微信文章中完整下载并存档，包含 18 张原版页面渲染图与逐页识别的 1.2 万字文字大纲。7 款 PPT 生成 skill 参赛，统一在两条赛道下测试：")
body("纯文字赛道——仅提供文字大纲，考察从文字到成品的转化能力；图片复刻赛道——提供 18 张原版页面图，逐页多模态对照复刻，考察版式还原能力。两条赛道均硬性要求生成与原稿等量的 18 页原生可编辑 PPTX（文本框可修改、形状可编辑），整页截图贴入判负分。评测核心标准：「可编辑但丑 = 废物」。")
body("7 款 skill × 2 条赛道 = 14 套 18 页产物，共 252 张页面渲染图，全部经 LibreOffice 渲染为 PNG 后，由达仔以打散匿名盲评方式逐张打分（1–5 分制），252/252 全部评完无遗漏。最终名次与全部结论均以该盲评数据为准；执行工具的过程分仅作过程存档与交叉参考。")
body("终榜结论：GordenPPTSkill 以模板级质感与全场最高信息密度双冠——图片复刻 4.22、纯文字 4.39，综合 4.31/5，领先第二名 1.33 分；ppt-master 综合第二（复刻 3.00、文字 2.94，综合 2.97），是唯一自带硬质检门禁（零越界零重叠零图片贴片）且复杂图形全部矢量原生重绘的工程化路线；claude-pptx 综合 2.78 居第三。推荐前三名：GordenPPTSkill、ppt-master、claude-pptx。elite-ppt-pro 综合 2.47，以 0.03 分之差跌破 2.5 推荐线，不再推荐。后三名 deck-dna（1.69）、guizang-ppt-skill（1.50）、codex-slides（1.17）不推荐使用。", True)
pagebreak()

# ================= 第一章 =================
heading("第一章 评测背景与目标样例", 1)
heading("1.1 基准来源", 2)
body("评测基准取自以下微信公众号文章，该文章内嵌 18 页完整 PPT 内容图：")
body("文章链接：https://mp.weixin.qq.com/s/qWMxezxrJu3P_pe0aTCEhw", 12)
body("文章标题：《某股份制银行零售部2026年二季度重点工作运营方案》。评测执行方使用 Python 脚本自动下载全部 18 张页面渲染图（正则提取微信 CDN 图片链接，mmbiz.qpic.cn 域名），逐页保存为 PNG（分辨率 1080p），存档于 sample/images/，命名为 slide01–slide18.png；同时逐页识别全部文字与版式，形成约 1.2 万字文字大纲 sample/transcript_full.md，作为纯文字赛道的唯一输入。")
heading("1.2 目标样例 18 页拼接总览", 2)
img(os.path.join("eval", "assets", "sample_grid.png"), 15.2)
note("图 1  18 页目标样例拼接总览（原始图片：sample/images/slide01–18.png）")
heading("1.3 原稿特征分析", 2)
bullets([
    "配色体系：主色酱红 #B01F24 系（深红 #7E1714、正红 #B01F24、浅红 #C46662、粉色 #E0A5A2），辅色深灰 #1C1917，底色纯白，分割线浅灰 #E7DFDA。",
    "信息密度：每页 5–15 个信息块，单页文字量 200–500 字，典型中文咨询级高密度汇报风格。",
    "版式多样：KPI 卡片、三维复盘三角、目标仪表盘、客户分层金字塔、获客漏斗、旅程×渠道矩阵、财富工程三段递进、消费金融三引擎、中心辐射图、RACI 协同表、数字化四层架构、智慧服务三列、全员数字化五步、人才梯队四层、风险防火墙七关口、战役三月份列、90 天路线、管理闭环与风险响应矩阵。",
    "页面元素：每页带红色页码章（左上角方块内白色数字）与标题带。",
])
heading("1.4 评测目标", 2)
body("核心问题：当 7 款 PPT 生成 skill 分别以「纯文字大纲」和「原版 18 页图片」为输入时，谁能复刻出与原图更像、或比原图更好看的 18 页成品，且成品可在 PowerPoint 中编辑。美观是主赛道，可编辑是二筛，相似度是参考项。")

# ================= 第二章 =================
pagebreak()
heading("第二章 参赛 Skill 详述与安装方法", 1)
heading("2.0 为什么是这 7 款（选型理由）", 2)
body("7 款参赛 skill 按三条原则选定，合起来覆盖当前 AI 生成 PPT 的全部五条主流技术路线：")
bullets([
    "原则一 · 技术路线全覆盖：① 模板换字路线（GordenPPTSkill，中文市场最常见，「模板即成品」）；② HTML 网页 PPT 路线（guizang-ppt-skill 的瑞士国际主义风、deck-dna 的设计 DNA 驱动，美学体系强但原生非 PPTX）；③ 原生 PPTX 代码直出路线（elite-ppt-pro、claude-pptx，经 pptxgenjs 生成，可编辑性最好）；④ AI Studio 应用路线（codex-slides，应用级产品：AI 出图 + 场景工作流 + 版本管理）；⑤ SVG 中间格式工程化路线（ppt-master，手写 SVG 经官方转换器出原生 DrawingML，自带硬质检门禁）。",
    "原则二 · 出品方与热度代表性：claude-pptx 为 Anthropic 官方开源；elite-ppt-pro 来自 vercel-labs 官方 skills 仓库；guizang-ppt-skill 为国内知名 AI 博主歸藏出品；GordenPPTSkill、deck-dna 为社区快速上升的新秀；ppt-master 以 GitHub 53.4k star 为全场热度第一。",
    "原则三 · 热度入围：全网检索「什么 skill 做 PPT 最火」，ppt-master 以 53.4k star 居首，按同一评测流程入围（A7）。",
])
body("落选说明：PPTAgent（中科院 VILA 团队，约 3.7k star）曾列入候选。其为需要 litellm LLM 后端 + 参考 PPTX 模板的完整 Python 框架（服务级部署），不属于「可独立安装使用的 skill」，与本轮评测口径不符，放弃，不计入名次。", 11)
heading("2.1 GordenPPTSkill（A1）", 2)
body("GordenPPTSkill 是中文模板换字式 PPT 生成器，内置 21 套覆盖中文场景的成品模板（党政红、商务蓝、学术绿、教育橙等），核心纪律是「只换文字、不破排版」。工作流：读模板 INDEX 选定模板 → 读 detail.json（每页每个文本槽位的容量字段：chars_per_line / max_lines / max_chars）→ 撰写 edits.json（页选择 + 逐槽位文字替换）→ 运行 build_pptx.py 构建。构建自带出框检测（文字超出槽位宽度/高度时告警），默认不阻断保存。")
body("优势：模板本身即成品级设计，换字后即为可交付质量；原生 PPTX、文本框级可编辑；信息密度与原稿咨询风格高度匹配。劣势：版式由模板决定、无法按内容自适应；槽位容量与超密度文案不匹配时会出现重叠或截断，需要按出框检测反复打磨文字；无原生封面页的模板需以装饰页替代。")
subheading("来源与安装")
body("GitHub 仓库：github.com/GordenSun/GordenPPTSkill。安装：将 GordenPPTSkill 整个目录放入 Agent skills 目录（本机路径：~/.workbuddy/skills/GordenPPTSkill）。首次使用建议运行 scripts/apply_update.py 增量更新模板库（需 GitHub 连通，不可达时跳过不影响使用）。运行时依赖：python-pptx、LibreOffice（渲染预览）。")
heading("2.2 ppt-master（A7）", 2)
body("ppt-master（v6.3.2，GitHub 53.4k star，作者 hugohe3）是目前热度最高的 PPT 生成 skill，路线与其余各款完全不同：以 SVG 为中间格式——先按其 shared-standards-core 规范逐页手写 1280×720 的版式 SVG（根分组强制声明 data-pptx-bounds 版面区域且互不重叠、段落文本必须 text+tspan 结构、仅允许白名单内的着色与滤镜语法），再经官方转换器 svg_to_pptx.py 转成原生 DrawingML PPTX（形状即 PowerPoint 形状、文本即文本框，非图片贴片）。体系自带两道硬门禁：svg_quality_checker.py（版面越界/文字溢出/规范符合性检查）与导出后 postflight 校验，本次双赛道均 0 error 通过。")
body("优势：① 产物为原生可编辑 DrawingML（两套 18 页各约 2500 个形状、920+ 个文本框、0 张图片贴片）；② 版式还原能力强——漏斗、金字塔、仪表盘、RACI 徽章矩阵、环形循环、风险响应矩阵等复杂图形均为矢量重绘，信息密度高；③ 质检硬门禁保证零越界零重叠，工程纪律全场最严格。劣势：① 工作流极重（约 30 万字规范文档，学习与执行成本显著高于其他 skill）；② 输出为规整商务编辑风，模板质感不及 GordenPPTSkill 的成品模板；③ 照片类内容只能矢量简化重绘（其高保真 Image to PPTX 管线依赖 Codex 宿主的参考图生成能力，本机为 ZCode 宿主，按其规范改走确定性矢量重绘）。")
subheading("来源与安装")
body("GitHub 仓库：github.com/hugohe3/ppt-master。安装：GitHub 直连不可用，经 gh-proxy.com 镜像下载整仓（本机路径：eval/tools/ppt-master），无需 pip 安装（依赖 python-pptx / lxml 等常用库）。使用 Quick Generate / Image to PPTX 工作流：project_manager.py init 建项目 → import-sources 导入大纲/图片 → 逐页手写 SVG → svg_quality_checker.py 门禁 → svg_to_pptx.py 导出。")
heading("2.3 claude-pptx（A6，Anthropic 官方）", 2)
body("claude-pptx 是 Anthropic 官方开源的 document-skills:pptx skill，以 pptxgenjs 为主的原生 PPTX 生成（13.33×7.5 英寸宽幅画布）。核心纪律：结论式标题（不写话题）、每页一个视觉焦点、原生可编辑图表（addChart）、禁用装饰条与 AI 味排版（禁用 edge stripe、下划线、emoji、纯色封面），并附溢出/重叠的程序化 QA 要求。")
body("优势：最专业稳健，原生可编辑图表（双击可改数据），零溢出零截断零重叠。劣势：信息密度主动降低（skill 认为高密度影响美观），多页留白偏多，金字塔/仪表盘/循环等图形元素简化为文字条。")
subheading("来源与安装")
body("GitHub 仓库：github.com/anthropics/skills（document-skills/pptx）。安装：ZCode 官方插件内置（.zcode/cli/plugins/cache/document-skills/），随插件缓存自动加载；运行时依赖 pptxgenjs（npm 安装）。")
heading("2.4 elite-ppt-pro（A5）", 2)
body("elite-ppt-pro 是顶级咨询风格 PPT 生成器（vercel-labs/skills 仓库 PR #813），内置 7 大咨询风主题 × 38 种版式与代码模板，pptxgenjs 直出原生 PPTX。核心纪律：标题即结论（不写话题）、数据先行（先收集 15+ 数据点再动手）、每页 ≥4 个内容区、至多 2 个强调色、来源标注（每个数据点必须有来源）、双输出（PPTX + HTML）。")
body("优势：零溢出零截断零重叠，纪律性全场最强。劣势：视觉平淡——大量纯色块 + 文字为主，封面无设计元素，卡片下半空置。")
subheading("来源与安装")
body("GitHub 来源：vercel-labs/skills 仓库 PR #813。安装：将 elite-ppt-pro 目录放入 skills 目录（本机路径：~/.workbuddy/skills/elite-ppt-pro）；运行时依赖 pptxgenjs（npm install pptxgenjs）。")
heading("2.5 deck-dna（A3）", 2)
body("PPT-Design-DNA（V3.1）是设计 DNA 驱动的 HTML 演示文稿生成器。核心理念：先从参考图提取五层设计 DNA（Mood / Composition / Visual / Content Strategy / Presentation），形成可复用的设计档案，再通过规格驱动（Page Specs + Layout Guard）生成 1920×1080 定舞台 HTML deck。附带静态布局守卫脚本 ppt-layout-guard.js（行高、孤行、区域碰撞、导航安全区等 48 项指标）。")
body("优势：数据保真度全场最高（指标零失真），红白编辑风与原稿色系一致，守卫脚本保证无溢出。劣势：原生输出为 HTML，无 PPTX 产出路径；版面偏素，视觉冲击力有限。")
subheading("来源与安装")
body("GitHub 仓库：github.com/dakjdakd/PPT-Design-DNA。安装：GitHub 直连不可用，经 gh-proxy.com 镜像下载 zip 解压（本机路径：eval/tools/PPT-Design-DNA），无需安装依赖即可运行守卫脚本（需 Node.js）。")
heading("2.6 guizang-ppt-skill（A2）", 2)
body("guizang-ppt-skill 由歸藏出品，生成瑞士国际主义风格的横向翻页网页 PPT（单文件 HTML），自带 22 个登记版式（S01–S22）、5 套杂志风 + 4 套瑞士风主题色预设、Motion One 动效与校验脚本 validate-swiss-deck.mjs。核心美学：大字号对比（标题与正文比 ≥ 8:1）、单一锚点色（IKB 蓝 / 柠檬黄 / 柠檬绿 / 安全橙）、直角纯色、无渐变无阴影无圆角。")
body("优势：体系内美学极强，大字账单页与时间线页视觉冲击力出众。劣势：原生输出为单文件 HTML（非 PPTX），需要额外转换；瑞士橙色体系与目标原稿的红色商务风格气质冲突；种子模板缺少 21 个登记版式的 CSS 类（含 ledger-row、brief-grid、matrix-fill、three-forces 等），需按 SKILL.md 规范手工补入 style 块，否则对应版式会散架。")
subheading("来源与安装")
body("GitHub 仓库：github.com/op7418/guizang-ppt-skill。安装：将 skill 目录放入 skills-marketplace 目录（本机路径：~/.workbuddy/skills-marketplace/skills/guizang-ppt-skill）。运行时依赖：无（浏览器直接打开 HTML）。")
heading("2.7 codex-slides（A4）", 2)
body("codex-slides 是 nexu-io 出品的 AI slide studio（Next.js 应用 + Electron + MCP），支持场景工作流、品牌系统、版本管理与 PPTX/PDF 导出。核心架构：出图由 AI 逐页生成（硬编码依赖 chatgpt.com 后端 + ChatGPT OAuth 登录态），文字由 AI 或本地 CLI 生成（依赖 codex CLI 的 spawn）。")
body("本机阻塞详情：① AI 出图管线硬编码依赖 chatgpt.com 后端 + ChatGPT OAuth（本机网络不可达且无凭据）；② 文字管线经本地 codex CLI 时 spawn 报 EINVAL（Windows 新版 Node 禁止直接 spawn .cmd 文件）；③ 确定性 REST 路径可建项目与写入大纲（已验证成功），但空页导出 PPTX 返回 500。证据文件存于 eval/A4_codex-slides/。本次产物为按评测规则以 pptxgenjs 原生补全的 18 页。")
subheading("来源与安装")
body("GitHub 仓库：github.com/nexu-io/codex-slides。安装：GitHub 直连不可用，经 gh-proxy 镜像下载 zip；本机完成 npm install + next build（需设 ELECTRON_SKIP_BINARY_DOWNLOAD=1），运行 next start -p 4311 启动服务。完整运行需 chatgpt.com 可达 + ChatGPT OAuth 登录。")
heading("2.8 参赛资格裁定汇总", 2)
table(["Skill", "原生形态", "本次参赛形态", "资格"], [
    ["GordenPPTSkill", "原生 PPTX（模板换字）", "原生 PPTX", "参赛"],
    ["ppt-master", "SVG → 原生 DrawingML", "原生 PPTX（官方转换器直出）", "参赛"],
    ["claude-pptx", "原生 PPTX（pptxgenjs）", "原生 PPTX", "参赛"],
    ["elite-ppt-pro", "原生 PPTX（pptxgenjs）", "原生 PPTX", "参赛"],
    ["deck-dna", "HTML（网页 PPT）", "pptxgenjs 原生重写（红白编辑风）", "参赛（标注转换）"],
    ["guizang-ppt-skill", "HTML（网页 PPT）", "pptxgenjs 原生重写（瑞士体系）", "参赛（标注转换）"],
    ["codex-slides", "AI 图片页 + Electron", "本机管线阻塞，pptxgenjs 原生补全", "降级参赛（补全）"]],
    widths=[3.4, 3.4, 4.6, 2.4])
body("注：A2/A3 的转换产物由操作者以 pptxgenjs 重写，设计体系忠实于原 skill，但转换效率与原 skill 的 HTML 工作流不可比。", 10, False, MUT)
pagebreak()

# ================= 第三章 =================
heading("第三章 评测设计与方法", 1)
heading("3.1 双赛道设计", 2)
table(["赛道", "输入", "输出要求", "考察点"],
      [["纯文字赛道", "18 页内容识别后的文字大纲（约 1.2 万字，sample/transcript_full.md）", "18 页原生可编辑 PPTX，覆盖原稿全部核心主题", "文字→结构化成片的转化能力、信息组织、视觉表达"],
       ["图片复刻赛道", "18 张原版页面图（多模态逐页查看）", "18 页原生可编辑 PPTX，1:1 对应原稿页序与主题", "版式还原、风格相似度、信息保真、多模态理解"]],
      widths=[2.6, 5.6, 4.2, 3.4])
heading("3.2 硬性规则", 2)
bullets([
    "页数铁律：两条赛道均须 18 页整，1:1 对应原稿页序与主题，不许合并、拆分或缩水。初版不足 18 页的产物全部补齐重写。",
    "交付物铁律：必须是原生可编辑 PPTX（python-pptx / pptxgenjs / 原生 DrawingML 转换器产出，文本框可改）。整页截图贴入 PPTX 判负分。",
    "多模态铁律：图片复刻赛道每页生成前必须重新查看对应原图，按所见复刻。",
], 11)
heading("3.3 评审口径与评分规则", 2)
para("评审分两层。第一层为执行工具（ZCode/GLM-5.3-Flash）逐页审阅全部渲染图，按 10 分制打分，权重为美观 60% + 相似度/覆盖 25% + 缺陷 15%，仅作过程存档与交叉参考。第二层为人工盲评：252 张页面图全部打散、以随机五字符盲码匿名（如 BYQGD-07），达仔逐张按 1–5 分打分，252/252 全部评完无遗漏，评分实时存档并导出 CSV。最终排名以盲评分为准。", 12)
para("美观评分的主赛道权重最高，体现「可编辑但丑 = 废物」的核心评测标准。可编辑性全员达标（14 套产物均为原生可编辑 PPTX），不再作为区分项单独计分。", 11)
heading("3.4 排名口径", 2)
para("图片复刻赛道总分 = 美观 60% + 相似度 25% + 缺陷 15%；纯文字赛道总分 = 美观 55% + 内容覆盖 25% + 缺陷 15% + 结构完整 5%（仅用于执行工具过程分）。人工盲评榜单 = 各套 18 页得分的算术平均；综合排名 = 两赛道均分的算术平均。", 11)

# ================= 第四章 =================
pagebreak()
heading("第四章 评测流程", 1)
body("整个评测按以下 9 个步骤执行：")
for tag, name, desc in [
    ("Step 1", "基准采集", "Python 脚本下载微信文章全文 HTML（3.3MB），正则提取全部 18 张 mmbiz CDN 图片链接并逐页下载存档为 PNG（1080p）；逐页多模态识别全部文字与版式，产出约 1.2 万字逐页转写与合并大纲。"),
    ("Step 2", "Skill 就位", "7 款 skill 全部安装/下载/构建就绪：GordenPPTSkill / guizang / elite-ppt-pro / claude-pptx 为本地已有；deck-dna 与 ppt-master 经 gh-proxy 镜像下载；codex-slides 经镜像下载并完成 npm install + Next.js build。"),
    ("Step 3", "文字赛道生成", "以 sample/transcript_full.md 为唯一输入，分别驱动 7 款 skill 产出完整 18 页 deck；初版不足 18 页的（A1 11 页、A5 10 页、A6 13 页）全部补齐重写至 18 页。"),
    ("Step 4", "复刻赛道生成", "逐页查看 18 张原版图片后多模态复刻，每套 18 页 1:1 对应原稿页序；初版仅做 6 页的全部重做为 18 页完整版。"),
    ("Step 5", "渲染核验", "全部 14 套产物统一经 LibreOffice（pptx→PDF→PNG）或 Chromium（HTML→PNG）渲染为图片；python-pptx 程序化核验页数（14 套 × 18 页全过）。发现并修复评测工具两起缺陷（A2 截图未翻页、A3 视口裁切）与 A7 封面装饰组缺失。"),
    ("Step 6", "执行工具评审", "ZCode/GLM-5.3-Flash 逐页审阅渲染图并打 10 分制过程分。该分仅作过程存档与交叉参考，不参与排名。"),
    ("Step 7", "盲评工作台", "搭建浏览器打分台（eval/workbench/index.html）：252 张图打散、随机五字符盲码匿名、进度存档/缩略图总览/CSV/JSON 导出、一键揭盲。"),
    ("Step 8", "人工盲评", "达仔逐张 1–5 分，252/252 全部评完，导出 CSV（存档 eval/盲评结果_252.csv）。打分期间不看任何盲码对应关系。"),
    ("Step 9", "揭盲统计与报告", "揭盲对应关系，计算均分与排名，撰写终版报告。"),
]:
    p = doc.add_paragraph()
    r = p.add_run(tag + "　" + name + "　"); set_cn(r, 11.5, True, RED)
    r2 = p.add_run(desc); set_cn(r2, 11)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = 1.4

# ================= 第五章 =================
heading("第五章 素材准备", 1)
body("评测涉及的全部素材按用途分类如下：")
table(["素材", "内容与规格", "位置"],
      [["18 页基准原图", "slide01–18.png，1080p 渲染，盲评与复刻的对照基准", "sample/images/"],
       ["文章原始 HTML", "微信文章全文存档（3.3MB），含原始图片链接", "sample/article.html"],
       ["图片链接清单", "18 张 mmbiz 图直链列表", "sample/image_urls.txt"],
       ["逐页文字识别", "18 页逐页文字与版式识别结果", "sample/text/"],
       ["合并文字大纲", "文字赛道唯一输入（约 1.2 万字）", "sample/transcript_full.md"],
       ["18 页拼接总览图", "3×6 拼接（1968×2238），报告与 PPT 素材", "eval/assets/sample_grid.png"],
       ["排行榜柱状图", "双赛道均分分组柱状图（matplotlib，STKaiti）", "eval/assets/bar_ranking.png"],
       ["四象限定位图", "复刻×文字双维象限图", "eval/assets/quadrant.png"],
       ["14 套成品拼图", "每套 3×6 十八页拼接（PIL 生成）", "eval/assets/collages/"],
       ["盲评工作台", "252 张匿名图 + 浏览器打分页（localStorage 存档、CSV/JSON 导出、揭盲功能）", "eval/workbench/"],
       ["盲评原始 CSV", "盲评结果_252.csv", "252 行评分数据（14 套 × 18 页）"],
       ["渲染工具链", "render_pptx.py（pptx→PNG）、render_html.js / render_html_stack.js（HTML→PNG）", "eval/ 根目录"],
       ["页数核验脚本", "python-pptx 程序化核验 14 套 deck 页数", "eval/verify18.py、verify_pptm.py"],
       ["ppt-master 框架", "官方整仓（53.4k star），含质检器/转换器/工作流规范", "eval/tools/ppt-master/"],
       ["ppt-master 项目", "SVG 源文件、质检报告、导出记录、重建清单", "eval/tools/ppt-master/projects/"]],
      widths=[3.6, 6.6, 5.0])
body("渲染工具链说明：HTML deck 截图必须直接驱动 deck.style.transform（CDN 被断时键盘翻页绑定失效）；截图视口必须 ≥1920×1080，否则出现裁切错位。", 10, False, MUT)

# ================= 第六章 =================
pagebreak()
heading("第六章 测试内容", 1)
heading("6.1 纯文字赛道", 2)
body("统一输入 sample/transcript_full.md（覆盖原稿 18 页全部核心主题：经营复盘、关键挑战、经营主线 1-2-3、目标体系、客户分层、生态获客、服务体验、财富工程、消费金融、协同联动、数字化基座、智慧服务、全员数字化、人才梯队、风险防火墙、运营战役、90 天路线、管理闭环）。各 skill 自行决定信息取舍与版式，产出 18 页成品。考察：内容覆盖度、美观专业度、结构完整度与缺陷率。")
heading("6.2 图片复刻赛道", 2)
body("统一输入 18 张原版页面图。要求逐页 1:1 复刻：页序对应原稿页序、页面标题与数据一致、版式结构参照原图（多模态）。产出 18 页成品。考察：版式还原度、信息保真度、风格相似度。")
heading("6.3 各 Skill 实际生成内容", 2)
table(["Skill", "文字赛道产出", "复刻赛道产出"],
      [["GordenPPTSkill", "premium-corp/mckinsey 模板 18 页：装饰封面、复盘维恩、挑战木桶、主线阶梯、目标统计、分层金字塔、获客漏斗、服务五段、财富递进、消金箭头、协同双S、数字化PDCA、智慧三段重点工作、全员五段流程、人才两段、风控六段、战役甘特/S环、90天、闭环无穷环", "mckinsey-style 模板 18 页 1:1 对照"],
       ["ppt-master", "正红编辑风 SVG 18 页：丝带封面、KPI+三角复盘、神庙主线、目标仪表盘、分层金字塔表、获客漏斗、旅程×渠道矩阵、财富三面板、消金三列、协同辐射+RACI、数据四层、智慧三列、能力阶梯+工具箱、人才五区、风控七关口、月份战役、90 天旗帜、管理闭环+风险矩阵", "正红复刻 18 页：同上结构与原稿逐页对位（封面丝带+天际线构图，正红色系贴近原稿）"],
       ["claude-pptx", "pptxgenjs 宽幅 18 页", "pptxgenjs 宽幅 18 页逐页对照"],
       ["elite-ppt-pro", "pptxgenjs 18 页", "pptxgenjs 18 页逐页对照（v2 满版）"],
       ["deck-dna", "红白编辑风 18 页 HTML → pptxgenjs 18 页", "红白编辑风 18 页逐页对照"],
       ["guizang-ppt-skill", "瑞士风 18 页 HTML → pptxgenjs 18 页：封面、复盘对照、主线123、目标账单、三引擎、12周条形、战役六卡、90天时间线、数字化矩阵、收束宣言", "瑞士风 18 页逐页对照（标注对照原图页码）"],
       ["codex-slides", "midnight 深色风 18 页（原生补全）", "midnight 深色风 18 页逐页对照（原生补全）"]],
      widths=[3.4, 6.4, 5.4])

# ================= 第七章 =================
pagebreak()
heading("第七章 测试结果", 1)
body("以下所有分数均为达仔对 252 张渲染图的盲评实测均分（5 分制）：14 套产物 × 18 页 = 252 张全部逐张打完，无遗漏。")
heading("7.1 排行榜柱状图", 2)
img(os.path.join("eval", "assets", "bar_ranking.png"), 16.0)
note("图 2  双赛道盲评均分排行榜（红色柱 = 图片复刻赛道，深灰柱 = 纯文字赛道；金色虚线 = 推荐线 2.5 分）")
heading("7.2 四象限定位图", 2)
img(os.path.join("eval", "assets", "quadrant.png"), 14.0)
note("图 3  四象限定位：双优区三家——GordenPPTSkill 领跑，ppt-master 越线入围，claude-pptx 压线")
heading("7.3 打分矩阵总表", 2)
table(["Skill", "图片复刻", "纯文字", "综合", "结论"],
      [["GordenPPTSkill", "4.22", "4.39", "4.31", "推荐 双冠"],
       ["ppt-master", "3.00", "2.94", "2.97", "推荐 第二"],
       ["claude-pptx", "2.72", "2.83", "2.78", "推荐 第三"],
       ["elite-ppt-pro", "2.39", "2.56", "2.47", "跌破推荐线，不再推荐"],
       ["deck-dna（转换）", "1.39", "2.00", "1.69", "不推荐"],
       ["guizang-ppt-skill（转换）", "1.39", "1.61", "1.50", "不推荐"],
       ["codex-slides（补全）", "1.17", "1.17", "1.17", "不推荐"]],
      widths=[4.6, 2.2, 2.2, 2.2, 3.6])
heading("7.4 图片复刻赛道逐套详解", 2)
subheading("7.4.1 GordenPPTSkill（均分 4.22）")
body("全场唯一做到「模板级好看 + 咨询级密度 + 原生可编辑」的产物。维恩图复盘页与原 P02 四 KPI 卡布局同构；3D 金字塔分层页与原 P05 金字塔+四列表格同构；无穷环闭环页与原 P18 管理闭环神似；漏斗生态页与原 P06 漏斗结构对应。信息密度全场最高。缺陷：漏斗页个别竖排标签乱序、装饰字母 J/L/S/Z 残留、少数标签换行局促——经两轮修复后 13 处缺陷清零，残留均为装饰级。")
subheading("7.4.2 ppt-master（均分 3.00）")
body("复刻赛道唯一在 GordenPPTSkill 之外站上 3 分档的产物。强项：官方质检门禁保证零越界零重叠、零图片贴片，漏斗、金字塔、仪表盘、RACI 徽章矩阵、环形循环、风险响应矩阵全部矢量原生重绘；正红色系贴近原稿，封面丝带+天际线构图按原图重做；原生 DrawingML 文本框/形状级可编辑。失分点：无成品级模板质感，版式规整但视觉冲击力一般——与第一名差距 1.22 分，印证「模板级好看」仍是硬通货。")
subheading("7.4.3 claude-pptx（均分 2.72）")
body("逐页对位最准：页码章+标题带+四维目标块+RACI 表+风险矩阵全对应；零溢出零截断零重叠；带原生可编辑柱状图。但信息密度主动降低是设计取舍，多页留白偏多，金字塔/仪表盘/循环等图形元素简化为文字条。整体评价：规范稳健但不够惊艳。")
subheading("7.4.4 elite-ppt-pro（均分 2.39）")
body("v2 满版重写后空洞感消除：分层四列表格对照原 P05、5×6 服务矩阵对照原 P07、RACI 表对照原 P10、进度条可视化增强。零溢出零截断零捏造。但视觉表现力仍弱：大量纯色块+文字为主，封面无设计元素。")
subheading("7.4.5 deck-dna 转换版（均分 1.39）")
body("红白编辑风与原稿色系一致、数据零失真。分层表对照原 P05 的还原度较高。但版面偏素、留白多，页码标注有小偏移。原生形态为 HTML，本次以 pptxgenjs 规格驱动重写转换。")
subheading("7.4.6 guizang-ppt-skill 转换版（均分 1.39）")
body("瑞士黑金账单页体系内惊艳，时间线页干净利落。但橙色体系与红色原稿气质冲突；部分页稀疏（右列悬空、约 70% 页面空白）；部分卡片文字对比度不足。原生形态为 HTML。")
subheading("7.4.7 codex-slides 补全版（均分 1.17）")
body("自身 AI 管线在 Windows/无外网环境完全不可用（出图硬绑 chatgpt.com、文字管线 spawn 受阻、导出报错），按评测规则以 pptxgenjs 原生补全 18 页。midnight 深色风面板化版式统一但缺乏变化。垫底的原因是补全产物本身的信息密度与视觉质量均远低于其他 skill 的直接产出。")
heading("7.5 纯文字赛道逐套详解", 2)
para("GordenPPTSkill（4.39）：premium-corp 模板 18 页。模板质感+密度最高，早期 9 处重叠/截断已修复；无原生封面页（模板限制）。", 11)
para("ppt-master（2.94，第二）：正红编辑风 SVG 18 页。统一页码章+标题带版式系统，KPI 卡、神庙主线图、目标仪表盘、获客漏斗、旅程×渠道矩阵、管理闭环与风险矩阵等图形齐全，密度高而整；短板是视觉冲击力不及成品模板，执行成本全场最高。", 11)
para("claude-pptx（2.83）：宽幅 13.33×7.5 英寸 18 页。结论式标题+原生图表+旅程矩阵+RACI 表，专业稳健；密度主动降低是设计取舍。", 11)
para("elite-ppt-pro（2.56）：18 页，零溢出零截断零捏造；视觉平淡、卡片下半空置（后经 v2 满版重写改善）。", 11)
para("deck-dna 转换版（2.00）：红白编辑风 18 页 HTML。数据零失真；版面偏素、卡片下部留白偏多。", 11)
para("guizang-ppt-skill 转换版（1.61）：瑞士风 18 页 HTML。大字时间线/账单页出色；橙灰色系偏离原稿。", 11)
para("codex-slides 补全版（1.17）：midnight 深色风 18 页。面板化版式统一但缺乏变化；自身管线阻塞，垫底。", 11)
heading("7.6 执行工具过程分（参考，非最终）", 2)
body("执行工具（ZCode/GLM-5.3-Flash）曾对产物逐页审阅打 10 分制过程分。以 ppt-master 为例：美观 8.4–8.5、相似度/覆盖 8.8–9.0、缺陷控制 8.6，折 5 分制约 4.3，显著高于盲评实分 2.97——过程分在绝对值上系统性偏宽，但其「ppt-master 可进推荐席前二」的名次判断与盲评终审一致。过程分仅可用于横向相对比较，不可与盲评分直接比较。完整记录见 7.6 附表与 eval/A7_pptmaster/。", 11)
img(os.path.join("eval", "assets", "a7_process_scores.png"), 15.5)
note("图 4  ppt-master 执行工具过程分（10 分制；非盲评分，不可与表 7.3 直接比较）")

# ================= 第八章 =================
pagebreak()
heading("第八章 评测工具缺陷与产物缺陷台账", 1)
heading("8.1 评测工具误判（已修复）", 2)
body("评测过程中发现并修复了评测工具自身的两起缺陷，曾导致早期报告出现误判：")
bullets([
    "A2 截图未翻页：HTML deck 的键盘翻页绑定依赖 CDN 加载（Motion One 库从 CDN 加载，本机不可达时绑定失效），导致 10 张截图全是封面。修复：直接驱动 deck.style.transform 跳页。",
    "A3 视口裁切：HTML deck 画布 1920px 宽，截图视口只有 1600px，导致渲染图被裁切错位。修复：视口调整为 1960×1200 并使用元素级截图。",
    "两起均为评测工具问题而非产物问题。修复后相关页面全部重新渲染并重新评审。",
], 11)
heading("8.2 产物缺陷台账（全部已修复并重新渲染）", 2)
table(["Skill", "缺陷描述", "修复方式"],
      [["A1 复刻", "副标题与6%大字重叠；桶签竖排乱序；底部截断×2；标题孤字AUM；错别字「六月冲冲刺」；漏映射的智慧服务页标题；漏斗页标题孤字「营」；残留槽位「团队协作低效」", "缩短副标题使6%语义成立；桶签改2字；截断文本重写入槽；标题槽改「高净值」；错别字更正；补映射缺失槽位"],
       ["A1 文字", "复盘页副标题压字；木桶页桶签竖排乱序+底部截断×2；主线页标题孤字「系」+面板压字；金字塔中央标签方向颠倒", "同口径修复：缩短副标题、桶签2字化、标题改写、截断文本重写"],
       ["A5 复刻", "P07–P11 五页底部区块与来源行碰撞（布局溢出）；v1 卡片下半空置", "布局压缩（5处y坐标调整）+ v2 满版重写"],
       ["A3 转换", "构建脚本数组嵌套错误导致目标页丢内容；KPI 卡片塌缩（.kpi 类 position:absolute 在 grid 容器中）", "整脚本规格化重写（build_a3_v2.js），通过 node --check 与渲染核验"],
       ["A2 转换", "种子模板缺 21 个登记版式 CSS 类；timeline 构建参数错", "按 skill 规范补 CSS 到模板 style 块 + 修参重渲染"],
       ["A7 双赛道", "封面底部丝带/天际线装饰组被构建脚本重构漏加（渲染页整块缺失）；≥95%/75%/15%/+5/AUM 等 ASCII 密集文本在 LibreOffice 下折行", "补回装饰组调用并重新导出；窄空格（U+2009）加固文本帧后重新导出，官方门禁复检 0 error"],
       ["评测工具", "A2 截图未翻页（CDN 断→键盘绑定失效）；A3 视口裁切（1600 < 1920）", "render_html.js 改为直接驱动 transform；render_html_stack.js 视口改 1920"]],
      widths=[2.6, 7.0, 5.5])
heading("8.3 方法局限", 2)
bullets([
    "A4 codex-slides 的排名反映的是「合规补全产物」的水平——自身 AI 管线在本机完全不可用，在线完整能力未测。",
    "A2/A3 的转换产物由操作者以 pptxgenjs 重写，设计体系忠实于原 skill，但转换效率与原 skill 的 HTML 工作流不可比（操作者编写时间未计入 skill 评测）。",
    "美观评分含主观成分；已用「打散匿名 + 双层评审 + 逐页台账」把主观性压到最低，但无法完全消除。",
    "文字赛道页数下限为 18 页；部分 skill 的 13 页版本可能在信息密度上更有优势——18 页扩展过程中不可避免地引入了填充性内容。",
], 10.5)
pagebreak()

# ================= 第九章 =================
heading("第九章 结论与推荐", 1)
body("基于达仔 252 张盲评实测数据（14 套 × 18 页）与执行工具逐页审阅的交叉验证，最终结论如下：")
heading("9.1 推荐名单（前三名）", 2)
subheading("综合第一名：GordenPPTSkill（综合 4.31/5）")
body("盲评双赛道双冠：图片复刻 4.22、纯文字 4.39，均为第一（领先第二名 1.22–1.45 分）。推荐理由：① 酱红咨询模板与目标原稿的红色商务审美同源，「模板级好看」是全场唯一的模板质感产出；② 信息密度全场最高，最接近原稿的咨询级密度；③ 金字塔/漏斗/闭环等核心图形与原图同构，不是简单的文字条替换；④ 原生 PPTX 文本框级可编辑，继承模板主题母版。注意事项：模板槽位容量与超密度文案不匹配时会出现重叠/截断，需要按出框检测反复打磨文字（本轮共修复 13 处），这是该路线的必要工序而非可选步骤。", 11)
subheading("综合第二名：ppt-master（综合 2.97/5）")
body("纯文字 2.94（第二）、图片复刻 3.00（第二）。推荐理由：① 7 款中唯一自带硬质检门禁（svg_quality_checker 最终门禁 + 导出 postflight，双赛道 0 error），零越界零重叠、零图片贴片；② 复杂图形矢量还原厚度全场第二——漏斗、金字塔、仪表盘、RACI 徽章矩阵、环形循环、风险响应矩阵均为原生形状重绘，不是文字条凑数；③ 原生 DrawingML 输出，文本框/形状级可编辑。注意事项：执行成本全场最高（规范体系极重，单套 18 页工时显著高于其他 skill）；无成品级模板质感。", 11)
subheading("综合第三名：claude-pptx（综合 2.78/5）")
body("纯文字 2.83（第三）、图片复刻 2.72（第三）。推荐理由：① 零溢出零截断零重叠；② 原生可编辑图表（双击可改数据）是独有优势；③ 逐页对位最准（页码章+标题带+四维目标块+RACI+矩阵全对应）。注意事项：信息密度主动降低是设计取舍，多页留白偏多；图形元素简化为文字条。", 11)
body("上榜线说明：第四名 elite-ppt-pro 综合 2.47（复刻 2.39、文字 2.56），以 0.03 分之差跌破 2.5 推荐线，退出推荐席，列为中性备选：零缺陷纪律仍在，但在「可编辑但丑 = 废物」的标准下，美观主赛道的短板是硬伤。", 11)
heading("9.2 不推荐名单（后三名）", 2)
table(["名次", "Skill", "综合", "不推荐理由"], [
    ["5", "deck-dna（转换）", "1.69", "原生形态为 HTML、无 PPTX 产出路径；转换版依赖操作者重写；版面偏素、留白多"],
    ["6", "guizang-ppt-skill（转换）", "1.50", "原生形态为 HTML；瑞士橙体系与红色商务原稿气质冲突；复刻均分仅 1.39"],
    ["7", "codex-slides（补全）", "1.17", "自身 AI 管线在 Windows/无外网环境完全不可用（出图硬绑 chatgpt.com、spawn 受阻、导出报错）；补全产物双赛道均垫底"]],
    widths=[1.2, 4.2, 1.8, 8.0])
para("注：第 5、6 名的复刻均分并列 1.39（deck-dna 与 guizang），综合排序由文字赛道拉开；两者均为转换产物，排名差异对实际选型无影响。", 10, False, MUT)
heading("9.3 组合使用建议", 2)
bullets([
    "综合最强、要「最好看且能编辑」→ GordenPPTSkill（A1）：模板质感 + 咨询级密度，双赛道双冠。",
    "要「版面零越界零重叠的硬保证 + 复杂图形原生重绘」→ ppt-master（A7）：唯一自带质检门禁的工程化路线，代价是全场最高的执行成本。",
    "文字大纲快速出规范骨架 → claude-pptx（A6）：从大纲到成片稳、快、规范，原生可编辑图表。",
    "对视觉不满意的页面 → 用 GordenPPTSkill 的模板体系重做外观，逐页换字；A1 / A7 / A6 产物均为原生 PPTX，可在 PowerPoint 中互相拷页。",
    "如果需要网页 PPT（非 PPTX 场景），A2 的瑞士风和 A3 的编辑风仍有体系内价值，但不在本次「可编辑 PPTX」评测范畴内。",
], 11)
heading("9.4 一句话结论", 2)
p = doc.add_paragraph()
r = p.add_run("7 款终榜：GordenPPTSkill 以 4.31 双冠断层领先；ppt-master 以 2.97 摘得综合第二（硬质检门禁 + 复杂图形原生重绘）；claude-pptx 2.78 第三。elite-ppt-pro 2.47 跌破推荐线不再推荐；后三名 deck-dna（1.69）、guizang-ppt-skill（1.50）、codex-slides（1.17）不推荐。")
set_cn(r, 13, True, DARK)
pagebreak()

# ================= 第十章 =================
heading("第十章 证据与复现索引", 1)
table(["证据/工具", "位置", "说明"],
      [["14 套 18 页 PPTX 产物", "eval/{skill}/{mode}/deck18*.pptx", "全部 18 页原生可编辑"],
       ["全部渲染 PNG", "同目录 render18/ 子目录", "LibreOffice / Chromium 渲染"],
       ["逐页亲审笔记", "eval/review_notes.md", "含逐页打分与修复记录"],
       ["页数核验脚本", "eval/verify18.py、verify_pptm.py", "python-pptx 程序化核验"],
       ["渲染工具链", "eval/render_pptx.py、render_html*.js", "pptx→PNG、HTML→PNG"],
       ["A4 阻塞证据链", "eval/A4_codex-slides/", "generate_sse.txt、project.json 等"],
       ["盲评工作台", "eval/workbench/index.html", "可清空复用"],
       ["盲评原始 CSV", "eval/盲评结果_252.csv", "252 行评分数据（14 套 × 18 页）"],
       ["盲码映射", "eval/workbench/answers.js", "盲码→真实身份对应"],
       ["基准与文字识别", "sample/", "18 张原图 + 逐页转写 + 合并大纲"],
       ["样例拼接图", "eval/assets/sample_grid.png", "3×6 拼接（1968×2238）"],
       ["图表素材", "eval/assets/bar_ranking.png、quadrant.png、a7_process_scores.png", "matplotlib 生成"],
       ["14 套成品拼图", "eval/assets/collages/*.png", "每套 3×6 十八页拼接"],
       ["ppt-master 框架与项目", "eval/tools/ppt-master/", "质检报告 validation/、导出 exports/、重建清单 analysis/"],
       ["历史草稿存档", "eval/report_superseded_drafts.md", "含被更正的早期结论"]],
      widths=[4.5, 6.5, 4.2])
body("渲染注意：HTML deck 截图须直接驱动 deck.style.transform（CDN 被断时键盘翻页绑定失效）；截图视口必须 ≥1920×1080，否则出现裁切错位。", 10, False, MUT)

# ================= 附录 =================
pagebreak()
heading("附录 全部 14 套成品 PPT 截图拼接总览", 1)
body("以下为 7 款 skill × 2 条赛道 = 14 套 18 页成品的全量渲染截图拼接图。每套按原稿页序 01–18 排列（3 列 × 6 行），渲染分辨率 1080p；盲评与执行工具评审均以这些渲染图为唯一评审对象。图片源文件见 eval/assets/collages/。")
_collages = [
    ("A1 · GordenPPTSkill — 纯文字赛道（盲码 BYQGD）", "A1_mode1_collage.png"),
    ("A1 · GordenPPTSkill — 图片复刻赛道（盲码 AW9ZV）", "A1_mode2_collage.png"),
    ("A2 · guizang-ppt-skill（转换版）— 纯文字赛道（盲码 YW957）", "A2_mode1_collage.png"),
    ("A2 · guizang-ppt-skill（转换版）— 图片复刻赛道（盲码 7SCGZ）", "A2_mode2_collage.png"),
    ("A3 · deck-dna（转换版）— 纯文字赛道（盲码 CSRUZ）", "A3_mode1_collage.png"),
    ("A3 · deck-dna（转换版）— 图片复刻赛道（盲码 VBYVC）", "A3_mode2_collage.png"),
    ("A4 · codex-slides（补全版）— 纯文字赛道（盲码 M4ZUS）", "A4_mode1_collage.png"),
    ("A4 · codex-slides（补全版）— 图片复刻赛道（盲码 4BR89）", "A4_mode2_collage.png"),
    ("A5 · elite-ppt-pro — 纯文字赛道（盲码 TDSRC）", "A5_mode1_collage.png"),
    ("A5 · elite-ppt-pro — 图片复刻赛道（盲码 EHEVD，v2 满版）", "A5_mode2_collage.png"),
    ("A6 · claude-pptx — 纯文字赛道（盲码 HHMF2）", "A6_mode1_collage.png"),
    ("A6 · claude-pptx — 图片复刻赛道（盲码 M2E8D）", "A6_mode2_collage.png"),
    ("A7 · ppt-master — 纯文字赛道（盲码 FZHT8）", "A7_mode1_collage.png"),
    ("A7 · ppt-master — 图片复刻赛道（盲码 4RDJM）", "A7_mode2_collage.png"),
]
for _cap, _fn in _collages:
    subheading(_cap)
    img(os.path.join("eval", "assets", "collages", _fn), 14.2)

doc.save(os.path.join("eval", "PPT_Skill横向评测报告.docx"))
print("docx v6 (unified 7-skill rewrite) saved, paragraphs:", len(doc.paragraphs))
