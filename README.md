# ppt-skill-benchmark

对多款 PPT 生成 skill / AI PPT 工具做横向评测的可复用 ZCode 技能：双赛道（纯文字大纲 + 原版图片复刻）生成 18 页原生可编辑 PPTX、渲染逐页核验、打散盲评打分工作台、排行榜/象限图、楷体 DOCX+PDF 报告与全部成品截图附录。

## 使用

安装到 `~/.zcode/skills/ppt-skill-benchmark/`（本仓库根即该目录），然后在 ZCode 里说：

> 用 ppt-skill-benchmark 评测这几个 skill：…（附上 skill 名单或目录），基准是 …

## 结构

- `SKILL.md`：硬规则 + 九步流程 + 交付清单
- `scripts/`：渲染、核验、拼接图、图表、盲评工作台增量加图、报告模板
- `references/`：各 skill 安装实战笔记、ppt-master 最小闭环跑法、报告 14 章规范

评测实例（7 款终榜与全部产物）见配套仓库。
