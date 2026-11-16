# paper-to-powerpoint

将科研论文制作成中文文献汇报与可编辑 PowerPoint 的 Agent Skill。

通过论文精读、HTML 排版、浏览器布局提取、SVG 重建和 Microsoft PowerPoint 原生转形，保留可编辑文字与完整论文原图。默认约 20 页、16:9、蓝白配色与少量红色重点，可按任务调整。

## 安装

支持 Agent Skills 的客户端可使用 Skills CLI 安装。

```bash
npx skills add wanghailang2005-alt/paper-to-powerpoint
```

Codex 也可将仓库克隆到技能目录。

```bash
git clone https://github.com/wanghailang2005-alt/paper-to-powerpoint.git ~/.codex/skills/paper-to-powerpoint
```

安装后重新加载客户端，并提供论文 PDF、汇报要求及可选模板。

## 使用

向支持该技能的 AI 客户端发送以下请求，并附上论文。

> 使用 paper-to-powerpoint 精读这篇论文，制作约 20 页中文组会汇报。解释研究问题、方法、公式和实验，保留完整原图并标注来源，交付可编辑 PPTX 和预览。

已有 HTML 或 SVG 时，也可以要求从该阶段继续。

## 运行环境

- Python 3，PowerPoint 自动化使用 `pywin32`，截图使用 `Pillow`。
- Windows 桌面版 Microsoft PowerPoint 和可交互桌面，用于 SVG 转形、保存及重开检查。
- PowerShell 7，以及 Node.js/npm 和已安装的 `agent-browser`，用于默认浏览器布局捕获。
- 论文 PDF 的读取可使用 PyMuPDF。

```powershell
python -m pip install pywin32 Pillow PyMuPDF
```

HTML 和 SVG 阶段可独立运行。最终可编辑性与视觉效果需要在 Microsoft PowerPoint 中逐页检查。使用其他浏览器时，参照工具链说明运行相同的布局提取器。

## 文件与流程

| 文件 | 用途 |
| --- | --- |
| [SKILL.md](SKILL.md) | AI 执行入口与交付要求 |
| [内容与布局](references/content-and-layout.md) | 论文证据组织与页面设计 |
| [工具链](references/toolchain.md) | 输入 JSON、运行命令和输出说明 |
| [PowerPoint 兼容性](references/powerpoint-compatibility.md) | SVG 转形与视觉核验 |
| `scripts/` | HTML、浏览器布局、SVG 和 PowerPoint 脚本 |
| `assets/deck.css` | 默认页面样式 |

内容由 AI 根据提供的论文整理，脚本负责渲染和转换。交付前需核对论文事实、图表来源、公式、文字边界和最终预览。

## 许可证

本仓库以 [MIT License](LICENSE) 发布。使用时引入的论文、图表、字体及模板按各自权利人的授权使用。
