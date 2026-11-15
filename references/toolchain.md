# 可重复执行的工具链

所有路径按当前环境选择。不要硬编码历史桌面、微信临时路径、RUL论文名、DOI、20页循环或Notion页面ID。Python示例中的`python`指已检查依赖的解释器；Windows原生环节需要`pywin32`，截图需要`Pillow`，论文读取可用PyMuPDF。不要把项目专用venv路径写成Skill全局约束。

Windows中文文件建议设置`$env:PYTHONUTF8='1'`。浏览器捕获脚本默认面向PowerShell 7；旧版PowerShell应显式设置UTF-8管道编码，防止中文JavaScript/JSON被改成问号。

## 内容JSON与HTML

`render_html.py`使用1280×720蓝白红框架。用户不同版式可自行编写HTML，保留以下提取契约：每页`.slide`、当前页`.slide.active`、`#slide-picker`选择器用1开始的页码且change事件切页，标题可存`data-title`。提取器读取实际画布尺寸，不依赖固定页数。

JSON示意（内容由论文核验后填写；这里不是科研事实）：

```json
{
  "title": "本次文献汇报",
  "assets": ["figures/figure-1.png"],
  "references": [{"title": "真实完整论文题名", "doi": null}],
  "slides": [{
    "title": "本页要回答的问题",
    "claim": "核心判断与<span class=\"v2-red\">关键证据</span>。",
    "body": "<div class=\"v2-columns v2-58\"><section class=\"v2-panel\"><h2>证据</h2><p>据原文核验的内容。</p></section><section class=\"v2-panel v2-pale\"><h2>阅读分析</h2><p>与作者结论分开。</p></section></div>",
    "takeaway": "本页要点",
    "locator": "PDF第几页 / 图表编号",
    "note": "讲解备注"
  }]
}
```

`assets`相对输入JSON目录；脚本复制图片到输出HTML的`assets`，正文引用`assets/figure-1.png`。同名不同内容会报错，避免覆盖错图。每页可用`references`覆盖共同来源。没有DOI标明未提供。

常用HTML类：`v2-columns v2-58`（58/42双栏）、`v2-66`、`v2-70`、`v2-stack`、`v2-panel`、`v2-pale`、`v2-table v2-numeric`、`v2-vertical`、`v2-bottom-pair`、`v2-bottom-triplet`、`v2-equation`、`v2-note`、`v2-red`。原图结构：

```html
<div class="v2-evidence"><figure class="paper-figure v2-figure">
<div class="figure-zoom"><img src="assets/figure-1.png" alt="原图的准确说明"></div>
<figcaption>Fig.1 · 来源页码与实验条件</figcaption>
</figure></div>
```

正文body是静态HTML片段，不是完整网页；禁止动态脚本。公式上下标可用sup/sub，转换后复核。普通文本保持整句、整行；不要在输入HTML就逐字包装。

```powershell
python '<SKILL>/scripts/render_html.py' --deck '<WORK>/deck.json' --out '<WORK>/html'
& '<SKILL>/scripts/capture_layout.ps1' -Html '<WORK>/html/index.html' -Out '<WORK>/layout'
python '<SKILL>/scripts/build_svg.py' --layout '<WORK>/layout/layout.json' --html-dir '<WORK>/html' --out '<WORK>/svg'
```

浏览器捕获使用已安装的`npx --offline --no-install agent-browser`，不自动下载。工具不可用时用可用浏览器运行同一个`extract_layout.js`并保存JSON；不能省略实际渲染。`unsupported styles`错误意味着需要修正HTML或手工重建相应SVG，不要直接抹掉警告强行转换。回看每页PNG并修复布局后继续。

## PowerPoint原生转换

```powershell
python '<SKILL>/scripts/powerpoint_export.py' --svg-dir '<WORK>/svg' --template '<SOURCE>/template.pptx' --out '<WORK>/final' --notes '<WORK>/html/slides.json'
```

无模板时省略`--template`，默认新建960×540 pt。可用`--size 720 405`指定无模板的真实尺寸；有模板则由PageSetup决定，SVG长宽比不匹配会报错。不要拉伸论文图来凑比例。

只需SVG两态验证而不生成PPTX时加`--validate-only`。默认最终保存`presentation.pptx`和PDF，另存两态PNG、最终预览、编辑窗口截图与对象报告；PDF失败会单列错误，不抹掉PPTX。

可选轻度微调（只在完整PPTX终点使用）：

```json
{"title_pt":24,"source_min_pt":9.5,"page_number_pt":11}
```

传入`--adjustments '<WORK>/adjustments.json'`。参数按实际模板设计选择，不能不看模板直接照抄。来源字号仅作用于页面底部来源区；单独页码按右边界对齐。复杂版面修正应回到源HTML/SVG；脚本的微调不替代逐页目视检查。

脚本输出路径已有成品时默认拒绝覆盖。修订另用新目录，或在会话授权覆盖时显式传`--overwrite`。它不会关闭用户其他演示文稿，不会保存原模板。

## 最后检查

核验PowerPoint报告的text_content_verified、pictures、pptx_svg_media_count、text_bounds/overflow、figure_bytes_preserved与reopened_in_powerpoint。脚本的通过状态只代表结构检查，视觉检查需读取所有预览并单独记录。检查全部截图和重点页原尺寸，确认没有重复验证页、正文逐字成框、图片碎片或模板原有正文。做对比测量时分开记录插入态/转形态与最终轻度微调后的成品，不把授权微调混成转换漂移。

数据表与公式数值还需回到论文验证；脚本只检验传递和编辑性，不证明论文事实正确。交付文件链接；若只有部分阶段成功，按实际状态报告。
