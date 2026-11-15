# PowerPoint SVG转形：已验证的细节

这些是Windows PowerPoint 16.0上观察到的行为。每次按当前PowerPoint版本、实际字体与真实模板复测，不把一次经验当作跨平台保证。

## 文字

1. `SVGEdit`是真实“转换为形状”的Ribbon命令。先选中SVG，再检查`GetEnabledMso('SVGEdit')`，执行`ExecuteMso('SVGEdit')`。取消选择后导出。不能以直接生成DrawingML冒充执行过此命令。
2. 同一text内，先bold后normal的tspan可能使后半行也加粗；数字400/700、normal/bold与内联style的重置均曾失败。优先整行一致字重；确需混合则按完整短语分成少量text，不能拆正文单字。只变颜色的tspan可保留，但仍检查实际转形。
3. 一些小数字号在转形时按SVG像素取整。脚本将字号校准到整数SVG单位，再在真实尺寸插入测试；不能只改成pt或加transform后假定解决。
4. 只有同字体/字号/字重且空间连续的片段才能合并。不连续的公式主基线片段之间往往有上下标占位，合并会吞掉间隔并造成符号重叠。
5. PowerPoint原生文本框带边距。检查TextFrame2.TextRange的可见边界与实际导出图，不能仅检查Shape.Left/Top。设置字号后复核右边界与换行。
6. 文本行转形为可编辑对象不意味着整段是一个TextFrame；报告实际粒度。表格默认由文本和线条组成，不冒称Office原生Table；公式文本也不冒称Office Equation。

## 环境、导出与复查

- Python使用pywin32；窗口截图使用Pillow。PowerPoint必须有可交互桌面。不要终止用户现有进程或任意关闭其演示文稿。
- 截图前在进程开始设置DPI awareness，否则GetWindowRect的逻辑坐标与ImageGrab物理坐标可能不一致。窗口用PPTFrameClass定位；不要依赖不可用的Application.HWND属性。
- ImageGrab截取屏幕可见像素，必须先激活PowerPoint并确认前台窗口；否则可能截到遮挡它的应用。查看截图内容，不能仅凭文件存在判定编辑视图检查完成。
- 先读取模板尺寸。例：1280×720 viewBox插入960×540 pt得到0.75 pt/单位；720×405 pt只有0.5625，二者不能共用字号测试结论。
- 打开模板ReadOnly，修改仅发生在内存副本，SaveAs到新路径。保存前后比对模板哈希。关闭时仅关闭本脚本打开的对象。
- Python动态COM调用ExportAsFixedFormat可能发生参数封送错误；已验证的PDF后备路径是先SaveAs(PPTX,24)，再SaveAs(PDF,32)。PDF导出失败不能抹掉已成功保存的PPTX，也不能声称PDF成功。
- 保存后重新用PowerPoint打开交付文件再导出预览；检查可编辑文本内容、原图数量、残留SVG对象、页数、字体和可见越界。不要把验证用两态副本也放进最终成品。

## 不支持自动转换的HTML

附带提取器针对本Skill的平面幻灯片HTML，不是任意CSS转换器。不可静默忽略渐变、阴影、圆角、滤镜、伪元素、旋转、RTL/复杂脚本或复杂裁切。先简化为已确认的等价基础形状，或手工重建相应SVG并核验。文字、图片及表格清晰度优先于装饰。

局部像素墨迹框可能包含相邻公式上下标，不应用这一数值判断字体本身扩大了几十个百分点。自动异常须回到原尺寸图复核，并保留具体例外。

## 工具包验证范围

2026-11-15在Microsoft PowerPoint 16.0完成两页端到端样例：HTML捕获→结构化SVG→原生SVGEdit→字号微调→保存PPTX/PDF→重开核验。得到38个可编辑文本对象和1张完整图片，原图字节保持一致，备注保存完整，无残留SVG。另以只读模板完成两页仅验证模式，模板哈希不变且未输出PPTX。检查了HTML、两态/最终PNG及PowerPoint编辑窗口；中英文、混合字重、上下标与无DOI标注可用。

这是工具链样例验证，不代表所有论文、字体、复杂公式或PowerPoint版本都已验证。新任务仍需按实际内容和尺寸逐页检查。
