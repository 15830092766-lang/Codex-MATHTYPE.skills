---
name: mathtype-word-reference
description: 在 Word/WPS 中用 LaTeX 插入可编辑 MathType 公式，或复用参考文档的 MathType 对象，统一主公式和正文小符号的字体、粗细、上下标及居中右编号。适用于有参考格式的专利、论文和技术文档排版。
---

# MathType 参考格式插入

交付可编辑的 Word 文档。既有参考公式时优先复制其原生对象；新增公式时使用已安装的 MathType 编辑器。完整流程见 [references/workflow.md](references/workflow.md)。技能独立使用，不要求安装其他 Codex skill。

## 选择操作方式

- **新公式**：准备语义核对后的 LaTeX，在 Word 的 MathType 插件选择“右编号”，向新打开的 MathType 编辑器粘贴 LaTeX，确认已解析成公式，然后关闭该编辑窗口并返回 Word 保存。普通正文中的小符号使用行内公式，不加编号。如果用户明确要求该操作顺序，保留它。
- **参考文档已有相同公式**：复制完整 MathType OLE 对象及其缓存预览，保留字母字形、粗细、斜体、上下标和运算符。使用下方脚本按明确映射处理。
- **公式相同、格式不同**：以用户指定的参考公式为准；分别核对主公式和正文同名符号。只改 Word 字体或对象外框无法改变 MathType 内部字体。

## 必须保留的约束

1. 把原稿和参考文档作为只读来源，写到新的输出路径。用户要求覆盖既有生成稿时，先备份。旧 `.doc` 只读打开后另存临时 `.docx`，不保存回原文件。
2. 不改变公式数学内容，也不为排版而整体替换说明段落。逐项替换符号，保留周围中文、标点、变量定义和权利要求内容。参考文档自身也可能有符号错误，复制前先核对语义。
3. 一个主公式可能包含多个相邻 OLE 对象；全部保留。正文行内公式不增加制表位、独立段落或编号。
4. 以参考公式为格式基准，避免强制设置通用字号比例。变量、向量、矩阵的粗斜体服从用户和参考定义，不把所有字母统一加粗。
5. MathType 生成的原生右编号可以保留。若通过 OpenXML 重建排版，使用“居中制表位→对象组→右制表位→普通编号”；不要手工拼接 `MTPlaceRef`/`SEQ MTEqn` 嵌套域。普通编号不会自动更新，需要插删公式后重新核对。
6. 权利要求书、说明书等不同部分按用户要求独立编号。不要仅凭相同的编号匹配两份文档的公式。
7. 新增公式需要 MathType 正常安装、注册、可用的授权及实际可操作的 Word/WPS；不要承诺未验证的 GUI 自动化能力。
8. 自动化仅关闭本次创建的后台文档和进程，不关闭用户正在编辑的窗口。

## 脚本工作流

离线复制与检查只需要 Python 3.10+、`lxml`、`olefile`。脚本不调用 Word，不生成新公式，也不提供 LaTeX 到 MathType 的转换。

```powershell
python -m pip install -r "<skill-folder>\requirements.txt"
python "<skill-folder>\scripts\reference_math.py" inspect --docx "<reference.docx>" --json "<inventory.json>"
python "<skill-folder>\scripts\reference_math.py" apply --input "<candidate.docx>" --reference "<reference.docx>" --manifest "<mapping.json>" --output "<result.docx>"
python "<skill-folder>\scripts\reference_math.py" audit --docx "<result.docx>" --json "<audit.json>"
```

清单格式见 [references/mapping.example.json](references/mapping.example.json)。`source_paragraph` 是 `inspect` 返回的 **OpenXML 零起始段落序号**，`source_object` 是该段内 **零起始 MathType 对象序号**。不得把它们当作 Word COM 的段落序号。

`display` 用一个完整占位段落替换主公式；`inline` 按原文字串精确替换正文符号，可跨多个 Word run。重复次数必须符合 `expected_count`，否则在写入前报错。先复制主公式与符号，再渲染，不整段粘贴参考解释文本。

## 最终验收

- 映射和数学意义相符；正文内容及符号定义未被排版操作改写。
- 审计分别报告含公式段落数和对象总数。每个对象必须有 `Equation.DSMT4`、真正的复合存储、原生公式流和有效预览关系。不能只看 ProgID 就宣称可编辑。
- 审计流结构不等于验证全部可编辑行为；有条件时双击代表性公式确认 MathType 可逐项编辑。
- 用实际 Word/WPS 导出 PDF，或可用的 DOCX 渲染器输出逐页图像；检查所有页面：主公式居中、编号右对齐、行内符号基线合适、无剪切、无意外粗体、无问号和域代码。渲染器与实际 Office 有差异时以用户实际应用复核为准。
- 检查用户指定只读原稿的校验和或文件属性未变化。交付最终 Word 路径，明确已做的检查和仍未验证的部分。
