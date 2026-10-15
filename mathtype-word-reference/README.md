# mathtype-word-reference

用于 Codex 的 MathType/Word 技能：从 LaTeX 插入可编辑公式，或从参考 Word 文档复用公式与行内小符号，统一字体、粗斜体、上下标及居中右编号。

## 安装

将本仓库文件夹命名为 `mathtype-word-reference`，放到个人技能目录：

```text
Windows: %USERPROFILE%\.codex\skills\mathtype-word-reference\
macOS/Linux: ~/.codex/skills/mathtype-word-reference/
```

如果使用自定义 `CODEX_HOME`，放在其 `skills` 子目录。重新加载技能列表；列表仍未更新时重新打开应用。下载 `.skill` 文件时可把它当作 ZIP 压缩包解压，再复制其中同名目录。

调用示例：

> 使用 $mathtype-word-reference，将目标 Word 的公式及正文小符号统一为参考文档的 MathType 格式，公式居中、编号靠右，保存新文件。

## 环境与能力

- 离线检查、参考对象移植：Python 3.10+，`pip install -r requirements.txt`；不依赖其他 skill。
- 新增 LaTeX 公式的实际编辑器插入：Windows、Word/WPS、已安装且可用的 MathType，以及当前代理能够操作该应用的接口。脚本本身不执行 GUI、不把 LaTeX 转成 MathType。
- 无参考对象时，按 `SKILL.md` 和 `references/workflow.md` 使用 MathType 编辑器；不把公式截图或普通 Word 数学对象冒充 MathType。
- 不附带 MathType 软件、商业字体、Office 组件、专利原稿或公式二进制样本。

## 脚本用法

```powershell
python scripts/reference_math.py inspect --docx reference.docx --json inventory.json
python scripts/reference_math.py apply --input candidate.docx --reference reference.docx --manifest mapping.json --output result.docx
python scripts/reference_math.py audit --docx result.docx --json audit.json
```

先查看清单，再按 `references/mapping.example.json` 建立准确映射。来源序号是 OpenXML 零起始序号，不是 Word COM 序号。目标显示公式使用整段占位符，行内符号按文字精确匹配并验证出现次数。输出路径必须与输入、参考不同。

脚本会保留原生公式数据、复制预览及关系、重分配对象标识；主公式编号使用普通文本，不自动重排。最终仍需渲染逐页检查，抽查 MathType 实际编辑效果。结构审计不代替数学语义和视觉验收。

## 发布到 GitHub

直接上传本目录内容作为仓库根目录，确保根目录存在 `SKILL.md`。发布压缩包时保留顶层 `mathtype-word-reference/` 文件夹。运行 `python scripts/test_reference_math.py` 可检查对象完整移植、跨 run 的行内替换，以及只读来源保护。

本次交付未代替作者选择开源许可证。需要授予他人再分发或修改权利时，发布前在仓库增加你选定的 `LICENSE`。
