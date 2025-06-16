# DOCX ➜ Markdown 转换器（支持 LaTeX 数学公式）

一个体积轻巧但功能强大的脚本，能将 Microsoft Word **.docx** 文档转换为友好的 **Markdown**，并保留：

- 🖋️ 文字 & 标题
- 🖼️ 内嵌图片
- 📊 表格
- 🧮 数学公式（自动转换为 **LaTeX**）

所有元素按照在 Word 中出现的顺序导出，生成的 Markdown 与原文档排版保持高度一致。

---

## ✨ 主要特性

1. **完整内容转换**：文字、标题、表格、图片、公式依次处理，不遗漏任何部分。
2. **OMML ➜ LaTeX**：自研转换器，支持分数、积分、求和、根号、矩阵、重音、希腊字母等。
3. **行内 / 显示公式智能判断**：复杂公式使用 `$$ … $$`，简单公式使用 `$ … $` 包裹。
4. **图片提取**：所有图片保存到自定义目录，并在 Markdown 中以**绝对路径**引用，确保可见性。
5. **表格转换**：Word 表格 → Markdown 表格，自动调整列宽。
6. **转换统计**：结束后输出图片数量、行内公式 / 显示公式数量等信息。
7. **CLI & 库双模式**：既可命令行调用，也可在 Python 代码中直接调用函数。

---

## 🛠️ 安装

```bash
pip install -r requirements.txt
```

依赖包：`python-docx`、`Pillow`、`lxml`、`mammoth`、`sympy`（上述命令会一并安装）。

---

## 🚀 命令行用法

```bash
python docx2md.py 输入.docx [-o 输出.md] [-i 图片目录]
```

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `输入.docx` | 待转换的 Word 文件 | — |
| `-o, --output` | 输出 Markdown 路径 | `<输入文件名>_with_formulas.md` |
| `-i, --image_dir` | 图片保存目录 | `images` |

示例：

```bash
python docx2md.py 论文.docx -o 论文.md -i assets
```

转换结束后将看到类似输出：

```
Markdown file saved to: 论文.md
Images saved to: assets/
Total images found: 12
Formula conversion completed.
Conversion Statistics:
  - Inline formulas: 64
  - Display formulas: 19
  - Total formulas: 83
```

---

## 🐍 代码中调用

```python
from docx2md import docx_to_markdown_with_formulas

docx_to_markdown_with_formulas(
    "输入.docx",          # Word 文件
    "输出.md",           # 目标 Markdown
    image_dir="images"   # 图片目录
)
```

---

## 📁 项目结构

```
├── docx2md.py            # CLI & 函数入口
├── omml_to_latex.py      # OMML ➜ LaTeX 转换器
├── example.py            # 使用示例脚本
├── requirements.txt      # 依赖列表
└── README.md             # 当前说明文档
```

---

## ⚠️ 已知限制

- 仅支持使用 **Word 公式编辑器**（OMML）的公式，旧版 Equation 3.0 不支持。
- 复杂版式（脚注、文本框、SmartArt 等）目前会被忽略。


---

祝转换愉快 🎉
