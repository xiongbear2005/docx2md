# DOCX to Markdown Converter with Math Formula Support

一个强大的DOCX到Markdown转换器，支持数学公式转换。将Microsoft Word文档中的文本、图片、表格和数学公式完美转换为Markdown格式，公式自动转换为LaTeX格式。

## ✨ 功能特性

- 📝 **完整内容转换**：支持文本、图片、表格的完整转换
- 🧮 **数学公式支持**：将Word中的数学公式转换为LaTeX格式
- 📍 **位置保持**：公式在文档中的位置与原文档保持一致
- 🎯 **智能分类**：自动区分行内公式（`$...$`）和显示公式（`$$...$$`）
- 🔧 **格式优化**：LaTeX命令格式规范，如 `\geq ` 后自动添加空格
- 🖼️ **图片提取**：自动提取并保存图片到指定目录，使用绝对路径确保显示
- 📊 **表格转换**：将Word表格转换为Markdown表格格式
- 🚀 **高效处理**：支持大型文档的快速转换

## 🛠️ 安装要求

### 依赖包
```bash
pip install -r requirements.txt
```

或手动安装：
```bash
pip install python-docx lxml mammoth sympy pillow
```

## 🚀 快速开始

### 基本用法

```bash
python docx2md.py input.docx
```

### 指定输出文件和图片目录

```bash
python docx2md.py input.docx -o output.md -i images
```

### 参数说明

- `input_file`：输入的DOCX文件路径
- `-o, --output`：输出Markdown文件路径（可选）
- `-i, --image_dir`：保存图片的目录名（默认：`images`）

## 📖 使用示例

### 示例1：基本转换
```bash
python docx2md.py sample.docx
# 输出：sample_with_formulas.md 和 images/ 目录
```

### 示例2：自定义输出
```bash
python docx_to_md_with_formulas.py document.docx -o result.md -i pictures
# 输出：result.md 和 pictures/ 目录
```

## 🧮 数学公式支持

### 支持的数学元素

#### 基本运算
- **分数**：`\frac{a}{b}`
- **上标**：`x^{2}`
- **下标**：`x_{1}`
- **根号**：`\sqrt{x}`, `\sqrt[n]{x}`

#### 高级运算
- **积分**：`\int`, `\int_{a}^{b}`
- **求和**：`\sum`, `\sum_{i=1}^{n}`
- **乘积**：`\prod`
- **极限**：`\lim`

#### 希腊字母
- **小写**：α, β, γ, δ, ε, θ, λ, μ, π, σ, τ, φ, ψ, ω
- **大写**：Γ, Δ, Θ, Λ, Ξ, Π, Σ, Υ, Φ, Ψ, Ω

#### 运算符和符号
- **关系**：≤, ≥, ≠, ≈, ≡
- **集合**：∈, ∉, ⊂, ⊆, ∪, ∩, ∅
- **箭头**：→, ←, ↔, ⇒, ⇐, ⇔
- **其他**：∞, ∂, ∇, ±, ×, ÷, ·

### 转换示例

**输入（Word文档）**：包含数学公式的Word文档

**输出（Markdown）**：
```markdown
在传统的MoE架构中，给定输入 $x$，门控网络 $G(x)$ 会输出一个权重向量 $p=[p_1,p_2,…,p_n]$。

复杂公式将显示为：
$$
Attention(Q,K,V) = softmax(\frac{QK^T}{\sqrt{d_k}})V
$$
```

## 📊 转换统计

转换完成后，程序会显示详细统计信息：

```
Markdown file saved to: output.md
Images saved to: images/
Total images found: 19
Formula conversion completed.

Conversion Statistics:
  - Inline formulas: 64
  - Display formulas: 19
  - Total formulas: 83
```

## 🖼️ 查看建议

为了获得最佳的数学公式渲染效果，推荐使用以下Markdown查看器：

- **Visual Studio Code** + "Markdown Preview Enhanced" 扩展
- **Typora** Markdown编辑器
- **Jupyter Notebook**
- **在线查看器**：StackEdit、Dillinger

## 🔧 进一步处理

### 转换为PDF
```bash
pandoc output.md -o document.pdf
```

### 转换为HTML
```bash
pandoc output.md -o document.html
```

## 📁 项目结构

```
project/
├── docx_to_md_with_formulas.py  # 主转换脚本
├── omml_to_latex.py             # OMML到LaTeX转换器
├── requirements.txt             # 依赖包列表
├── input.docx                   # 示例输入文件
└── README.md                    # 项目说明
```

## 🔍 工作原理

1. **文档解析**：使用 `python-docx` 读取DOCX文档结构
2. **内容提取**：按顺序提取文本、图片、表格和数学公式
3. **公式转换**：使用自定义OMML到LaTeX转换器处理数学公式
4. **格式优化**：优化LaTeX命令格式，添加适当空格
5. **位置保持**：确保所有元素在Markdown中保持正确位置
6. **文件生成**：生成Markdown文件和图片目录

## ⚠️ 注意事项

- 确保输入的DOCX文件没有损坏
- 数学公式必须使用Word的公式编辑器创建
- 复杂的格式可能无法完全保留
- 建议在转换前备份原始文件
- **图片路径**：转换器使用绝对路径引用图片，确保在任何位置打开Markdown文件都能正确显示图片

## 🐛 故障排除

### 常见问题

1. **公式显示为 "[Math Formula]"**
   - 检查DOCX文件是否包含有效的数学公式
   - 确保公式是使用Word的公式编辑器创建的

2. **图片无法显示**
   - 检查图片目录路径是否正确
   - 确保有足够的磁盘空间保存图片

3. **转换失败**
   - 检查是否安装了所有必需的依赖包
   - 确保DOCX文件没有损坏
---

**享受无缝的DOCX到Markdown转换体验！** 🎉
