#!/usr/bin/env python3
"""
DOCX to Markdown Converter - 使用示例

这个脚本演示了如何使用DOCX到Markdown转换器。
"""

import os
import sys
from docx2md import docx_to_markdown_with_formulas


def main():
    """示例：转换input.docx文件"""
    
    print("DOCX to Markdown Converter - 使用示例")
    print("=" * 50)
    
    # 检查示例文件
    input_file = "input.docx"
    if not os.path.exists(input_file):
        print(f"❌ 找不到示例文件: {input_file}")
        print("请确保当前目录下有input.docx文件")
        return
    
    # 设置输出参数
    output_file = "example_output.md"
    image_dir = "example_images"
    
    print(f"📁 输入文件: {input_file}")
    print(f"📁 输出文件: {output_file}")
    print(f"📁 图片目录: {image_dir}")
    print()
    
    try:
        print("🔄 开始转换...")
        
        # 执行转换
        docx_to_markdown_with_formulas(input_file, output_file, image_dir)
        
        print("✅ 转换完成！")
        print()
        print("📖 查看结果:")
        print(f"  - Markdown文件: {output_file}")
        print(f"  - 图片目录: {image_dir}/")

        
    except Exception as e:
        print(f"❌ 转换失败: {e}")
        return


if __name__ == "__main__":
    main()
