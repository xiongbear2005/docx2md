import os
from docx import Document
from docx.document import Document as _Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
import argparse
from docx2md.omml_to_latex import convert_omml_to_latex


def save_image(rel, image_dir, image_id):
    """Save image from relationship and return the filename."""
    try:
        image_bytes = rel.target_part.blob
        image_ext = os.path.splitext(rel.target_ref)[-1]
        if not image_ext:
            image_ext = '.png'  # Default extension if none found
            
        image_filename = f"image_{image_id}{image_ext}"
        image_path = os.path.join(image_dir, image_filename)
        
        with open(image_path, 'wb') as f:
            f.write(image_bytes)
            
        return image_filename
    except Exception as e:
        print(f"Error extracting image: {e}")
        return None


def iter_block_items(parent):
    """Generator to iterate through all block items (paragraphs and tables) in order."""
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Expected a Document or a Cell")
        
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def get_element_text(element):
    """Extract text from an XML element."""
    text = ""
    for child in element.iter():
        if child.tag.endswith('}t'):  # Text element in Word XML
            if child.text:
                text += child.text
    return text


def table_to_markdown(table):
    """Convert a docx table to Markdown format."""
    if not table.rows:
        return ""
        
    md_table = []
    
    # Extract header row
    header = []
    for cell in table.rows[0].cells:
        header.append(cell.text.strip() or " ")
    
    # Calculate column widths
    col_widths = [max(len(header[i]), 3) for i in range(len(header))]
    
    # Adjust column widths based on content
    for row in table.rows[1:]:
        for i, cell in enumerate(row.cells):
            if i < len(col_widths):
                col_widths[i] = max(col_widths[i], len(cell.text.strip() or " "))
    
    # Create header row
    header_formatted = "| " + " | ".join(h.ljust(col_widths[i]) for i, h in enumerate(header)) + " |"
    md_table.append(header_formatted)
    
    # Create separator row
    separator = "|" + "|".join("-" * (w + 2) for w in col_widths) + "|"
    md_table.append(separator)
    
    # Create content rows
    for row in table.rows[1:]:
        row_cells = []
        for i, cell in enumerate(row.cells):
            if i < len(col_widths):
                row_cells.append((cell.text.strip() or " ").ljust(col_widths[i]))
        md_table.append("| " + " | ".join(row_cells) + " |")
    
    return "\n".join(md_table)


def find_embedded_image_ids(element):
    """Find embedded image IDs in an element."""
    image_ids = []
    
    # We need to look for drawing elements in the XML
    for child in element.iter():
        if child.tag.endswith('}drawing'):
            # Look for blip elements that contain image references
            for subchild in child.iter():
                if subchild.tag.endswith('}blip'):
                    # Get the embed attribute which is the relationship ID
                    for key, value in subchild.attrib.items():
                        if key.endswith('}embed'):
                            image_ids.append(value)
    
    return image_ids


def extract_math_from_element(element):
    """Extract math elements (OMML) from a paragraph element."""
    math_elements = []
    
    # Look for math elements in the XML
    for child in element.iter():
        if child.tag.endswith('}oMath'):
            math_elements.append(child)
    
    return math_elements


def omml_to_latex_basic(omml_element):
    """Convert OMML (Office Math Markup Language) to LaTeX format using the advanced converter."""
    return convert_omml_to_latex(omml_element)


def process_paragraph_with_math(paragraph, image_dir, image_id_counter, relationship_map):
    """Process a paragraph that may contain text, images, and math formulas."""
    # Check for heading style first
    if paragraph.style.name.startswith('Heading'):
        heading_level = int(paragraph.style.name[-1]) if paragraph.style.name[-1].isdigit() else 1
        para_text = paragraph.text.strip()
        if para_text:
            return ['#' * heading_level + ' ' + para_text]

    # Process the paragraph element directly to maintain order
    result_text = process_paragraph_element_recursively(paragraph._element)

    # Handle images in the paragraph
    image_content = []
    for image_id in find_embedded_image_ids(paragraph._element):
        if image_id in relationship_map:
            rel = relationship_map[image_id]
            image_filename = save_image(rel, image_dir, image_id_counter[0])
            if image_filename:
                # Use absolute path for the image
                image_path = os.path.abspath(os.path.join(image_dir, image_filename))
                # Convert backslashes to forward slashes for markdown compatibility
                image_path = image_path.replace('\\', '/')
                image_content.append(f"![image_{image_id_counter[0]}]({image_path})")
                image_id_counter[0] += 1

    result = []
    if result_text and result_text.strip():
        # Clean up extra spaces
        result_text = ' '.join(result_text.split())
        result.append(result_text)
    result.extend(image_content)

    return result


def print_xml_structure(element, level=0):
    """Print the XML structure of an element for debugging."""
    indent = "  " * level
    tag = element.tag.split('}')[-1] if '}' in element.tag else element.tag
    attrs = []
    for key, value in element.attrib.items():
        key = key.split('}')[-1] if '}' in key else key
        attrs.append(f"{key}='{value}'")
    attrs_str = " ".join(attrs)
    
    if element.text and element.text.strip():
        print(f"{indent}<{tag} {attrs_str}>{element.text.strip()}")
    else:
        print(f"{indent}<{tag} {attrs_str}>")
    
    for child in element:
        print_xml_structure(child, level + 1)
    
    if element.tail and element.tail.strip():
        print(f"{indent}{element.tail.strip()}")


def process_paragraph_element_recursively(element):
    """Recursively process paragraph element to extract text and math in correct order."""
    result_parts = []

    # Process all child elements in order
    for child in element:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag == 'r':  # Run element
            # Process run content
            run_text = process_run_element(child)
            if run_text:
                result_parts.append(run_text)

        elif tag == 'oMath':  # Math element
            # (debug prints removed)
            
            latex_formula = omml_to_latex_basic(child)
            if latex_formula and latex_formula != "[Math Formula]":
                # Determine if it's inline or display math
                if len(latex_formula) > 50 or any(cmd in latex_formula for cmd in ['\\frac', '\\sum', '\\int', '\\prod']):
                    result_parts.append(f" $$\n{latex_formula}\n$$ ")
                else:
                    result_parts.append(f" ${latex_formula}$ ")

        else:
            # Recursively process other elements
            child_text = process_paragraph_element_recursively(child)
            if child_text:
                result_parts.append(child_text)

    return ''.join(result_parts)


def process_run_element(run_element):
    """Process a run element to extract text and inline math."""
    result_parts = []

    for child in run_element:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag == 't':  # Text element
            if child.text:
                result_parts.append(child.text)

        elif tag == 'oMath':  # Inline math in run
            latex_formula = omml_to_latex_basic(child)
            if latex_formula and latex_formula != "[Math Formula]":
                # For math in runs, prefer inline format
                result_parts.append(f" ${latex_formula}$ ")

        else:
            # Recursively process other elements
            child_text = process_run_element(child)
            if child_text:
                result_parts.append(child_text)

    return ''.join(result_parts)


def docx_to_markdown_with_formulas(docx_path, output_md_path, image_dir="images"):
    """Convert DOCX file to Markdown, preserving text, images, tables, and math formulas in order."""
    # Create image directory if it doesn't exist
    if not os.path.exists(image_dir):
        os.makedirs(image_dir)
    
    doc = Document(docx_path)
    md_content = []
    
    # Use a counter wrapped in a list to track the image_id through function calls
    image_id_counter = [1]
    formula_count = {'inline': 0, 'display': 0}
    
    # Build a map of relationship IDs to relationships
    relationship_map = {}
    for rel_id, rel in doc.part.rels.items():
        relationship_map[rel_id] = rel
    
    # Process document blocks (paragraphs and tables) in order
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            para_content = process_paragraph_with_math(block, image_dir, image_id_counter, relationship_map)
            
            # Count formulas for statistics
            for content in para_content:
                if content.startswith('$$') and content.endswith('$$'):
                    formula_count['display'] += 1
                elif '$' in content and not content.startswith('$$'):
                    formula_count['inline'] += content.count('$') // 2
            
            if para_content:
                md_content.extend(para_content)
                
        elif isinstance(block, Table):
            md_table = table_to_markdown(block)
            if md_table:
                md_content.append(md_table)
    
    # Note: All images should be processed within paragraphs above
    # No need to check for remaining images as they are handled in paragraph processing
    
    # Write to markdown file - ensure UTF-8 encoding
    try:
        with open(output_md_path, 'w', encoding='utf-8') as f:
            f.write('\n\n'.join(md_content))
        print(f"Markdown file saved to: {output_md_path}")
    except UnicodeEncodeError:
        # Fallback to write with explicit error handling
        with open(output_md_path, 'w', encoding='utf-8', errors='xmlcharrefreplace') as f:
            f.write('\n\n'.join(md_content))
        print(f"Markdown file saved to: {output_md_path} (with character encoding workaround)")
    
    print(f"Images saved to: {image_dir}/")
    print(f"Total images found: {image_id_counter[0] - 1}")
    print(f"Formula conversion completed.")
    print(f"Conversion Statistics:")
    print(f"  - Inline formulas: {formula_count['inline']}")
    print(f"  - Display formulas: {formula_count['display']}")
    print(f"  - Total formulas: {formula_count['inline'] + formula_count['display']}")

    # # Save a debug copy with BOM for troubleshooting encoding issues
    # with open(output_md_path + '.debug', 'w', encoding='utf-8-sig') as f:
    #     f.write('\n\n'.join(md_content))
    # print(f"Debug copy with BOM saved to: {output_md_path}.debug")


def main():
    parser = argparse.ArgumentParser(description='Convert DOCX file to Markdown with math formula support')
    parser.add_argument('docx_file', help='Input DOCX file path')
    parser.add_argument('-o', '--output', help='Output Markdown file path')
    parser.add_argument('-i', '--image_dir', default='images', help='Directory to save images')
    
    args = parser.parse_args()
    
    docx_path = args.docx_file
    output_path = args.output if args.output else os.path.splitext(docx_path)[0] + '_with_formulas.md'
    
    docx_to_markdown_with_formulas(docx_path, output_path, args.image_dir)


if __name__ == "__main__":
    main() 