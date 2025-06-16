"""
OMML (Office Math Markup Language) to LaTeX converter
This module provides functions to convert Microsoft Word math equations to LaTeX format.
"""

import re
import xml.etree.ElementTree as ET
from lxml import etree


class OmmlToLatexConverter:
    """Converter class for OMML to LaTeX transformation."""
    
    def __init__(self):
        self.symbol_map = {
            # Greek letters
            'α': '\\alpha', 'β': '\\beta', 'γ': '\\gamma', 'δ': '\\delta',
            'ε': '\\epsilon', 'ζ': '\\zeta', 'η': '\\eta', 'θ': '\\theta',
            'ι': '\\iota', 'κ': '\\kappa', 'λ': '\\lambda', 'μ': '\\mu',
            'ν': '\\nu', 'ξ': '\\xi', 'ο': 'o', 'π': '\\pi',
            'ρ': '\\rho', 'σ': '\\sigma', 'τ': '\\tau', 'υ': '\\upsilon',
            'φ': '\\phi', 'χ': '\\chi', 'ψ': '\\psi', 'ω': '\\omega',
            
            # Capital Greek letters
            'Α': 'A', 'Β': 'B', 'Γ': '\\Gamma', 'Δ': '\\Delta',
            'Ε': 'E', 'Ζ': 'Z', 'Η': 'H', 'Θ': '\\Theta',
            'Ι': 'I', 'Κ': 'K', 'Λ': '\\Lambda', 'Μ': 'M',
            'Ν': 'N', 'Ξ': '\\Xi', 'Ο': 'O', 'Π': '\\Pi',
            'Ρ': 'P', 'Σ': '\\Sigma', 'Τ': 'T', 'Υ': '\\Upsilon',
            'Φ': '\\Phi', 'Χ': 'X', 'Ψ': '\\Psi', 'Ω': '\\Omega',
            
            # Mathematical operators
            '∞': '\\infty', '∑': '\\sum', '∫': '\\int', '∂': '\\partial',
            '∇': '\\nabla', '∆': '\\Delta', '∏': '\\prod',
            
            # Relations
            '≤': '\\leq', '≥': '\\geq', '≠': '\\neq', '≈': '\\approx',
            '≡': '\\equiv', '∝': '\\propto', '∼': '\\sim',
            
            # Set theory
            '∈': '\\in', '∉': '\\notin', '⊂': '\\subset', '⊆': '\\subseteq',
            '⊃': '\\supset', '⊇': '\\supseteq', '∪': '\\cup', '∩': '\\cap',
            '∅': '\\emptyset', '∀': '\\forall', '∃': '\\exists',
            
            # Arrows
            '→': '\\rightarrow', '←': '\\leftarrow', '↔': '\\leftrightarrow',
            '⇒': '\\Rightarrow', '⇐': '\\Leftarrow', '⇔': '\\Leftrightarrow',
            '↑': '\\uparrow', '↓': '\\downarrow', '↕': '\\updownarrow',
            
            # Other symbols
            '±': '\\pm', '∓': '\\mp', '×': '\\times', '÷': '\\div',
            '·': '\\cdot', '∘': '\\circ', '√': '\\sqrt', '∝': '\\propto',
            '∠': '\\angle', '⊥': '\\perp', '∥': '\\parallel',
        }
    
    def convert_element(self, element):
        """Convert an OMML element to LaTeX."""
        if element is None:
            return ""
        
        tag = element.tag.split('}')[-1] if '}' in element.tag else element.tag
        
        if tag == 'oMath':
            return self.convert_omath(element)
        elif tag == 'f':
            return self.convert_fraction(element)
        elif tag == 'sSup':
            return self.convert_superscript(element)
        elif tag == 'sSub':
            return self.convert_subscript(element)
        elif tag == 'sSubSup':
            return self.convert_subsuperscript(element)
        elif tag == 'rad':
            return self.convert_radical(element)
        elif tag == 'nary':
            return self.convert_nary(element)
        elif tag == 'd':
            return self.convert_delimiter(element)
        elif tag == 'm':
            return self.convert_matrix(element)
        elif tag == 'func':
            return self.convert_function(element)
        elif tag == 'acc':
            return self.convert_accent(element)
        elif tag == 'bar':
            return self.convert_bar(element)
        elif tag == 'box':
            return self.convert_box(element)
        elif tag == 'borderBox':
            return self.convert_border_box(element)
        elif tag == 'groupChr':
            return self.convert_group_char(element)
        elif tag == 'limLow':
            return self.convert_limit_lower(element)
        elif tag == 'limUpp':
            return self.convert_limit_upper(element)
        elif tag == 'r':
            return self.convert_run(element)
        elif tag == 't':
            return self.convert_text(element)
        else:
            # For unknown elements, try to process children
            result = ""
            for child in element:
                result += self.convert_element(child)
            return result
    
    def convert_omath(self, element):
        """Convert oMath element."""
        result = ""
        for child in element:
            result += self.convert_element(child)
        return result
    
    def convert_fraction(self, element):
        """Convert fraction element."""
        num = ""
        den = ""
        
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'num':
                num = self.convert_element(child)
            elif tag == 'den':
                den = self.convert_element(child)
        
        return f"\\frac{{{num}}}{{{den}}}"
    
    def convert_superscript(self, element):
        """Convert superscript element."""
        base = ""
        sup = ""
        
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'e':
                base = self.convert_element(child)
            elif tag == 'sup':
                sup = self.convert_element(child)
        
        return f"{{{base}}}^{{{sup}}}"
    
    def convert_subscript(self, element):
        """Convert subscript element."""
        base = ""
        sub = ""
        
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'e':
                base = self.convert_element(child)
            elif tag == 'sub':
                sub = self.convert_element(child)
        
        return f"{{{base}}}_{{{sub}}}"
    
    def convert_subsuperscript(self, element):
        """Convert subscript and superscript element."""
        base = ""
        sub = ""
        sup = ""
        
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'e':
                base = self.convert_element(child)
            elif tag == 'sub':
                sub = self.convert_element(child)
            elif tag == 'sup':
                sup = self.convert_element(child)
        
        return f"{{{base}}}_{{{sub}}}^{{{sup}}}"
    
    def convert_radical(self, element):
        """Convert radical (square root) element."""
        deg = ""
        base = ""
        
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'deg':
                deg = self.convert_element(child)
            elif tag == 'e':
                base = self.convert_element(child)
        
        if deg:
            return f"\\sqrt[{deg}]{{{base}}}"
        else:
            return f"\\sqrt{{{base}}}"
    
    def convert_nary(self, element):
        """Convert n-ary operators (sum, integral, etc.)."""
        char = ""
        sub = ""
        sup = ""
        base = ""
        
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'naryPr':
                for prop_child in child:
                    prop_tag = prop_child.tag.split('}')[-1] if '}' in prop_child.tag else prop_child.tag
                    if prop_tag == 'chr':
                        char = prop_child.get('val', '')
            elif tag == 'sub':
                sub = self.convert_element(child)
            elif tag == 'sup':
                sup = self.convert_element(child)
            elif tag == 'e':
                base = self.convert_element(child)
        
        # Map common n-ary operators
        operator_map = {
            '∑': '\\sum',
            '∫': '\\int',
            '∏': '\\prod',
            '⋃': '\\bigcup',
            '⋂': '\\bigcap',
            '⋁': '\\bigvee',
            '⋀': '\\bigwedge',
        }
        
        latex_op = operator_map.get(char, char)
        
        if sub and sup:
            return f"{latex_op}_{{{sub}}}^{{{sup}}} {base}"
        elif sub:
            return f"{latex_op}_{{{sub}}} {base}"
        elif sup:
            return f"{latex_op}^{{{sup}}} {base}"
        else:
            return f"{latex_op} {base}"
    
    def convert_delimiter(self, element):
        """Convert delimiter element."""
        # This is a simplified implementation
        result = ""
        for child in element:
            result += self.convert_element(child)
        return f"\\left( {result} \\right)"
    
    def convert_matrix(self, element):
        """Convert matrix element."""
        # This is a simplified implementation
        result = "\\begin{matrix}\n"
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'mr':  # matrix row
                row_content = []
                for cell in child:
                    cell_content = self.convert_element(cell)
                    row_content.append(cell_content)
                result += " & ".join(row_content) + " \\\\\n"
        result += "\\end{matrix}"
        return result
    
    def convert_function(self, element):
        """Convert function element."""
        func_name = ""
        base = ""
        
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'fName':
                func_name = self.convert_element(child)
            elif tag == 'e':
                base = self.convert_element(child)
        
        return f"\\{func_name}{{{base}}}"
    
    def convert_accent(self, element):
        """Convert accent element."""
        # Simplified implementation
        base = ""
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'e':
                base = self.convert_element(child)
        return f"\\hat{{{base}}}"
    
    def convert_bar(self, element):
        """Convert bar element."""
        base = ""
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'e':
                base = self.convert_element(child)
        return f"\\overline{{{base}}}"
    
    def convert_box(self, element):
        """Convert box element."""
        return self.convert_element(element)
    
    def convert_border_box(self, element):
        """Convert border box element."""
        base = ""
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'e':
                base = self.convert_element(child)
        return f"\\boxed{{{base}}}"
    
    def convert_group_char(self, element):
        """Convert group character element."""
        # Simplified implementation
        base = ""
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'e':
                base = self.convert_element(child)
        return f"\\underbrace{{{base}}}"
    
    def convert_limit_lower(self, element):
        """Convert limit lower element."""
        base = ""
        lim = ""
        
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'e':
                base = self.convert_element(child)
            elif tag == 'lim':
                lim = self.convert_element(child)
        
        return f"\\underset{{{lim}}}{{{base}}}"
    
    def convert_limit_upper(self, element):
        """Convert limit upper element."""
        base = ""
        lim = ""
        
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'e':
                base = self.convert_element(child)
            elif tag == 'lim':
                lim = self.convert_element(child)
        
        return f"\\overset{{{lim}}}{{{base}}}"
    
    def convert_run(self, element):
        """Convert run element."""
        result = ""
        for child in element:
            result += self.convert_element(child)
        return result
    
    def convert_text(self, element):
        """Convert text element."""
        text = element.text or ""

        # Replace symbols with LaTeX equivalents first
        for symbol, latex in self.symbol_map.items():
            text = text.replace(symbol, latex)

        # Don't escape special characters in math mode as they might be part of LaTeX commands
        # Just remove problematic equation numbering patterns
        import re

        # Remove equation numbers like #(2-1), #(3-4), etc.
        text = re.sub(r'#\([^)]+\)', '', text)

        # Remove standalone # that aren't part of LaTeX commands
        text = re.sub(r'(?<!\\)#(?![a-zA-Z])', '', text)

        return text

    def add_spaces_after_latex_commands(self, text):
        """Add spaces after LaTeX commands for proper formatting."""
        import re

        # List of LaTeX commands that should have spaces after them
        latex_commands = [
            r'\\geq', r'\\leq', r'\\neq', r'\\approx', r'\\equiv', r'\\propto', r'\\sim',
            r'\\in', r'\\notin', r'\\subset', r'\\subseteq', r'\\supset', r'\\supseteq',
            r'\\cup', r'\\cap', r'\\emptyset', r'\\forall', r'\\exists',
            r'\\rightarrow', r'\\leftarrow', r'\\leftrightarrow', r'\\Rightarrow',
            r'\\Leftarrow', r'\\Leftrightarrow', r'\\uparrow', r'\\downarrow', r'\\updownarrow',
            r'\\pm', r'\\mp', r'\\times', r'\\div', r'\\cdot', r'\\circ', r'\\sqrt',
            r'\\angle', r'\\perp', r'\\parallel', r'\\infty', r'\\partial', r'\\nabla',
            r'\\alpha', r'\\beta', r'\\gamma', r'\\delta', r'\\epsilon', r'\\zeta', r'\\eta',
            r'\\theta', r'\\iota', r'\\kappa', r'\\lambda', r'\\mu', r'\\nu', r'\\xi',
            r'\\pi', r'\\rho', r'\\sigma', r'\\tau', r'\\upsilon', r'\\phi', r'\\chi',
            r'\\psi', r'\\omega', r'\\Gamma', r'\\Delta', r'\\Theta', r'\\Lambda', r'\\Xi',
            r'\\Pi', r'\\Sigma', r'\\Upsilon', r'\\Phi', r'\\Psi', r'\\Omega'
        ]

        # Add space after LaTeX commands if not already present
        for cmd in latex_commands:
            # Pattern: command followed by non-space, non-brace character
            pattern = f'({cmd})(?=[a-zA-Z0-9])'
            text = re.sub(pattern, r'\1 ', text)

        return text
    
    def clean_latex_output(self, latex_text):
        """Clean and post-process LaTeX output."""
        if not latex_text:
            return latex_text

        import re

        # Remove equation numbers and references that cause issues
        # Pattern like #(2-1), #(3-4), #\left( 2−1 \right), etc.
        latex_text = re.sub(r'#\([^)]+\)', '', latex_text)
        latex_text = re.sub(r'#\\left\([^)]+\\right\)', '', latex_text)

        # Remove standalone # characters that aren't part of LaTeX commands
        latex_text = re.sub(r'(?<!\\)#(?![a-zA-Z])', '', latex_text)

        # Add proper spacing after LaTeX commands
        latex_text = self.add_spaces_after_latex_commands(latex_text)

        # Clean up extra spaces and commas at the end
        latex_text = re.sub(r'\s*,\s*$', '', latex_text)
        latex_text = re.sub(r'\s+', ' ', latex_text).strip()

        return latex_text

    def omml_to_latex(self, omml_element):
        """Main conversion function."""
        try:
            result = self.convert_element(omml_element)
            return self.clean_latex_output(result)
        except Exception as e:
            print(f"Error converting OMML to LaTeX: {e}")
            return "[Math Formula]"


def convert_omml_to_latex(omml_element):
    """Convenience function to convert OMML to LaTeX."""
    converter = OmmlToLatexConverter()
    return converter.omml_to_latex(omml_element)
