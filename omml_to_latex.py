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
            '~': '\\sim',  # ASCII tilde mapped to \sim (within math)
            # Additional mappings for calligraphic/blackboard symbols and variants used in formulas
            'ℒ': '\\mathcal{L}',  # Script L
            '𝒟': '\\mathcal{D}',  # Script D (uppercase)
            'ℰ': '\\mathbb{E}',  # Blackboard bold E (alternative)
            '𝔼': '\\mathbb{E}',  # Blackboard bold E (common)
            'ϕ': '\\varphi',      # Variant phi
        }
    
    def _get_attr(self, element, attr_name):
        """Helper to fetch an attribute value ignoring namespaces."""
        if attr_name in element.attrib:
            return element.attrib.get(attr_name)
        for key, val in element.attrib.items():
            if key.endswith('}' + attr_name):
                return val
        return None

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
        elif tag == 'sym':
            return self.convert_symbol(element)
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
        
        # Special-case expectation operator E to \mathbb{E}
        if base.strip('{}') == 'E':
            base = '\\mathbb{E}'
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
                        char = self._get_attr(prop_child, 'val') or ''
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
            'max': '\\operatorname*{max}',
            'min': '\\operatorname*{min}',
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
    
    def handle_conditional_probability(self, element):
        """Handle conditional probability expressions with vertical bar."""
        parts = []
        has_bar = False
        
        for child in element.iter():
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 't' and child.text == '|':
                has_bar = True
            elif tag in ['oMath', 'r', 't']:
                text = self.convert_element(child)
                if text:
                    parts.append(text)
        
        if has_bar and len(parts) >= 2:
            # Join parts around the vertical bar
            left_parts = []
            right_parts = []
            found_bar = False
            for part in parts:
                if '|' in part:
                    found_bar = True
                    sub_parts = part.split('|')
                    if len(sub_parts) == 2:
                        if left_parts:
                            left_parts.append(sub_parts[0])
                        else:
                            left_parts = [sub_parts[0]]
                        right_parts = [sub_parts[1]]
                elif not found_bar:
                    left_parts.append(part)
                else:
                    right_parts.append(part)
            
            left = ''.join(left_parts).strip()
            right = ''.join(right_parts).strip()
            return f"{left}\\mid {right}"
        
        return None

    def convert_delimiter(self, element):
        """Convert delimiter element."""
        # If not a conditional probability, proceed with normal delimiter handling
        left_delim = '('  # default
        right_delim = ')'

        # First, inspect properties for custom delimiters
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'dPr':
                for pr in child:
                    pr_tag = pr.tag.split('}')[-1] if '}' in pr.tag else pr.tag
                    if pr_tag == 'begChr':
                        left_delim = self._get_attr(pr, 'val') or left_delim
                    elif pr_tag == 'endChr':
                        right_delim = self._get_attr(pr, 'val') or right_delim

        # Check if there's a separator character (e.g., "|") between expressions
        sep_char = None
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'dPr':
                # sepChr may be an attribute of <m:dPr> *or* nested <m:sepChr> element
                sep_char = self._get_attr(child, 'sepChr') or sep_char
                for pr in child:
                    pr_tag = pr.tag.split('}')[-1] if '}' in pr.tag else pr.tag
                    if pr_tag == 'sepChr':
                        sep_char = self._get_attr(pr, 'val') or sep_char
                    elif pr_tag == 'val' and '|' in (self._get_attr(pr, 'val') or ''):
                        sep_char = '|'

        # Collect the expressions inside the delimiter
        expr_parts = []
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'e':
                expr = self.convert_element(child)
                # Check if this expression contains a vertical bar that should be treated as a separator
                if '|' in expr and not sep_char:
                    parts = expr.split('|')
                    if len(parts) == 2:  # Only split if there's exactly one vertical bar
                        expr_parts.extend(parts)
                        sep_char = '|'
                        continue
                expr_parts.append(expr)

        # Forced special handling for p_θ(y|x,I) patterns - this is a common case in ML papers
        # Check if we have exactly 2 expressions and no explicit separator
        if len(expr_parts) == 2 and not sep_char and left_delim == '(' and right_delim == ')':
            # Get the parent context to see if this is part of a probability expression
            parent_context = ""
            for p in element.iter():
                if p != element:  # Skip self
                    parent_text = p.text or ""
                    if parent_text:
                        parent_context += parent_text
            
            # Look for typical probability notations in parent context
            prob_indicators = ['p', 'P', 'Pr', 'θ', 'log']
            
            # First part is typically a single variable like y
            first_part = expr_parts[0].strip()
            # Second part often contains x, context, etc.
            second_part = expr_parts[1].strip()
            
            # If the first part is a single letter (like y) and second part contains typical variables
            if ((len(first_part) <= 2 and any(x in second_part for x in ['x', 'X', 'I', 'c'])) or
                (any(p in parent_context for p in prob_indicators))):
                # This looks like a conditional probability p(y|x)
                sep_char = '|'  # Force using vertical bar as separator

        # Join with separator if specified
        if sep_char:
            # Use \mid for vertical bar to get proper spacing; otherwise literal char
            latex_sep = ' \\mid ' if sep_char == '|' else f' {sep_char} '
            inner_expr = latex_sep.join(expr_parts)
        else:
            inner_expr = ''.join(expr_parts)

        # --- Fix 1: ensure delimiters that are special characters in LaTeX (like curly braces)
        #             are escaped when used after \left/\right -----------------------------
        def _escape_delim(ch):
            """Return a LaTeX-safe delimiter for use after \left/\right."""
            if not ch:
                return ''
            if ch in ['{', '}']:
                return f'\\{ch}'  # e.g. '{' -> '\\{'
            return ch  # other delimiters (|, (, ), [, ], etc.) are fine as-is

        left_delim_tex = _escape_delim(left_delim)
        right_delim_tex = _escape_delim(right_delim)

        # In OMML, an omitted right delimiter may be encoded as an empty string. In LaTeX we
        # need some delimiter after \right; a period (.) is the conventional choice for
        # an invisible delimiter.
        if right_delim_tex == '':
            right_delim_tex = '.'

        # LaTeX requires \left and \right before certain delimiters
        return f"\\left{left_delim_tex} {inner_expr} \\right{right_delim_tex}"
    
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
        
        # Detect common operators like max/min to use operatorname* with subscript
        base_stripped = base.strip('{}')
        if base_stripped in {'max', 'min'} and lim:
            # Remove any extra backslashes before operatorname
            base_stripped = re.sub(r'\\+', '', base_stripped)
            return f"\\operatorname*{{{base_stripped}}}_{{{lim}}}"
        else:
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

        # Special handling for vertical bar in math mode
        if text == '|':
            return '\\mid'

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
        # Note: The short command \\in is deliberately excluded to avoid interfering
        # with longer commands like \\infty or \\int. If needed, callers should
        # insert explicit spaces around \\in themselves.
        latex_commands = [
            r'\\rightarrow', r'\\leftarrow', r'\\leftrightarrow', r'\\Rightarrow',
            r'\\Leftarrow', r'\\Leftrightarrow', r'\\uparrow', r'\\downarrow', r'\\updownarrow',
            r'\\subseteq', r'\\supseteq', r'\\subset', r'\\supset',
            r'\\notin', r'\\neq', r'\\approx', r'\\equiv', r'\\propto',
            r'\\parallel', r'\\emptyset', r'\\forall', r'\\exists',
            r'\\geq', r'\\leq', r'\\pm', r'\\mp', r'\\times', r'\\div',
            r'\\cdot', r'\\circ', r'\\sqrt', r'\\angle', r'\\perp',
            r'\\infty', r'\\partial', r'\\nabla',
            # Greek letters and variants
            r'\\Gamma', r'\\Delta', r'\\Theta', r'\\Lambda', r'\\Xi', r'\\Pi',
            r'\\Sigma', r'\\Upsilon', r'\\Phi', r'\\Psi', r'\\Omega',
            r'\\alpha', r'\\beta', r'\\gamma', r'\\delta', r'\\epsilon', r'\\zeta',
            r'\\eta', r'\\theta', r'\\iota', r'\\kappa', r'\\lambda', r'\\mu',
            r'\\nu', r'\\xi', r'\\pi', r'\\rho', r'\\sigma', r'\\tau',
            r'\\upsilon', r'\\phi', r'\\chi', r'\\psi', r'\\omega',
            r'\\cup', r'\\cap', r'\\sim'
        ]

        # Process longer commands first to reduce partial-match issues
        latex_commands.sort(key=len, reverse=True)

        # Add space after LaTeX commands if they are immediately followed by
        # an alphanumeric character *and* the command itself is not a prefix
        # of a longer command (handled above via ordering and exclusion).
        for cmd in latex_commands:
            pattern = f'({cmd})(?=[a-zA-Z0-9])'
            text = re.sub(pattern, r'\1 ', text)

        # Special-case: ensure a space after membership operator "\\in" when followed by
        # an uppercase identifier (e.g. "\\inD" -> "\\in D").  This will *not* match
        # when the next letters form longer commands like "\\infty" or "\\int" because
        # they start with lowercase letters.
        text = re.sub(r'\\in([A-Z])', r'\\in \1', text)

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

        # Fix double backslashes in LaTeX commands (except for line breaks)
        latex_text = re.sub(r'\\\\(?!\\|$)', r'\\', latex_text)

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

    def convert_symbol(self, element):
        """Convert <m:sym> elements that contain a single symbol specified via the 'char' attribute."""
        char_val = self._get_attr(element, 'char') or ''
        if not char_val:
            return ''
        # Map to LaTeX if available
        return self.symbol_map.get(char_val, char_val)


def convert_omml_to_latex(omml_element):
    """Convenience function to convert OMML to LaTeX."""
    converter = OmmlToLatexConverter()
    return converter.omml_to_latex(omml_element)
