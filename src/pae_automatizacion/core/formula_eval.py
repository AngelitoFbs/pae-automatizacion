"""
Simple formula evaluator for Excel formulas commonly used in PAE templates.
Supports basic arithmetic, SUM, and cell references.
"""
import re
from typing import Dict, Any, Optional
import openpyxl


class FormulaEvaluator:
    """Simple formula evaluator for Excel formulas in PAE templates."""
    
    def __init__(self, worksheet):
        self.ws = worksheet
        self._cache = {}
    
    def evaluate(self, formula: str) -> Any:
        """Evaluate a formula string and return the computed value."""
        if not formula or not formula.startswith('='):
            return formula
        
        # Check cache
        if formula in self._cache:
            return self._cache[formula]
        
        # Remove = prefix
        expr = formula[1:].strip()
        
        # Handle common functions
        result = self._evaluate_expression(expr)
        
        # Cache and return
        self._cache[formula] = result
        return result
    
    def _evaluate_expression(self, expr: str) -> Any:
        """Evaluate an expression with cell references and basic operations."""
        # Replace cell references with their values
        # Pattern: $A$1, A1, $A1, A$1, etc.
        def replace_ref(match):
            ref = match.group(0)
            try:
                return str(self._get_cell_value(ref))
            except:
                return '0'
        
        # Replace cell references
        expr = re.sub(r'\$?[A-Z]{1,3}\$?\d+', replace_ref, expr)
        
        # Handle SUM function: SUM(A1:B10) or SUM(A1, B2)
        def replace_sum(match):
            args = match.group(1)
            refs = [r.strip() for r in args.split(',')]
            total = 0
            for ref in refs:
                if ':' in ref:
                    # Range like A1:B10
                    total += self._sum_range(ref)
                else:
                    try:
                        total += float(self._get_cell_value(ref))
                    except:
                        pass
            return str(total)
        
        expr = re.sub(r'SUM\(([^)]+)\)', replace_sum, expr, flags=re.IGNORECASE)
        
        # Evaluate the resulting arithmetic expression
        try:
            # Only allow safe operations
            allowed = set('0123456789.+-*/() ')
            if all(c in allowed for c in expr):
                return eval(expr)
        except:
            pass
        
        return 0
    
    def _get_cell_value(self, ref: str) -> float:
        """Get numeric value of a cell reference like A1, $A$1, etc."""
        # Clean reference
        ref = ref.replace('$', '')
        match = re.match(r'([A-Z]+)(\d+)', ref)
        if not match:
            return 0
        col_letters, row_str = match.groups()
        col = 0
        for ch in col_letters:
            col = col * 26 + (ord(ch) - ord('A') + 1)
        row = int(row_str)
        
        val = self.ws.cell(row=int(row_str), column=col).value
        if val is None:
            return 0
        try:
            return float(val)
        except (ValueError, TypeError):
            return 0
    
    def _sum_range(self, range_ref: str) -> float:
        """Sum a range like A1:B10."""
        match = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', range_ref)
        if not match:
            return 0
        col1, row1, col2, row2 = match.groups()
        
        def col_to_num(col):
            num = 0
            for ch in col:
                num = num * 26 + (ord(ch) - ord('A') + 1)
            return num
        
        c1 = col_to_num(col1)
        c2 = col_to_num(col2)
        r1 = int(row1)
        r2 = int(row2)
        
        total = 0
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                val = self.ws.cell(row=r, column=c).value
                try:
                    total += float(val) if val is not None else 0
                except:
                    pass
        return total


def evaluate_cell_value(ws, row: int, col: int) -> Any:
    """Get evaluated value of a cell, evaluating formulas if needed."""
    cell = ws.cell(row=row, column=col)
    val = cell.value
    
    if val is None:
        return None
    
    # If it's already a value (not formula), return it
    if not isinstance(val, str) or not val.startswith('='):
        return val
    
    # Evaluate formula
    evaluator = FormulaEvaluator(worksheet)
    return evaluator.evaluate(val)


# Monkey-patch for easy use
def patch_worksheet_for_eval(ws):
    """Add evaluate_value method to worksheet."""
    ws.evaluate_value = lambda row, col: evaluate_cell_value(ws, row, col)
    return ws