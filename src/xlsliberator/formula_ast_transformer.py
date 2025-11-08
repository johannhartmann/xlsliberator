"""AST-based formula transformation using Lark parser.

This module uses Lark to parse LibreOffice Calc formulas, transform INDIRECT(ADDRESS(...))
patterns to OFFSET(...), and rebuild the formula.
"""

from lark import Lark, Token, Transformer, Tree
from loguru import logger

CALC_FORMULA_GRAMMAR = r"""
    ?start: "=" expr

    ?expr: comparison

    ?comparison: term
        | comparison "=" term     -> eq
        | comparison "<>" term    -> ne
        | comparison "<" term     -> lt
        | comparison "<=" term    -> le
        | comparison ">" term     -> gt
        | comparison ">=" term    -> ge

    ?term: factor
        | term "+" factor         -> add
        | term "-" factor         -> sub
        | term "&" factor         -> concat

    ?factor: power
        | factor "*" power        -> mul
        | factor "/" power        -> div

    ?power: atom
        | atom "^" power          -> pow

    ?atom: NUMBER               -> number
        | STRING                -> string
        | cell_ref              -> cell
        | function_call
        | "(" expr ")"

    function_call: NAME "(" [expr (";" expr)*] ")"

    cell_ref: /\$?[A-Z]+\$?[0-9]+/
            | /\$?[A-Z]+\$?[0-9]+:\$?[A-Z]+\$?[0-9]+/
            | /\$?'?[\w\-]+'?\.\$?[A-Z]+\$?[0-9]+/

    STRING: /"[^"]*"/
    NUMBER: /\-?\d+(\.\d+)?/
    NAME: /[A-Z_][A-Z0-9_]*/i

    %import common.WS
    %ignore WS
"""


class IndirectAddressTransformer(Transformer):
    """Transforms INDIRECT(ADDRESS(...)) to OFFSET(...)."""

    def __init__(self, sheet_mapping: dict[str, str] | None = None):
        super().__init__()
        self.sheet_mapping = sheet_mapping or {}

    def function_call(self, children: list) -> Tree:
        """Transform function_call nodes.

        Args:
            children: [Token(name), *args] where args are Tree objects

        Returns:
            Tree object (possibly transformed)
        """
        if not children:
            return Tree("function_call", [])

        name_token = children[0]
        name = str(name_token).upper()
        args = children[1:]

        # Check for INDIRECT(ADDRESS(...))
        if name == "INDIRECT" and len(args) == 1:
            arg = args[0]
            # Check if arg is an ADDRESS function call
            if isinstance(arg, Tree) and arg.data == "function_call":
                addr_children = arg.children
                if addr_children and str(addr_children[0]).upper() == "ADDRESS":
                    # Found it! Transform
                    return self._transform_indirect_address(addr_children)

        # Return unchanged
        return Tree("function_call", children)

    def _transform_indirect_address(self, address_children: list) -> Tree:
        """Transform ADDRESS(...) to OFFSET(...).

        Args:
            address_children: [Token(ADDRESS), row, col, abs, a1, sheet]

        Returns:
            Tree for OFFSET function call
        """
        if len(address_children) < 6:
            logger.warning(f"ADDRESS has only {len(address_children) - 1} args, need 5. Skipping.")
            # Return INDIRECT(ADDRESS(...)) unchanged
            return Tree(
                "function_call",
                [Token("NAME", "INDIRECT"), Tree("function_call", address_children)],
            )

        # Extract arguments: ADDRESS(row, col, abs, a1, sheet)
        row_tree = address_children[1]
        col_tree = address_children[2]
        sheet_tree = address_children[5]

        # Extract sheet name from string tree
        if not (isinstance(sheet_tree, Tree) and sheet_tree.data == "string"):
            logger.warning(f"Sheet is not a string: {sheet_tree}. Skipping.")
            return Tree(
                "function_call",
                [Token("NAME", "INDIRECT"), Tree("function_call", address_children)],
            )

        # Get the STRING token value
        sheet_token = sheet_tree.children[0]
        sheet_name = str(sheet_token).strip('"')

        # Get quoted sheet reference
        if sheet_name in self.sheet_mapping:
            sheet_ref = self.sheet_mapping[sheet_name]
        else:
            needs_quote = sheet_name[0].isdigit() or "-" in sheet_name or " " in sheet_name
            sheet_ref = f"'{sheet_name}'" if needs_quote else sheet_name

        # Build OFFSET(Sheet.A1, row-1, col-1)
        # ADDRESS is 1-indexed, OFFSET is 0-indexed
        base_ref = Tree("cell", [Token("__ANON_0", f"{sheet_ref}.A1")])

        # Create (row - 1) and (col - 1) trees
        offset_row = Tree("sub", [row_tree, Tree("number", [Token("NUMBER", "1")])])
        offset_col = Tree("sub", [col_tree, Tree("number", [Token("NUMBER", "1")])])

        logger.debug(
            f"Transformed INDIRECT(ADDRESS(..., {sheet_name})) → OFFSET({sheet_ref}.A1, ...)"
        )

        return Tree("function_call", [Token("NAME", "OFFSET"), base_ref, offset_row, offset_col])


def needs_parens(tree: Tree, parent_op: str | None = None) -> bool:
    """Check if expression needs parentheses based on operator precedence."""
    if not isinstance(tree, Tree):
        return False

    # Operators by precedence (lower number = lower precedence)
    precedence = {
        "eq": 1,
        "ne": 1,
        "lt": 1,
        "le": 1,
        "gt": 1,
        "ge": 1,
        "concat": 2,
        "add": 3,
        "sub": 3,
        "mul": 4,
        "div": 4,
        "pow": 5,
    }

    if tree.data not in precedence or parent_op not in precedence:
        return False

    return precedence[tree.data] < precedence[parent_op]


def tree_to_formula(tree: Tree | Token, parent_op: str | None = None) -> str:
    """Rebuild formula string from Lark tree.

    Args:
        tree: Lark Tree or Token
        parent_op: Parent operator (for precedence checking)

    Returns:
        Formula string (with semicolons)
    """
    if isinstance(tree, Token):
        return str(tree)

    if not isinstance(tree, Tree):
        return str(tree)

    if tree.data in ("number", "string"):
        # These nodes have a single Token child
        return str(tree.children[0])
    elif tree.data == "cell":
        # Cell may have a cell_ref child tree or a Token
        child = tree.children[0]
        if isinstance(child, Tree) and child.data == "cell_ref":
            # Nested: cell -> cell_ref -> Token
            return str(child.children[0])
        else:
            # Direct Token
            return str(child)
    elif tree.data == "cell_ref":
        # cell_ref has a Token child
        return str(tree.children[0])
    elif tree.data == "function_call":
        name = str(tree.children[0])
        args = tree.children[1:]
        # Filter out None args (empty function calls)
        args_str = ";".join(tree_to_formula(arg) for arg in args if arg is not None)
        return f"{name}({args_str})"
    elif tree.data in (
        "add",
        "sub",
        "mul",
        "div",
        "pow",
        "concat",
        "eq",
        "ne",
        "lt",
        "le",
        "gt",
        "ge",
    ):
        left = tree_to_formula(tree.children[0], tree.data)
        right = tree_to_formula(tree.children[1], tree.data)
        ops = {
            "add": "+",
            "sub": "-",
            "mul": "*",
            "div": "/",
            "pow": "^",
            "concat": "&",
            "eq": "=",
            "ne": "<>",
            "lt": "<",
            "le": "<=",
            "gt": ">",
            "ge": ">=",
        }
        op = ops[tree.data]

        # Only add parentheses if needed based on precedence
        if needs_parens(tree, parent_op):
            return f"({left}{op}{right})"
        else:
            return f"{left}{op}{right}"
    else:
        # Default: recursively process children
        if tree.children:
            return "".join(tree_to_formula(child) for child in tree.children)
        return ""


class FormulaASTTransformer:
    """Main interface for formula transformation."""

    def __init__(self, sheet_mapping: dict[str, str] | None = None):
        self.sheet_mapping = sheet_mapping or {}
        self.parser = Lark(CALC_FORMULA_GRAMMAR, start="start", parser="lalr")

    def transform_indirect_address_to_offset(self, formula: str) -> str:
        """Transform INDIRECT(ADDRESS(...)) to OFFSET(...).

        Args:
            formula: LibreOffice Calc formula (with semicolons)

        Returns:
            Transformed formula

        Raises:
            FormulaTransformError: If parsing fails
        """
        try:
            logger.debug(f"Parsing: {formula[:80]}...")
            tree = self.parser.parse(formula)

            transformer = IndirectAddressTransformer(self.sheet_mapping)
            transformed = transformer.transform(tree)

            result = "=" + tree_to_formula(transformed)
            logger.debug(f"Result: {result[:80]}...")
            return result

        except Exception as e:
            logger.error(f"Transform failed: {e}")
            raise FormulaTransformError(f"Failed: {e}") from e


class FormulaTransformError(Exception):
    """Raised when transformation fails."""

    pass
