from __future__ import annotations

import ast
import re
from pathlib import Path
from typing import Any


FUNCTION_NAMES = {"clean_text", "_company_lookup_key", "_company_name_key", "parse_peer_company_queries"}


def _load_peer_query_parser() -> Any:
    app_path = Path("app.py")
    module = ast.parse(app_path.read_text())
    functions = [node for node in module.body if isinstance(node, ast.FunctionDef) and node.name in FUNCTION_NAMES]
    namespace: dict[str, Any] = {"Any": Any, "re": re}
    exec(compile(ast.Module(body=functions, type_ignores=[]), str(app_path), "exec"), namespace)
    return namespace["parse_peer_company_queries"]


def test_peer_entry_parser_preserves_company_names_and_splits_ticker_lists() -> None:
    parse_queries = _load_peer_query_parser()

    assert parse_queries("AAPL MSFT AMD") == ["AAPL", "MSFT", "AMD"]
    assert parse_queries("Lattice Semiconductor, Microchip Technology; AMD\nMarvell Technology") == [
        "Lattice Semiconductor",
        "Microchip Technology",
        "AMD",
        "Marvell Technology",
    ]


def test_peer_entry_parser_deduplicates_name_and_ticker_entries() -> None:
    parse_queries = _load_peer_query_parser()

    assert parse_queries("LSCC, lscc; Lattice Semiconductor") == ["LSCC", "Lattice Semiconductor"]


def test_company_name_key_normalizes_common_legal_suffixes() -> None:
    app_path = Path("app.py")
    module = ast.parse(app_path.read_text())
    functions = [node for node in module.body if isinstance(node, ast.FunctionDef) and node.name in FUNCTION_NAMES]
    namespace: dict[str, Any] = {"Any": Any, "re": re}
    exec(compile(ast.Module(body=functions, type_ignores=[]), str(app_path), "exec"), namespace)

    assert namespace["_company_name_key"]("Lattice Semiconductor Corporation") == namespace["_company_name_key"](
        "LATTICE SEMICONDUCTOR CORP"
    )
