"""LangChain tools that wrap local Excel file operations."""

import json
from langchain_core.tools import tool
import graph_client


@tool
def list_files() -> str:
    """List all available Excel files in the data directory."""
    files = graph_client.list_files()
    return f"Available files: {', '.join(files)}"


@tool
def list_sheets(filename: str) -> str:
    """
    List all worksheet names in a specific Excel file.

    Args:
        filename: Excel file name (e.g. "invoice_jan_2026.xlsx")
    """
    sheets = graph_client.list_worksheets(filename)
    return f"Worksheets in {filename}: {', '.join(sheets)}"


@tool
def read_excel(filename: str, sheet_name: str, address: str) -> str:
    """
    Read data from an Excel range.

    Args:
        filename: Excel file name (e.g. "invoice_jan_2026.xlsx")
        sheet_name: Worksheet name (e.g. "Invoices")
        address: Cell range to read (e.g. "A1:E10")
    """
    data = graph_client.read_range(filename, sheet_name, address)
    return json.dumps(data, indent=2)


@tool
def write_excel(filename: str, sheet_name: str, address: str, values: str) -> str:
    """
    Write data to an Excel range. Overwrites existing content.

    Args:
        filename: Excel file name (e.g. "invoice_jan_2026.xlsx")
        sheet_name: Worksheet name (e.g. "Invoices")
        address: Cell range to write (e.g. "A1:B2"). Must match the shape of values.
        values: JSON string of 2D array, e.g. '[["Name","Score"],["Alice",95]]'
    """
    parsed = json.loads(values)
    graph_client.update_range(filename, sheet_name, address, parsed)
    return f"Successfully wrote to {filename} → {sheet_name}!{address}"


@tool
def append_rows(filename: str, sheet_name: str, start_column: str, rows: str) -> str:
    """
    Append rows at the bottom of existing data in a worksheet.

    Args:
        filename: Excel file name (e.g. "invoice_jan_2026.xlsx")
        sheet_name: Worksheet name
        start_column: First column letter (e.g. "A")
        rows: JSON string of 2D array of rows to add, e.g. '[["ORD-001","Laptop","2026-01-15","2026-01-20",999.99]]'
    """
    parsed = json.loads(rows)
    graph_client.add_rows(filename, sheet_name, start_column, parsed)
    return f"Appended {len(parsed)} row(s) to {filename} → {sheet_name}"


# All tools exposed to the agent
all_tools = [list_files, list_sheets, read_excel, write_excel, append_rows]
