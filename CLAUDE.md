# CLAUDE.md

This file provides guidance to Claude Code when working with code in this repository.

## Project Overview

This is a Model Context Protocol (MCP) server for Excel file manipulation. It provides 20 tools for working with Excel workbooks without needing Microsoft Excel installed.

**Key Feature**: Integrates [HyperFormula](https://hyperformula.handsontable.com/) for real-time formula calculation. Formulas like `=SUM(A1:A10)` are evaluated immediately, returning calculated results to agents.

## Tools Available (20 total)

### Workbook Management
- `create_workbook` - Create new workbook in memory, returns workbook ID
- `open_workbook` - Open existing .xlsx file by path
- `save_workbook` - Save workbook to file path
- `close_workbook` - Close workbook and free memory
- `list_workbooks` - List all open workbooks

### Sheet Operations
- `list_sheets` - List all sheets in a workbook
- `create_sheet` - Add new sheet with optional name
- `delete_sheet` - Remove sheet by name or index
- `rename_sheet` - Rename a sheet
- `get_sheet_info` - Get dimensions, row/column counts

### Cell Operations
- `read_cell` - Read value from cell (returns both raw value and **calculated value** from HyperFormula)
- `write_cell` - Write value or formula (values starting with `=` are formulas)
- `read_range` - Read values from range (returns **calculated values**)
- `write_range` - Write 2D array to range
- `list_columns` - List columns with headers from first row

### Formula Support
- `get_formula` - Get formula string from cell (returns `=SUM(...)` not result)
- `recalculate` - Force recalculation of all formulas using HyperFormula

### Charts
- `create_chart` - Create bar, line, pie, or scatter chart
- `delete_chart` - Remove chart by name

### Pivot Tables
- `create_pivot_table` - Create pivot table from data range

## Formula Calculation Engine

This MCP uses **HyperFormula** (GPLv3 license) for formula evaluation:

- **400+ Excel functions** supported (SUM, AVERAGE, VLOOKUP, IF, etc.)
- **Real-time calculation** - formulas evaluate immediately when cells change
- **No Excel required** - calculations happen in-process via HyperFormula's AST parser
- **Secure** - uses grammar-based parsing, no `eval()` or code injection risk

### How it works:
1. Each workbook has a paired HyperFormula instance
2. When you write to a cell, both ExcelJS and HyperFormula are updated
3. When you read a cell, `calculatedValue` comes from HyperFormula
4. When you save, calculated results are synced back to the .xlsx file

## Development Commands

- `npm run build` - Compile TypeScript to JavaScript in dist/
- `npm run dev` - Run in development mode with tsx
- `npm start` - Run the compiled version from dist/
- `npm run clean` - Remove dist/ directory

## Architecture

### Core Components
- `src/index.ts` - Main MCP server entry point, handles tool registration and request routing
- `src/excel-client.ts` - Excel operations class using exceljs + HyperFormula

### Dependencies
- `exceljs` - Read/write .xlsx files
- `hyperformula` - Formula calculation engine (GPLv3)
- `@modelcontextprotocol/sdk` - MCP protocol implementation

### MCP Integration
The server implements the Model Context Protocol using `@modelcontextprotocol/sdk`:
- Uses StdioServerTransport for communication
- Registers tools via ListToolsRequestSchema handler
- Processes tool calls via CallToolRequestSchema handler

### State Management
- Workbooks are kept in memory using a Map with unique IDs (wb_1, wb_2, etc.)
- Each workbook has a paired HyperFormula instance for calculation
- Multiple workbooks can be open simultaneously

## Usage Pattern

1. Create or open a workbook to get a workbook ID
2. Use the workbook ID with other tools to manipulate sheets and cells
3. Formulas are calculated automatically when reading cells
4. Save the workbook to persist changes (including calculated results)
5. Close the workbook when done to free memory

## Example Workflow

```
create_workbook → wb_1
write_cell(wb_1, "Sheet1", "A1", "Price")
write_cell(wb_1, "Sheet1", "B1", "Qty")
write_cell(wb_1, "Sheet1", "C1", "Total")
write_range(wb_1, "Sheet1", "A2", [[100, 5, "=A2*B2"], [200, 3, "=A3*B3"]])
write_cell(wb_1, "Sheet1", "C4", "=SUM(C2:C3)")

read_cell(wb_1, "Sheet1", "C4")
# Returns: { value: {formula: "SUM(C2:C3)"}, calculatedValue: 1100, formula: "=SUM(C2:C3)" }

save_workbook(wb_1, "/path/to/file.xlsx")
close_workbook(wb_1)
```
