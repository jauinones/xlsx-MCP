# Excel MCP Server

A Model Context Protocol (MCP) server for Excel file manipulation with real-time formula calculation.

## Why a Formula Engine?

**Problem**: Both humans and LLMs are inconsistent at math. When an AI writes `=SUM(A1:A10)` to a spreadsheet, it's just storing a string—the calculation never happens unless Microsoft Excel opens the file.

**Solution**: This MCP integrates [HyperFormula](https://hyperformula.handsontable.com/), a headless spreadsheet engine that evaluates formulas in real-time. When you read a cell, you get the **calculated result**, not just the formula text.

```
write_cell(wb, "A1", 100)
write_cell(wb, "A2", 200)
write_cell(wb, "A3", "=SUM(A1:A2)")

read_cell(wb, "A3")
→ { calculatedValue: 300, formula: "=SUM(A1:A2)" }
```

This ensures mathematical accuracy for:
- Financial calculations
- Data aggregations
- Report generation
- Any workflow where precision matters

## Features

- **No Excel Required** - Works with .xlsx files using pure JavaScript
- **400+ Excel Functions** - SUM, AVERAGE, VLOOKUP, IF, and more via HyperFormula
- **Multi-Workbook** - Open and manipulate multiple workbooks simultaneously

## Tools (21 total)

### Workbook Management
| Tool | Description |
|------|-------------|
| `create_workbook` | Create new workbook in memory, returns ID |
| `open_workbook` | Open existing .xlsx file |
| `save_workbook` | Save workbook to disk |
| `close_workbook` | Close workbook and free memory |
| `list_workbooks` | List all open workbooks |

### Sheet Operations
| Tool | Description |
|------|-------------|
| `list_sheets` | List all sheets in workbook |
| `create_sheet` | Add new sheet |
| `delete_sheet` | Remove sheet |
| `rename_sheet` | Rename sheet |
| `get_sheet_info` | Get dimensions and metadata |

### Cell Operations
| Tool | Description |
|------|-------------|
| `read_cell` | Read cell value (returns calculated result) |
| `write_cell` | Write value or formula (prefix with `=`) |
| `read_range` | Read range of cells |
| `write_range` | Write 2D array to range |
| `list_columns` | List columns with headers |

### Formula Support
| Tool | Description |
|------|-------------|
| `get_formula` | Get raw formula string from cell |
| `recalculate` | Force recalculation of all formulas |

### Charts & Pivot Tables
| Tool | Description |
|------|-------------|
| `create_chart` | Create bar, line, pie, or scatter chart |
| `delete_chart` | Remove chart |
| `create_pivot_table` | Create pivot table from data |

### Import
| Tool | Description |
|------|-------------|
| `import_markdown_table` | Import markdown table to Excel |

## Installation

```bash
git clone https://github.com/jauinones/excel-mcp.git
cd excel-mcp
npm install
npm run build
```

## Configuration

Add to your Claude Code MCP settings (`~/.claude.json`):

```json
{
  "mcpServers": {
    "excel": {
      "type": "stdio",
      "command": "node",
      "args": ["/path/to/excel-mcp/dist/index.js"]
    }
  }
}
```

## Example Usage

```javascript
// Create a simple invoice
create_workbook → "wb_1"

write_cell(wb_1, "Sheet1", "A1", "Item")
write_cell(wb_1, "Sheet1", "B1", "Price")
write_cell(wb_1, "Sheet1", "C1", "Qty")
write_cell(wb_1, "Sheet1", "D1", "Total")

write_range(wb_1, "Sheet1", "A2", [
  ["Widget", 25.00, 10, "=B2*C2"],
  ["Gadget", 15.50, 20, "=B3*C3"],
  ["Thing",  8.99,  50, "=B4*C4"]
])

write_cell(wb_1, "Sheet1", "D5", "=SUM(D2:D4)")

read_cell(wb_1, "Sheet1", "D5")
// → { calculatedValue: 1009.5, formula: "=SUM(D2:D4)" }

save_workbook(wb_1, "/path/to/invoice.xlsx")
```

## Tech Stack

- **TypeScript** - Type-safe implementation
- **ExcelJS** - Read/write .xlsx files
- **HyperFormula** - Formula calculation engine (GPLv3)
- **MCP SDK** - Model Context Protocol integration

## License

GPLv3 (due to HyperFormula dependency)
