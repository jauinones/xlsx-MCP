#!/usr/bin/env node

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { ExcelClient } from "./excel-client.js";

const server = new Server(
  {
    name: "excel-mcp",
    version: "1.0.0",
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

const excelClient = new ExcelClient();

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      // Workbook Management
      {
        name: "create_workbook",
        description: "Create a new Excel workbook in memory. Returns a workbook ID to use with other tools.",
        inputSchema: {
          type: "object",
          properties: {},
          required: [],
        },
      },
      {
        name: "open_workbook",
        description: "Open an existing Excel file (.xlsx) from disk",
        inputSchema: {
          type: "object",
          properties: {
            filePath: {
              type: "string",
              description: "The full path to the Excel file to open",
            },
          },
          required: ["filePath"],
        },
      },
      {
        name: "save_workbook",
        description: "Save a workbook to disk. If filePath is not provided, saves to the original location.",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "The workbook ID (e.g., 'wb_1')",
            },
            filePath: {
              type: "string",
              description: "The file path to save to (optional if workbook was opened from a file)",
            },
          },
          required: ["workbookId"],
        },
      },
      {
        name: "close_workbook",
        description: "Close a workbook and free memory. Does not save automatically.",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "The workbook ID to close",
            },
          },
          required: ["workbookId"],
        },
      },
      {
        name: "list_workbooks",
        description: "List all currently open workbooks with their IDs and file paths",
        inputSchema: {
          type: "object",
          properties: {},
          required: [],
        },
      },

      // Sheet Operations
      {
        name: "list_sheets",
        description: "List all sheets in a workbook with their names and indices",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "The workbook ID",
            },
          },
          required: ["workbookId"],
        },
      },
      {
        name: "create_sheet",
        description: "Create a new sheet in a workbook",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "The workbook ID",
            },
            name: {
              type: "string",
              description: "Name for the new sheet (optional, auto-generated if not provided)",
            },
          },
          required: ["workbookId"],
        },
      },
      {
        name: "delete_sheet",
        description: "Delete a sheet from a workbook",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "The workbook ID",
            },
            sheet: {
              type: "string",
              description: "Sheet name or index (1-based)",
            },
          },
          required: ["workbookId", "sheet"],
        },
      },
      {
        name: "rename_sheet",
        description: "Rename a sheet in a workbook",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "The workbook ID",
            },
            sheet: {
              type: "string",
              description: "Current sheet name or index (1-based)",
            },
            newName: {
              type: "string",
              description: "The new name for the sheet",
            },
          },
          required: ["workbookId", "sheet", "newName"],
        },
      },
      {
        name: "get_sheet_info",
        description: "Get information about a sheet including dimensions and row/column counts",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "The workbook ID",
            },
            sheet: {
              type: "string",
              description: "Sheet name or index (1-based)",
            },
          },
          required: ["workbookId", "sheet"],
        },
      },

      // Cell Operations
      {
        name: "read_cell",
        description: "Read the value from a single cell",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "The workbook ID",
            },
            sheet: {
              type: "string",
              description: "Sheet name or index (1-based)",
            },
            cell: {
              type: "string",
              description: "Cell address (e.g., 'A1', 'B5', 'AA100')",
            },
          },
          required: ["workbookId", "sheet", "cell"],
        },
      },
      {
        name: "write_cell",
        description: "Write a value or formula to a cell. Values starting with '=' are treated as formulas.",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "The workbook ID",
            },
            sheet: {
              type: "string",
              description: "Sheet name or index (1-based)",
            },
            cell: {
              type: "string",
              description: "Cell address (e.g., 'A1')",
            },
            value: {
              type: ["string", "number", "boolean"],
              description: "Value to write. Use '=' prefix for formulas (e.g., '=SUM(A1:A10)')",
            },
          },
          required: ["workbookId", "sheet", "cell", "value"],
        },
      },
      {
        name: "read_range",
        description: "Read values from a range of cells",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "The workbook ID",
            },
            sheet: {
              type: "string",
              description: "Sheet name or index (1-based)",
            },
            range: {
              type: "string",
              description: "Cell range (e.g., 'A1:C10', 'B2:D5')",
            },
          },
          required: ["workbookId", "sheet", "range"],
        },
      },
      {
        name: "write_range",
        description: "Write a 2D array of values to a range starting at a specific cell",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "The workbook ID",
            },
            sheet: {
              type: "string",
              description: "Sheet name or index (1-based)",
            },
            startCell: {
              type: "string",
              description: "Top-left cell address (e.g., 'A1')",
            },
            values: {
              type: "array",
              items: {
                type: "array",
                items: {
                  type: ["string", "number", "boolean", "null"],
                },
              },
              description: "2D array of values. Each inner array is a row. Use null for empty cells.",
            },
          },
          required: ["workbookId", "sheet", "startCell", "values"],
        },
      },
      {
        name: "list_columns",
        description: "List columns with headers from the first row of a sheet",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "The workbook ID",
            },
            sheet: {
              type: "string",
              description: "Sheet name or index (1-based)",
            },
          },
          required: ["workbookId", "sheet"],
        },
      },

      // Formula Support
      {
        name: "get_formula",
        description: "Get the formula string from a cell (returns the formula itself, not the calculated result)",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "The workbook ID",
            },
            sheet: {
              type: "string",
              description: "Sheet name or index (1-based)",
            },
            cell: {
              type: "string",
              description: "Cell address (e.g., 'A1')",
            },
          },
          required: ["workbookId", "sheet", "cell"],
        },
      },

      {
        name: "recalculate",
        description: "Force recalculation of all formulas in a workbook using HyperFormula engine",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "The workbook ID",
            },
          },
          required: ["workbookId"],
        },
      },

      // Charts
      {
        name: "create_chart",
        description: "Create a chart from data in a sheet",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "The workbook ID",
            },
            sheet: {
              type: "string",
              description: "Sheet name or index (1-based)",
            },
            type: {
              type: "string",
              enum: ["bar", "line", "pie", "scatter"],
              description: "Type of chart to create",
            },
            dataRange: {
              type: "string",
              description: "Range containing chart data (e.g., 'A1:B10')",
            },
            title: {
              type: "string",
              description: "Chart title (optional)",
            },
          },
          required: ["workbookId", "sheet", "type", "dataRange"],
        },
      },
      {
        name: "delete_chart",
        description: "Delete a chart by its title/name",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "The workbook ID",
            },
            sheet: {
              type: "string",
              description: "Sheet name or index (1-based)",
            },
            chartName: {
              type: "string",
              description: "The chart title/name to delete",
            },
          },
          required: ["workbookId", "sheet", "chartName"],
        },
      },

      // Import
      {
        name: "import_markdown_table",
        description: "Import a markdown table into an Excel workbook. Creates a new workbook or uses existing one.",
        inputSchema: {
          type: "object",
          properties: {
            markdown: {
              type: "string",
              description: "The markdown table string to import (with | delimiters)",
            },
            workbookId: {
              type: "string",
              description: "Optional workbook ID to import into. Creates new workbook if not provided.",
            },
            sheetName: {
              type: "string",
              description: "Optional sheet name for the imported data (defaults to 'Sheet1')",
            },
          },
          required: ["markdown"],
        },
      },

      // Pivot Tables
      {
        name: "create_pivot_table",
        description: "Create a pivot table from source data. Creates a new sheet with the aggregated results.",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "The workbook ID",
            },
            sourceSheet: {
              type: "string",
              description: "Sheet containing source data (name or index)",
            },
            sourceRange: {
              type: "string",
              description: "Range of source data including headers (e.g., 'A1:D100')",
            },
            rowFields: {
              type: "array",
              items: { type: "string" },
              description: "Column headers to use as row grouping fields",
            },
            dataFields: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  field: {
                    type: "string",
                    description: "Column header to aggregate",
                  },
                  aggregation: {
                    type: "string",
                    enum: ["sum", "count", "average", "min", "max"],
                    description: "Aggregation function",
                  },
                },
                required: ["field", "aggregation"],
              },
              description: "Fields to aggregate with their aggregation functions",
            },
            destinationSheet: {
              type: "string",
              description: "Name for the pivot table sheet (optional, auto-generated if not provided)",
            },
          },
          required: ["workbookId", "sourceSheet", "sourceRange", "rowFields", "dataFields"],
        },
      },
    ],
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    switch (name) {
      // Workbook Management
      case "create_workbook": {
        const id = excelClient.createWorkbook();
        return {
          content: [{ type: "text", text: JSON.stringify({ workbookId: id }) }],
        };
      }

      case "open_workbook": {
        const { filePath } = args as { filePath: string };
        const id = await excelClient.openWorkbook(filePath);
        return {
          content: [{ type: "text", text: JSON.stringify({ workbookId: id, filePath }) }],
        };
      }

      case "save_workbook": {
        const { workbookId, filePath } = args as { workbookId: string; filePath?: string };
        const savedPath = await excelClient.saveWorkbook(workbookId, filePath);
        return {
          content: [{ type: "text", text: JSON.stringify({ saved: true, filePath: savedPath }) }],
        };
      }

      case "close_workbook": {
        const { workbookId } = args as { workbookId: string };
        excelClient.closeWorkbook(workbookId);
        return {
          content: [{ type: "text", text: JSON.stringify({ closed: true, workbookId }) }],
        };
      }

      case "list_workbooks": {
        const workbooks = excelClient.listWorkbooks();
        return {
          content: [{ type: "text", text: JSON.stringify(workbooks) }],
        };
      }

      // Sheet Operations
      case "list_sheets": {
        const { workbookId } = args as { workbookId: string };
        const sheets = excelClient.listSheets(workbookId);
        return {
          content: [{ type: "text", text: JSON.stringify(sheets) }],
        };
      }

      case "create_sheet": {
        const { workbookId, name: sheetName } = args as { workbookId: string; name?: string };
        const result = excelClient.createSheet(workbookId, sheetName);
        return {
          content: [{ type: "text", text: JSON.stringify(result) }],
        };
      }

      case "delete_sheet": {
        const { workbookId, sheet } = args as { workbookId: string; sheet: string };
        const sheetId = isNaN(Number(sheet)) ? sheet : Number(sheet);
        excelClient.deleteSheet(workbookId, sheetId);
        return {
          content: [{ type: "text", text: JSON.stringify({ deleted: true, sheet }) }],
        };
      }

      case "rename_sheet": {
        const { workbookId, sheet, newName } = args as { workbookId: string; sheet: string; newName: string };
        const sheetId = isNaN(Number(sheet)) ? sheet : Number(sheet);
        excelClient.renameSheet(workbookId, sheetId, newName);
        return {
          content: [{ type: "text", text: JSON.stringify({ renamed: true, oldName: sheet, newName }) }],
        };
      }

      case "get_sheet_info": {
        const { workbookId, sheet } = args as { workbookId: string; sheet: string };
        const sheetId = isNaN(Number(sheet)) ? sheet : Number(sheet);
        const info = excelClient.getSheetInfo(workbookId, sheetId);
        return {
          content: [{ type: "text", text: JSON.stringify(info) }],
        };
      }

      // Cell Operations
      case "read_cell": {
        const { workbookId, sheet, cell } = args as { workbookId: string; sheet: string; cell: string };
        const sheetId = isNaN(Number(sheet)) ? sheet : Number(sheet);
        const result = excelClient.readCell(workbookId, sheetId, cell);
        return {
          content: [{ type: "text", text: JSON.stringify(result) }],
        };
      }

      case "write_cell": {
        const { workbookId, sheet, cell, value } = args as {
          workbookId: string;
          sheet: string;
          cell: string;
          value: string | number | boolean;
        };
        const sheetId = isNaN(Number(sheet)) ? sheet : Number(sheet);
        excelClient.writeCell(workbookId, sheetId, cell, value);
        return {
          content: [{ type: "text", text: JSON.stringify({ written: true, cell, value }) }],
        };
      }

      case "read_range": {
        const { workbookId, sheet, range } = args as { workbookId: string; sheet: string; range: string };
        const sheetId = isNaN(Number(sheet)) ? sheet : Number(sheet);
        const result = excelClient.readRange(workbookId, sheetId, range);
        return {
          content: [{ type: "text", text: JSON.stringify(result) }],
        };
      }

      case "write_range": {
        const { workbookId, sheet, startCell, values } = args as {
          workbookId: string;
          sheet: string;
          startCell: string;
          values: Array<Array<string | number | boolean | null>>;
        };
        const sheetId = isNaN(Number(sheet)) ? sheet : Number(sheet);
        excelClient.writeRange(workbookId, sheetId, startCell, values);
        return {
          content: [{ type: "text", text: JSON.stringify({ written: true, startCell, rowCount: values.length }) }],
        };
      }

      case "list_columns": {
        const { workbookId, sheet } = args as { workbookId: string; sheet: string };
        const sheetId = isNaN(Number(sheet)) ? sheet : Number(sheet);
        const columns = excelClient.listColumns(workbookId, sheetId);
        return {
          content: [{ type: "text", text: JSON.stringify(columns) }],
        };
      }

      // Formula Support
      case "get_formula": {
        const { workbookId, sheet, cell } = args as { workbookId: string; sheet: string; cell: string };
        const sheetId = isNaN(Number(sheet)) ? sheet : Number(sheet);
        const formula = excelClient.getFormula(workbookId, sheetId, cell);
        return {
          content: [{ type: "text", text: JSON.stringify({ cell, formula }) }],
        };
      }

      case "recalculate": {
        const { workbookId } = args as { workbookId: string };
        const result = excelClient.recalculate(workbookId);
        return {
          content: [{ type: "text", text: JSON.stringify({ recalculated: true, ...result }) }],
        };
      }

      // Charts
      case "create_chart": {
        const { workbookId, sheet, type, dataRange, title } = args as {
          workbookId: string;
          sheet: string;
          type: "bar" | "line" | "pie" | "scatter";
          dataRange: string;
          title?: string;
        };
        const sheetId = isNaN(Number(sheet)) ? sheet : Number(sheet);
        excelClient.createChart(workbookId, sheetId, { type, dataRange, title });
        return {
          content: [{ type: "text", text: JSON.stringify({ created: true, type, dataRange, title }) }],
        };
      }

      case "delete_chart": {
        const { workbookId, sheet, chartName } = args as { workbookId: string; sheet: string; chartName: string };
        const sheetId = isNaN(Number(sheet)) ? sheet : Number(sheet);
        excelClient.deleteChart(workbookId, sheetId, chartName);
        return {
          content: [{ type: "text", text: JSON.stringify({ deleted: true, chartName }) }],
        };
      }

      // Pivot Tables
      case "create_pivot_table": {
        const { workbookId, sourceSheet, sourceRange, rowFields, dataFields, destinationSheet } = args as {
          workbookId: string;
          sourceSheet: string;
          sourceRange: string;
          rowFields: string[];
          dataFields: Array<{ field: string; aggregation: "sum" | "count" | "average" | "min" | "max" }>;
          destinationSheet?: string;
        };
        const sheetId = isNaN(Number(sourceSheet)) ? sourceSheet : Number(sourceSheet);
        const result = excelClient.createPivotTable(workbookId, sheetId, {
          sourceRange,
          rowFields,
          dataFields,
          destinationSheet,
        });
        return {
          content: [{ type: "text", text: JSON.stringify({ created: true, ...result }) }],
        };
      }

      // Import
      case "import_markdown_table": {
        const { markdown, workbookId, sheetName } = args as {
          markdown: string;
          workbookId?: string;
          sheetName?: string;
        };
        const result = excelClient.importMarkdownTable(markdown, workbookId, sheetName);
        return {
          content: [{ type: "text", text: JSON.stringify({ imported: true, ...result }) }],
        };
      }

      default:
        throw new Error(`Unknown tool: ${name}`);
    }
  } catch (error) {
    return {
      content: [
        {
          type: "text",
          text: `Error: ${error instanceof Error ? error.message : String(error)}`,
        },
      ],
      isError: true,
    };
  }
});

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((error) => {
  console.error("Server error:", error);
  process.exit(1);
});
