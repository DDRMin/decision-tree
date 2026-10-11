# DecisionTree Studio

An offline Excel add-in for building and solving decision trees. It's a modern replacement for the BYTreePlan VBA add-in, adding reliable tree undo/redo, live Excel formulas, cell-linked inputs and import of existing BYTreePlan trees.

**Status:** Phase 0 (host checks). The add-in skeleton runs in Excel; see [docs/PHASE0.md](docs/PHASE0.md).

## Stack

- C# with [Excel-DNA](https://excel-dna.net). The add-in targets .NET Framework 4.8, which is built into Windows.
- Ribbon plus a WPF side panel; the diagram is drawn on the worksheet.
- Windows desktop Excel, 64-bit (Microsoft 365 / Office 2024).

## Layout

| Path | Purpose |
|---|---|
| `src/DecisionTree.Core` | Tree model, validation, calculation, undo history. No Excel dependency. |
| `src/DecisionTree.Excel` | The Excel-DNA add-in: ribbon, panel, Excel access, rendering. |
| `tests/DecisionTree.Core.Tests` | Unit and property tests for the core. |
| `tools/` | Scripts that check the add-in against a real Excel instance. |
| `docs/` | Phase notes and guides. |

## Build

Requires the [.NET 8 SDK](https://dotnet.microsoft.com/download).

```bash
dotnet build
dotnet test
```

The add-in is written to `src/DecisionTree.Excel/bin/Debug/net48/publish/DecisionTreeStudio64-packed.xll`. To load it in Excel, go to **File → Options → Add-ins → Manage: Excel Add-ins → Go… → Browse…**.
