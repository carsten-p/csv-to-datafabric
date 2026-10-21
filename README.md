# CSV → Data Fabric

Turns any CSV file into a UiPath Data Fabric entity and loads its rows, using the `uip` CLI.
Header names are fixed to Data Fabric's rules (originals kept as display names), column types are inferred from all rows, and rows are imported in chunks.

| File | Runs on |
|---|---|
| `Import-CsvToDataFabric.ps1` | Windows (PowerShell 5.1 / 7) |
| `import-csv-to-datafabric.mjs` | macOS, Linux, Windows (Node.js 18+) |

## Quick start

Requires the UiPath CLI, logged in (`uip login`), and `uip tools install @uipath/data-fabric-tool@latest`.

```powershell
.\Import-CsvToDataFabric.ps1 -CsvPath .\data.csv -DryRun
.\Import-CsvToDataFabric.ps1 -CsvPath .\data.csv
```

```bash
node import-csv-to-datafabric.mjs ./data.csv --dry-run
node import-csv-to-datafabric.mjs ./data.csv
```

All parameters and the column-name rules: download and open [`README.html`](README.html), or run `node import-csv-to-datafabric.mjs --help` / `Get-Help .\Import-CsvToDataFabric.ps1 -Full`.
