# XLOOKUP for Excel 2019

This is an open-source implementation of the `XLOOKUP` function for Excel 2019 (and earlier versions) using C# and Excel-DNA.

## Features
- **Native Performance**: Built as a high-performance C++ XLL add-in.
- **Full Compatibility**: Supports all standard `XLOOKUP` arguments:
  - Lookup Value
  - Lookup Array
  - Return Array
  - If Not Found
  - Match Mode (0, -1, 1, 2)
  - Search Mode (1, -1)
- **Zero Dependencies**: Uses .NET Framework 4.7.2 (built-in to Windows).

## Installation

1. Download the latest release from the `Release` folder.
2. Open Excel -> File -> Options -> Add-ins.
3. Manage: Excel Add-ins -> Go...
4. Browse -> Select `XLookupAddIn-AddIn64.xll`.
5. Done! Use `=XLOOKUP(...)` in any cell.

## Build From Source

1. Clone the repository.
2. Open in Visual Studio or VS Code.
3. Run `dotnet build -c Release`.
