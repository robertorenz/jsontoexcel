# JsonToExcel

A command-line application that converts JSON files to Excel format (.xlsx) with optional professional formatting.

## Features

### Core Functionality
- Convert JSON arrays to Excel tables with column headers
- Convert JSON objects to key-value pair tables
- Support for various data types (strings, numbers, booleans, dates, arrays, objects)
- Handles nested objects and arrays by converting them to JSON strings

### Professional Formatting (Default)
- **Professional Headers**: Blue background with white text
- **Alternating Row Colors**: Gray and white stripes for better readability
- **Data Type Formatting**:
  - Numbers: Right-aligned with thousand separators (#,##0)
  - Decimals: Right-aligned with 2 decimal places (#,##0.00)
  - Dates: Center-aligned with mm/dd/yyyy format
  - Booleans: Center-aligned, green/bold for true, red for false
  - Email addresses: Blue and underlined
  - JSON objects/arrays: Left-aligned, italic, gray text
- **Borders**: Professional borders around all cells
- **Auto-sizing**: Columns automatically sized to fit content (with max width limit)

### Plain Mode (--no-format)
- Simple Excel output with minimal formatting
- Headers are bold only
- No colors, borders, or special formatting
- Basic auto-sized columns

## Installation

### Prerequisites
- .NET 9.0 or later (for development)
- No prerequisites for the standalone executable

### Download
Download the latest standalone executable from the releases:
- `JsonToExcel-Final.exe`  - Complete standalone version (no .NET installation required)

## Usage

### Basic Usage
```bash
# With formatting (default)
JsonToExcel input.json output.xlsx

# Without formatting (plain Excel)
JsonToExcel input.json output.xlsx --no-format

# Show help
JsonToExcel --help
```

### Examples

#### JSON Array Example
```json
[
  {
    "id": 1,
    "name": "John Doe",
    "email": "john@example.com",
    "salary": 75000.50,
    "active": true
  },
  {
    "id": 2,
    "name": "Jane Smith", 
    "email": "jane@example.com",
    "salary": 68000.25,
    "active": false
  }
]
```
Converts to a table with columns: id, name, email, salary, active

#### JSON Object Example
```json
{
  "name": "John Doe",
  "age": 30,
  "email": "john@example.com",
  "active": true
}
```
Converts to a two-column table with Property and Value columns

## Command Line Options

- `--no-format` - Disable formatting (creates plain Excel output)
- `--help`, `-h` - Show help message

## Development

### Build from Source
```bash
# Clone the repository
git clone https://github.com/robertorenz/jsontoexcel.git
cd jsontoexcel

# Build the project
cd JsonToExcel
dotnet build

# Run with sample data
dotnet run sample.json output.xlsx

# Publish standalone executable
dotnet publish -c Release
```

### Dependencies
- **Newtonsoft.Json**: JSON parsing and manipulation
- **EPPlus**: Excel file creation and formatting

## License
This project uses EPPlus with a non-commercial license. For commercial use, please ensure proper EPPlus licensing.

## Contributing
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## Changelog

### v1.0.0
- Initial release with JSON to Excel conversion
- Professional formatting with optional plain mode
- Support for arrays and objects
- Data type-specific formatting
- Standalone executable distribution
