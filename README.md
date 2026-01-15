# Google Sheets to PowerPoint Automation

Automatically update PowerPoint presentations with data from Google Sheets. This tool pulls values from a Google Sheet and updates named shapes and table cells in PowerPoint files, preserving formatting.

## Features

- **Cross-platform**: Works on both Windows and macOS
- **Named shape updates**: Update text boxes and shapes by name
- **Table cell updates**: Update specific cells in PowerPoint tables
- **Format preservation**: Maintains font styles, colors, and sizes
- **Value formatting**: Currency, percentage, dates, and more
- **Dry-run mode**: Preview changes without modifying files
- **Batch processing**: Efficient single API call for multiple values

## Requirements

- Python 3.9 or higher
- Google Cloud project with Sheets API enabled
- OAuth credentials from Google Cloud Console

## Installation

1. Clone or download this project
2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Set up Google Sheets API credentials:
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select existing
   - Enable the Google Sheets API
   - Create OAuth 2.0 credentials (Desktop application)
   - Download the credentials JSON file as `credentials.json`
   - Place it in the project directory

## Configuration

### config.yaml

```yaml
google:
  spreadsheet_id: "YOUR_SPREADSHEET_ID_HERE"
  mapping_sheet: "KeynoteMap"
  credentials_file: "credentials.json"
  token_file: "token.json"

powerpoint:
  file_path: "Investor Report.pptx"

defaults:
  empty_value: ""
  error_value: "#ERROR"

dry_run: false
```

### KeynoteMap Sheet Structure

Create a tab named "KeynoteMap" in your Google Sheet with these columns:

| Column | Name | Description |
|--------|------|-------------|
| A | ID | Unique identifier for the mapping |
| B | Sheet Range | Source cell in A1 notation (e.g., "DataVault!B12") |
| C | Slide | Slide number (1-indexed) |
| D | Target Type | "shape" or "table_cell" |
| E | Object Name | Name of the shape or table in PowerPoint |
| F | Row | Row number for table cells (1-indexed, leave empty for shapes) |
| G | Col | Column number for table cells (1-indexed, leave empty for shapes) |
| H | Format | Format type (see below) |
| I | Prefix | Text to prepend to value |
| J | Suffix | Text to append to value |
| K | Notes | Optional notes |

### Format Types

| Format | Description | Example |
|--------|-------------|---------|
| `text` | Plain text (default) | "Hello World" |
| `currency0` | Currency, no decimals | "$5,000" |
| `currency1` | Currency, 1 decimal | "$5,000.0" |
| `currency2` | Currency, 2 decimals | "$5,000.00" |
| `percent0` | Percentage, no decimals | "13%" |
| `percent1` | Percentage, 1 decimal | "13.3%" |
| `percent2` | Percentage, 2 decimals | "13.30%" |
| `integer` | Whole number with commas | "10,000" |
| `decimal1` | 1 decimal place | "5.0" |
| `decimal2` | 2 decimal places | "5.00" |
| `date_mdy` | Full date | "April 1, 2021" |
| `date_short` | Short date | "1/2027" |
| `text_number` | Number as word (1-20) | "Ten" |

## Usage

### Basic Update

```bash
python update_presentation.py
```

### With Options

```bash
# Dry run - preview changes without saving
python update_presentation.py --dry-run

# Use a specific config file
python update_presentation.py --config my_config.yaml

# Override presentation file
python update_presentation.py --presentation "My Deck.pptx"

# Save to different output file
python update_presentation.py --output "Updated Deck.pptx"

# Verbose logging
python update_presentation.py --verbose
```

### Discovery Commands

List shapes and tables in your presentation to find their names:

```bash
# List all shapes on slide 1
python update_presentation.py -p "My Deck.pptx" --list-shapes 1

# List all tables on slide 3
python update_presentation.py -p "My Deck.pptx" --list-tables 3
```

## Setting Up PowerPoint Shapes

### Naming Shapes

1. Open your PowerPoint presentation
2. Select a shape or text box
3. Go to the Selection Pane (Home > Select > Selection Pane)
4. Click on the shape name to rename it
5. Use descriptive names like "RevenueValue" or "GrowthRate"

### Naming Tables

Tables are automatically named when created (e.g., "Table 1"). You can rename them in the Selection Pane for clarity.

## Example Workflow

1. **Prepare your PowerPoint**:
   - Name the shapes you want to update
   - Note the slide numbers and shape names

2. **Set up Google Sheet**:
   - Create a "DataVault" tab with your source data
   - Create a "KeynoteMap" tab with mappings

3. **Configure**:
   - Update `config.yaml` with your spreadsheet ID
   - Place `credentials.json` in the project directory

4. **Test**:
   ```bash
   python update_presentation.py --dry-run
   ```

5. **Run**:
   ```bash
   python update_presentation.py
   ```

## File Structure

```
keynote_ppt/
├── config.yaml              # Configuration file
├── credentials.json         # Google OAuth credentials (you provide)
├── token.json               # OAuth token (auto-generated)
├── requirements.txt         # Python dependencies
├── update_presentation.py   # Main script
├── powerpoint_bridge.py     # PowerPoint operations
├── sheets_client.py         # Google Sheets API client
├── format_value.py          # Value formatting utilities
├── create_sample_pptx.py    # Creates sample presentation for testing
├── test_powerpoint_bridge.py # Unit tests for PowerPoint module
├── test_full_update.py      # Integration tests
├── sample_investor_deck.pptx # Sample presentation for testing
└── sample_keynote_map.csv   # Example mapping configuration
```

## Testing

Run the test suite to verify everything is working:

```bash
# Test PowerPoint bridge module
python test_powerpoint_bridge.py

# Test full integration (uses mock data)
python test_full_update.py
```

## Troubleshooting

### "Shape not found" errors
- Use `--list-shapes` to see actual shape names
- Check the slide number is correct (1-indexed)
- Shape names are case-sensitive

### "Table not found" errors
- Use `--list-tables` to see table names
- Verify the shape has a table (not just a grouped shape)

### Google API errors
- Verify spreadsheet ID is correct
- Check sheet names match exactly (including spaces)
- Ensure OAuth credentials are valid

### Format errors
- Values should be raw numbers (not pre-formatted)
- For percentages: use decimals (0.15) or whole numbers (15)
- For dates: use serial numbers or standard date formats

## License

MIT License - See LICENSE file for details.
