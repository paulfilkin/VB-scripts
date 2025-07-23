# SDLXLIFF Translation Attributes Extractor

## Overview
This VBA script extracts translation metadata and attributes from SDLXLIFF (Trados Studio) files into Excel. It's designed to help translation project managers, linguists, and localization engineers analyze translation memory matches, segment statuses, and other translation-related metadata.

## Features
- Extracts source and target text from translation units
- Captures segment-level metadata including:
  - Translation status (Draft, Translated, etc.)
  - Origin and origin system
  - Match percentages
  - Text and structure match indicators
  - Creation and modification dates
  - User information
- Creates a formatted Excel table with filters
- Handles XLIFF namespaces properly
- Skips non-translatable segments

## Requirements
- Microsoft Excel (2010 or later)
- Windows operating system
- VBA macros must be enabled

## Installation

1. Open Microsoft Excel
2. Press `Alt + F11` to open the VBA Editor
3. Insert a new module: `Insert > Module`
4. Copy and paste the entire script into the module
5. Save the workbook as a macro-enabled file (.xlsm)

## Usage

1. Run the macro by:
   - Press `Alt + F8` in Excel
   - Select `ExtractTranslationAttributes` from the list
   - Click `Run`

2. Select your SDLXLIFF file when prompted

3. The script will create a new worksheet with the extracted data

## Output Columns

| Column | Description |
|--------|-------------|
| Trans-Unit ID | Unique identifier for each translation unit |
| Source Text | Original source language text |
| Target Text | Translated target language text |
| Segment ID | Segment identifier within the translation unit |
| Status | Translation status (Draft, Translated, Approved, etc.) |
| Origin | Where the translation came from |
| Origin System | System that provided the translation (e.g., TM name) |
| Percent Match | Translation memory match percentage |
| Text Match | Whether text matching was used |
| Struct Match | Whether structural matching was used |
| TM Name | Translation memory name (if applicable) |
| Created Date | When the translation was created |
| Modified Date | When the translation was last modified |
| Last Used Date | When the translation was last used |
| User ID | User who created/modified the translation |

## Troubleshooting

### "Error loading file"
- Ensure the file is a valid SDLXLIFF or XLIFF file
- Check that the file is not corrupted
- Verify you have read permissions for the file

### No data extracted
- The file may not contain translatable segments
- Check if all segments are marked with `translate="no"`
- Verify the file structure matches standard SDLXLIFF format

### Missing data in some columns
- Not all segments contain all metadata fields
- Some attributes are only present for TM matches
- Manual translations may lack certain metadata

## Customization

To add additional attributes:

1. Add new column headers to the `headers` array
2. In the segment processing section, add:
   ```vba
   ws.Cells(row, [column_number]).Value = GetAttributeValue(segNode, "[attribute_name]")

To extract different metadata:

- Modify the XPath queries to target different nodes
- Add additional namespace declarations if needed

## Known Limitations

- Only processes SDLXLIFF format
- Does not extract inline tag information
- Limited to attributes at the segment level
- Does not process comment or revision data

## Example Use Cases

1. **Quality Analysis**: Review segments by match percentage and status
2. **TM Analysis**: Identify which translation memories were used
3. **Project Tracking**: Monitor translation progress by status
4. **Vendor Management**: Track which users worked on translations
5. **Leverage Reporting**: Analyze match percentages across files
