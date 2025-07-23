# SDLXLIFF Translation Attributes Extractor

## Overview
This VBA script extracts comprehensive translation metadata, attributes, comments, and track changes from SDLXLIFF (Trados Studio) files into Excel. It's designed to help translation project managers, linguists, and localization engineers analyze translation memory matches, segment statuses, review comments, track changes, and other translation-related metadata.

## Features
- Extracts source and target text from translation units
- **Captures translation comments** including:
  - Comment text
  - Comment author
  - Comment date and time
  - Comment severity (Low, Medium, High)
- **Tracks revision changes** including:
  - Added and deleted text
  - Revision author
  - Revision date
  - Change type (addition/deletion)
- Captures segment-level metadata including:
  - Translation status (Draft, Translated, etc.)
  - Confirmation level with descriptive text
  - Origin and origin system
  - Match percentages (percent, text, structure, context)
  - Creation and modification dates
  - User information
  - Locked segment status
- Creates a formatted Excel table with filters
- Handles XLIFF namespaces properly
- Handles SDL-specific namespaced attributes
- Skips non-translatable segments
- Supports multiple fallback methods for element detection

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
   - Select `ExtractSDLXLIFFNamespaceAware` from the list
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
| Status | Translation status (conf attribute value) |
| Confirmation Level | Human-readable status (Draft, Translated, Approved, etc.) |
| Origin | Where the translation came from |
| Origin System | System that provided the translation (e.g., TM name) |
| Percent Match | Translation memory match percentage |
| Text Match | Whether text matching was used |
| Struct Match | Whether structural matching was used |
| Context Match | Context match percentage |
| Created Date | When the translation was created |
| Created By | User who created the translation |
| Modified Date | When the translation was last modified |
| Modified By | User who last modified the translation |
| Comment Text | Review comment text (if any) |
| Comment Author | Who wrote the comment |
| Comment Date | When the comment was added |
| Comment Severity | Comment priority (Low, Medium, High) |
| Has Track Changes | Whether segment has revisions (Yes/No) |
| Revision Author | Who made the revision |
| Revision Date | When the revision was made |
| Deleted Text | Text that was deleted |
| Added Text | Text that was added |
| Locked | Whether segment is locked |

## Troubleshooting

### "Error loading file"
- Ensure the file is a valid SDLXLIFF or XLIFF file
- Check that the file is not corrupted
- Verify you have read permissions for the file

### No trans-units found
- The script tries multiple methods to find trans-units
- Check if the file uses a different namespace structure
- Verify the file contains actual translation units

### Missing comments or track changes
- Comments are stored in `doc-info/cmt-defs/cmt-def/Comments/Comment`
- Track changes are marked with `mtype="x-sdl-deleted"` or `"x-sdl-added"`
- Ensure your SDLXLIFF file contains these elements

### Missing data in some columns
- Not all segments contain all metadata fields
- Some attributes are only present for TM matches
- Manual translations may lack certain metadata
- Comments and track changes are optional features

## Technical Details

### Namespace Handling
The script properly handles XML namespaces:
- XLIFF namespace: `urn:oasis:names:tc:xliff:document:1.2`
- SDL namespace: `http://sdl.com/FileTypes/SdlXliff/1.0`

### Comment Structure
Comments are stored in the document header and referenced by ID:
```xml
<cmt-def id="...">
  <Comments>
    <Comment severity="..." user="..." date="...">Comment text</Comment>
  </Comments>
</cmt-def>
```

### Track Changes Structure

Revisions are marked with `mrk` elements:

xml

```xml
<mrk mtype="x-sdl-deleted" sdl:revid="...">deleted text</mrk>
<mrk mtype="x-sdl-added" sdl:revid="...">added text</mrk>
```

## Customization

To add additional attributes:

1. Add new column headers to the `headers` array
2. In the segment processing section, add extraction logic
3. For namespaced attributes, use XPath with proper namespace handling

To extract different metadata:

- Modify the XPath queries to target different nodes
- Add additional namespace declarations if needed
- Use fallback methods for robust element detection

## Example Use Cases

1. **Quality Review**: Analyze translator comments and severity levels
2. **Change Tracking**: Monitor what text was modified during review
3. **TM Analysis**: Identify which translation memories were used
4. **Project Tracking**: Monitor translation progress by confirmation status
5. **Vendor Management**: Track which users worked on translations and reviews
6. **Compliance Auditing**: Document all changes and comments for audit trails
7. **Leverage Reporting**: Analyze match percentages across files
8. **Review Efficiency**: Identify segments with high-severity comments

## Version History

### v2.0 (Current)

- Added comment extraction support
- Added track changes detection
- Improved namespace handling
- Added multiple fallback methods for element detection
- Expanded from 15 to 25 output columns

### v1.0

- Initial release with basic metadata extraction
