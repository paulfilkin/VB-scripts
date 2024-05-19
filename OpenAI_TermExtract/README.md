# OpenAI TermExtract

This repository contains a number of VB scripts for Excel that use OpenAI to take an export from Trados Studio (or any tool that supports extracting a bilingual translation to Excel) and then extracts the terminology and provides a definition of the terms. This information is saved in Excel in a format suitable for conversion to a termbase or glossary, or some intermediate format such as TBX by using the Glossary Converter from the Trados AppStore.

## Scripts

There are three scripts.

### JsonConverter.bas

Taken as-is from the code provided for a tool developed by Tim Hall called [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) which is used for JSON conversion and parsing for VBA (Windows and Mac Excel, Access, and other Office applications). 

It was originally tested by Tim in Windows Excel 2013 and Excel for Mac 2011, but should apply to 2007+. I have used the JasonConverter.bas in Excel as part of Office 365 without any changes at all.

This is used within the CallOpenAI script to parse the JSON responses received from OpenAI's API, converting them into VBA objects that can be easily accessed and manipulated within Excel. This functionality is critical for extracting the AI-generated text responses needed for the TermExtraction script, which processes terms from a worksheet, sends them to the OpenAI API using the CallOpenAI function, and then inserts the responses back into the Excel workbook, replacing the API call formulas with their resulting values. 

### CallOpenAI.bas

I took the idea for this from an article written by Ed Twomey called [How to Integrate GPT-4 into Excel](https://medium.com/@ed.twomey1/how-to-integrate-gpt-4-into-excel-23954e4d60a6). I have extensively changed this for my needs but the original explanation of how to get OpenAI working in Excel came from this article.

The changes I made were to:

- read the API key from a cell in Excel
- read the model from a dropdown in Excel (GPT-4 or GPT-4o)
- read the prompt from a cell in Excel
- corrected some issues with the content sent to OpenAI having unescaped quotes
- added simple error handling 

### TermExtraction.bas

This script just brings things together and takes the output of the OpenAI calls, putting the content into a new worksheet where it is formatted quite simply as source term, target term, definition of the term.

This script does the following:

- sends the source and target content in Excel to OpenAI and extracts a list of source and matching target terms for each row (prompt is in an Excel worksheet for this)
- creates a new worksheet and adds the term pairs into separate columns
- sends the source terms to OpenAI and requests a definition for each one
- it also replaces the formula used with the plain text result to ensure any editing in the excel afterwards won't kick-off anymore calls to OpenAI
