File2XL: Open a csv file into MS-Excel with pre-formatted cells
---

File2XL lets you open any text file into MS-Excel and creates two sheets: one in text format, and a second one in standard format for numeric columns. File2XL adds a context menu titled: "Open in MS-Excel using File2XL" for any file in Windows File Explorer. Partially created workbook can be viewed without having to wait for the opening of the entire file.

# Keywords
Excel Text Import Wizard, Csv2Excel, Csv to Excel, Txt2Excel, Txt to Excel, Text2Excel, Text to Excel.

<!-- TOC -->

- [Keywords](#keywords)
- [Features](#features)
- [Explanations](#explanations)
    - [Context menu](#context-menu)
    - [Multiple delimiter](#multiple-delimiter)
    - [Settings](#settings)
- [Projects](#projects)
- [Versions](#versions)
- [Links](#links)

<!-- /TOC -->

# Features
- Two sheets are created: one in text format, and one in standard format for numeric columns, because sometimes you need to see the original text before it was converted to numeric (for conversion problem investigation);
- Partially workbook viewing: check big file quickly before waiting the complete workbook to be created (pause/continue/cancel/show buttons are available);
- Excel limitations are checked: 256 columns and 65536 lines for Excel 2003, and 16384 columns and 1048576 lines for Excel 2007 (or >), and 32767 characters max. in one cell for both versions; colored and text alerts are displayed within the sheet if these limits are exceeded, and you are only prompted once by this kind of limit exceed;
- Source file encoding is detected (UTF7, UTF8, Unicode, BigEndianUnicode, UTF32 and ASCII);
- Temporary Excel file is removed after closing Excel, if you agree to delete it;
- Delimiter detection: a few delimiters are counted (at the top of the file): ,;| and tabulation; possible delimiters are configurable;
- Special delimiter: "," or ";" is supported (not configurable);
- Minus sign at the end of the value is supported, e.g.: 0.72- -> -0.72;
- Using Excel 2003 (or 2000/2002) and/or Excel 2007 (or >) is configurable;
- Autofilter on the header, the first line, is yet enabled (not configurable);
- Frozen column is configurable (1 column left is always visible by default, but 0 is possible too);
- Autosizing columns is configurable;
- The number of header lines analyzed is configurable;
- The standard sheet can be disabled (only text sheet is then created);
- Removing NULL value in standard sheet is configurable (for example PhpMyAdmin NULL value in csv export).

# Explanations

## Context menu
The first time, run File2XL in administrator privilege (run as admin.), and add (or remove) context menu using the + (or - respectively) button;
After that, use the context menu "Open in MS-Excel using File2XL" for any text file in Windows File Explorer.

## Multiple delimiter
There are only two multiple delimiters (not configurable): "," and ";"

Only a quick parsing is performed (splitting with "," or ";"), not a deep parsing. If a deep parsing is required (like the slow one that Excel use in his Text Import Wizard), there is a second context menu to choose for example comma (,) instead of ",": "Open in MS-Excel using File2XL (single delimiter)", otherwise the default context menu gives chance to choose the multiple delimiter ",".

The SingleDelimiter menu means that we do not add weight to the separators "," and ";" to favor the comma separator (,) if that is sufficient. Without this SingleDelimiter menu (normal menu), we add weight (3 characters therefore 3 times more weight) to the multiple separators.

Example of a file that is generated with multiple delimiter: phpmyadmin csv export (null value doesn't have "", so you should use the second context menu for it: single delimiter, if you have nullable fields).

## Settings
There is no user interface to configure settings, simply edit the config. file in the notepad: File2XL.exe.config


# Projects
- Options menu: before loading the file in Excel, show a panel to set encoding option, and to set delimiter option (instead of the 'single delimiter' menu);
- Numeric field: count how many decimal digits of precision are required (actually, no decimal is shown by default, but you can change it afterward in Excel as you want);
- Date field: show date (and time) fields in formatted and colored cells in the standard sheet.


# Versions

See [Changelog.md](Changelog.md)