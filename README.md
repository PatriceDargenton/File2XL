File2XL : Open a csv file into MS-Excel with pre-formatted cells
---

[File2XL.html](http://patrice.dargenton.free.fr/CodesSources/File2XL.html)  
[File2XL.vbproj.html](http://patrice.dargenton.free.fr/CodesSources/File2XL.vbproj.html)  
[File2XL on GitHub](https://github.com/PatriceDargenton/File2XL)  
[File2XL on CodeProject](https://www.codeproject.com/Tips/1108923/File-XL-Open-a-csv-file-into-MS-Excel-with-pre-for)  
By Patrice Dargenton (patrice.dargenton@free.fr)  
[My website](http://patrice.dargenton.free.fr/index.html)  
[My source codes](http://patrice.dargenton.free.fr/CodesSources/index.html)  

File2XL lets you open any text file into MS-Excel and creates two sheets : one in text format, and a second one in standard format for numeric columns. File2XL adds a context menu titled : "Open in MS-Excel using File2XL" for any file in Windows File Explorer. Partially created workbook can be viewed without having to wait for the opening of the entire file.

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
    - [Version 1.05 - 25/01/2019](#version-105---25012019)
    - [Version 1.04 - 05/01/2018](#version-104---05012018)
    - [Version 1.03 - 20/05/2017](#version-103---20052017)
    - [Version 1.02 - 08/05/2017](#version-102---08052017)
    - [Version 1.01 - 25/06/2016 : First version](#version-101---25062016--first-version)
- [Links](#links)
    - [See also](#see-also)

<!-- /TOC -->

# Features
- Two sheets are created : one in text format, and one in standard format for numeric columns, because sometimes you need to see the original text before it was converted to numeric (for conversion problem investigation) ;
- Partially workbook viewing : check big file quickly before waiting the complete workbook to be created (pause/continue/cancel/show buttons are available) ;
- Excel limitations are checked : 256 columns and 65536 lines for Excel 2003, and 16384 columns and 1048576 lines for Excel 2007 (or >), and 32767 characters max. in one cell for both versions ; colored and text alerts are displayed within the sheet if these limits are exceeded, and you are only prompted once by this kind of limit exceed ;
- Source file encoding is detected (UTF7, UTF8, Unicode, BigEndianUnicode, UTF32 and ASCII) ;
- Temporary Excel file is removed after closing Excel, if you agree to delete it ;
- Delimiter detection : a few delimiters are counted (at the top of the file) : ,;| and tabulation ; possible delimiters are configurable ;
- Special delimiter : "," or ";" is supported (not configurable) ;
- Minus sign at the end of the value is supported, e.g.: 0.72- -> -0.72 ;
- Using Excel 2003 (or 2000/2002) and/or Excel 2007 (or >) is configurable ;
- Autofilter on the header, the first line, is yet enabled (not configurable) ;
- Frozen column is configurable (1 column left is always visible by default, but 0 is possible too) ;
- Autosizing columns is configurable ;
- The number of header lines analyzed is configurable ;
- The standard sheet can be disabled (only text sheet is then created) ;
- Removing NULL value in standard sheet is configurable (for example PhpMyAdmin NULL value in csv export).

# Explanations

## Context menu
The first time, run File2XL in administrator privilege (run as admin.), and add (or remove) context menu using the + (or - respectively) button ;
After that, use the context menu "Open in MS-Excel using File2XL" for any text file in Windows File Explorer.

## Multiple delimiter
There are only two multiple delimiters (not configurable) : "," and ";"

Only a quick parsing is performed (splitting with "," or ";"), not a deep parsing. If a deep parsing is required (like the slow one that Excel use in his Text Import Wizard), there is a second context menu to choose for example comma (,) instead of "," : "Open in MS-Excel using File2XL (single delimiter)", otherwise the default context menu gives chance to choose the multiple delimiter ",".

Example of a file that is generated with multiple delimiter : phpmyadmin csv export (null value doesn't have "", so you should use the second context menu for it : single delimiter, if you have nullable fields).

## Settings
There is no user interface to configure settings, simply edit the config. file in the notepad : File2XL.exe.config


# Projects
- Numeric field : count how many decimal digits of precision are required (actually, no decimal is shown by default, but you can change it afterward in Excel as you want) ;
- Date field : show date (and time) fields in formatted and colored cells in the standard sheet ;
- Event handler for the writing of the Excel workbook (which may be cancelled) for large files : suggestion have been submitted to NPOI team but not yet implemented (possible way to do it : counting every row or every line of each sheet to be written).


# Versions

## Version 1.05 - 25/01/2019
- Detecting UTF8 Encoding : one case added.

## Version 1.04 - 05/01/2018
- Encoding reading : no need to read with write access right, just read access.

## Version 1.03 - 20/05/2017
- LogFile setting added : to log conversion time of each file ;
- Visual Studio 2017 code analysis : almost all rules are respected ;
- UTF8 encoding added in encoding detection ;
- MinColumnWidth and MaxColumnWidth settings added ;
- Bug fixed (from 1.02 version) : object variable not set : fs.position while fs is null.

## Version 1.02 - 08/05/2017
- RemoveNULL setting : remove NULL in field value in standard sheet, for example PhpMyAdmin NULL value in csv export ;
- SingleDelimiter : disable multiple delimiter (not simply prefer single one).

## Version 1.01 - 25/06/2016 : First version


# Links
- Library used : [NPOI](https://github.com/tonyqus/npoi) (2.2.1.0 version, may 2016)  
  NPOI : a .NET library that can read/write Office formats without Microsoft Office installed. No COM+, no interop.  
  Only one add (2.2.1.1) : GetColumnWidth : iNumRow++; if (iNumRow > iNbRowMax) break;  
  const int iNbRowMax = 100;  
  in order to perform a fast column autosize based on only the top 100 lines (suggestion have been submitted to NPOI team but not yet implemented in the github repository).  

- [Neuzilla User Group](https://www.linkedin.com/groups/6655065) (Tony Qu from Neuzilla is the main NPOI contributor)


## See also
- (french) [XL2Csv](http://patrice.dargenton.free.fr/CodesSources/XL2Csv.html) : Convertir un fichier Excel en fichiers Csv (ou en 1 fichier txt)  
  Source code : [XL2Csv.vbproj.html](http://patrice.dargenton.free.fr/CodesSources/XL2Csv.vbproj.html)  