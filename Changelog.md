# Changelog

All notable changes to the File2XL project will be documented in this file.

## [Unreleased]

## [1.07] - 01/05/2023
### Added
- Log probable delimiter detection results in the log file;
- Option detectEncodingFromByteOrderMarks used in StreamReader (don't try to detect the encoding on your own, instead use the standard .net option for this complex task);
- Test text-encoding-detect from https://github.com/AutoItConsulting/text-encoding-detect.

### Fixed
- Indent code for Visual Studio 2022;
- .Net45 -> .Net472.

## [1.06] - 22/10/2021
### Added
- Detecting UTF8 Encoding: one case added.

## [1.05] - 25/01/2019
### Added
- Detecting UTF8 Encoding: one case added.

## [1.04] - 05/01/2018
### Fixed
- Encoding reading: no need to read with write access right, just read access.

## [1.03] - 20/05/2017
### Added
- LogFile setting added: to log conversion time of each file;
- Visual Studio 2017 code analysis: almost all rules are respected;
- UTF8 encoding added in encoding detection;
- MinColumnWidth and MaxColumnWidth settings added.

### Fixed
- Bug fixed (from 1.02 version): object variable not set: fs.position while fs is null.

## [1.02] - 08/05/2017
### Fixed
- RemoveNULL setting: remove NULL in field value in standard sheet, for example PhpMyAdmin NULL value in csv export;
- SingleDelimiter: disable multiple delimiter (not simply prefer single one).

## [1.01] - 2016-06-25 First version