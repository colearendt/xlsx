# NEWS for package `xlsx`

## Changes in development version


## Changes in version 0.6.5 (released 2020-11)

 - Remove dependency on `RefManageR` (it was archived, and we only used it for a vignette)

## Changes in version 0.6.4.2 (released 2020-09)

 - Fix path expanding for `saveWorkbook`, `write.xlsx`, `write.xlsx2` for home directories / `~` (#108)

 - Remove Java version checks, to avoid collisions with Java 14 (#154, #156, #159)

## Changes in version 0.6.3 (released 2020-02)

 - Add `set_java_tmp_dir` and `get_java_tmp_dir` to address CRAN feedback

 - Change default Java temp directory to be under the R temp directory by default

 - Move a time/CPU consuming example to `\donttest{}` to comply with CRAN feedback

## Changes in verison 0.6.2 (released 2020-02)

 - Fix bug in `saveWorkbook` with a password (#107)

 - Switch to roxygen based documentation

 - Fix bug in version comparison for Java 10+

## Changes in version 0.6.1 (released 2018-06-10)

 - Export S3 methods for CellBlock, CellStyle

 - Do not depend (just import) rJava and xlsxjars

 - Minor fixes to pass CRAN again.

 - Add Cole Arendt as an author

## Changes in version 0.6.0 (released 2015-11-29)

 - Added support to read and write password protected xlsx files.
   Note that this is for xlsx format ONLY (no binary xls format).
   Thanks to Heather Turner for code contribution.  Note that
   on Linux LibreOffice does not open password protected workbooks
   as of now.  See [issue 49](https://github.com/colearendt/xlsx/issues/49).

 - Fixed a bug in read.xlsx when an empty rows was triggering an
   error.  Reported by maxnax, see [issue 57](https://github.com/colearendt/xlsx/issues/57).

## Changes in version 0.5.7 (released 2014-08-01)

 - Fixed [issue 41](https://github.com/colearendt/xlsx/issues/41).  There was a bug in the '+' method for CellStyle
   that was triggered when Font + Fill was not working properly.

 - Added a new function called setRowHeight (wish of Sven Neulinger).
   See [issue 43](https://github.com/colearendt/xlsx/issues/43).

 - Function write.xlsx did not check properly for POSIXct column classes and
   this resulted in spreadsheets not properly formatted.  See [issue
   45](https://github.com/colearendt/xlsx/issues/45).

 - Added a new function addAutoFilter to set a filter on a sheet.


## Changes in version 0.5.6 (released 2014-04-01)

 - Functions read.xlsx, read.xlsx2 crashed when trying to read empty sheets.
   Now they return NULL if the sheet exists but it's empty.  See [Issue
   28](https://github.com/colearendt/xlsx/issues/28).

 - Fixed a corner case in addDataFrame when the data.frame has zero columns.

 - Allow colors to be specified as hex strings (e.g. "#FF0000" for red).
   Previously, only R colors were supported.  This change applies only to the
   xlsx format.  Old style binary xls format still only supports a limited number
   of colors (POI limitation).

 - Change the .onLoad functions from both xlsxjars and xlsx package to allow
   for other users to initialize the Java VM themselves (useful for other
   packages that depend on xlsxjars).  This is not a user visible change. See
   [Issue 36](https://github.com/colearendt/xlsx/issues/36).

## Changes in version 0.5.5 (released 2013-12-01)

 - Allow CB.setMatrixData to accept character matrices too.  It only worked
   with numeric matrices in earlier versions.  See [Issue
   19](https://github.com/colearendt/xlsx/issues/19).

 - Fixed an issuewith read.xlsx2 when you wanted columns of type character, but
   the cell was of numeric type and the value was an exact integer (say 12).
   In R the result was imported as "12.0" instead of "12".  See [Issue
   25](https://github.com/colearendt/xlsx/issues/25).

 - DateTime and Date formatting when using write.xlsx and addDataFrame was
   hardcoded to the Java formats "m/d/yyyy h:mm:ss" and "m/d/yyyy"
   respectively.  These have been added as package options to allow for other
   formats when using these functions.  See [Issue
   26](https://github.com/colearendt/xlsx/issues/26).

 - For write.xlsx(2) when append==TRUE, check if file exists and load it only
   if it exists, otherwise create the workbook.  As suggested by Richard
   Cotton.  See [Issue 27](https://github.com/colearendt/xlsx/issues/27).

## Changes in version 0.5.3 (released 2013-07-11)

 - Fix a bug introduced in 0.5.1 version in read.xlsx.  The rowIndex argument
   was not respected anymore.

 - Cleaned up the code in readColumns responsible for guessing the colClasses.

## Changes in version 0.5.1 (released 2013-03-31)

 - Fix an annoyance with read.xlsx2, so it doesn't return column names of the
   form c..............

 - Add a startRow, endRow argument to read.xlsx.  These are convenience
   arguments for better control.  Also, it brings the function signature closer
   to that of read.xlsx2.

 - A NullPointerException was triggered for read.xlsx when the table due to
   incorrect indexing of the header columns in certain circumstances.  See
   [Issue 9](https://github.com/colearendt/xlsx/issues/9).

 - Fix an error with readColumns that was affecting read.xlsx2.  By default
   readColumns was using the max number of rows in the sheet to determine how
   many rows to read.  If the max number of rows in sheet was greater than the
   existing number of rows in the column you were reading, you were getting a
   NullPointerException.  This may happen for malformed spreadsheets (created by
   some other proprietary software for example).

## Changes in version 0.5.0 (released 2012-09-23)

 - New functionality: Introduce the concept of CellBlock to allow faster
   writes/updates of cell blocks.  Writing a spreadsheet is now +50% faster for
   medium sized spreadsheets compared to version 0.4.2.  This functionality allows
   you to style a block of cells much easier than before.  Thanks to Alexey
   Stukalov for his significant contribution on the Java side and for the original
   impetus.

 - Absolute paths are no longer needed for saveWorkbook and friends.  Reported
   by Mike Cheetham.

 - Functions write.xlsx, write.xlsx2 now allow you to create xls documents too.
   Previous versions only allowed xlsx documents.

 - Bug fix: read.xlsx2 doesn't read an extra empty column anymore.  Also
   colClasses were not mapped correctly to the corresponding colIndex.

 - Deprecated: getMatrixValues.  Use the readColumns function instead.

## Changes in version 0.4.2 (released 2012-02-08)

 - New function `readRows()` similar to `readColumns()` for reading accross the
   columns.

 - Removed a stray `browser()` in function `.guess_cell_type`

## Changes in version 0.4.1 (released 2012-01-22)

 - Added a test for file existence in loadWorkbook.  If file does not
   exist, the error message from read.xlsx, read.xlsx2 and
   loadWorkbook is now more informative.  Also added path expansion in
   file names. (Suggested by Dirk Eddelbuettel.)

 - Fixed bug in addDataFrame to allow it to add a df to existing
   sheets, not only to new sheets.

## Changes in version 0.4.0 (released 2012-01-15)

 - BACWARDS INCOMPATIBLE CHANGES!  A complete rewrite of functionality
   to deal with cell styles.  Although I want to minimize API breaking
   changes, I belive that these changes are for the better as the old
   function createCellStyle had a whopping 11 arguments and was still
   incomplete.  The new functionality defines an S3 object CellStyle
   on which you can "add" DataFormat, Font, Fill, Border, Alignment,
   Protection.  On the note of backward compatibility, I don't want to
   promise anything as the package is still young.  As more people
   starts using the package, the API will freeze and if breaking
   changes are contemplated, a clear deprecation path will be
   provided.

 - A Google project has been created http://code.google.com/p/rexcel/
   to expose the development branch of the package and manage the
   social interaction.  Please report all issues through this venue.
   Also a Google groups http://groups.google.com/group/R-package-xlsx
   has been created for announcements, etc.  Do register if you are
   interested in the development of the package.  It may give me more
   impetus to resolve an issue if I know many people are using this
   package.

 - New function addHyperlink to add hyperlinks (urls, emails) to
   a cell.

 - New function addDataFrame to add a data.frame to an existing
   sheet.  It alows the user to style the header, rownames, or
   individual columns.  This is now used internally in the write.xlsx2
   function.

 - New function getColumns to read a rectangular shape of cells into
   an R data.frame.  This is now used internally in the read.xlsx2
   function.

 - Rename the default in read.xlsx, createSheet and removeSheet to
   sheetName="Sheet1" from "Sheet 1".  This makes it consistent with
   Excel 2007 names of an empty workbook.

 - Added ... arguments to read.xlsx2 function to mirror read.xlsx.

 - Thanks to Neal Richardson and James Ward for submitting some code and
   suggestions.

## Changes in version 0.3.0 (released 2011-03-03)

 - Effort has been made to make all functions of this package to be agnostic
   between Excel versions.  You can now read, write and format files in Excel
   versions 97/2000/XP/2003 (not 95!) with file extension xls, in addition to
   Excel 2007 with file extension xlsx.  Please report issues you encounter.  Note
   that Colors are limited for xls workbooks (see ?CellStyle).

 - Read strings in a different encoding.  Thanks to Wincent Huang for code
   contribution. Function ?getCellValue now has an encoding argument.

 - Add support for Excel ranges.  Thanks to Wolfgang Abele for contributing
   preliminary code.  See ?Range.

 - New function read.xlsx2 for reading spreadsheets.  By moving the looping
   into java one gets a speed bump of one order of magnitude or better over
   read.xlsx.

 - Documentation fixes.

 - Move to version 3.7 for POI jars.  See http://poi.apache.org/.

## Changes in version 0.2.4 (released 2010-10-20)

 - New function write.xlsx2 in which the writing is done on the java side.
   Speed improvements of one order of magnitude over write.xlsx on moderately
   large data.frames (100,000 elements).

## Changes in version 0.2.3 (released 2010-08-26)

 - Fix the hAlign and vAlign arguments in createCellStyle.  The internal method
   call was lacking a cast to jshort.  Reported by Douglas Rivers.

 - Fix getCellValue when you have formulas with String values (it assumed
   Numeric values and errored out for other types).  Support now Strings and
   Booleans.  Error out for other cell types.  Don't know a sound solution.
   Reported by ravi(?) rv15i@yahoo.se.

## Changes in version 0.2.2 (released 2010-07-14)

 - Added a colIndex argument to read.xlsx to facilitate reading only specific
   columns.

 - setCellValue now tests for NA, and if value is NA, it will fill the cell
   with #N/A.

 - Fixed bug in getCellValue.  It now returns NA for all the error codes in the
   cell.  It used to return a numeric code which was confusing to the R user.
   Reported by Ralf Tautenhahn.

## Changes in version 0.2.1 (released 2010-05-15)

 - Fixed bug with write.xlsx.  It does not write colnames even if
   col.names=TRUE.  Reported by Ralf Tautenhahn.

 - Added an ... arg to read.xlsx that is passed to the data.frame constructor,
   for example to control the stringsAsFactors option.

## Changes in version 0.2.0 (released 2010-05-01)

 - Switched to POI 3.6.  This resulted in significant memory improvements but
   will still run into memory issues when reading/writing large xlsx files.

 - Added addPicture function for embedding pictures into xlsx files.

 - Added removeRow function for conveniently removing existing rows from the
   spreadsheet.

 - Added/Fixed comments for cells.  See ?Comment

 - Fixed bug in read.xlsx for the case when the file contains only one column
   issue reported by Hans Petersen), a corner case when drop=TRUE wrecked
   havoc.

 - Fixed bug in createRow.  If rowIndex did not start at 1, it created spurious
   NULL entries.

## Changes in version 0.1.3  (released 2010-03-15)

 - Added indCol argument to getCells in case you want to get only a subset of
   columns.

 - Added function getMatrixValue to extract blocks of data from the sheet.

 - Improved and expanded the unit tests.

 - On Mac, you cannot set colors directly using createCellStyle.  You can still
   do it manually, please see the javadocs.

## Changes in version 0.1.2  (released 2010-01-02)

 - Fixed getRows, getCells so it does not error out for empty rows/cells.
   Modified read.xlsx too.

 - Added append argument to write.xlsx to be able to export to multiple
   worksheets of a file.  (Suggestion by rlearnr@gmail.com.)
