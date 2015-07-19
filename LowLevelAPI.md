

Using the low-level functions gives you the ability to control what
you read into R or what and how you export to a spreadsheet.

# Create a spreadsheet from scratch #

You want to export your `R` data to Excel.  Here are the steps you
need to follow.  These steps mimic the Apache POI approach but made
idiomatic to `R`.

```
wb    <- createWorkbook(type="xlsx")           # create an empty workbook
sheet <- createSheet(wb, sheetName="Sheet1")   # create an empty sheet 
rows  <- createRow(sheet, rowIndex=1:12)       # 12 rows
cells <- createCell(rows, colIndex=1:10)       # 10 columns in each row
```
The workbook object `wb`, `sheet` are just a java reference.
```
> wb
[1] "Java-Object{Name: /xl/workbook.xml - Content Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml}"
> sheet
[1] "Java-Object{Name: /xl/worksheets/sheet1.xml - Content Type: application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml}"
```

`rows` is a list of java references with names equal to the row
number and `cells` is a matrix of size 20x10 of java references.
```
> head(rows)
$`1`
[1] "Java-Object{<xml-fragment r=\"1\"/>}"

$`2`
[1] "Java-Object{<xml-fragment r=\"2\"/>}"

$`3`
[1] "Java-Object{<xml-fragment r=\"3\"/>}"

$`4`
[1] "Java-Object{<xml-fragment r=\"4\"/>}"

$`5`
[1] "Java-Object{<xml-fragment r=\"5\"/>}"

$`6`
[1] "Java-Object{<xml-fragment r=\"6\"/>}"

> head(cells)
  1 2 3 4 5 6 7 8 9 10
1 ? ? ? ? ? ? ? ? ? ? 
2 ? ? ? ? ? ? ? ? ? ? 
3 ? ? ? ? ? ? ? ? ? ? 
4 ? ? ? ? ? ? ? ? ? ? 
5 ? ? ? ? ? ? ? ? ? ? 
6 ? ? ? ? ? ? ? ? ? ? 
>
```


For example, to populate the first column with the names of the month
do
```
mapply(setCellValue, cells[1:12,1], month.name)   # set one column 
setCellValue(cells[[2,2]], "I'm a special cell!") # set one cell
```


When you're done populating your spreadsheet, you need to save your
workbook to a file
```
saveWorkbook(wb, file="/tmp/test.xlsx")
```


You are now done!  The function `write.xlsx` in the high level API,
follows the steps above.


# Customize cell appearance #

The code shown in the section above just writes content to a
spreadsheet using the default spreadsheet styling.  One of the
appealing reasons to use a spreadsheet is that you can customize its
appearance.  You can highlight some data, add borders, use different
fonts for important results, etc.  However to customize the appearance
of your spreadsheet programatically is not easy to achieve.  The way
it is done with Apache POI API is verbose and not very intuitive.

Formatting options in Excel are grouped in six categories: DataFormat,
Alignment, Font, Border, Fill and Protection.  I've decided to emulate
this grouping.  Also, inspired by package **`ggplot2`** I've decided to overload
the `+` operator to work with these categories.

For a workbook object `wb` a call to `CellStyle(wb)` creates the
default (empty) cell style on which you can "add" properties.  I won't
list here all possible combination (see the help) but here are some
examples of what you can do

```
cs0 <- CellStyle(wb)                        # default/empty style
cs1 <- CellStyle(wb) + 
  Font(wb, name="Courier New", isBold=TRUE)     # add a Font
cs2 <- CellStyle(wb) +    
  Font(wb, name="Courier New", isBold=TRUE) +   # add a Font
  Borders(col="blue", position=c("TOP", "BOTTOM"), pen="BORDER_THICK") +  # add borders    
  Alignment(h="ALIGN_RIGHT")                # add Alignment
cs3 <- CellStyle(wb) + 
  DataFormat("m/d/yyyy h:mm:ss;@")          # format datetimes
```

You can now use these defined cell styles to "style" your
spreadsheet.  For example, if you have a matrix of cells as in
[LowLevelAPI#Create\_a\_spreadsheet\_from\_scratch](LowLevelAPI#Create_a_spreadsheet_from_scratch.md), you can format them with

```
lapply(cells[,1], setCellStyle, cs3) # format first column as a datetime
setCellStyle(cells[[2,2]], cs1)      # format one cell in Courier New, bold
```


# Details about setCellValue and getCellValue #

These two functions are the workhorse of the package.  They work with
one cell only, to set or get its value respectively.

Function `setCellValue` supports only five R types, `integer`,
`numeric`, `Date`, `POSIXct`, `character`.  All other types are
converted to character.  Note that Excel datetimes are in the GMT
timezone.

Function `getCellValue` supports reading `numeric`, `character`,
`logical` values.  Blank cells are imported as `NA`.  Errors are
imported as `NA` also.  If there is an unknown cell type it is
imported as `Error`.

# Customize a spreadsheet with addDataFrame #

Behind the scene the function `addDataFrame` creates a Java object of
class RInterface.  This is used to minimize the interaction between R
and Java and speed up the operations on the spreadsheet.

By default, when using the function `addDataFrame`, a column of type `Date` will
be formated with `DataFormat("m/d/yyyy")`, and a column
of type `POSIXt` will be formated with `DataFormat("m/d/yyyy h:mm:ss;@")`.


## Details about the RInterface ##

You can inspect in details the java code from the package sources in
folder `other/`.  The `RInterface` object has several methods, of which I'm showing the most
important below
```
>  Rintf <- .jnew("dev/RInterface")   # create an RInterface object
> .jmethods(Rintf)
 [1] "public void dev.RInterface.createCells(org.apache.poi.ss.usermodel.Sheet,int,int)"                                                             
 [2] "public double[] dev.RInterface.readColDoubles(org.apache.poi.ss.usermodel.Sheet,int,int,int)"                                                  
 [3] "public java.lang.String[] dev.RInterface.readColStrings(org.apache.poi.ss.usermodel.Sheet,int,int,int)"                                        
 [4] "public java.lang.String[] dev.RInterface.readRowStrings(org.apache.poi.ss.usermodel.Sheet,int,int,int)"                                        
 [5] "public void dev.RInterface.writeColDoubles(org.apache.poi.ss.usermodel.Sheet,int,int,double[])"                                                
 [6] "public void dev.RInterface.writeColDoubles(org.apache.poi.ss.usermodel.Sheet,int,int,double[],org.apache.poi.ss.usermodel.CellStyle)"          
 [7] "public void dev.RInterface.writeColInts(org.apache.poi.ss.usermodel.Sheet,int,int,int[])"                                                      
 [8] "public void dev.RInterface.writeColInts(org.apache.poi.ss.usermodel.Sheet,int,int,int[],org.apache.poi.ss.usermodel.CellStyle)"                
 [9] "public void dev.RInterface.writeColStrings(org.apache.poi.ss.usermodel.Sheet,int,int,java.lang.String[])"                                      
[10] "public void dev.RInterface.writeColStrings(org.apache.poi.ss.usermodel.Sheet,int,int,java.lang.String[],org.apache.poi.ss.usermodel.CellStyle)"
[11] "public void dev.RInterface.writeRowStrings(org.apache.poi.ss.usermodel.Sheet,int,int,java.lang.String[],org.apache.poi.ss.usermodel.CellStyle)"
[12] "public void dev.RInterface.writeRowStrings(org.apache.poi.ss.usermodel.Sheet,int,int,java.lang.String[])"                                      
```

The `[1]` method creates cells, and the other are used to read/write
columns and rows of data.  Writing columns of data is possible with a
given `CellStyle`.

You can look at the source code for the function `addDataFrame` for a
practical example of how to add a `data.frame` to a sheet.  The basic
idea is simple.  Create an `RInterface` object, create the cells, and
start populating the cells by column.