[![AppVeyor Build Status](https://ci.appveyor.com/api/projects/status/github/colearendt/xlsx?branch=master&svg=true)](https://ci.appveyor.com/project/colearendt/xlsx)
[![Travis-CI Build Status](https://travis-ci.org/colearendt/xlsx.svg?branch=master)](https://travis-ci.org/colearendt/xlsx)
[![CRAN Version](http://www.r-pkg.org/badges/version-last-release/xlsx)](https://cran.r-project.org/web/packages/xlsx/index.html)

xlsx
========

**An R package to read, write, format Excel 2007 and Excel 97/2000/XP/2003 files**

The package provides R functions to read, write, and format Excel files.  It depends 
on Java, but this makes it available on most operating systems. 

### Install

Stable version from CRAN

```r
install.packages('xlsx')
```

Or development version from GitHub

```r
devtools::install_github('dragua/xlsx')
```

## Quick start

To read the first sheet from spreadsheet into a data.frame 

```r
read.xlsx2('file.xlsx', 1)
```
To write a data.frame to a spreadsheet 
```r
write.xlsx2(iris, file='iris.xlsx')
```

The package has many functions that make it easy to style and
formalize output into Excel, as well.

```r
wb <- createWorkbook()
s <- createSheet(wb,'test')

cs <- CellStyle(wb) + 
  Font(wb,heightInPoints = 16, isBold = TRUE) +
  Alignment(horizontal='ALIGN_CENTER')
  

r <- createRow(s,1)
cell <- createCell(r,1:ncol(iris))

setCellValue(cell[[1]],'Title for Iris')
for (i in cell) {
  setCellStyle(i,cs)
}

addMergedRegion(s, 1,1, 1,ncol(iris))

addDataFrame(iris, s, row.names=FALSE, startRow=3)

saveWorkbook(wb,'iris_pretty.xlsx')
```

## Issues/Mailing list

To report a bug, use the Issues page at 
[issues](https://github.com/dragua/xlsx/issues): https://github.com/dragua/xlsx/issues

Questions should be asked on the dedicated mailing
[list](http://groups.google.com/group/R-package-xlsx): http://groups.google.com/group/R-package-xlsx

## Acknowledgements

The package is made possible thanks to the excellent
work on [Apache POI](https://poi.apache.org/spreadsheet/index.html).
