xlsx
========

**An R package to read, write, format Excel 2007 and Excel 97/2000/XP/2003 files**

The package provides R functions to read, write, and format Excel files.  Depends 
on Java, but this makes it available on most OS'es. 

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
read.xlsx2('C:/temp/file.xlsx', 1)
```
To write a data.frame to a spreadsheet 
```r
write.xlsx2(iris, file='C:/temp/iris.xlsx')
```

## Issues/Mailing list

To report a bug, use the Issues page at [issues]
[issues]: https://github.com/dragua/xlsx/issues

Questions should be asked on the dedicated mailing [list]
[list]: http://groups.google.com/group/R-package-xlsx

