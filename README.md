  <!-- badges: start -->
  [![CRAN Version](https://www.r-pkg.org/badges/version-last-release/xlsx)](https://CRAN.R-project.org/package=xlsx)
  [![Codecov test coverage](https://codecov.io/gh/colearendt/xlsx/branch/main/graph/badge.svg)](https://codecov.io/gh/colearendt/xlsx?branch=main)
  [![R-CMD-check](https://github.com/colearendt/xlsx/workflows/R-CMD-check/badge.svg)](https://github.com/colearendt/xlsx/actions)
  [![CRAN Activity](https://cranlogs.r-pkg.org/badges/xlsx)](https://CRAN.R-project.org/package=xlsx)
  [![CRAN History](https://cranlogs.r-pkg.org/badges/grand-total/xlsx)](https://CRAN.R-project.org/package=xlsx)
  <!-- badges: end -->

xlsx
========

**An R package to read, write, format Excel 2007 and Excel 97/2000/XP/2003 files**

The package provides R functions to read, write, and format Excel files.  It depends 
on Java, but this makes it available on most operating systems. 

## Install

Stable version from CRAN

```r
install.packages('xlsx')
```

Or development version from GitHub

```r
devtools::install_github('colearendt/xlsx')
```

## Common Problems

This package depends on Java and the [`rJava`](https://www.rforge.net/rJava/) package to make the connection between R and Java seamless.  In order to use the `xlsx` package, you will need to:

- Ensure you have a `jdk` (Java Development Kit, version >= 1.5) installed for your Operating System.  More information can be found on [Oracle's website](https://www.oracle.com/java/technologies/javase-downloads.html)

- Ensure that the system environment variable `JAVA_HOME` is configured appropriately and points to your `jdk` of choice.  Typically, this will be included in your `PATH` environment variable as well.  Options and system environmental variables that are available from `R` can be seen with `Sys.getenv()`

- Particularly on UNIX systems, if you continue experiencing issues, you may need to reconfigure `R`'s support for Java on your system.  From a terminal, use the command `R CMD javareconf`.  You may need to run this as root or prepended with `sudo` to ensure it has appropriate permission.

More detail can be found in the [`rJava` docs](https://www.rforge.net/rJava/).

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

To report a bug, use the Issues page at: https://github.com/colearendt/xlsx/issues

If you are wrestling with the Java dependency, there are some very good
alternatives that do not require Java. Your choice will vary depending on what
you are trying to accomplish.

- [openxlsx](https://github.com/awalker89/openxlsx)
- [readxl](https://readxl.tidyverse.org/)
- [writexl](https://docs.ropensci.org/writexl/)

## Acknowledgements

The package is made possible thanks to the excellent
work on [Apache POI](https://poi.apache.org/components/spreadsheet/index.html).
 
## Code of Conduct

Please note that the xlsx project is released with a [Contributor Code of Conduct](https://contributor-covenant.org/version/2/0/CODE_OF_CONDUCT.html). By contributing to this project, you agree to abide by its terms.
