# Setup

## Platform

|setting  |value                        |
|:--------|:----------------------------|
|version  |R version 3.4.4 (2018-03-15) |
|system   |x86_64, linux-gnu            |
|ui       |RStudio (1.1.447)            |
|language |(EN)                         |
|collate  |en_US.UTF-8                  |
|tz       |Etc/UTC                      |
|date     |2018-06-01                   |

## Packages

|package    |*  |version    |date       |source                     |
|:----------|:--|:----------|:----------|:--------------------------|
|covr       |   |3.1.0      |2018-05-18 |cran (@3.1.0)              |
|knitr      |   |1.20       |2018-02-20 |cran (@1.20)               |
|RefManageR |   |1.2.0      |2018-04-25 |cran (@1.2.0)              |
|rJava      |   |0.9-10     |2018-05-29 |cran (@0.9-10)             |
|rmarkdown  |   |1.9        |2018-03-01 |cran (@1.9)                |
|rprojroot  |   |1.3-2      |2018-01-03 |cran (@1.3-2)              |
|testthat   |   |2.0.0      |2017-12-13 |cran (@2.0.0)              |
|tibble     |   |1.4.2      |2018-01-22 |cran (@1.4.2)              |
|xlsx       |   |0.6.0.9000 |2018-06-01 |local (colearendt/xlsx@NA) |
|xlsxjars   |   |0.6.1      |2014-08-22 |cran (@0.6.1)              |

# Check results

12 packages with problems

|package       |version | errors| warnings| notes|
|:-------------|:-------|------:|--------:|-----:|
|caRpools      |0.83    |      1|        0|     0|
|compareGroups |3.4.0   |      1|        0|     0|
|diveRsity     |1.9.90  |      1|        0|     0|
|flowQB        |2.6.0   |      1|        0|     0|
|geoSpectral   |0.17.4  |      1|        0|     0|
|ImportExport  |1.1     |      1|        0|     0|
|InterfaceqPCR |1.0     |      1|        0|     0|
|ISM           |0.1.0   |      1|        0|     0|
|IsoGeneGUI    |2.14.0  |      1|        0|     0|
|polmineR      |0.7.8   |      1|        0|     0|
|polyPK        |3.0.0   |      1|        0|     0|
|survHE        |1.0.6   |      1|        0|     0|

## caRpools (0.83)
Maintainer: Jan Winter <jan.winter@dkfz-heidelberg.de>  
Bug reports: https://github.com/boutroslab/caRpools

1 error  | 0 warnings | 0 notes

```
checking package dependencies ... ERROR
Package required but not available: ‘DESeq2’

See section ‘The DESCRIPTION file’ in the ‘Writing R Extensions’
manual.
```

## compareGroups (3.4.0)
Maintainer: Isaac Subirana <isubirana@imim.es>

1 error  | 0 warnings | 0 notes

```
checking package dependencies ... ERROR
Packages required but not available: ‘Hmisc’ ‘SNPassoc’

See section ‘The DESCRIPTION file’ in the ‘Writing R Extensions’
manual.
```

## diveRsity (1.9.90)
Maintainer: Kevin Keenan <kkeenan02@qub.ac.uk>

1 error  | 0 warnings | 0 notes

```
checking package dependencies ... ERROR
Package required but not available: ‘qgraph’

Package suggested but not available for checking: ‘sendplot’

See section ‘The DESCRIPTION file’ in the ‘Writing R Extensions’
manual.
```

## flowQB (2.6.0)
Maintainer: Josef Spidlen <jspidlen@gmail.com>

1 error  | 0 warnings | 0 notes

```
checking package dependencies ... ERROR
Package required but not available: ‘extremevalues’

See section ‘The DESCRIPTION file’ in the ‘Writing R Extensions’
manual.
```

## geoSpectral (0.17.4)
Maintainer: Servet Ahmet Cizmeli <ahmet@pranageo.com>

1 error  | 0 warnings | 0 notes

```
checking package dependencies ... ERROR
Package required but not available: ‘rgdal’

See section ‘The DESCRIPTION file’ in the ‘Writing R Extensions’
manual.
```

## ImportExport (1.1)
Maintainer: Isaac Subirana <isubirana@imim.es>

1 error  | 0 warnings | 0 notes

```
checking package dependencies ... ERROR
Packages required but not available: ‘Hmisc’ ‘RODBC’

Package suggested but not available for checking: ‘compareGroups’

See section ‘The DESCRIPTION file’ in the ‘Writing R Extensions’
manual.
```

## InterfaceqPCR (1.0)
Maintainer: Olivier LE GOFF <olivierlegoff1@gmail.com>

1 error  | 0 warnings | 0 notes

```
checking package dependencies ... ERROR
Package required but not available: ‘tkrplot’

See section ‘The DESCRIPTION file’ in the ‘Writing R Extensions’
manual.
```

## ISM (0.1.0)
Maintainer: Gunjan Bansal <gunjan.1512@gmail.com>

1 error  | 0 warnings | 0 notes

```
checking examples ... ERROR
Running examples in ‘ISM-Ex.R’ failed
The error most likely occurred in:

> base::assign(".ptime", proc.time(), pos = "CheckExEnv")
> ### Name: ISM
> ### Title: Interpretive Structural Modeling (ISM).
> ### Aliases: ISM
> 
> ### ** Examples
> 
> ISM(fname=matrix(c(1,1,1,1,1,0,1,1,1,1,0,0,1,0,0,0,1,1,1,1,0,1,1,0,1),5,5,byrow=TRUE),Dir=tempdir())
Error in .jcall(sheet, "V", "autoSizeColumn", as.integer(ic), TRUE) : 
  java.lang.InternalError: Can't connect to X11 window server using '' as the value of the DISPLAY variable.
Calls: ISM -> <Anonymous> -> .jcall -> .jcheck -> .Call
Execution halted
```

## IsoGeneGUI (2.14.0)
Maintainer: Setia Pramana <setia.pramana@ki.se>

1 error  | 0 warnings | 0 notes

```
checking package dependencies ... ERROR
Package required but not available: ‘tkrplot’

See section ‘The DESCRIPTION file’ in the ‘Writing R Extensions’
manual.
```

## polmineR (0.7.8)
Maintainer: Andreas Blaette <andreas.blaette@uni-due.de>  
Bug reports: https://github.com/PolMine/polmineR/issues

1 error  | 0 warnings | 0 notes

```
checking examples ... ERROR
Running examples in ‘polmineR-Ex.R’ failed
The error most likely occurred in:

> base::assign(".ptime", proc.time(), pos = "CheckExEnv")
> ### Name: textstat-class
> ### Title: S4 textstat class
> ### Aliases: textstat-class as.data.frame,textstat-method
> ###   show,textstat-method dim,textstat-method colnames,textstat-method
> ###   rownames,textstat-method names,textstat-method
... 15 lines ...
... get encoding: latin1
... get cpos and strucs
... getting counts for p-attribute(s): word
... using RcppCWB
> y <- cooccurrences(P, query = "Arbeit")
> y[1:25]
Error in get("View", envir = .GlobalEnv)(.Object@stat) : 
  View() should not be used in examples etc
Calls: <Anonymous> ... <Anonymous> -> view -> view -> .local -> <Anonymous>
Execution halted
** found \donttest examples: check also with --run-donttest
```

## polyPK (3.0.0)
Maintainer: Tianlu Chen <chentianlu@sjtu.edu.cn>

1 error  | 0 warnings | 0 notes

```
checking package dependencies ... ERROR
Packages required but not available: ‘imputeLCMD’ ‘Hmisc’

See section ‘The DESCRIPTION file’ in the ‘Writing R Extensions’
manual.
```

## survHE (1.0.6)
Maintainer: Gianluca Baio <gianluca@stats.ucl.ac.uk>  
Bug reports: https://github.com/giabaio/survHE/issues

1 error  | 0 warnings | 0 notes

```
checking package dependencies ... ERROR
Package required but not available: ‘rms’

Package suggested but not available for checking: ‘INLA’

See section ‘The DESCRIPTION file’ in the ‘Writing R Extensions’
manual.
```

