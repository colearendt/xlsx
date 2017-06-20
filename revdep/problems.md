# Setup

## Platform

|setting  |value                        |
|:--------|:----------------------------|
|version  |R version 3.4.0 (2017-04-21) |
|system   |x86_64, linux-gnu            |
|ui       |RStudio (1.0.143)            |
|language |en_US                        |
|collate  |en_US.UTF-8                  |
|tz       |America/New_York             |
|date     |2017-06-20                   |

## Packages

|package   |*  |version |date       |source                     |
|:---------|:--|:-------|:----------|:--------------------------|
|covr      |   |2.2.2   |2017-01-05 |cran (@2.2.2)              |
|rJava     |*  |0.9-8   |2016-01-07 |cran (@0.9-8)              |
|rprojroot |   |1.2     |2017-01-16 |cran (@1.2)                |
|testthat  |*  |1.0.2   |2016-04-23 |cran (@1.0.2)              |
|tibble    |   |1.3.3   |2017-05-28 |cran (@1.3.3)              |
|xlsx      |*  |0.6.0   |2017-06-20 |local (colearendt/xlsx@NA) |
|xlsxjars  |*  |0.6.1   |2014-08-22 |cran (@0.6.1)              |

# Check results

9 packages with problems

|package       |version | errors| warnings| notes|
|:-------------|:-------|------:|--------:|-----:|
|caRpools      |0.83    |      1|        0|     0|
|compareGroups |3.2.4   |      1|        0|     0|
|diveRsity     |1.9.90  |      1|        0|     0|
|geoSpectral   |0.17.3  |      1|        0|     0|
|ImportExport  |1.1     |      1|        0|     0|
|InterfaceqPCR |1.0     |      1|        0|     0|
|OpenRepGrid   |0.1.10  |      1|        0|     0|
|qdap          |2.2.5   |      0|        1|     0|
|RJafroc       |0.1.1   |      0|        1|     0|

## caRpools (0.83)
Maintainer: Jan Winter <jan.winter@dkfz-heidelberg.de>  
Bug reports: https://github.com/boutroslab/caRpools

1 error  | 0 warnings | 0 notes

```
checking package dependencies ... ERROR
Packages required but not available: ‘DESeq2’ ‘biomaRt’

Package suggested but not available for checking: ‘BiocGenerics’

See section ‘The DESCRIPTION file’ in the ‘Writing R Extensions’
manual.
```

## compareGroups (3.2.4)
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

## geoSpectral (0.17.3)
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

## OpenRepGrid (0.1.10)
Maintainer: Mark Heckmann <heckmann@uni-bremen.de>

1 error  | 0 warnings | 0 notes

```
checking package dependencies ... ERROR
Package required but not available: ‘rgl’

See section ‘The DESCRIPTION file’ in the ‘Writing R Extensions’
manual.
```

## qdap (2.2.5)
Maintainer: Tyler Rinker <tyler.rinker@gmail.com>  
Bug reports: http://github.com/trinker/qdap/issues

0 errors | 1 warning  | 0 notes

```
checking re-building of vignette outputs ... WARNING
Error in re-building vignettes:
  ...
Warning: running command 'kpsewhich framed.sty' had status 1
Warning in test_latex_pkg("framed", system.file("misc", "framed.sty", package = "knitr")) :
  unable to find LaTeX package 'framed'; will use a copy from knitr
Error in texi2dvi(file = file, pdf = TRUE, clean = clean, quiet = quiet,  : 
  Running 'texi2dvi' on 'cleaning_and_debugging.tex' failed.
Messages:
sh: 1: /usr/bin/texi2dvi: not found
Calls: buildVignettes -> texi2pdf -> texi2dvi
Execution halted

```

## RJafroc (0.1.1)
Maintainer: Xuetong Zhai <xuetong.zhai@gmail.com>

0 errors | 1 warning  | 0 notes

```
checking re-building of vignette outputs ... WARNING
Error in re-building vignettes:
  ...
Warning: running command 'kpsewhich framed.sty' had status 1
Warning in test_latex_pkg("framed", system.file("misc", "framed.sty", package = "knitr")) :
  unable to find LaTeX package 'framed'; will use a copy from knitr
Loading required package: xlsx
Loading required package: rJava
Loading required package: xlsxjars
Loading required package: ggplot2
Loading required package: stringr
Loading required package: shiny
Error in texi2dvi(file = file, pdf = TRUE, clean = clean, quiet = quiet,  : 
  Running 'texi2dvi' on 'RJafroc.tex' failed.
Messages:
sh: 1: /usr/bin/texi2dvi: not found
Calls: buildVignettes -> texi2pdf -> texi2dvi
Execution halted

```

