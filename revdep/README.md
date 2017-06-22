# Setup

## Platform

|setting  |value                        |
|:--------|:----------------------------|
|version  |R version 3.4.0 (2017-04-21) |
|system   |x86_64, linux-gnu            |
|ui       |X11                          |
|language |en_US                        |
|collate  |en_US.UTF-8                  |
|tz       |America/New_York             |
|date     |2017-06-21                   |

## Packages

|package   |*  |version    |date       |source                            |
|:---------|:--|:----------|:----------|:---------------------------------|
|covr      |   |2.2.2      |2017-01-05 |cran (@2.2.2)                     |
|rJava     |   |0.9-8      |2016-01-07 |cran (@0.9-8)                     |
|rprojroot |   |1.2        |2017-01-16 |cran (@1.2)                       |
|testthat  |   |1.0.2      |2016-04-23 |cran (@1.0.2)                     |
|tibble    |   |1.3.3      |2017-06-21 |Github (tidyverse/tibble@b2275d5) |
|xlsx      |   |0.6.0.9000 |2017-06-21 |local (colearendt/xlsx@NA)        |
|xlsxjars  |   |0.6.1      |2014-08-22 |cran (@0.6.1)                     |

# Check results

24 packages

|package         |version | errors| warnings| notes|
|:---------------|:-------|------:|--------:|-----:|
|caRpools        |0.83    |      1|        0|     0|
|compareGroups   |3.2.4   |      1|        0|     0|
|condformat      |0.6.0   |      0|        0|     1|
|DataClean       |1.0     |      0|        0|     0|
|DataLoader      |1.3     |      0|        0|     0|
|diveRsity       |1.9.90  |      1|        0|     0|
|ecoseries       |0.1.3   |      0|        0|     0|
|ELT             |1.6     |      0|        0|     0|
|G2Sd            |2.1.5   |      0|        0|     0|
|geoSpectral     |0.17.3  |      1|        0|     0|
|ImportExport    |1.1     |      1|        0|     0|
|InterfaceqPCR   |1.0     |      1|        0|     0|
|lar             |0.1-2   |      0|        0|     0|
|ODMconverter    |2.3     |      0|        0|     0|
|OpenRepGrid     |0.1.10  |      1|        0|     0|
|PairwiseD       |0.9.62  |      0|        0|     0|
|polmineR        |0.7.3   |      0|        0|     1|
|ProjectTemplate |0.7     |      0|        0|     1|
|psData          |0.2.2   |      0|        0|     0|
|qdap            |2.2.5   |      0|        0|     0|
|repmis          |0.5     |      0|        0|     0|
|RJafroc         |0.1.1   |      0|        0|     0|
|rpcdsearch      |1.0     |      0|        0|     0|
|rrepast         |0.6.0   |      0|        0|     0|

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

## condformat (0.6.0)
Maintainer: Sergio Oller Moreno <sergioller@gmail.com>  
Bug reports: http://github.com/zeehio/condformat/issues

0 errors | 0 warnings | 1 note 

```
checking dependencies in R code ... NOTE
Missing or unexported object: ‘xlsx::CellBlock.default’
```

## DataClean (1.0)
Maintainer: Xiaorui(Jeremy) Zhu <zhuxiaorui1989@gmail.com>

0 errors | 0 warnings | 0 notes

## DataLoader (1.3)
Maintainer: Srivenkatesh Gandhi <srivenkateshg@sase.ssn.edu.in>

0 errors | 0 warnings | 0 notes

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

## ecoseries (0.1.3)
Maintainer: Fernando Teixeira <fernando.teixeira@fgv.br>  
Bug reports: https://github.com/fernote7/ecoseries/issues

0 errors | 0 warnings | 0 notes

## ELT (1.6)
Maintainer: Wassim Youssef <Wassim.G.Youssef@gmail.com>

0 errors | 0 warnings | 0 notes

## G2Sd (2.1.5)
Maintainer: Regis K. Gallon <reg.gallon@gmail.com>

0 errors | 0 warnings | 0 notes

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

## lar (0.1-2)
Maintainer: Richard Zijdeman <richard.zijdeman@iisg.nl>

0 errors | 0 warnings | 0 notes

## ODMconverter (2.3)
Maintainer: Martin Dugas <dugas@uni-muenster.de>

0 errors | 0 warnings | 0 notes

## OpenRepGrid (0.1.10)
Maintainer: Mark Heckmann <heckmann@uni-bremen.de>

1 error  | 0 warnings | 0 notes

```
checking package dependencies ... ERROR
Package required but not available: ‘rgl’

See section ‘The DESCRIPTION file’ in the ‘Writing R Extensions’
manual.
```

## PairwiseD (0.9.62)
Maintainer: Marcin Stryjek <pairwised@post.pl>

0 errors | 0 warnings | 0 notes

## polmineR (0.7.3)
Maintainer: Andreas Blaette <andreas.blaette@uni-due.de>  
Bug reports: https://github.com/PolMine/polmineR/issues

0 errors | 0 warnings | 1 note 

```
checking package dependencies ... NOTE
Packages suggested but not available for checking:
  ‘polmineR.Rcpp’ ‘europarl.en’ ‘rcqp’
```

## ProjectTemplate (0.7)
Maintainer: Kenton White <jkentonwhite@gmail.com>  
Bug reports: https://github.com/johnmyleswhite/ProjectTemplate/issues

0 errors | 0 warnings | 1 note 

```
checking package dependencies ... NOTE
Packages suggested but not available for checking: ‘RMySQL’ ‘RODBC’
```

## psData (0.2.2)
Maintainer: Christopher Gandrud <christopher.gandrud@gmail.com>  
Bug reports: https://github.com/christophergandrud/psData/issues

0 errors | 0 warnings | 0 notes

## qdap (2.2.5)
Maintainer: Tyler Rinker <tyler.rinker@gmail.com>  
Bug reports: http://github.com/trinker/qdap/issues

0 errors | 0 warnings | 0 notes

## repmis (0.5)
Maintainer: Christopher Gandrud <christopher.gandrud@gmail.com>  
Bug reports: https://github.com/christophergandrud/repmis/issues

0 errors | 0 warnings | 0 notes

## RJafroc (0.1.1)
Maintainer: Xuetong Zhai <xuetong.zhai@gmail.com>

0 errors | 0 warnings | 0 notes

## rpcdsearch (1.0)
Maintainer: David Springate <daspringate@gmail.com>

0 errors | 0 warnings | 0 notes

## rrepast (0.6.0)
Maintainer: Antonio Prestes Garcia <antonio.pgarcia@alumnos.upm.es>

0 errors | 0 warnings | 0 notes

