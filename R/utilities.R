# .guess_cell_type
# .onAttach
# .onLoad
# .Rcolor2XLcolor
# .splitBlocks
# .hssfcolor
# .xssfcolor  



###################################################################
# given a vector of cells, return the "appropriate" R type
# you cannot decide if it's a Date or a POSIXct
#
# only returns "numeric", "character", "POSIXct"
#
.guess_cell_type <- function(cells)
{
  cellType <- sapply(cells, function(x){x$getCellType()}) + 1

  dateUtil   <- .jnew("org/apache/poi/ss/usermodel/DateUtil")
  isDateTime <- mapply(function(x,y) {
    if (y != 1) {  # if it's not numeric, forget about it
      FALSE
    } else {
      dateUtil$isCellDateFormatted(x)
    } 
  }, cells, cellType)

  guessedColClasses <- ifelse(cellType==1,"numeric", "character")

  guessedColClasses[isDateTime] <- "POSIXct"
  
  guessedColClasses
}


####################################################################
#
.onAttach <- function(libname, pkgname)
{
  .jpackage(pkgname)  # needed to load RInterface.java
  
  # what's your java  version?  Need > 1.5.0.
  jversion <- .jcall('java.lang.System','S','getProperty','java.version')
  if (jversion < "1.5.0")
    stop(paste("Your java version is ", jversion,
                 ".  Need 1.5.0 or higher.", sep=""))
  
  wb <- createWorkbook()   # load/initialize jars here as it takes 
  rm(wb)                   # a few seconds ...
 }


####################################################################
#
.onLoad <- function(libname, pkgname)
{
  .jpackage("xlsxjars")
  
  options(xlsx.date.format = "m/d/yyyy")   # e.g. 3/18/2013
  options(xlsx.datetime.format = "m/d/yyyy h:mm:ss")  # e.g. 3/18/2013 05:25:51
  
}

####################################################################
# Converts R color into Excel compatible color
# rcolor <- c("red", "blue")
#
.Rcolor2XLcolor <- function( rcolor, isXSSF=TRUE )
{
  if ( isXSSF ) {
    sapply( rcolor, .xssfcolor )
  } else {                     
    sapply( rcolor, .hssfcolor ) 
  }    
}


###################################################################
# Split a vector into contiguous chunks c(1,2,3, 6,7, 9,10,11)
# @param x an integer vector
# @return a list of integer vectors, each vector being a group.
#
.splitBlocks <- function(x)
{
  blocks <- list(x[1])
  if (length(x)==1)
    return(blocks)
  
  k <- 1
  for (i in 2:length(x)) {
    if (x[i]==(x[i-1]+1)) {
      blocks[[k]] <- c(blocks[[k]], x[i])
    } else {
      k <- k+1
      blocks[[k]] <- x[i]
    }
  }

  blocks
}


########################################################################
# Take an R color and return an HSSFColor object
# @param Rcolor is a string
#   NOTE: Only a limited set of colors are supported.  See the java API.
# 
.hssfcolor <- function(Rcolor)
{
  .jshort(INDEXED_COLORS_[toupper( Rcolor )])
}


########################################################################
# Take an R color and return an XSSFColor object
# @param Rcolor is a string returned from "colors()" or a hex string
#   starting with #, e.g. "#FF0000" is red.
#
.xssfcolor <- function(Rcolor)
{
  if (grepl("^#", Rcolor)) {
    rgb <- c(strtoi(substring(Rcolor, 2, 3), base=16L),
             strtoi(substring(Rcolor, 4, 5), base=16L),
             strtoi(substring(Rcolor, 6, 7), base=16L)) 
  } else {     
    rgb <- as.integer(col2rgb(Rcolor))
  }
  jcol <- .jnew("java.awt.Color", rgb[1], rgb[2], rgb[3])
  
  .jnew("org.apache.poi.xssf.usermodel.XSSFColor", jcol)
}




  HALIGN_STYLES_ <- c(2,6,4,0,5,1,3)
  names(HALIGN_STYLES_) <- c('ALIGN_CENTER' ,'ALIGN_CENTER_SELECTION'
    ,'ALIGN_FILL' ,'ALIGN_GENERAL' ,'ALIGN_JUSTIFY' ,'ALIGN_LEFT'
    ,'ALIGN_RIGHT')

  VALIGN_STYLES_<- c(2,1,3,0)
  names(VALIGN_STYLES_) <- c('VERTICAL_BOTTOM' ,'VERTICAL_CENTER',
                             'VERTICAL_JUSTIFY' ,'VERTICAL_TOP')
              
  BORDER_STYLES_ <- c(9,11,3,7,6,4,2,10,12,8,0,13,5,1)
  names(BORDER_STYLES_) <- c("BORDER_DASH_DOT", "BORDER_DASH_DOT_DOT",
    "BORDER_DASHED", "BORDER_DOTTED", "BORDER_DOUBLE", "BORDER_HAIR",
    "BORDER_MEDIUM", "BORDER_MEDIUM_DASH_DOT",
    "BORDER_MEDIUM_DASH_DOT_DOT", "BORDER_MEDIUM_DASHED",
    "BORDER_NONE", "BORDER_SLANTED_DASH_DOT", "BORDER_THICK",
    "BORDER_THIN")

  FILL_STYLES_<- c(3,9,10,16,2,18,17,0,1,4,15,7,8,5,6,13,14,11,12)
  names(FILL_STYLES_) <- c('ALT_BARS', 'BIG_SPOTS','BRICKS' ,'DIAMONDS'
    ,'FINE_DOTS' ,'LEAST_DOTS' ,'LESS_DOTS' ,'NO_FILL'
    ,'SOLID_FOREGROUND' ,'SPARSE_DOTS' ,'SQUARES'
    ,'THICK_BACKWARD_DIAG' ,'THICK_FORWARD_DIAG' ,'THICK_HORZ_BANDS'
    ,'THICK_VERT_BANDS' ,'THIN_BACKWARD_DIAG' ,'THIN_FORWARD_DIAG'
    ,'THIN_HORZ_BANDS' ,'THIN_VERT_BANDS')

  # from org.apache.poi.ss.usermodel.CellStyle  
  CELL_STYLES_ <- c(HALIGN_STYLES_, VALIGN_STYLES_,
     BORDER_STYLES_, FILL_STYLES_)

  INDEXED_COLORS_ <- c(8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,
    25,26,28,29,30,31,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,
    56,57,58,59,60,61,62,63,64)
  names(INDEXED_COLORS_) <- c('BLACK' ,'WHITE' ,'RED' ,'BRIGHT_GREEN'
    ,'BLUE' ,'YELLOW' ,'PINK' ,'TURQUOISE' ,'DARK_RED' ,'GREEN'
    ,'DARK_BLUE' ,'DARK_YELLOW' ,'VIOLET' ,'TEAL' ,'GREY_25_PERCENT'
    ,'GREY_50_PERCENT' ,'CORNFLOWER_BLUE' ,'MAROON' ,'LEMON_CHIFFON'
    ,'ORCHID' ,'CORAL' ,'ROYAL_BLUE' ,'LIGHT_CORNFLOWER_BLUE'
    ,'SKY_BLUE' ,'LIGHT_TURQUOISE' ,'LIGHT_GREEN' ,'LIGHT_YELLOW'
    ,'PALE_BLUE' ,'ROSE' ,'LAVENDER' ,'TAN' ,'LIGHT_BLUE' ,'AQUA'
    ,'LIME' ,'GOLD' ,'LIGHT_ORANGE' ,'ORANGE' ,'BLUE_GREY'
    ,'GREY_40_PERCENT' ,'DARK_TEAL' ,'SEA_GREEN' ,'DARK_GREEN'
    ,'OLIVE_GREEN' ,'BROWN' ,'PLUM' ,'INDIGO' ,'GREY_80_PERCENT'
    ,'AUTOMATIC')
  
                   



# probably should use
# .jfield("org.apache.poi.ss.usermodel.CellStyle", NULL, "ALIGN_CENTER") # 2
## CELL_STYLES <- function(x) {
## }

 


##   aux <- 
##    "BLACK(8),
##     WHITE(9),
##     RED(10),
##     BRIGHT_GREEN(11),
##     BLUE(12),
##     YELLOW(13),
##     PINK(14),
##     TURQUOISE(15),
##     DARK_RED(16),
##     GREEN(17),
##     DARK_BLUE(18),
##     DARK_YELLOW(19),
##     VIOLET(20),
##     TEAL(21),
##     GREY_25_PERCENT(22),
##     GREY_50_PERCENT(23),
##     CORNFLOWER_BLUE(24),
##     MAROON(25),
##     LEMON_CHIFFON(26),
##     ORCHID(28),
##     CORAL(29),
##     ROYAL_BLUE(30),
##     LIGHT_CORNFLOWER_BLUE(31),
##     SKY_BLUE(40),
##     LIGHT_TURQUOISE(41),
##     LIGHT_GREEN(42),
##     LIGHT_YELLOW(43),
##     PALE_BLUE(44),
##     ROSE(45),
##     LAVENDER(46),
##     TAN(47),
##     LIGHT_BLUE(48),
##     AQUA(49),
##     LIME(50),
##     GOLD(51),
##     LIGHT_ORANGE(52),
##     ORANGE(53),
##     BLUE_GREY(54),
##     GREY_40_PERCENT(55),
##     DARK_TEAL(56),
##     SEA_GREEN(57),
##     DARK_GREEN(58),
##     OLIVE_GREEN(59),
##     BROWN(60),
##     PLUM(61),
##     INDIGO(62),
##     GREY_80_PERCENT(63),
##     AUTOMATIC(64)"
## bux <- gsub("^ *", "", strsplit(aux, ",\n")[[1]])
## fields <- paste(gsub("(.*)\\([[:digit:]]+\\)", "\\1", bux), sep="",
##   collapse="' ,'")
## values <- paste(as.numeric(gsub(".*\\(([[:digit:]]+)\\)", "\\1", bux)),
##   sep="", collapse=",")


## aux <- "
## public static final byte	ANSI_CHARSET	0
## public static final short	BOLDWEIGHT_BOLD	700
## public static final short	BOLDWEIGHT_NORMAL	400
## public static final short	COLOR_NORMAL	32767
## public static final short	COLOR_RED	10
## public static final byte	DEFAULT_CHARSET	1
## public static final short	SS_NONE	0
## public static final short	SS_SUB	2
## public static final short	SS_SUPER	1
## public static final byte	SYMBOL_CHARSET	2
## public static final byte	U_DOUBLE	2
## public static final byte	U_DOUBLE_ACCOUNTING	34
## public static final byte	U_NONE	0
## public static final byte	U_SINGLE	1
## public static final byte	U_SINGLE_ACCOUNTING	33"
## bux <- strsplit(aux, " ")[[1]])
## fields <- paste(gsub("(.*)\\([[:digit:]]+\\)", "\\1", bux), sep="",
##   collapse="' ,'")
## values <- paste(as.numeric(gsub(".*\\(([[:digit:]]+)\\)", "\\1", bux)),
##   sep="", collapse=",")

