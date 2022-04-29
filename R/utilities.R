# .guess_cell_type
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
.onLoad <- function(libname, pkgname)
{
  rJava::.jpackage("xlsxjars")
  rJava::.jpackage(pkgname, lib.loc = libname)  # needed to load RInterface.java

  set_java_tmp_dir(getOption("xlsx.tempdir", tempdir()))

  wb <- createWorkbook()   # load/initialize jars here as it takes
  rm(wb)                   # a few seconds ...

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
#' @importFrom grDevices col2rgb
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

  # default_map <- .jnew("org.apache.poi.xssf.usermodel.DefaultIndexedColorMap")

  .jnew("org.apache.poi.xssf.usermodel.XSSFColor", .jarray(.jbyte(rgb[1]), .jbyte(rgb[2]), .jbyte(rgb[3])))
}


#' Set Java Temp Directory
#'
#' Java sets the java temp directory to `/tmp` by default. However, this is
#' usually not desirable in R. As a result, this function allows changing that
#' behavior. Further, this function is fired on package load to ensure all
#' temp files are written to the R temp directory.
#'
#' On package load, we use `getOption("xlsx.tempdir", tempdir())` for the
#' default value, in case you want to have this value set by an option.
#'
#' @param tmp_dir optional. The new temp directory. Defaults to the R temp
#'   directory
#'
#' @return The previous java temp directory (prior to any changes).
#'
#' @export
set_java_tmp_dir <- function(tmp_dir = tempdir()) {
  rJava::.jcall(
    "java/lang/System",
    "Ljava/lang/String;",
    "setProperty",
    "java.io.tmpdir", tmp_dir
  )
}

#' @rdname set_java_tmp_dir
#' @export
get_java_tmp_dir <- function() {
  rJava::.jcall(
    "java/lang/System",
    "Ljava/lang/String;",
    "getProperty",
    "java.io.tmpdir"
  )
}

#' @title Constants used in the project.
#'
#' @description Document some Apache POI constants used in the project.
#'
#' @return A named vector.
#' @seealso \code{\link{CellStyle}} for using the \code{POI_constants}.
#' @name POI_constants
NULL


# Horizontal Alignment ---------------------------------------

  halign_styles_raw <- .jcall(.jfindClass("org.apache.poi.ss.usermodel.HorizontalAlignment"), "[Ljava/lang/Object;", "getEnumConstants")

  halign_styles_names <- as.character(lapply(halign_styles_raw, function(.x) .jstrVal(.x)))
  halign_styles_names_old <- paste0("ALIGN_", halign_styles_names)

  #' @rdname POI_constants
  #' @export
  style_horizontal <- setNames(halign_styles_raw, halign_styles_names)

  #' @rdname POI_constants
  #' @export
  HALIGN_STYLES_ <- setNames(halign_styles_raw, halign_styles_names_old)

  style_horizontal_all <- c(style_horizontal, HALIGN_STYLES_)

# Vertical Alignment ---------------------------------------

  valign_styles_raw <- .jcall(
    .jfindClass("org.apache.poi.ss.usermodel.VerticalAlignment"),
    "[Ljava/lang/Object;",
    "getEnumConstants"
  )

  valign_styles_names <- as.character(lapply(valign_styles_raw, function(.x) .jstrVal(.x)))
  valign_styles_names_old <- paste0("VERTICAL_", valign_styles_names)

  #' @rdname POI_constants
  #' @export
  VALIGN_STYLES_ <- setNames(valign_styles_raw, valign_styles_names_old)

  #' @rdname POI_constants
  #' @export
  style_vertical <- setNames(valign_styles_raw, valign_styles_names)

  style_vertical_all <- c(style_vertical, VALIGN_STYLES_)

# Border Styles ---------------------------------------

  border_styles_raw <- .jcall(
    .jfindClass("org.apache.poi.ss.usermodel.BorderStyle"),
    "[Ljava/lang/Object;",
    "getEnumConstants"
  )

  border_styles_names <- as.character(lapply(border_styles_raw, function(.x) .jstrVal(.x)))
  border_styles_names_old <- paste0("BORDER_", border_styles_names)

  #' @rdname POI_constants
  #' @export
  BORDER_STYLES_ <- setNames(border_styles_raw, border_styles_names_old)

  #' @rdname POI_constants
  #' @export
  style_border <- setNames(border_styles_raw, border_styles_names)

  style_border_all <- c(style_border, BORDER_STYLES_)

# Fill Pattern Styles ---------------------------------------

  fill_pattern_styles_raw <- .jcall(
    .jfindClass("org.apache.poi.ss.usermodel.FillPatternType"),
    "[Ljava/lang/Object;",
    "getEnumConstants"
  )

  fill_pattern_styles_names <- as.character(lapply(fill_pattern_styles_raw, function(.x) .jstrVal(.x)))

  #' @rdname POI_constants
  #' @export
  FILL_STYLES_ <- setNames(fill_pattern_styles_raw, fill_pattern_styles_names)

  #' @rdname POI_constants
  #' @export
  style_fill_pattern <- setNames(fill_pattern_styles_raw, fill_pattern_styles_names)

# All Cell Styles ---------------------------------------

  # from org.apache.poi.ss.usermodel.CellStyle
  #' @rdname POI_constants
  #' @export
  CELL_STYLES_ <- c(
    HALIGN_STYLES_, style_horizontal,
    VALIGN_STYLES_, style_vertical,
    BORDER_STYLES_, style_border,
    FILL_STYLES_, style_fill_pattern
    )

# Hyperlink Types

  hyperlink_types_raw <- .jcall(
    .jfindClass("org.apache.poi.common.usermodel.HyperlinkType"),
    "[Ljava/lang/Object;",
    "getEnumConstants"
  )

  hyperlink_types_names <- as.character(lapply(hyperlink_types_raw, function (.x) .jstrVal(.x)))

  #' @rdname POI_constants
  #' @export
  types_hyperlink <- setNames(hyperlink_types_raw, hyperlink_types_names)

# Indexed Colors ---------------------------------------

  #' @rdname POI_constants
  #' @export
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
