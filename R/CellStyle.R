######################################################################
# Create a cell style.  It needs a workbook object!
#

#' Functions to manipulate cells.
#'
#' Create and set cell styles.
#'
#' \code{setCellStyle} sets the \code{CellStyle} to one \code{Cell} object.
#'
#' You need to have a \code{Workbook} object to attach a \code{CellStyle}
#' object to it.
#'
#' Since OS X 10.5 Apple dropped support for AWT on the main thread, so
#' essentially you cannot use any graphics classes in R on OS X 10.5 since R is
#' single-threaded. (verbatim from Simon Urbanek).  This implies that setting
#' colors on Mac will not work as is!  A set of about 50 basic colors are still
#' available please see the javadocs.
#'
#' For Excel 95/2000/XP/2003 the choice of colors is limited.  See
#' \code{INDEXED_COLORS_} for the list of available colors.
#'
#' Unspecified values for arguments are taken from the system locale.
#'
#' @param wb a workbook object as returned by \code{\link{createWorkbook}} or
#' \code{\link{loadWorkbook}}.
#' @param dataFormat a \code{\link{DataFormat}} object.
#' @param alignment a \code{\link{Alignment}} object.
#' @param border a \code{\link{Border}} object.
#' @param fill a \code{\link{Fill}} object.
#' @param font a \code{\link{Font}} object.
#' @param cellProtection a \code{\link{CellProtection}} object.
#' @param x a \code{CellStyle} object.
#' @param cell a \code{\link{Cell}} object.
#' @param cellStyle a \code{CellStyle} object.
#' @return
#'
#' \code{createCellStyle} creates a CellStyle object.
#'
#' \code{is.CellStyle} returns \code{TRUE} if the argument is of class
#' "CellStyle" and \code{FALSE} otherwise.
#' @author Adrian Dragulescu
#' @examples
#'
#' \dontrun{
#'   wb <- createWorkbook()
#'   sheet <- createSheet(wb, "Sheet1")
#'
#'   rows  <- createRow(sheet, rowIndex=1)
#'
#'   cell.1 <- createCell(rows, colIndex=1)[[1,1]]
#'   setCellValue(cell.1, "Hello R!")
#'
#'   cs <- CellStyle(wb) +
#'     Font(wb, heightInPoints=20, isBold=TRUE, isItalic=TRUE,
#'       name="Courier New", color="orange") +
#'     Fill(backgroundColor="lavender", foregroundColor="lavender",
#'       pattern="SOLID_FOREGROUND") +
#'     Alignment(h="ALIGN_RIGHT")
#'
#'   setCellStyle(cell.1, cellStyle1)
#'
#'   # you need to save the workbook now if you want to see this art
#' }
#' @export
#' @name CellStyle
CellStyle <- function(wb, dataFormat=NULL, alignment=NULL,
  border=NULL, fill=NULL, font=NULL, cellProtection=NULL) UseMethod("CellStyle")

#' @export
#' @rdname CellStyle
is.CellStyle <- function(x) {inherits(x, "CellStyle")}


######################################################################
# Create a cell style.  It needs a workbook object!
#
#' @export
#' @rdname CellStyle
CellStyle.default <- function(wb, dataFormat=NULL, alignment=NULL,
  border=NULL, fill=NULL, font=NULL, cellProtection=NULL)
{
  if ((!is.null(dataFormat)) & (!inherits(dataFormat, "DataFormat")))
    stop("Argument dataFormat needs to be of class DataFormat.")
  if ((!is.null(alignment)) & (!inherits(alignment, "Alignment")))
    stop("Argument alignment needs to be of class Alignment.")
  if ((!is.null(border)) & (!inherits(border, "Border")))
    stop("Argument border needs to be of class Border.")
  if ((!is.null(fill)) & (!inherits(fill, "Fill")))
    stop("Argument fill needs to be of class Fill.")
  if ((!is.null(font)) & (!inherits(font, "Font")))
    stop("Argument font needs to be of class Font.")

  cellStyle <- .jcall(wb, "Lorg/apache/poi/ss/usermodel/CellStyle;",
    "createCellStyle")

  CS <- CELL_STYLES_

  if (!is.null(dataFormat)){
    fmt <- .jcall(wb, "Lorg/apache/poi/ss/usermodel/DataFormat;",
             "createDataFormat")
    .jcall(cellStyle, "V", "setDataFormat",
           .jshort(fmt$getFormat(dataFormat$dataFormat)))
  }


  # the alignment
  if (!is.null(alignment)) {
    if (!is.null(alignment$horizontal))
      .jcall(cellStyle, "V", "setAlignment", .jshort(CS[alignment$horizontal]))
    if (!is.null(alignment$vertical))
      .jcall(cellStyle, "V", "setVerticalAlignment", .jshort(CS[alignment$vertical]))
    if (alignment$wrapText)
      .jcall(cellStyle, "V", "setWrapText", TRUE)
    if (alignment$rotation != 0)
      .jcall(cellStyle, "V", "setRotation", .jshort(alignment$rotation))
    if (alignment$indent != 0)
      .jcall(cellStyle, "V", "setIndention", .jshort(alignment$indent))
  }

  # the border
  if (!is.null(border)) {
    for (i in 1:length(border$position)) {
      bcolor <- if (grepl("XSSF", wb$getClass()$getName())) {
        .xssfcolor(border$color[i])
       } else {
        idcol <- INDEXED_COLORS_[toupper(border$color[i])]
        .jshort(idcol)
      }
      switch(border$position[i],
          BOTTOM = {
            .jcall(cellStyle, "V", "setBorderBottom",.jshort(CS[border$pen[i]]))
            .jcall(cellStyle, "V", "setBottomBorderColor", bcolor)
            },
          LEFT = {
            .jcall(cellStyle, "V", "setBorderLeft", .jshort(CS[border$pen[i]]))
            .jcall(cellStyle, "V", "setLeftBorderColor", bcolor)
           },
          TOP = {
            .jcall(cellStyle, "V", "setBorderTop", .jshort(CS[border$pen[i]]))
            .jcall(cellStyle, "V", "setTopBorderColor", bcolor)
            },
          RIGHT = {
            .jcall(cellStyle, "V", "setBorderRight",.jshort(CS[border$pen[i]]))
            .jcall(cellStyle, "V", "setRightBorderColor", bcolor)
          }
       )
    }
  }

  # the fill
  if (!is.null(fill)) {
    isXSSF <-grepl("XSSF", wb$getClass()$getName())
    if (isXSSF) {
       .jcall(cellStyle, "V", "setFillForegroundColor",
         .xssfcolor(fill$foregroundColor))
       .jcall(cellStyle, "V", "setFillBackgroundColor",
         .xssfcolor(fill$backgroundColor))
    } else {
       .jcall(cellStyle, "V", "setFillForegroundColor",
         .jshort(INDEXED_COLORS_[toupper(fill$foregroundColor)]))
       .jcall(cellStyle, "V", "setFillBackgroundColor",
         .jshort(INDEXED_COLORS_[toupper(fill$backgroundColor)]))
    }
    .jcall(cellStyle, "V", "setFillPattern", .jshort(CS[fill$pattern]))
  }


  if (!is.null(font))
    .jcall(cellStyle, "V", "setFont", font$ref)

  if (!is.null(cellProtection)) {
    .jcall(cellStyle, "V", "setHidden", cellProtection$hidden)
    .jcall(cellStyle, "V", "setLocked", cellProtection$locked)
  }

  structure(list(ref=cellStyle, wb=wb, dataFormat=dataFormat,
    alignment=alignment, border=border, fill=fill, font=font,
    cellProtection=cellProtection),
    class="CellStyle")
}


######################################################################
#
#' CellStyle construction.
#'
#' Create cell styles.
#'
#' The style of the argument object takes precedence over the style of argument
#' cs1.
#'
#' @param cs1 a \code{\link{CellStyle}} object.
#' @param object an object to add.  The object can be another
#' \code{\link{CellStyle}}, a \code{\link{DataFormat}}, a
#' \code{\link{Alignment}}, a \code{\link{Border}}, a \code{\link{Fill}}, a
#' \code{\link{Font}}, or a \code{\link{CellProtection}} object.
#' @return A CellStyle object.
#' @author Adrian Dragulescu
#' @examples
#'
#' \dontrun{
#'   cs <- CellStyle(wb) +
#'     Font(wb, heightInPoints=20, isBold=TRUE, isItalic=TRUE,
#'       name="Courier New", color="orange") +
#'     Fill(backgroundColor="lavender", foregroundColor="lavender",
#'       pattern="SOLID_FOREGROUND") +
#'     Alignment(h="ALIGN_RIGHT")
#'
#'   setCellStyle(cell.1, cellStyle1)
#'
#'   # you need to save the workbook now if you want to see this art
#' }
#'
#' @export
#' @name CellStyle-plus
"+.CellStyle" <- function(cs1, object)
{
  if (is.null(object)) return(cs1)

  cs <- if (is.CellStyle(object)) {
    dataFormat <- if (is.null(object$dataFormat)){cs1$dataFormat}
    alignment  <- if (is.null(object$alignment)){cs1$aligment}
    border     <- if (is.null(object$border)){cs1$border}
    fill       <- if (is.null(object$fill)){cs1$fill}
    font       <- if (is.null(object$font)){cs1$font}
    cellProtection <- if (is.null(object$cellProtection)){cs1$cellProtection}

    CellStyle.default(object$wb, dataFormat=dataFormat,
      alignment=alignment, border=border, fill=fill,
      font=font, cellProtection=cellProtection)

  } else if (is.DataFormat(object)) {
    CellStyle.default(cs1$wb, dataFormat=object,
      alignment=cs1$alignment, border=cs1$border, fill=cs1$fill,
      font=cs1$font, cellProtection=cs1$cellProtection)
  } else if (is.Alignment(object)) {
    CellStyle.default(cs1$wb, dataFormat=cs1$dataFormat,
      alignment=object, border=cs1$border, fill=cs1$fill,
      font=cs1$font, cellProtection=cs1$cellProtection)
  } else if (is.Border(object)) {
    CellStyle.default(cs1$wb, dataFormat=cs1$dataFormat,
      alignment=cs1$alignment, border=object, fill=cs1$fill,
      font=cs1$font, cellProtection=cs1$cellProtection)
  } else if (is.Fill(object)) {
    CellStyle.default(cs1$wb, dataFormat=cs1$dataFormat,
      alignment=cs1$alignment, border=cs1$border, fill=object,
      font=cs1$font, cellProtection=cs1$cellProtection)
  } else if (is.Font(object)) {
    CellStyle.default(cs1$wb, dataFormat=cs1$dataFormat,
      alignment=cs1$alignment, border=cs1$border, fill=cs1$fill,
      font=object, cellProtection=cs1$cellProtection)
  } else if (is.CellProtection(object)) {
    CellStyle.default(cs1$wb, dataFormat=cs1$dataFormat,
      alignment=cs1$alignment, border=cs1$border, fill=cs1$fill,
      font=cs1$font, cellProtection=object)
  } else {
     stop("Don't know how to add ", deparse(substitute(object)), " to a plot",
       call. = FALSE)
  }

  cs
}

## ######################################################################
## # Set the cell style for one cell.
## # Only one cell and one value.
## #
## setCellStyle <- function(x, cellStyle, ...)
##   UseMethod("setCellStyle", x)

######################################################################
# Set the cell style for one cell.
# Only one cell and one value.
#
#' @export
#' @rdname CellStyle
setCellStyle <- function(cell, cellStyle)
{
  .jcall(cell, "V", "setCellStyle", cellStyle$ref)
  invisible(NULL)
}

## # for a CellBlock  - NOT USED FOR NOW
## setCellStyle.CellBlock <- function(cellBlock, cellStyle, rowIndex=NULL,
##   colIndex=NULL)
## {
##   if (is.null(rowIndex) & is.null(colIndex)) {
##     cellBlock$setCellStyle( style )   # apply it to all the cells
##   } else {
##     cellBlock$setCellStyle( style, .jarray( as.integer( rowIndex-1L ) ),
##       .jarray( as.integer( colIndex-1L ) ) )
##   }

##   invisible()
## }


######################################################################
# Get the cell style for one cell.
# Only one cell and one value.
#
#' @export
#' @rdname CellStyle
getCellStyle <- function(cell)
{
  .jcall(cell,  "Lorg/apache/poi/ss/usermodel/CellStyle;", "getCellStyle")
}



