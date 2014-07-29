######################################################################
# Create a cell style.  It needs a workbook object!
#

CellStyle <- function(wb, dataFormat=NULL, alignment=NULL,
  border=NULL, fill=NULL, font=NULL, cellProtection=NULL) UseMethod("CellStyle")

is.CellStyle <- function(x) {inherits(x, "CellStyle")}


######################################################################
# Create a cell style.  It needs a workbook object!
#
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
getCellStyle <- function(cell)
{ 
  .jcall(cell,  "Lorg/apache/poi/ss/usermodel/CellStyle;", "getCellStyle")
}



