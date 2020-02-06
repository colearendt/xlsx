# functions to work with binary plots
#
#' Functions to manipulate images in a spreadsheet.
#'
#' For now, the following image types are supported: dib, emf, jpeg, pict, png,
#' wmf.  Please note, that scaling works correctly only for workbooks with the
#' default font size (Calibri 11pt for .xlsx). If the default font is changed
#' the resized image can be streched vertically or horizontally.
#'
#' Don't know how to remove an existing image yet.
#'
#' @param file the absolute path to the image file.
#' @param sheet a worksheet object as returned by \code{createSheet} or by
#' subsetting \code{getSheets}.  The picture will be added on this sheet at
#' position \code{startRow}, \code{startColumn}.
#' @param scale a numeric specifying the scale factor for the image.
#' @param startRow a numeric specifying the row of the upper left corner of the
#' image.
#' @param startColumn a numeric specifying the column of the upper left corner
#' of the image.
#' @return \code{addPicture} a java object references pointing to the picture.
#' @author Adrian Dragulescu
#' @examples
#'
#'
#' file <- system.file("tests", "log_plot.jpeg", package = "xlsx")
#'
#' wb <- createWorkbook()
#' sheet <- createSheet(wb, "Sheet1")
#'
#' addPicture(file, sheet)
#'
#' # don't forget to save the workbook!
#'
#'
#' @export
#' @name Picture
addPicture <- function(file, sheet, scale=1, startRow=1, startColumn=1)
{
  # get the picture extension
  pic_ext <- paste("PICTURE_TYPE_",
    toupper(gsub(".*\\.(.*)$", "\\1", file)), sep="")

  iStream <- .jnew("java/io/FileInputStream", file)
  ioutils <- .jnew("org/apache/poi/util/IOUtils")
  bytes <- .jcall(ioutils, "[B", "toByteArray",
    .jcast(iStream, "java/io/InputStream"))
  .jcall(iStream, "V", "close")

  wb <- .jcall(sheet,
    "Lorg/apache/poi/ss/usermodel/Workbook;", "getWorkbook")

  extId <- .jfield(wb, "I", pic_ext)
  picId <- .jcall(wb, "I", "addPicture", bytes, as.integer(extId))

  drawing <- .jcall(sheet,
    "Lorg/apache/poi/ss/usermodel/Drawing;", "createDrawingPatriarch")
  factory <- .jcall(wb,
    "Lorg/apache/poi/ss/usermodel/CreationHelper;", "getCreationHelper")
  anchor <- .jcall(factory,
    "Lorg/apache/poi/ss/usermodel/ClientAnchor;", "createClientAnchor")
  .jcall(anchor, "V", "setCol1", as.integer(startColumn-1))
  .jcall(anchor, "V", "setRow1", as.integer(startRow-1))

  # create the picture
  pict <- .jcall(drawing, "Lorg/apache/poi/ss/usermodel/Picture;",
    "createPicture", anchor, as.integer(picId))
  .jcall(pict, "V", "resize", scale)

  invisible(pict)
}
