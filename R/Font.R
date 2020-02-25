######################################################################
# Deal with Fonts
#' @export
#' @rdname Font
is.Font <- function(x) inherits(x, "Font")


######################################################################
# Create a Font.  It needs a workbook object!
#  - color is an R color string.
#
#' Create a Font object.
#'
#' Create a Font object.
#'
#' Default values for \code{NULL} parameters are taken from Excel.  So the
#' default font color is black, the default font name is "Calibri", and the
#' font height in points is 11.
#'
#' For Excel 95/2000/XP/2003, it is impossible to set the font to bold.  This
#' limitation may be removed in the future.
#'
#' NOTE: You need to have a \code{Workbook} object to attach a \code{Font}
#' object to it.
#'
#' @aliases Font is.Font
#' @param wb a workbook object as returned by \code{\link{createWorkbook}} or
#' \code{\link{loadWorkbook}}.
#' @param color a character specifiying the font color.  Any color names as
#' returned by \code{\link[grDevices]{colors}} can be used.  Or, a hex
#' character, e.g. "#FF0000" for red.  For Excel 95 workbooks, only a subset of
#' colors is available, see the constant \code{INDEXED_COLORS_}.
#' @param heightInPoints a numeric value specifying the font height.  Usual
#' values are 10, 12, 14, etc.
#' @param name a character value for the font to use.  All values that you see
#' in Excel should be available, e.g. "Courier New".
#' @param isItalic a logical indicating the font should be italic.
#' @param isStrikeout a logical indicating the font should be stiked out.
#' @param isBold a logical indicating the font should be bold.
#' @param underline a numeric value specifying the thickness of the underline.
#' Allowed values are 0, 1, 2.
#' @param boldweight a numeric value indicating bold weight.  Normal is 400,
#' regular bold is 700.
#' @param x A Font object, as returned by \code{Font}.
#' @return \code{Font} returns a list with a java reference to a \code{Font}
#' object, and a class attribute "Font".
#'
#' \code{is.Font} returns \code{TRUE} if the argument is of class "Font" and
#' \code{FALSE} otherwise.
#' @author Adrian Dragulescu
#' @seealso \code{\link{CellStyle}} for using the a \code{Font} object.
#' @examples
#'
#' \dontrun{
#'   font <-  Font(wb, color="blue", isItalic=TRUE)
#' }
#'
#' @export
Font <- function(wb, color=NULL, heightInPoints=NULL, name=NULL,
  isItalic=FALSE, isStrikeout=FALSE, isBold=FALSE, underline=NULL,
  boldweight=NULL) # , setFamily=NULL
{
  font <- .jcall(wb, "Lorg/apache/poi/ss/usermodel/Font;",
    "createFont")

  if (!is.null(color))
    if (grepl("XSSF", wb$getClass()$getName())){
      .jcall(font, "V", "setColor", .xssfcolor(color))
    } else {
      .jcall(font, "V", "setColor",
        .jshort(INDEXED_COLORS_[toupper(color)]))
    }

  if (!is.null(heightInPoints))
    .jcall(font, "V", "setFontHeightInPoints", .jshort(heightInPoints))

  if (!is.null(name))
    .jcall(font, "V", "setFontName", name)

  if (isItalic)
    .jcall(font, "V", "setItalic", TRUE)

  if (isStrikeout)
    .jcall(font, "V", "setStrikeout", TRUE)

  if (isBold & grepl("XSSF", wb$getClass()$getName()))
    .jcall(font, "V", "setBold", TRUE)

  if (!is.null(underline))
    .jcall(font, "V", "setUnderline", .jbyte(underline))

  if (!is.null(boldweight))
    .jcall(font, "V", "setBoldweight", .jshort(boldweight))

  structure(list(ref=font), class="Font")
}
