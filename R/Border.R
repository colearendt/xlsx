######################################################################
# Deal with Borders
#' @export
#' @rdname Border
is.Border <- function(x) inherits(x, "Border")


######################################################################
# Create a Border.  It needs a workbook object!
#  - color is an R color string.
#
#' Create an Border object.
#' 
#' Create an Border object, useful when working with cell styles.
#' 
#' The values for the color, position, or pen arguments are replicated to the
#' longest of them.
#' 
#' @param color a character vector specifiying the font color.  Any color names
#' as returned by \code{\link[grDevices]{colors}} can be used.  Or, a hex
#' character, e.g. "#FF0000" for red.  For Excel 95 workbooks, only a subset of
#' colors is available, see the constant \code{INDEXED_COLORS_}.
#' @param position a character vector specifying the border position.  Valid
#' values are "BOTTOM", "LEFT", "TOP", "RIGHT".
#' @param pen a character vector specifying the pen style.  Valid values come
#' from constant \code{BORDER_STYLES_}.
#' @param x An Border object, as returned by \code{Border}.
#' @return \code{Border} returns a list with components from the input
#' argument, and a class attribute "Border".  Border objects are used when
#' constructing cell styles.
#' 
#' \code{is.Border} returns \code{TRUE} if the argument is of class "Border"
#' and \code{FALSE} otherwise.
#' @author Adrian Dragulescu
#' @seealso \code{\link{CellStyle}} for using the a \code{Border} object.
#' @examples
#' 
#' 
#'   border <-  Border(color="red", position=c("TOP", "BOTTOM"),
#'     pen=c("BORDER_THIN", "BORDER_THICK"))
#' 
#' @export
Border <- function(color="black", position="BOTTOM", pen="BORDER_THIN")
{
  if (any(!(position %in% c("BOTTOM", "LEFT", "TOP", "RIGHT"))))
    stop("Not a valid postion value.  See help page.")

  if (any(!(pen %in% names(BORDER_STYLES_))))
    stop("Not a valid pen value.  See help page.")

  len <- c(length(position), length(color), length(pen))
  if (length(color) < max(len))
    color <- rep(color, length.out=max(len))
  if (length(position) < max(len))
    position <- rep(position, length.out=max(len))
  if (length(pen) < max(len))
    pen <- rep(pen, length.out=max(len))
  
  structure(list(color=color, position=position, pen=pen),
            class="Border")
}
