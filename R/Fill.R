######################################################################
# Deal with Fill properties for cells
#' @export
#' @rdname Fill
is.Fill <- function(x) inherits(x, "Fill")


######################################################################
# Create a Fill.
#  - colors is an R color string.
#
#' Create an Fill object.
#'
#' Create an Fill object, useful when working with cell styles.
#'
#' @param foregroundColor a character vector specifiying the foreground color.
#' Any color names as returned by \code{\link[grDevices]{colors}} can be used.
#' Or, a hex character, e.g. "#FF0000" for red.  For Excel 95 workbooks, only a
#' subset of colors is available, see the constant \code{INDEXED_COLORS_}.
#' @param backgroundColor a character vector specifiying the foreground color.
#' Any color names as returned by \code{\link[grDevices]{colors}} can be used.
#' Or, a hex character, e.g. "#FF0000" for red.  For Excel 95 workbooks, only a
#' subset of colors is available, see the constant \code{INDEXED_COLORS_}.
#' @param pattern a character vector specifying the fill pattern style.  Valid
#' values come from constant \code{FILL_STYLES_}.
#' @param x a Fill object, as returned by \code{Fill}.
#' @return \code{Fill} returns a list with components from the input argument,
#' and a class attribute "Fill".  Fill objects are used when constructing cell
#' styles.
#'
#' \code{is.Fill} returns \code{TRUE} if the argument is of class "Fill" and
#' \code{FALSE} otherwise.
#' @author Adrian Dragulescu
#' @seealso \code{\link{CellStyle}} for using the a \code{Fill} object.
#' @examples
#'
#'   fill <-  Fill()
#'
#' @export
Fill <- function(foregroundColor="lightblue", backgroundColor="lightblue",
  pattern="SOLID_FOREGROUND")
{
  if (!(pattern %in% names(FILL_STYLES_)))
    stop("Not a valid pattern value.  See help page.")

  structure(list(foregroundColor=foregroundColor,
    backgroundColor=backgroundColor, pattern=pattern),
    class="Fill")
}
