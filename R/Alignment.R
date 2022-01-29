######################################################################
# Deal with Alignment
#' @export
#' @rdname Alignment
is.Alignment <- function(x) inherits(x, "Alignment")


######################################################################
# Create an Alignment.
#
#' Create an Alignment object.
#'
#' Create an Alignment object, useful when working with cell styles.
#'
#'
#' @param horizontal a character value specifying the horizontal alignment.
#' Valid values come from constant \code{HALIGN_STYLES_}.
#' @param vertical a character value specifying the vertical alignment.  Valid
#' values come from constant \code{VALIGN_STYLES_}.
#' @param wrapText a logical indicating if the text should be wrapped.
#' @param rotation a numerical value indicating the degrees you want to rotate
#' the text in the cell.
#' @param indent a numerical value indicating the number of spaces you want to
#' indent the text in the cell.
#' @param x An Alignment object, as returned by \code{Alignment}.
#' @return \code{Alignment} returns a list with components from the input
#' argument, and a class attribute "Alignment".  Alignment objects are used
#' when constructing cell styles.
#'
#' \code{is.Alignment} returns \code{TRUE} if the argument is of class
#' "Alignment" and \code{FALSE} otherwise.
#' @author Adrian Dragulescu
#' @seealso \code{\link{CellStyle}} for using the a \code{Alignment} object.
#' @examples
#'
#'
#'   # you can just use h for horizontal, since R does the matching for you
#'   a1 <-  Alignment(h="ALIGN_CENTER", rotation=90) # centered and rotated!
#'
#' @export
Alignment <- function(horizontal=NULL, vertical=NULL, wrapText=FALSE,
  rotation=0, indent=0)
{
  if (
    !is.null(horizontal) &&
    !(horizontal %in% names(style_horizontal))
    # !.jinherits(horizontal, "org.apache.poi.ss.usermodel.HorizontalAlignment")
    ) {
    # TODO: figure out a way to allow java objects to be passed...
    if (horizontal %in% names(HALIGN_STYLES_)) {
      lifecycle::deprecate_soft("0.7.0", "HALIGN_STYLES_()", "style_horizontal()", "Try removing the 'ALIGN_' prefix")
    } else {
      stop("Not a valid horizontal value. See `style_horizontal` or `?POI_constants`")
    }
  }

  if (
    !is.null(vertical) &&
    !(vertical %in% names(style_vertical))
    # !.jinherits(vertical, "org.apache.poi.ss.usermodel.VerticalAlignment")
  ) {
    # TODO: figure out a way to allow java objects to be passed...
    if (vertical %in% names(VALIGN_STYLES_)) {
      lifecycle::deprecate_soft("0.7.0", "VALIGN_STYLES_()", "style_vertical()")
    } else {
      stop("Not a valid vertical value.  See `style_vertical` or `?POI_constants`")
    }
  }

  structure(list(horizontal=horizontal, vertical=vertical,
    wrapText=wrapText, rotation=rotation, indent=indent),
    class="Alignment")
}
