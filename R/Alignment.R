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
  if (!is.null(horizontal) && !(horizontal %in% names(HALIGN_STYLES_)))
    stop("Not a valid horizontal value.  See help page.")

  if (!is.null(vertical) && !(vertical %in% names(VALIGN_STYLES_)))
    stop("Not a valid vertical value.  See help page.")
  
  structure(list(horizontal=horizontal, vertical=vertical,
    wrapText=wrapText, rotation=rotation, indent=0),
    class="Alignment")  
}
