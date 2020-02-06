######################################################################
# Deal with CellProtection
#' @rdname CellProtection
#' @export
is.CellProtection <- function(x) inherits(x, "CellProtection")


######################################################################
# Create an CellProtection
#
#' Create a CellProtection object.
#'
#' Create a CellProtection object used for cell styles.
#'
#' @param locked a logical indicating the cell is locked.
#' @param hidden a logical indicating the cell is hidden.
#' @param x A CellProtection object, as returned by \code{CellProtection}.
#' @return
#'
#' \code{CellProtection} returns a list with components from the input
#' argument, and a class attribute "CellProtection".  CellProtection objects
#' are used when constructing cell styles.
#'
#' \code{is.CellProtection} returns \code{TRUE} if the argument is of class
#' "CellProtection" and \code{FALSE} otherwise.
#' @author Adrian Dragulescu
#' @seealso \code{\link{CellStyle}} for using the a \code{CellProtection}
#' object.
#' @examples
#'
#'
#'   font <-  CellProtection(locked=TRUE)
#'
#' @export
CellProtection <- function(locked=TRUE, hidden=FALSE)
{
  structure(list(locked=locked, hidden=hidden),
    class="CellProtection")
}
