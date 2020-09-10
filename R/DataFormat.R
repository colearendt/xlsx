######################################################################
# Deal with Format properties for cells, needed for CellStyle
#' @export
#' @rdname DataFormat
is.DataFormat <- function(df) inherits(df, "DataFormat")


######################################################################
# Create a DataFormat
# maybe allow to change the Locale
# .jfield("java/util/Locale", "Ljava/util/Locale;","FRENCH")
# https://docs.oracle.com/javase/7/docs/api/java/util/Locale.html
#' Create an DataFormat object.
#'
#' Create an DataFormat object, useful when working with cell styles.
#'
#' Specifying the \code{dataFormat} argument allows you to format the cell.
#' For example, "#,##0.00" corresponds to using a comma separator for powers of
#' 1000 with two decimal places, "m/d/yyyy" can be used to format dates and is
#' the equivalent of 's MM/DD/YYYY format.  To format datetimes use "m/d/yyyy
#' h:mm:ss;@".  To show negative values in red within parantheses with two
#' decimals and commas after power of 1000 use "#,##0.00_);[Red](#,##0.00)".  I
#' am not aware of an official way to discover these strings.  I find them out
#' by recording a macro that formats a specific cell and then checking out the
#' resulting VBA code.  From there you can read the \code{dataFormat} code.
#'
#' @param x a character value specifying the data format.
#' @param df An DataFormat object, as returned by \code{DataFormat}.
#' @return \code{DataFormat} returns a list one component dataFormat, and a
#' class attribute "DataFormat".  DataFormat objects are used when constructing
#' cell styles.
#'
#' \code{is.DataFormat} returns \code{TRUE} if the argument is of class
#' "DataFormat" and \code{FALSE} otherwise.
#' @author Adrian Dragulescu
#' @seealso \code{\link{CellStyle}} for using the a \code{DataFormat} object.
#' @examples
#'
#'   df <-  DataFormat("#,##0.00")
#'
#' @export
DataFormat <- function(x)
{
  structure(list(dataFormat=as.character(x)),
    class="DataFormat")
}
