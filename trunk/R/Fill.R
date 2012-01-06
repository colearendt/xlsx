######################################################################
# Deal with Fill properties for cells
is.Fill <- function(x) inherits(x, "Fill")


######################################################################
# Create a Fill.
#  - colors is an R color string.
#
Fill <- function(foregroundColor="lightblue", backgroundColor="lightblue",
  pattern="SOLID_FOREGROUND")
{
  if (!(pattern %in% names(FILL_STYLES_)))
    stop("Not a valid pattern value.  See help page.")
  
  structure(list(foregroundColor=foregroundColor,
    backgroundColor=backgroundColor, pattern=pattern),
    class="Fill")
}
