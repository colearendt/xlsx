######################################################################
# Deal with CellProtection
is.CellProtection <- function(x) inherits(x, "CellProtection")


######################################################################
# Create an CellProtection 
#
CellProtection <- function(locked=TRUE, hidden=FALSE)
{
  structure(list(locked=locked, hidden=hidden),
    class="CellProtection")
}
