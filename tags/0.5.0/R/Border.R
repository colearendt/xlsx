######################################################################
# Deal with Borders
is.Border <- function(x) inherits(x, "Border")


######################################################################
# Create a Border.  It needs a workbook object!
#  - color is an R color string.
#
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
