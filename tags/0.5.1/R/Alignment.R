######################################################################
# Deal with Alignment
is.Alignment <- function(x) inherits(x, "Alignment")


######################################################################
# Create an Alignment. 
#
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
