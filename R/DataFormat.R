######################################################################
# Deal with Format properties for cells, needed for CellStyle
is.DataFormat <- function(x) inherits(x, "DataFormat")


######################################################################
# Create a DataFormat
#
DataFormat <- function(x)
{
  structure(list(dataFormat=as.character(x)), 
    class="DataFormat")
}
