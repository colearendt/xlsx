######################################################################
# Deal with Format properties for cells, needed for CellStyle
is.DataFormat <- function(df) inherits(df, "DataFormat")


######################################################################
# Create a DataFormat
# maybe allow to change the Locale
# .jfield("java/util/Locale", "Ljava/util/Locale;","FRENCH")
# http://docs.oracle.com/javase/7/docs/api/java/util/Locale.html
DataFormat <- function(x)
{
  structure(list(dataFormat=as.character(x)),
    class="DataFormat")
}
