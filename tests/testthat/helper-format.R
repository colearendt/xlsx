read_data_format <- function(cell) {
  cs <- getCellStyle(cell)
  
  .jcall(cs,'S','getDataFormatString')
}

read_fill_foreground <- function(cell) {
  cs <- getCellStyle(cell)
  
  col <- .jcall(cs,'Lorg/apache/poi/xssf/usermodel/XSSFColor;','getFillForegroundXSSFColor')
  
  if (!is.null(col)) {
    barr <- .jcall(col,'[B','getRgb')
    return(paste(as.character(barr),collapse=''))
  } else {
    return(NULL)
  }
}