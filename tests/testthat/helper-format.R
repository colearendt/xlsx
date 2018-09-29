read_data_format <- function(cell) {
  cs <- getCellStyle(cell)
  
  rJava::.jcall(cs,'S','getDataFormatString')
}

read_fill_foreground <- function(cell) {
  cs <- getCellStyle(cell)
  
  col <- rJava::.jcall(cs,'Lorg/apache/poi/xssf/usermodel/XSSFColor;','getFillForegroundXSSFColor')
  
  if (!is.null(col)) {
    barr <- rJava::.jcall(col,'[B','getRgb')
    return(paste(as.character(barr),collapse=''))
  } else {
    return(NULL)
  }
}
