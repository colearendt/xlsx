# Write a data.frame to a new xlsx file.
# with a java back-end
#

#' @export
#' @rdname write.xlsx
write.xlsx2 <- function(x, file, sheetName="Sheet1",
  col.names=TRUE, row.names=TRUE, append=FALSE,
  password=NULL, ...)
{
  if (append && file.exists(file)){
    wb <- loadWorkbook(file, password=password)
  } else {
    ext <- gsub(".*\\.(.*)$", "\\1", basename(file))
    wb  <- createWorkbook(type=ext)
  }
  sheet <- createSheet(wb, sheetName)

  addDataFrame(x, sheet, col.names=col.names, row.names=row.names,
    startRow=1, startColumn=1, colStyle=NULL, colnamesStyle=NULL,
    rownamesStyle=NULL)

  saveWorkbook(wb, file, password=password)

  invisible()
}


