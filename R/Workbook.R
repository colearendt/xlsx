# createWorkbook
# loadWorkbook
# saveWorkbook
# getSheets
# createSheet
# removeSheet


######################################################################
# 
createWorkbook <- function(type="xlsx")
{
  if (type=="xls") {
    wb <- .jnew("org/apache/poi/hssf/usermodel/HSSFWorkbook")
  } else if (type == "xlsx") {
    wb <- .jnew("org/apache/poi/xssf/usermodel/XSSFWorkbook")
  } else {
    stop(paste("Unknown format", type)) 
  }

  return(wb)
}

######################################################################
# 
loadWorkbook <- function(file)
{
  if (!file.exists(path.expand(file)))
    stop("Cannot find ", path.expand(file))
    
  inputStream <- .jnew("java/io/FileInputStream", path.expand(file))
  
  wbFactory <- .jnew("org/apache/poi/ss/usermodel/WorkbookFactory")
  
  wb <- wbFactory$create(inputStream)
    
  return(wb)
}

######################################################################
# 
saveWorkbook <- function(wb, file)
{
  jFile <- .jnew("java/io/File", file)
  fh <- .jnew("java/io/FileOutputStream", jFile)

  # write the workbook to the file
  wb$write(fh)
  
  # close the filehandle
  .jcall(fh, "V", "close")
}
## # make a file handle and write the xlsx to file
## filename <- "C:/Temp/junk.xlsx"
## fh <- .jnew("java/io/FileOutputStream", filename)
## .jcall(wb, "V", "write", .jcast(fh, "java/io/OutputStream"))


######################################################################
# return the sheets
#
getSheets <- function(wb)
{
  noSheets <- wb$getNumberOfSheets()
  if (noSheets==0){
    cat("Workbook has no sheets!\n")
    return()
  }
  
  res <- vector("list", length=noSheets)
  for (sh in 1:noSheets){
    names(res)[sh] <- wb$getSheetName(as.integer(sh-1))
    res[[sh]] <- wb$getSheetAt(as.integer(sh-1))
  }

  res
}

######################################################################
# 
createSheet <- function(wb, sheetName="Sheet1")
{
  sheet <- .jcall(wb, "Lorg/apache/poi/ss/usermodel/Sheet;",
    "createSheet", sheetName)
  
  return(sheet)
}

######################################################################
# 
removeSheet <- function(wb, sheetName="Sheet1")
{
  sheetInd <- wb$getSheetIndex(sheetName)
  
  .jcall(wb, "V", "removeSheetAt", as.integer(sheetInd))
  
  return(invisible())
}




















