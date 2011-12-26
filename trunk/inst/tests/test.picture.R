#
#
#

test.picture <- function(outdir="C:/Temp/", type="xlsx")
{
  cat("##################################################\n")
  cat("Test embedding an R picture into an xlsx document\n")
  cat("##################################################\n")

  cat("Add log_plot.jpeg to a new xlsx...")
  picname <- system.file("tests", "log_plot.jpeg", package="xlsx")
  wb <- createWorkbook(type=type)
  sheet <- createSheet(wb, "Sheet1")

  addPicture(picname, sheet)

  xlsx:::.write_block(wb, sheet, iris)
  
  file <- paste(outdir, "test_picture.", type, sep="")
  saveWorkbook(wb, file)
  cat("Wrote file:", file, "\n")
        
}

##  pic_ext <- paste("PICTURE_TYPE_",
##     toupper(gsub(".*\\.(.*)$", "\\1", picname)), sep="")
  
  
##   iStream <- .jnew("java/io/FileInputStream", picname)
##   ioutils <- .jnew("org/apache/poi/util/IOUtils")
##   #jbytes <- .jcall(ioutils, "[B", "toByteArray", iStream) # Error: [B not found
##   #bytes <- ioutils$toByteArray(iStream)
##   bytes <- .jcall(ioutils, "[B", "toByteArray",
##     .jcast(iStream, "java/io/InputStream"))
##   iStream$close()

##   #jbytes <- .jnew("[B", bytes)  # no Such
##   #jbytes <- .jarray(bytes, "[B")
  
##   wb <- createWorkbook()
##   sheet <- createSheet(wb, "Sheet1")
##   extId <- .jfield(wb, "I", pic_ext) 
##   picId <- .jcall(wb, "I", "addPicture", bytes, as.integer(extId))
  
##   drawing <- .jcall(sheet,
##     "Lorg/apache/poi/xssf/usermodel/XSSFDrawing;", "createDrawingPatriarch")
   
##   factory <- .jcall(wb,
##     "Lorg/apache/poi/xssf/usermodel/XSSFCreationHelper;", "getCreationHelper")
##   anchor <- .jcall(factory,
##     "Lorg/apache/poi/ss/usermodel/ClientAnchor;", "createClientAnchor")
##   .jcall(anchor, "V", "setCol1", as.integer(1))
##   .jcall(anchor, "V", "setRow1", as.integer(1))
##   pict <- .jcall(drawing, "Lorg/apache/poi/xssf/usermodel/XSSFPicture;",
##     "createPicture", anchor, as.integer(picId))

##   scale <- 1
##   .jcall(pict, "V", "resize", scale)









