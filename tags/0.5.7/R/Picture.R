# functions to work with binary plots
#


addPicture <- function(file, sheet, scale=1, startRow=1, startColumn=1)
{
  # get the picture extension
  pic_ext <- paste("PICTURE_TYPE_",
    toupper(gsub(".*\\.(.*)$", "\\1", file)), sep="")
  
  iStream <- .jnew("java/io/FileInputStream", file)
  ioutils <- .jnew("org/apache/poi/util/IOUtils")
  bytes <- .jcall(ioutils, "[B", "toByteArray",
    .jcast(iStream, "java/io/InputStream"))
  .jcall(iStream, "V", "close")

  wb <- .jcall(sheet,
    "Lorg/apache/poi/ss/usermodel/Workbook;", "getWorkbook")
  
  extId <- .jfield(wb, "I", pic_ext) 
  picId <- .jcall(wb, "I", "addPicture", bytes, as.integer(extId))
  
  drawing <- .jcall(sheet,
    "Lorg/apache/poi/ss/usermodel/Drawing;", "createDrawingPatriarch")   
  factory <- .jcall(wb,
    "Lorg/apache/poi/ss/usermodel/CreationHelper;", "getCreationHelper")
  anchor <- .jcall(factory,
    "Lorg/apache/poi/ss/usermodel/ClientAnchor;", "createClientAnchor")
  .jcall(anchor, "V", "setCol1", as.integer(startColumn-1))
  .jcall(anchor, "V", "setRow1", as.integer(startRow-1))

  # create the picture
  pict <- .jcall(drawing, "Lorg/apache/poi/ss/usermodel/Picture;",
    "createPicture", anchor, as.integer(picId))
  .jcall(pict, "V", "resize", scale)

  invisible(pict)
}
