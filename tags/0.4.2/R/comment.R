######################################################################
# create ONE comment

createCellComment <- function(cell, string,  
  author=NULL, visible=TRUE)
{
  sheet <- .jcall(cell,
    "Lorg/apache/poi/ss/usermodel/Sheet;", "getSheet")
  wb <- .jcall(sheet,
    "Lorg/apache/poi/ss/usermodel/Workbook;", "getWorkbook")

  factory <- .jcall(wb,
    "Lorg/apache/poi/ss/usermodel/CreationHelper;", "getCreationHelper")

  anchor <- .jcall(factory,
    "Lorg/apache/poi/ss/usermodel/ClientAnchor;", "createClientAnchor")
  
  drawing <- .jcall(sheet,
    "Lorg/apache/poi/ss/usermodel/Drawing;", "createDrawingPatriarch")
  
  comment <- .jcall(drawing,
    "Lorg/apache/poi/ss/usermodel/Comment;", "createCellComment", anchor)

  rtstring <- factory$createRichTextString(string)  # rich text string
 
  comment$setString(rtstring)
  if (!is.null(author))
    .jcall(comment, "V", "setAuthor", author)
  if (visible)
    .jcall(cell, "V", "setCellComment", comment)

  invisible(comment)
}


removeCellComment <- function(cell){
  .jcall(cell, "V", "removeCellComment")
}

## setCellComment <- function(cell, comment){
##   .jcall(cell, "V", "setCellComment", comment)
## }

getCellComment <- function(cell)
{
  comment <- .jcall(cell, "Lorg/apache/poi/ss/usermodel/Comment;",
    "getCellComment")

  comment
}



