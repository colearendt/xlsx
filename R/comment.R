######################################################################
# create ONE comment

#' Functions to manipulate cell comments.
#'
#' These functions are not vectorized.
#'
#' @param cell a \code{Cell} object.
#' @param string a string for the comment.
#' @param author a string with the author's name
#' @param visible a logical value.  If \code{TRUE} the comment will be visible.
#' @return
#'
#' \code{createCellComment} creates a \code{Comment} object.
#'
#' \code{getCellComment} returns a the \code{Comment} object if it exists.
#'
#' \code{removeCellComment} removes a comment from the given cell.
#' @author Adrian Dragulescu
#' @seealso For cells, see \code{\link{Cell}}.  To format cells, see
#' \code{\link{CellStyle}}.
#' @examples
#'
#'
#'   wb <- createWorkbook()
#'   sheet1 <- createSheet(wb, "Sheet1")
#'   rows   <- createRow(sheet1, rowIndex=1:10)     # 10 rows
#'   cells  <- createCell(rows, colIndex=1:8)       # 8 columns
#'
#'   cell1 <- cells[[1,1]]
#'   setCellValue(cell1, 1)   # add value 1 to cell A1
#'
#'   # create a cell comment
#'   createCellComment(cell1, "Cogito", author="Descartes")
#'
#'
#'   # extract the comments
#'   comment <- getCellComment(cell1)
#'   stopifnot(comment$getAuthor()=="Descartes")
#'   stopifnot(comment$getString()$toString()=="Cogito")
#'
#'   # don't forget to save your workbook!
#'
#' @export
#' @name Comment
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

#' @export
#' @rdname Comment
removeCellComment <- function(cell){
  .jcall(cell, "V", "removeCellComment")
}

## setCellComment <- function(cell, comment){
##   .jcall(cell, "V", "setCellComment", comment)
## }

#' @export
#' @rdname Comment
getCellComment <- function(cell)
{
  comment <- .jcall(cell, "Lorg/apache/poi/ss/usermodel/Comment;",
    "getCellComment")

  comment
}
