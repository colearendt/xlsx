#
#' Functions to manipulate rows of a worksheet.
#'
#' \code{removeRow} is just a convenience wrapper to remove the rows from the
#' sheet (before saving).  Internally it calls \code{lapply}.
#'
#' @param sheet a worksheet object as returned by \code{createSheet} or by
#' subsetting \code{getSheets}.
#' @param rowIndex a numeric vector specifying the index of rows to create.
#' For \code{getRows}, a \code{NULL} value will return all non empty rows.
#' @param rows a list of \code{Row} objects.
#' @param inPoints a numeric value to specify the height of the row in points.
#' @param multiplier a numeric value to specify the multiple of default row
#' height in points.  If this value is set, it takes precedence over the
#' \code{inPoints} argument.
#' @return For \code{getRows} a list of java object references each pointing to
#' a row.  The list is named with the row number.
#' @author Adrian Dragulescu
#' @seealso To extract the cells from a given row, see \code{\link{Cell}}.
#' @examples
#'
#'
#' file <- system.file("tests", "test_import.xlsx", package = "xlsx")
#'
#' wb <- loadWorkbook(file)
#' sheets <- getSheets(wb)
#'
#' sheet <- sheets[[2]]
#' rows  <- getRows(sheet)  # get all the rows
#'
#' # see all the available java methods that you can call
#' rJava::.jmethods(rows[[1]])
#'
#' # for example
#' rows[[1]]$getRowNum()   # zero based index in Java
#'
#' removeRow(sheet, rows)  # remove them all
#'
#' # create some row
#' rows  <- createRow(sheet, rowIndex=1:5)
#' setRowHeight( rows, multiplier=3)  # 3 times bigger rows than the default
#'
#'
#'
#' @name Row
NULL


######################################################################
# Return a list of row objects.
#
#' @export
#' @rdname Row
createRow <- function(sheet, rowIndex=1:5)
{
  rows <- vector("list", length(rowIndex))
  names(rows) <- rowIndex
  for (ir in seq_along(rowIndex))
    rows[[ir]] <- .jcall(sheet, "Lorg/apache/poi/ss/usermodel/Row;",
      "createRow", as.integer(rowIndex[ir]-1)) # in java the index starts from 0!

  rows
}


######################################################################
# Return a list of all row objects in a sheet.
# You can specify which rows you want to return by provinding a vector of
# rowIndices with rowInd.
#
#' @export
#' @rdname Row
getRows <- function(sheet, rowIndex=NULL)
{
  noRows <- sheet$getLastRowNum()+1

  keep <- as.integer(seq_len(noRows)-1)
  if (!is.null(rowIndex))
    keep <- keep[rowIndex]

  rows <- vector("list", length(keep))
  namesRow <- rep(NA, length(keep))

  for (ir in seq_along(keep)){
    aux <- .jcall(sheet, "Lorg/apache/poi/ss/usermodel/Row;",
      "getRow", keep[ir])
    if (!is.null(aux)){
      rows[[ir]] <- aux
      namesRow[ir] <- .jcall(aux, "I", "getRowNum") + 1
    }
  }
  names(rows) <- namesRow   # need this if rows are ragged

  rows <- rows[!is.na(namesRow)]   # skip the empty rows

  rows
}


######################################################################
# remove rows
#
#' @export
#' @rdname Row
removeRow <- function(sheet, rows=NULL)
{
  if (is.null(rows))
    rows <- getRows(sheet)   # delete them all

  lapply(rows, function(x){.jcall(sheet, "V", "removeRow", x)})

  invisible()
}


######################################################################
# set the Row height
#
#' @export
#' @rdname Row
setRowHeight <- function(rows, inPoints, multiplier=NULL)
{
  if ( !is.null(multiplier) ) {
    sheet <- rows[[1]]$getSheet()   # get the sheet
    inPoints <- multiplier * sheet$getDefaultRowHeightInPoints()
  }

  lapply(rows, function(row) {
    row$setHeightInPoints( .jfloat(inPoints) )
  })

  invisible()
}
















## setClass("Workbook",
##   representation(
##     INITIAL_CAPACITY = "integer",
##     sheetName       = "character",    # a vector
##     ref             = "jobjRef"),     # the java reference
##   prototype = prototype(
##     INITIAL_CAPACITY = as.integer(3),
##     ref = .jnull()
##     )
## )


## setMethod("initialize", "Workbook",
##   function(.Object, sheetName, ...)
##     callNextMethod(.Object, sheetName, ...))

## setMethod("show", "WorksheetEntry",
##   function(object){
##     callNextMethod(object)
##     cat("\n@nrow =", object@nrow,
##         "\n@ncol =", object@ncol)
##   })



## ##########################################################################
## #
## #setMethod("getWorksheets", "SpreadsheetEntry",
## #  function(xls, ...) .getWorksheetEntries(xls, ...))

## getWorksheets <- function(xls, ...)
## {
##   id <- .getUniqueId(xls@id)[[1]]
##   worksheetId <- paste("https://spreadsheets.google.com/feeds/spreadsheets/",
##                        id, sep="")

##   msg <- xls@con@ref$getWorksheetEntries(worksheetId)

##   msg <- strsplit(msg, "\n")[[1]]      # split by worksheets
##   if (length(msg)==1){
##     cat("No worksheets found in the spreadheet.\n")
##     return
##   }
##   msg <- strsplit(msg, "\t")           # split by fields
##   slotNames <- msg[[1]]
##   msg <- msg[-1]
##   msg <- lapply(msg, as.list)
##   msg <- lapply(msg, function(x, slotNames){names(x) <- slotNames; x},
##                 slotNames)

##   msg <- lapply(msg,
##     function(x, slotNames){
##       x$canEdit <- as.logical(toupper(x$canEdit))
##       x$nrow  <- as.integer(x$nrow)
##       x$ncol  <- as.integer(x$ncol)
##       x$published <- as.POSIXct(x$published, "%Y-%m-%dT%H:%M:%OSZ", tz="")
##       x
##     }, slotNames)

##   wks <- vector("list", length(msg))
##   for (w in 1:length(wks)){
##     listFields <- msg[[w]]
##     listFields$con <- xls@con
##     wks[[w]] <- new("WorksheetEntry", listFields)  # call the constructor
##   }

##   wks
## }


## ##########################################################################
## # convert to data.frame
## setAs("WorksheetEntry", "data.frame",
##   function(from){
##     res <- data.frame(
##       title = from@title,
##       nrow = from@nrow,
##       ncol = from@ncol,
##       authors = from@authors,
##       canEdit = from@canEdit,
## #      id = from@id,
## #      etag = from@etag,
##       published = from@published,
## #      content = from@content,
## #      con = from@con,
##       stringsAsFactors = FALSE)

##     res
##   })


## ##########################################################################
## #
## #setMethod("getListEntries", "WorksheetEntry",
## #  function(wks, ...) .getListEntries(wks, ...))

## getListEntries <- function(wks, ...)
## {
##   id <- gsub("worksheets", "list", wks@id)
##   worksheetId <- paste(id, "/private/full", sep="")

##   msg <- wks@con@ref$getListEntries(worksheetId)

##   if (wks@con@ref$getMsg() != "")
##     stop(wks@con@ref$getMsg())

##   msg <- strsplit(msg, "\n")[[1]]      # split by list entries
##   if (length(msg)==0){
##     cat("No list entries found in the worksheet.\n")
##     return
##   }
##   msg <- strsplit(msg, "\t")           # split by fields
##   slotNames <- msg[[1]]
##   msg <- msg[-1]
##   msg <- lapply(msg, as.list)
##   msg <- lapply(msg, function(x, slotNames){names(x) <- slotNames; x},
##                 slotNames)

##   msg <- lapply(msg,
##     function(x, slotNames){
##       x$canEdit <- as.logical(toupper(x$canEdit))
##       x$published <- as.POSIXct(x$published, "%Y-%m-%dT%H:%M:%OSZ", tz="")
##       x
##     }, slotNames)

##   entries <- vector("list", length(msg))
##   for (e in 1:length(entries)){
##     listFields       <- msg[[e]]
##     listFields$rowId <- sub(".*/(.*)$", "\\1", listFields$id)
##     listFields$con   <- wks@con
##     entries[[e]] <- new("ListEntry", listFields)  # call the constructor
##   }
##   names(entries) <- sapply(entries, slot, "rowId")

##   entries
## }


## ##########################################################################
## # getContent for a ListEntry
## # Returns a data.frame with one row
## #
## setMethod("getContent", "ListEntry",
##   function(obj, ...){

##     id  <- obj@id
##     ind <- gregexpr("/", id)[[1]]
##     id  <- paste(substr(id, 1, ind[length(ind)]), "private/full/",
##                 substr(id, ind[length(ind)]+1, nchar(id)), sep="")

##     msg <- obj@con@ref$getContentListEntry(id)
##     if (is.null(msg)) return(NULL)

##     msg <- strsplit(msg, "\t")[[1]]           # split by fields
##     N <- length(msg)
##     values <- matrix(msg[seq.int(2,N,2)], ncol=N/2, byrow=TRUE)
##     colnames(values) <- msg[seq.int(1,N,2)]
##     rownames(values) <- obj@rowId

##     values
##   }
## )




## ##########################################################################
## # add ListEntries to a worksheetEntry (wks)
## #
## addListEntry <- function(wks, contentList)
## {
##   id  <- wks@id
##   ind <- gregexpr("/", id)[[1]]
##   id  <- paste(substr(id, 1, ind[length(ind)]), "private/full/",
##                substr(id, ind[length(ind)]+1, nchar(id)), sep="")

##   # make content string. tag,values separated by \t, listEntries
##   # separated by \t.
##   content <- sapply(contentList, function(x){
##     paste(names(x), x, sep="\t", collapse="\t")})
##   content <- paste(content, sep="", collapse="\n")

##   wks@con@ref$addListEntry(id, content)

##   invisible(as.logical(wks@con@ref$getMsg()))
## }


## ##########################################################################
## # add ListEntries to a worksheetEntry (wks)
## #
## updateListEntry <- function(listEntries, contentList)
## {
##   if (length(contentList)==1 & !is.list(contentList[1]))
##     contentList <- list(contentList) # give the user a break

##   if (length(listEntries) != length(contentList))
##     stop("Non-equal length for the two arguments!")

##   # construct the correct list entries id
##   id  <- sapply(listEntries, slot, "id")
##   ind <- gregexpr("/", id)
##   id  <- mapply(function(id, ind){
##     paste(substr(id, 1, ind[length(ind)]), "private/full/",
##           substr(id, ind[length(ind)]+1, nchar(id)), sep="")}, id, ind)
##   ids <- paste(id, sep="", collapse="\n")

##   # make content string. tag,values separated by \t, listEntries
##   # separated by \t.
##   content <- sapply(contentList, function(x){
##     paste(names(x), x, sep="\t", collapse="\t")})
##   content <- paste(content, sep="", collapse="\n")

##   listEntries[[1]]@con@ref$updateListEntry(ids, content)

##   invisible(as.logical(listEntries[[1]]@con@ref$getMsg()))
## }


## ##########################################################################
## # delete ListEntries to a worksheetEntry (wks)
## #
## deleteListEntry <- function(listEntries)
## {

##   # construct the correct list entries id
##   id  <- sapply(listEntries, slot, "id")
##   ind <- gregexpr("/", id)
##   id  <- mapply(function(id, ind){
##     paste(substr(id, 1, ind[length(ind)]), "private/full/",
##           substr(id, ind[length(ind)]+1, nchar(id)), sep="")}, id, ind)
##   ids <- paste(id, sep="", collapse="\n")

##   listEntries[[1]]@con@ref$deleteListEntry(ids)

##   invisible(as.logical(listEntries[[1]]@con@ref$getMsg()))
## }


## ##########################################################################
## # create an empty worksheet.  Call it new instead of add for consistency
## # with other objects.
## addWorksheet <- function(xls, title, nrow=100, ncol=20)
## {
##   key <- gsub("spreadsheet%3A(.*)", "\\1", xls@key)
##   id <- paste("https://spreadsheets.google.com/feeds/spreadsheets/", key,
##               sep="")
##   xls@con@ref$addWorksheet(title, as.integer(nrow), as.integer(ncol), id)

##   invisible(as.logical(xls@con@ref$getMsg()))
## }

## ##########################################################################
## # delete an existing worksheet
## deleteWorksheet <- function(wks)
## {
##   id  <- wks@id
##   ind <- gregexpr("/", id)[[1]]
##   id  <- paste(substr(id, 1, ind[length(ind)]), "private/full/",
##                substr(id, ind[length(ind)]+1, nchar(id)), sep="")
##   wks@con@ref$deleteWorksheet(id)

##   invisible(as.logical(wks@con@ref$getMsg()))
## }

















