# createWorkbook
# loadWorkbook
# saveWorkbook
# getSheets
# createSheet
# removeSheet

#' Functions to manipulate Excel 2007 workbooks.
#'
#' \code{createWorkbook} creates an empty workbook object.
#'
#' \code{loadWorkbook} loads a workbook from a file.
#'
#' \code{saveWorkbook} saves an existing workook to an Excel 2007 file.
#'
#' Reading or writing of password protected workbooks is supported for Excel
#' 2007 OOXML format only.  Note that in Linux, LibreOffice is not able to read
#' password protected spreadsheets.
#'
#' @param type a String, either \code{xlsx} for Excel 2007 OOXML format, or
#' \code{xls} for Excel 95 binary format.
#' @param file the path to the file you intend to read or write.  Can be an xls
#' or xlsx format.
#' @param wb a workbook object as returned by \code{createWorkbook} or
#' \code{loadWorkbook}.
#' @param password a String with the password.
#' @return \code{createWorkbook} returns a java object reference pointing to an
#' empty workbook object.
#'
#' \code{loadWorkbook} creates a java object reference corresponding to the
#' file to load.
#' @author Adrian Dragulescu
#' @seealso \code{\link{write.xlsx}} for writing a \code{data.frame} to an
#' \code{xlsx} file.  \code{\link{read.xlsx}} for reading the content of a
#' \code{xlsx} worksheet into a \code{data.frame}.  To extract worksheets and
#' manipulate them, see \code{\link{Worksheet}}.
#' @examples
#'
#'
#' wb <- createWorkbook()
#'
#' # see all the available java methods that you can call
#' rJava::.jmethods(wb)
#'
#' # for example
#' wb$getNumberOfSheets()   # no sheet yet!
#'
#' # loadWorkbook("C:/Temp/myFile.xls")
#'
#' @name Workbook
NULL

######################################################################
#
#' @export
#' @rdname Workbook
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
#' @export
#' @rdname Workbook
loadWorkbook <- function(file, password=NULL)
{
  if (!file.exists(path.expand(file)))
    stop("Cannot find ", path.expand(file))

  if (is.null(password)) {
    inputStream <- .jnew("java/io/FileInputStream", path.expand(file))
    wbFactory <- .jnew("org/apache/poi/ss/usermodel/WorkbookFactory")
    wb <- wbFactory$create(inputStream)

  } else {
    file <- .jnew("java/io/File", path.expand(file))
    fileSystem <-new(J("org/apache/poi/poifs/filesystem/NPOIFSFileSystem"),
                     file)
    info <- new(J("org/apache/poi/poifs/crypt/EncryptionInfo"),
                 fileSystem)
    decryptor <- J(info, "getDecryptor")
    verification <- J(decryptor, "verifyPassword", password)
    if (!verification) stop("password does not verify")
    dataStream <- J(decryptor, "getDataStream", fileSystem)
    wb <- new(J("org/apache/poi/xssf/usermodel/XSSFWorkbook"),
              dataStream)
  }


  return(wb)
}


######################################################################
#
#' @export
#' @rdname Workbook
saveWorkbook <- function(wb, file, password=NULL)
{
  nfile <- normalizePath(file, mustWork = FALSE)
  jFile <- .jnew("java/io/File", nfile)
  fh <- .jnew("java/io/FileOutputStream", jFile)

  # write the workbook to the file
  wb$write(fh)

  # close the filehandle
  .jcall(fh, "V", "close")

  if ( !is.null(password) ) {
    fs <- .jnew("org/apache/poi/poifs/filesystem/POIFSFileSystem")
    encMode <- J("org/apache/poi/poifs/crypt/EncryptionMode", "valueOf", "agile")
    info <- .jnew("org/apache/poi/poifs/crypt/EncryptionInfo", fs, encMode)

    enc <- info$getEncryptor()
    enc$confirmPassword(password)

    access <- J("org.apache.poi.openxml4j.opc.PackageAccess", "valueOf", "READ_WRITE")
    opc <- J("org.apache.poi.openxml4j.opc.OPCPackage", "open", jFile, access)
    outputStream <- enc$getDataStream(fs)
    opc$save(outputStream)
    opc$close()

    fos <- .jnew("java/io/FileOutputStream", nfile)
    fs$writeFilesystem(fos)
    fos$close()
  }

  invisible()
}



## # make a file handle and write the xlsx to file
## filename <- "C:/Temp/junk.xlsx"
## fh <- .jnew("java/io/FileOutputStream", filename)
## .jcall(wb, "V", "write", .jcast(fh, "java/io/OutputStream"))

#' Functions to manipulate worksheets.
#'
#' @aliases Worksheet
#'
#' @param wb a workbook object as returned by \code{createWorksheet} or
#' \code{loadWorksheet}.
#' @param sheetName a character specifying the name of the worksheet to create,
#' or remove.
#' @return \code{createSheet} returns the created \code{Sheet} object.
#'
#' \code{getSheets} returns a list of java object references each pointing to
#' an worksheet.  The list is named with the sheet names.
#' @author Adrian Dragulescu
#' @seealso To extract rows from a given sheet, see \code{\link{Row}}.
#' @examples
#'
#'
#' file <- system.file("tests", "test_import.xlsx", package = "xlsx")
#'
#' wb <- loadWorkbook(file)
#' sheets <- getSheets(wb)
#'
#' sheet  <- sheets[[2]]  # extract the second sheet
#'
#' # see all the available java methods that you can call
#' rJava::.jmethods(sheet)
#'
#' # for example
#' sheet$getLastRowNum()
#'
#' @name Sheet
NULL


######################################################################
# return the sheets
#
#' @export
#' @rdname Sheet
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
#' @export
#' @rdname Sheet
createSheet <- function(wb, sheetName="Sheet1")
{
  sheet <- .jcall(wb, "Lorg/apache/poi/ss/usermodel/Sheet;",
    "createSheet", sheetName)

  return(sheet)
}

######################################################################
#
#' @export
#' @rdname Sheet
removeSheet <- function(wb, sheetName="Sheet1")
{
  sheetInd <- wb$getSheetIndex(sheetName)

  .jcall(wb, "V", "removeSheetAt", as.integer(sheetInd))

  return(invisible())
}
