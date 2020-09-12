# This file contains other utilities to make this a more useful package

# If you have an Excel file template and use xlsx to fill in some data and write
# it back to a file (which is a general technique to produce Excel-based reports),
# then it works beautifully except that Excel will not refresh formulae and pivot
# tables when you open the file.
#
# The first problem of forcing formula calculation is easy to solve but the pivot
# table issue is not fixable by Apache POI library.  It turns out that
# Excel can refresh pivot tables when you open the file if the pivot cache definition
# XML file has the refreshOnLoad flag set to 1.
# See https://stackoverflow.com/questions/11670816/how-to-refresh-pivot-cache-of-excel-2010-using-open-xml#16624292
#
# This method automates the hack.  Basically, unzip the excel file, update
# the pivot cache definition files and writes it back.  It's really a hack.
# This hack should be temporary when apache-poi eventually allows adding the
# refreshOnLoad attribute for existing pivot tables in the workbook.

# Make Excel refresh all formula when the file is open
# If output is NULL then overwrite the original file
#' @title Force Refresh Pivot Tables and Formulae
#'
#' @description Functions to force formula calculation or refresh of pivot
#' tables when the Excel file is opened.
#'
#' @details
#' \code{forcePivotTableRefresh} forces pivot tables to be refreshed when the Excel file is opened.
#' \code{forceFormulaRefresh} forces formulae to be recalculated when the Excel file is opened.
#'
#' @param file the path of the source file where formulae/pivot table needs to be refreshed
#' @param output the path of the output file.  If it is \code{NULL} then the source file will be overwritten
#' @param verbose Whether to make logging more verbose
#'
#' @return Does not return any results
#'
#' @examples
#' # Patch a file where its pivot tables are not recalculated when the file is opened
#' \dontrun{
#' forcePivotTableRefresh("/tmp/file.xlsx")
#' forcePivotTableRefresh("/tmp/file.xlsx", "/tmp/fixed_file.xlsx")
#' }
#' # Patch a file where its formulae are not recalculated when the file is opened
#' \dontrun{
#' forceFormulaRefresh("/tmp/file.xlsx")
#' forceFormulaRefresh("/tmp/file.xlsx", "/tmp/fixed_file.xlsx")
#' }
#'
#'
#' @author Tom Kwong
#'
#' @export
#' @rdname autoRefresh
forceFormulaRefresh <- function(file, output=NULL, verbose=FALSE) {

  # redirect output to source location?
  if (is.null(output)) {
    output <- file
    if (verbose) {
      cat(sprintf("Overwriting source file at %s\n", file))
    }
  }

  wb <- loadWorkbook(file)
  wb$setForceFormulaRecalculation(TRUE)
  saveWorkbook(wb, output)

  if (verbose) {
    cat(sprintf("Successfully patched file to auto calculate formulae. File saved at %s\n", output))
  }
}

#' @export
#' @rdname autoRefresh
forcePivotTableRefresh <- function(file, output=NULL, verbose=FALSE) {

  if (!file.exists(file)) {
    stop("File does not exist ", file)
  }

  # redirect output to source location?
  if (is.null(output)) {
    output <- file
    if (verbose) {
      cat(sprintf("Overwriting source file at %s\n", file))
    }
  }

  # create a temp directory to hold the unzip'ed Excel content
  tmpDir <- tempfile()
  dir.create(tmpDir)
  #cat(sprintf("Temp directory: %s", tmpDir))

  # unzip the excel file
  unzip(file, exdir = tmpDir)

  # find pivot cache definition files & patch them
  pivotTablesPatched <- 0
  pivotCacheDir <- file.path(tmpDir, "xl", "pivotCache")
  if (dir.exists(pivotCacheDir)) {
    pivotCacheFiles <- list.files(path = pivotCacheDir)
    pivotCacheDefFiles <- pivotCacheFiles[grepl("pivotCacheDefinition", pivotCacheFiles)]
    result <- lapply(pivotCacheDefFiles,
                     function(defFile) {
                       # Read pivot cache definition file
                       workingFile <- file.path(pivotCacheDir, defFile)
                       text <- readLines(workingFile, warn = FALSE)

                       # Add refreshOnLoad attribute
                       text <- gsub('refreshedBy=', 'refreshOnLoad="1" refreshedBy=', text)

                       # Write back to the file now.
                       # Excel is particular about the file format... cannot use writeLines
                       # or else it will be a unix-style (newline, has EOL at last line)
                       text <- paste(text, collapse = '\r\n')
                       cat(text, file = workingFile)
                     })
    pivotTablesPatched <- length(result)
  }

  if (pivotTablesPatched > 0) {
    oldwd <- getwd()
    setwd(tmpDir)
    tmpOutputFile <- paste0(tempfile(), ".xlsx")
    status <- zip(tmpOutputFile,
                  files = list.files(tmpDir, recursive = TRUE, all.files = TRUE),
                  flags = "-r9q")
    setwd(oldwd)
    if (status != 0) {
      stop(sprintf("Unable to create zip file at %s", tmpOutputFile))
    }
    if (!file.copy(tmpOutputFile, output, overwrite = TRUE)) {
      stop(sprintf("Unable to save file to %s", output))
    }
    if (verbose) {
      cat(sprintf("Successfully patched file to auto refresh pivot tables. File saved at %s\n", output))
    }
  } else {
    if (file != output) {  # only make a copy if output is different from orig file
      file.copy(file, output)
    }
    warning(sprintf("This excel file has no pivot table %s", file))
    if (verbose) {
      cat("Nothing's done\n")
    }
  }
}

# # Unit testing
# # Run this file from the xlsx project home directory.
# input <- "resources/test_template1_stale.xlsx"
#
# # unit test 1.  Take source file output.xlsx and patch & save to output2.xlsx
# cat("Unit test 1\n")
# tmp <- "/tmp/temp.xlsx"  # intermediate file for two operations below
# output <- "/tmp/test_pivot_refresh_1.xlsx"
# cat(sprintf("output file: %s\n", output))
# if (file.exists(tmp)) { file.remove(tmp) }
# if (file.exists(output)) { file.remove(output) }
# forcePivotTableRefresh(input, tmp)
# forceFormulaRefresh(tmp, output)
#
# # unit test 2. Make a copy of the source file.  Then patch in-place.
# # this usage seems more natural
# cat("Unit test 2\n")
# output <- "/tmp/test_pivot_refresh_2.xlsx"
# cat(sprintf("output file: %s\n", output))
# file.copy(input, output, overwrite = TRUE)
# forcePivotTableRefresh(output)
# forceFormulaRefresh(output)
#
# # unit test 3.  Non-verbose.
# cat("Unit test 3 (expect no output)\n")
# output <- "/tmp/test_pivot_refresh_3.xlsx"
# cat(sprintf("output file: %s\n", output))
# file.copy(input, output, overwrite = TRUE)
# forcePivotTableRefresh(output, verbose = FALSE)
# forceFormulaRefresh(output, verbose = FALSE)
#
# cat("Unit test completed!\n")
#
#
# # experiment how to patch Excel file
# # lines <- readLines("/tmp/good_pivot_cache_def_file.xml", warn = FALSE)
# # lines <- gsub('refreshedBy=', 'refreshOnLoad="1" refreshedBy=', lines)
# # lines <- paste(lines, collapse = '\r\n')
# # cat(lines, file = "/tmp/test_pivot_cache_def_file.xml", fill = FALSE)
#
