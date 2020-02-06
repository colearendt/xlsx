######################################################################
# work with print setup
#
#' Function to manipulate print setup.
#'
#' Other settings are available but not exposed.  Please see the java docs.
#'
#' @aliases printSetup PrintSetup
#' @param sheet a worksheet object \code{\link{Worksheet}}.
#' @param fitHeight numeric value to set the number of pages high to fit the
#' sheet in.
#' @param fitWidth numeric value to set the number of pages wide to fit the
#' sheet in.
#' @param copies numeric value to set the number of copies.
#' @param draft logical indicating if it's a draft or not.
#' @param footerMargin numeric value to set the footer margin.
#' @param headerMargin numeric value to set the header margin.
#' @param landscape logical value to specify the paper orientation.
#' @param pageStart numeric value from where to start the page numbering.
#' @param paperSize character to set the paper size.  Valid values are
#' "A4_PAPERSIZE", "A5_PAPERSIZE", "ENVELOPE_10_PAPERSIZE",
#' "ENVELOPE_CS_PAPERSIZE", "ENVELOPE_DL_PAPERSIZE",
#' "ENVELOPE_MONARCH_PAPERSIZE", "EXECUTIVE_PAPERSIZE", "LEGAL_PAPERSIZE",
#' "LETTER_PAPERSIZE".
#' @param noColor logical value to indicate if the prints should be color or
#' not.
#' @return A reference to a java PrintSetup object.
#' @author Adrian Dragulescu
#' @examples
#'
#'
#' wb <- createWorkbook()
#' sheet <- createSheet(wb, "Sheet1")
#' ps   <- printSetup(sheet, landscape=TRUE, copies=3)
#'
#'
#' @export
#' @name PrintSetup
printSetup <- function(sheet, fitHeight=NULL,
  fitWidth=NULL, copies=NULL, draft=NULL, footerMargin=NULL,
  headerMargin=NULL, landscape=FALSE, pageStart=NULL, paperSize=NULL,
  noColor=NULL)
{
  ps <- .jcall(sheet, "Lorg/apache/poi/ss/usermodel/PrintSetup;",
               "getPrintSetup")

  if (!is.null(fitHeight))
    .jcall(ps, "V",  "setFitHeight", .jshort(fitHeight))

  if (!is.null(fitWidth))
    .jcall(ps, "V", "setFitWidth", .jshort(fitWidth))

  if (!is.null(copies))
    .jcall(ps, "V", "setCopies", .jshort(copies))

  if (!is.null(draft))
    .jcall(ps, "V", "setDraft", draft)

  if (!is.null(footerMargin))
    .jcall(ps, "V", "setFooterMargin", footerMargin)

  if (!is.null(headerMargin))
    .jcall(ps, "V", "setHeaderMargin", headerMargin)

  if (!is.null(landscape))
    .jcall(ps, "V", "setLandscape", as.logical(landscape))

  if (!is.null(pageStart))
    .jcall(ps, "V", "setPageStart", .jshort(pageStart))

  if (!is.null(paperSize)){
    pagesizeInd <- .jfield(ps, NULL, paperSize)
    .jcall(ps, "V", "setPaperSize", .jshort(pagesizeInd))
  }

  if (!is.null(noColor))
    .jcall(ps, "V", "setNoColor", noColor)

  ps
}


##   if (setAutobreaks)
##     .jcall(sheet, "V", "setAutobreaks", )

