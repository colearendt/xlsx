######################################################################
# work with print setup
#
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
  
