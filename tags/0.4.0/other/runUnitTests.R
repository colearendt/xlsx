
# Run some simple automated tests for the package
#
#


require(RUnit); require(xlsx)

dirUT <- paste(.libPaths()[1], "/xlsx/tests", sep="")
#dirUT <- "H:/user/R/Adrian/findataweb/temp/xlsx/trunk/inst/tests"


source(paste(dirUT, "/test.comments.R", sep=""))
test.comments(type="xlsx")
test.comments(type="xls")


source(paste(dirUT, "/test.dataFormats.R", sep=""))
test.dataFormats(type="xlsx")
test.dataFormats(type="xls") 


source(paste(dirUT, "/test.export.R", sep=""))
test.export(type="xlsx")
test.export(type="xls")

source(paste(dirUT, "/test.formats.R", sep=""))
test.formats(type="xlsx")
test.formats(type="xls")

source(paste(dirUT, "/test.import.R", sep=""))
test.import(type="xlsx")
test.import(type="xls")

#source(paste(dirUT, "/test.lowlevel.R", sep=""))
#test.lowlevel()

source(paste(dirUT, "/test.otherEffects.R", sep=""))
test.otherEffects(type="xlsx")
test.otherEffects(type="xls")

source(paste(dirUT, "/test.picture.R", sep=""))
test.picture(type="xlsx")
test.picture(type="xls")

source(paste(dirUT, "/test.workbook.R", sep=""))
test.workbook(type="xlsx")
test.workbook(type="xls")

source(paste(dirUT, "/test.ranges.R", sep=""))
test.ranges(type="xlsx")
test.ranges(type="xls")









