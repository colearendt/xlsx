#
#
#

test.export <- function(outdir="C:/Temp/", type="xlsx")
{
  cat("####################################################\n")
  cat("Test high level export ... \n")
  cat("####################################################\n")      

  x <- data.frame(mon=month.abb[1:10], day=1:10, year=2000:2009,
    date=seq(as.Date("2009-01-01"), by="1 month", length.out=10),
    bool=ifelse(1:10 %% 2, TRUE, FALSE))

  file <- paste(outdir, "test_export.", type, sep="")
  cat("Write an xlsx file with char, int, double, date, bool columns ...\n")
  write.xlsx(x, file)
  cat("Wrote file ", file, "\n\n")

## Used not to write col.names!  
##   test.df <- data.frame(mycol1=1:10, mycol2=as.character(1:10))
##   write.xlsx(test.df, file=paste(outdir, 'test.df.xlsx', sep=""),
##     col.names=TRUE, row.names=FALSE)
  
  
  cat("Test the append argument by adding another sheet \n")
  file <- paste(outdir, "test_export.xlsx", sep="")
  write.xlsx(USArrests, file, sheetName="usarrests", append=TRUE)
  cat("Wrote file ", file, "\n\n")


  cat("Test writing/reading data.frames with NA values ... \n") 
  file <- paste(outdir, "test_writeread_NA.xlsx", sep="")
  x <- data.frame(matrix(c(1.0, 2.0, 3.0, NA), 2, 2))
  write.xlsx(x, file, row.names=FALSE)
  xx <- read.xlsx(file, 1)
  if (identical(x,xx)){
    cat("OK \n\n")
  } else {
    cat("FAILED! \n\n")
  }

  
  cat("Test speed ...\n")
  file <- paste(outdir, "test_exportSpeed.xlsx", sep="")
  x <- expand.grid(ind=1:30, letters=letters, months=month.abb)
  x <- cbind(x, val=runif(nrow(x)))
  print(system.time(write.xlsx(x, file)))   # 206 seconds
  print(system.time(write.xlsx2(x, file)))  # 31 seconds
  cat("Wrote file ", file, "\n\n")

}


##   for (ic in 1:ncol(x)){
##     wb <- loadWorkbook(file)
##     sheets <- getSheets(wb)
##     sheet <- sheets[["Sheet1"]]
##     rows  <- createRow(sheet, 1:nrow(x))
##     cells <- createCell(rows, col=ic)
##     mapply(setCellValue, cells, x[,ic])
##     saveWorkbook(wb, file)
##   }


  ## cat("Test speed ...\n")
  ## file <- paste(outdir, "test_exportSpeed.xlsx", sep="")
  ## x <- expand.grid(ind=1:10, letters=letters, months=month.abb)
  ## (time <- system.time(write.xlsx(x, file)))
  ## cat("Wrote file ", file, "\n\n")

##   cat("Test memory ...\n")
##   file <- paste(outdir, "test_exportMemory.xlsx", sep="")
##   x <- expand.grid(ind=1:1000, letters=letters, months=month.abb)
##   cat("Writing object size: ", object.size(x), " uses all Java heap space\n")
##   (time <- system.time(write.xlsx2(x, file)))
##   cat("Wrote file ", file, "\n\n")




