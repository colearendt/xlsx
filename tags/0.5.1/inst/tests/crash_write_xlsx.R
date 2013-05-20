
options(java.parameters="-Xmx1024m")  
require(xlsx)

N <- 5000    # crash with 5000, at sheet 20

x <- as.data.frame(matrix(1:N, nrow=N, ncol=10))

wb <- createWorkbook()
for (k in 1:50) {
  cat("On sheet", k, "\n")
  sheetName <- paste("Sheet", k, sep="")
  sheet <- createSheet(wb, sheetName)

  addDataFrame(x, sheet)
}

saveWorkbook(wb, "C:/Temp/junk.xlsx")
 



