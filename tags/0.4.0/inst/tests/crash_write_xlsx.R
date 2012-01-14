
options(java.parameters="-Xmx1024m")  
require(xlsx)

N <- 1000    # crash with 5000

sample <- matrix(1:N, nrow=N, ncol=10)



Rintf <- .jnew("dev/RInterface")

wb <- createWorkbook()
for (k in 1:50) {
  cat("On sheet", k, "\n")
  sheetName <- paste("Sheet", k, sep="")
  sheet <- createSheet(wb, sheetName)

  x <- sample
  iOffset <- jOffset <- 0L
  Rintf$NCOLS <- ncol(x)
  Rintf$NROWS <- nrow(x)
  .jcall(Rintf, "V", "createCells", sheet, 0L, 0L)

  for (j in 1:ncol(x)) {
        if (class(x[, j]) == "integer") {
            aux <- .jarray(x[, j])
            .jcall(Rintf, "V", "writeColInts", sheet, iOffset, 
                as.integer(j - 1), aux)
        }
        else if (class(x[, j]) == "numeric") {
            aux <- .jarray(x[, j])
            .jcall(Rintf, "V", "writeColDoubles", sheet, iOffset, 
                as.integer(j - 1), aux)
        }
        else {
            aux <- .jarray(as.character(x[, j]))
            .jcall(Rintf, "V", "writeColStrings", sheet, iOffset, 
                as.integer(j - 1), aux)
        }
    }
}
saveWorkbook(wb, "C:/Temp/junk.xlsx")
 



