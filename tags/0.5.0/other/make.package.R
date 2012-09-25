# Automate the making of the package
#
#


##################################################################
#
.build.java <- function()
{
  # build maven project
  # and put the jars (package and its dependencies) to the inst/java/ directory
  system(paste("mvn -f ", file.path(pkgdir,'other', 'pom.xml'), ' package' ) )

  invisible()
}


##################################################################
#
.setEnv <- function(computer=c("HOME", "LAPTOP", "WORK"))
{
  if (computer=="WORK") {
    pkgdir  <<- "C:/google/rexcel/trunk/"
    outdir  <<- "H:/"
    Rcmd    <<- "S:/All/Risk/Software/R/R-2.12.1/bin/i386/Rcmd"
  } else if (computer == "LAPTOP") {
    pkgdir    <<- "/home/adrian/Documents/rexcel/trunk/"
    outdir    <<- "/tmp/"
    Rcmd      <<- "R CMD"
  } else if (computer == "HOME") {
    pkgdir    <<- "/home/adrian/Documents/rexcel/trunk/"
    outdir    <<- "/tmp"
    Rcmd      <<- "R CMD"
  } else if (computer == "WORK2") {  
    pkgdir  <<- "C:/google/rexcel/trunk/"
    outdir  <<- "H:/"
    Rcmd    <<- '"C:/Program Files/R/R-2.15.1/bin/i386/Rcmd"'
  } else {
  }

  invisible()
}

##################################################################
##################################################################

version    <- "0.5.0"      
package.gz <- paste("xlsx_", version, ".tar.gz", sep="")

.setEnv("HOME")   # "HOME" "WORK2" "LAPTOP"

.build.java() 

# make the package
setwd(outdir)
cmd <- paste(Rcmd, "build --force", pkgdir)
print(cmd)
system(cmd)

install.packages(package.gz, repos=NULL, type="source")

# Run the tests from inst/tests/lib_tests_xlsx.R


# make the package for CRAN
cmd <- paste(Rcmd, "build --compact-vignettes", pkgdir)
print(cmd); system(cmd)


# check source with --as-cran on the tarball before submitting it
cmd <- paste(Rcmd, "check --as-cran", package.gz)
print(cmd); system(cmd)










# change the version
#version <- .update.DESCRIPTION(pkgdir, version)

## ##################################################################
## #
## .update.DESCRIPTION <- function(packagedir, version)
## {
##   file <- paste(packagedir, "DESCRIPTION", sep="") 
##   DD  <- readLines(file)
##   ind  <- grep("Version: ", DD)
##   aux <- strsplit(DD[ind], " ")[[1]]
  
##   if (is.null(version)){   # increase by one 
##     vSplit    <- strsplit(aux[2], "\\.")[[1]]
##     vSplit[3] <- as.character(as.numeric(vSplit[3])+1) 
##     version <- paste(vSplit, sep="", collapse=".")
##   }   
##   DD[ind] <- paste(aux[1], version)

##   ind <- grep("Date: ", DD)
##   aux <- strsplit(DD[ind], " ")[[1]]
##   DD[ind] <- paste(aux[1], Sys.Date())
  
##   writeLines(DD, con=file)
##   return(version)
## }


