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
    Rcmd    <<- "S:/All/Risk/Software/R/R-2.15.2/bin/i386/Rcmd"
  } else if (computer == "LAPTOP") {
    pkgdir    <<- "/home/adrian/Documents/rexcel/trunk/"
    outdir    <<- "/tmp/"
    Rcmd      <<- "R CMD"
  } else if (computer == "HOME") {
    pkgdir    <<- "/home/adrian/Documents/rexcel/trunk/"
    outdir    <<- "/tmp"
    Rcmd      <<- "R CMD"
  } else if (computer == "WORK2") {
    Sys.setenv(R_HOME="C:/Program Files/R/R-3.0.1" )
    Sys.setenv(R_LIBS="H:/R_libs" )
    pkgdir  <<- "C:/google/rexcel/trunk/"
    outdir  <<- "H:/"
    Rcmd    <<- '"C:/Program Files/R/R-3.0.1/bin/x64/Rcmd"'
  } else {
  }

  invisible()
}

##################################################################
##################################################################

version    <- "0.5.4"      
package.gz <- paste("xlsx_", version, ".tar.gz", sep="")
computer <- "WORK2" #"WORK2" "LAPTOP"
.setEnv(computer)   

.build.java() 

# make & install the package
setwd(outdir)
cmd <- paste(Rcmd, "build --force --md5", pkgdir) 
print(cmd)
system(cmd)

if (computer %in% c("WORK2"))   #  Win7 needs special treatment, why?!
{
  # only x64 for now, not patience for figure out why rJava
  # doesn't install properly on i386
  install.packages(package.gz, repos=NULL, type="source",
    clean=TRUE, INSTALL_opts=c("--no-multiarch"))
  
} else {
  install.packages(package.gz, repos=NULL, type="source",
                   clean=TRUE)
}



# Run the tests from inst/tests/lib_tests_xlsx.R


# make the package for CRAN
cmd <- paste(Rcmd, "build --compact-vignettes", pkgdir)
print(cmd); system(cmd)


# check source with --as-cran on the tarball before submitting it
cmd <- paste(Rcmd, "check --as-cran", package.gz)
print(cmd); system(cmd)









