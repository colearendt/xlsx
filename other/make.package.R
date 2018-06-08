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
.setEnv <- function()
{
  pkgdir    <<- "/home/adrian/Documents/repos/git/dragua/xlsx/"
  outdir    <<- "/tmp"
  Rcmd      <<- "R CMD"
  invisible()
}

##################################################################
##################################################################

## .build.java() 

version    <- "0.6.1"      
package.gz <- paste("xlsx_", version, ".tar.gz", sep="")
.setEnv()   

# make the package
setwd(outdir)
cmd <- paste(Rcmd, "build --force --md5", pkgdir) 
print(cmd); system(cmd)

# check source with --as-cran on the tarball before submitting it
cmd <- paste(Rcmd, "check --as-cran", package.gz)
print(cmd); system(cmd)


# install the package
install.packages(package.gz, repos=NULL, type="source", clean=TRUE)

# Run the tests from inst/tests/lib_tests_xlsx.R


# make the package for CRAN
cmd <- paste(Rcmd, "build --compact-vignettes", pkgdir)
print(cmd); system(cmd)











