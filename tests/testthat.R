library(testthat)
library(xlsx)

invisible(
  rJava::.jcall(
    "java/lang/System",
    "Ljava/lang/String;",
    "setProperty",
      "java.io.tmpdir", tempdir()
    )
)

test_check("xlsx")
