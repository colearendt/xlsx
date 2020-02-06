# investigate specific files that turn up!
diff tests/testthat/test-write.xlsx.R /cloud/project/tests/testthat/test-write.xlsx.R -y --suppress-common-lines -w --ignore-blank-lines

# helpful overview of files
diff -warq --ignore-blank-lines tests /cloud/project/tests