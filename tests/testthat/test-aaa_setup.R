
## Remove tmp directory if already present
remove_tmp()

## Development version - add classPath
ijava <- rprojroot::is_r_package$find_file('inst/java')
rJava::.jaddClassPath(dir(ijava, full.names = TRUE))