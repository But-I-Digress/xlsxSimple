#' @import rJava
#' @import xlsxjars

xlsxSimple <- new.env()

.onLoad <- function(libname, pkgname) {
	# The xlsx package sets these options so let's save the values and then put them back after xlsx is loaded. 
	xlsx.datetime.format <- getOption("xlsx.datetime.format",  "d MMM yyyy")
	xlsx.date.format <- getOption("xlsx.date.format",  "d MMM yyyy")
	
	# Create a workbook and store it. This will allow xlsx to do its initialization. 
	assign("book", xlsx::createWorkbook(), envir = xlsxSimple)
	
	options(xlsx.datetime.format = xlsx.datetime.format, xlsx.date.format = xlsx.date.format)
}
