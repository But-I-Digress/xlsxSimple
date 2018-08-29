thisFile <- function () {
	if(interactive()) stop("saveWorkbook() needs a file path.")
	cmdArgs <- commandArgs(trailingOnly = FALSE)
	needle <- "--file="
	match <- grep(needle, cmdArgs)
	if (length(match) > 0) {
		# Rscript
		normalizePath(sub(needle, "", cmdArgs[match]))
	} else {
		# R console
		normalizePath(sys.frames()[[1]]$ofile)
	}
}

#' @title Save Work Book
#'
#' @description \code{saveWorkbook} saves an Excel workbook on the file system, setting the metadata. Optionally, the file is opened in Excel.
#' 
#' @param workbook An optional xlsxSimple workbook.
#' @param file.path An optional character scalar, the path and filename for the file.
#' @param title An optional character scalar, the title of the workbook.
#' @param subject An optional character scalar, the subject of the workbook.
#' @param creator An optional character scalar, the creator of the workbook. Defaults to the logged-on user name.
#' @param start An optional logical scalar, should the workbook be started in Excel.
#'
#' @return A character scalar, the path of the file, invisibly.
#' 
#' @note If the \code{file.path} does not end with ``.xlsx'' then the fiile extension will be set to ``.xlsx''. If \code{file,path} is omitted then the script file path will be used, but that extension substiuted. But an error it thrown if \code{file.path} is omitted in interactive mode.
#'
#' @examples
#'
#' book <- createWorkbook()
#' saveWorkbook(book, "foo.xlsx", "Foo", "Examples", start=TRUE)
#'
#' @export
saveWorkbook <- function (workbook = getWorkbook(), file.path = thisFile(), title = NA, subject = NA, creator = Sys.getenv("USERNAME"), start = FALSE) {
	#http://thinktibits.blogspot.com/2014/07/read-write-metadata-excel-poi-example.html
	if(tolower(tools::file_ext(file.path)) != "xlsx") file.path <- paste(tools::file_path_sans_ext(file.path), "xlsx", sep=".")
	
	# set Metadata
	
	props <- workbook$getProperties()
	cores <- props$getCoreProperties()
	if (!is.null(creator) && !is.na(creator)) cores$setCreator(creator)
	if (!is.null(title) && !is.na(title)) cores$setTitle(title)
	if (!is.null(subject) && !is.na(subject)) cores$setSubjectProperty(subject)
	
	jFile <- rJava::.jnew("java/io/File", file.path)
	fh <- rJava::.jnew("java/io/FileOutputStream", jFile)
	workbook$write(fh)
	rJava::.jcall(fh, "V", "close")
	
	if(missing(workbook)) assign("workbook", NULL, envir = xlsxSimple)
	
	if (start) shell.exec(file.path)
	invisible(file.path)
}