getColnamesStyle <- function (sheet) {
	book <- sheet$getWorkbook()
	xlsx::CellStyle(book) +
		xlsx::Alignment(h='ALIGN_CENTER') +
		xlsx::Fill(foregroundColor='#cccccc') +
		xlsx::Font(book, isBold=TRUE)
}

getSheetName <- function (workbook) sprintf("Sheet%1.0f", workbook$getNumberOfSheets() + 1)

#' @title Send something to a sheet
#'
#' @description \code{toSheet} sends tabular data or an image to a new xlsx work sheet.
#'
#' @param x A data frame, plot or path to a saved plot.
#' @param workbook An optional xlsx work book.
#' @param title Character, an optional title for the new sheet.
#'
#' @return A worksheet object, invisibly.
#'
#' @examples
#' \dontrun{
#'	book <- createWorkbook()
#'	toSheet(mtcars, book, "mtcars")
#'
#'	library(ggplot2)
#'	p <- ggplot(mtcars, aes(mpg))
#'	p <- p + geom_bar()
#'	toSheet(p)
#' }
#' @export 
toSheet <- function (x, workbook, title, ...) UseMethod("toSheet")

#' @describeIn toSheet Send a data frame to a work sheet.
#'
#' @param auto.filter A boolean, should the columns be filtered.
#' @param freeze.pane A boolean, should the top row be frozen.
#' @param row.names A boolean, should therow names of x be written along with x to the file.
#' @param startRow A scalar integer, the row to start the data frame.
#' @param ... Other parameters to pass to addDataFrame.
#'
#' @export 
toSheet.data.frame <- function (d, workbook = getWorkbook(), sheetName = getSheetName(workbook), auto.filter=TRUE, freeze.pane=TRUE, row.names=FALSE, startRow=1, ...) {
	worksheet <- xlsx::createSheet(workbook, sheetName)
	
	xlsx::addDataFrame(
		d,
		worksheet, 
		row.names=row.names, 
		colnamesStyle=getColnamesStyle(worksheet),
		startRow=startRow,
		...
	)

	if (freeze.pane && startRow == 1) xlsx::createFreezePane(worksheet, 2L, 1L)
	if (auto.filter && ncol(d) <= 676L) xlsx::addAutoFilter(worksheet, paste("A:", getColSpec(ncol(d)), sep=""))
	if (auto.filter && ncol(d) <= 676L) xlsx::addAutoFilter(worksheet, paste("A", startRow, ":", getColSpec(ncol(d)), nrow(d) + startRow, sep=""))
	xlsx::autoSizeColumn(worksheet, 1:ncol(d))
	
	attr(worksheet, "auto.filter") <- auto.filter
	attr(worksheet, "startRow") <- startRow	
	
	invisible(worksheet)
}

#' @describeIn toSheet Send a tibble to a work sheet.
#' @export 	
toSheet.tbl_df <- function (d, workbook = getWorkbook(), sheetName = getSheetName(workbook), ...) toSheet.data.frame(as.data.frame(d),  workbook, sheetName, ...)

#' @describeIn toSheet Send a data frame to a work sheet.
#' @export 	
toSheet.tbl <- function (d, workbook = getWorkbook(), sheetName = getSheetName(workbook), ...) toSheet.data.frame(as.data.frame(d),  workbook, sheetName, ...)

#' @describeIn toSheet Send a matrix to a work sheet.
#' @export 	
toSheet.matrix <- function (d, workbook = getWorkbook(), sheetName = getSheetName(workbook), row.names = TRUE, ...) toSheet.data.frame(as.data.frame(d), workbook, sheetName, row.names, ...)

#' @describeIn toSheet send a saved image to a work sheet. 
#'
#' @param landscape A boolean, should the sheet be oriented horizontally. 
#'
#' @export 	
toSheet.character <- function (x,  workbook = getWorkbook(), sheetName = getSheetName(workbook), landscape = TRUE) {
	sheet <- xlsx::createSheet(workbook, sheetName)
	xlsx::addPicture(x, sheet)	
	xlsx::printSetup(sheet, landscape = landscape)
	return (sheet)
}

#' @describeIn toSheet Send a ggplot object to a work sheet.
#'
#' @param filename Optional file path and name for the temporary file for the image of the plot.
#' @param height Numeric, optional height of the image in inches.
#' @param width Numeric, optional width of the image in inches.
#'
#' @export 	
toSheet.ggplot <- function (p, workbook = getWorkbook(), sheetName = getSheetName(workbook), filename = tempfile(fileext=".jpeg") , landscape = TRUE, height = if(landscape) 7 else 9.7, width = if (landscape) 8.5 else 5.99, ...) {
	ggplot2::ggsave(filename, p, height = height, width = width, ...)
	toSheet.character(filename, workbook, sheetName, landscape = landscape)
}

#' @export 	
toSheet.gtable <- function (p, workbook = getWorkbook(), sheetName = getSheetName(workbook), filename = tempfile(fileext=".jpeg") , landscape = TRUE, height = if(landscape) 7 else 9.7, width = if (landscape) 8.5 else 5.99, ...) {
	ggplot2::ggsave(filename, p, height = height, width = width, ...)	
	toSheet.character(filename, workbook, sheetName, landscape = landscape)
}

#' @export
toSheet.default <- function (x, ...) stop(sprintf("Error, you are trying to write to a sheet something of class '%s'.", class(x)))