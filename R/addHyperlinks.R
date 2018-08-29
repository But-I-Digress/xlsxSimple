#' @import xlsx

#' @title Add Hyper Links
#'
#' @description \code{addHyperlinks} adds hyperlinks to an \pkg{xlsx} or \pkg{xksxSimple}, Excel worksheet.
#' 
#' @param sheet An \pkg{xlsx} or \pkg{xksxSimple}, worksheet.
#' @param col.name A character scalar, the name of the column where the hyperlinks are to be added. The column will be created if it doesn't exist.
#' @param urls A character vector, the list of the hyperlink targets. 
#'
#' @return The worksheet invisibly.
#'
#' @examples
#' \dontrun{
#' toSheet(mtcars) -> sheet
#' addHyperLinks(sheet, "Google", "http://google.com")
#' }
#'	
#' @name addHyperLinks
#'
#' @export 	
setGeneric("addHyperLinks", function(sheet, col.name, urls) standardGeneric("addHyperLinks"))

#' @rdname addHyperLinks
#' @export 
setMethod("addHyperLinks", signature(sheet = "jobjRef"), function (sheet, col.name, urls) {
	if (attr(sheet, "startRow") > 1) stop("Cannot add hyperlinks to a table that does not start at row one.")
	if (sheet$getLastRowNum() < 1) {		
		warning("Attempting to add hyperlinks to a worksheet with now rows.")
	} else {
		auto.filter <- !is.null(attr(sheet, "auto.filter")) && attr(sheet, "auto.filter")
		b <- !is.na(urls)
		urls <- urls[b]
		
		r <- xlsx::getRows(sheet, 1)
		cells <- xlsx::getCells(r, colIndex=1:r[[1]]$getLastCellNum())
		col.names <- mapply(xlsx::getCellValue, cells)
		
		rows <- xlsx::getRows(sheet, 1:sheet$getLastRowNum() + 1) # Excel row numbering starts at 0
		rows <- rows[b] # Skip the rows missing URLs
		if (col.name %in% col.names) {			
			cells <- xlsx::getCells(rows, colIndex=which(col.names == col.name))		
		} else {	
			n <- length(cells) + 1
			cell <- xlsx::createCell(r, colIndex=n)
			xlsx::setCellStyle(cell[[1,1]], getColnamesStyle(sheet))
			xlsx::setCellValue(cell[[1,1]], col.name)
			cells <- xlsx::createCell(rows, colIndex=n)
			mapply(xlsx::setCellValue, cells, urls)
			if (auto.filter && n <= 676L) xlsx::addAutoFilter(sheet, paste("A:", getColSpec(n), sep=""))
			xlsx::autoSizeColumn(sheet, n)
		}		
		mapply(xlsx::addHyperlink, cells, sapply(urls, URLencode))
	}
	invisible(sheet)
})