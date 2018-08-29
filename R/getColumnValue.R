#' @import xlsx
	
#' @title Get Column Value
#'
#' @description Gets the values from column in an xlsx or xlsxsimple, Excel worksheet.
#' 
#' @param sheet An xlsx or xlsxsimple worksheet.
#' @param col.name A character scalar, the name of the column to return.
#'
#' @return A vector.	
#'
#' @examples 
#' \dontrun{
#' sheet <- toSheet(mtcars)
#' getColumnValue(sheet, "mpg")
#' }
#'
#' @name getColumnValue
#'
#' @export	
setGeneric("getColumnValue", function (sheet, col.name)standardGeneric("getColumnValue"))

#' @rdname getColumnValue
#' @export 
setMethod("getColumnValue", signature(sheet = "jobjRef"), function (sheet, col.name) {
		r <- xlsx::getRows(sheet, 1)
		cells <- xlsx::getCells(r, colIndex=1:r[[1]]$getLastCellNum())
		col.names <- mapply(xlsx::getCellValue, cells)
		
		if (!col.name %in% col.names) {
			warning("Column" , col.name, " does not exist.")
			return(NA)
		}
		
		rows <- xlsx::getRows(sheet, 1:sheet$getLastRowNum() + 1) # Excel row numbering starts at 0
		cells <- xlsx::getCells(rows, colIndex=which(col.names == col.name))
		mapply(xlsx::getCellValue, cells)		
	})