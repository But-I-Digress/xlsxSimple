
getWorkbook <- function () {
	if (is.null(get0("workbook", envir = xlsxSimple))) assign("workbook", xlsx::createWorkbook(), envir = xlsxSimple)
	get("workbook", envir = xlsxSimple)
}

#' Create an Empty Workbook object
#'
#' @param ... Arguments to pass through to \code{xlsx::createWorkbook}
#'
#' @seealso \code{\link[xlsx]{createWorkbook}}
#'
#' @return The worksheet object.
#'
#' @examples
#' \dontrun{
#' createWorkbook() -> book
#' }
#'
#' @export
createWorkbook <- xlsx::createWorkbook
