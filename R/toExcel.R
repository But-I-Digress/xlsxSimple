
#' @title toExcel
#'
#' @description Save something to a single sheet, Excel workbook. This is for when calling \code{toSheet()} and \code{saveWorkbook()} is too much work. Which happens when one is developing something. 
#'
#' @param x A data frame, plot or path to a saved plot.
#' @param filePath An optional character scalar, the path and filename for the file.
#' @param title An optional character scalar, the title of the workbook.
#' @param subject An optional character scalar, the subject of the workbook.
#' @param creator An optional character scalar, the creator of the workbook. Defaults to the logged-on user name.
#' @param start An optional logical scalar, should the workbook be started in Excel.
#'
#' @return A character scalar, the path of the file, invisibly.
#'
#' @examples
#'
#' toExcel(mtcars)
#'
#' @seealso See \code{\link{toSheet}} and \code{\link{saveWorkbook}} for details. This function calls and then the other.
#'
#' @export
toExcel <- function (x, filePath = thisFile(), title = NA, subject = NA, creator = Sys.getenv("USERNAME"), start = TRUE, ...) {
	toSheet(x, ...)
	saveWorkbook(filePath = filePath, title = title, subject = subject, creator = creator, start = start)
}