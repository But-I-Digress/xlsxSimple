#' @title xlxs Simple
#'
#' @description A simple interface to the \pkg{xlsx} package, simplifying the task of sending a data frame or a graph to Excel. 
#'
#' @details
#'
#' This requires the \pkg{xlxs} package and the Java run time environment. Which means that it will only run with the version of Java that you have installed and running. You may need to install with one of:
#'
#' \itemize{
#' \item \code{install.packages(file.choose(), INSTALL_opts="--no-multiarch")}
#' \item \code{install(path, args = "--no-multiarch")}
#' \item \code{install_github("But-I-Digress/xlsxSimple", args = "--no-multiarch")}
#' }
#'
#' @seealso \code{\link{xlsx}}, \code{\link{rJava}} 
#'
#' @section Changes With This Release:
#'
#' \itemize{
#' \item The ``book'' parameter has been renamed ``workbook'' to more closely follow the usage in the \pkg{xlsx} package and has been made optional. When \pkg{xlsxSimple} is loaded a default workbook is created. This workbook is then used wherever the ``workbook'' parameter is omitted. 
#' \item Error trapping has been added for the \code{toSheet} method. If there is no defined method for the object that you are trying to send to a sheet then an error is now thrown. \emph{Id est}, \code{toSheet(NULL)}.
#'	\item \code{saveWorkbook} no longer requires ``file.path'', ``title'' or ``subject'' parameters. If ``file.path'' is omitted then the workbook will be saved with the same path as the script but with an XLSX extension. 
#'}
#'
#' @examples
#' \dontrun{
#' createWorkbook() -> book
#' toSheet(book, mtcars, "mtcars") -> sheet
#' addHyperLinks(sheet, "Google", "http://google.com"); 
#' saveWokbook(book, start = TRUE)
#'
#' library(ggplot2)
#' (mtcars %>% ggplot(aes(mpg, cyl)) + geom_point()) %>% toSheet
#' saveWokbook(start = TRUE)
#' }
#' 
#' @references \url{https://poi.apache.org/apidocs/org/apache/poi/}
#'
"_PACKAGE"
#> [1] "_PACKAGE"