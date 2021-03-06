% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/toSheet.R
\name{toSheet}
\alias{toSheet}
\alias{toSheet.data.frame}
\alias{toSheet.tbl_df}
\alias{toSheet.tbl}
\alias{toSheet.matrix}
\alias{toSheet.character}
\alias{toSheet.ggplot}
\title{Send something to a sheet}
\usage{
toSheet(x, workbook, sheetName, ...)

\method{toSheet}{data.frame}(d, workbook = getWorkbook(),
  sheetName = getSheetName(workbook), auto.filter = TRUE,
  freeze.pane = TRUE, row.names = FALSE, startRow = 1, ...)

\method{toSheet}{tbl_df}(d, workbook = getWorkbook(),
  sheetName = getSheetName(workbook), ...)

\method{toSheet}{tbl}(d, workbook = getWorkbook(),
  sheetName = getSheetName(workbook), ...)

\method{toSheet}{matrix}(d, workbook = getWorkbook(),
  sheetName = getSheetName(workbook), row.names = TRUE, ...)

\method{toSheet}{character}(x, workbook = getWorkbook(),
  sheetName = getSheetName(workbook), landscape = TRUE)

\method{toSheet}{ggplot}(p, workbook = getWorkbook(),
  sheetName = getSheetName(workbook), filename = tempfile(fileext =
  ".jpeg"), landscape = TRUE, height = if (landscape) 7 else 9.7,
  width = if (landscape) 8.5 else 5.99, ...)
}
\arguments{
\item{x}{A data frame, plot or path to a saved plot.}

\item{workbook}{An optional xlsx work book.}

\item{sheetName}{Character, an optional title for the new sheet.}

\item{...}{Other parameters to pass to addDataFrame.}

\item{auto.filter}{A boolean, should the columns be filtered.}

\item{freeze.pane}{A boolean, should the top row be frozen.}

\item{row.names}{A boolean, should therow names of x be written along with x to the file.}

\item{startRow}{A scalar integer, the row to start the data frame.}

\item{landscape}{A boolean, should the sheet be oriented horizontally.}

\item{filename}{Optional file path and name for the temporary file for the image of the plot.}

\item{height}{Numeric, optional height of the image in inches.}

\item{width}{Numeric, optional width of the image in inches.}
}
\value{
A worksheet object, invisibly.
}
\description{
\code{toSheet} sends tabular data or an image to a new xlsx work sheet.
}
\section{Methods (by class)}{
\itemize{
\item \code{data.frame}: Send a data frame to a work sheet.

\item \code{tbl_df}: Send a tibble to a work sheet.

\item \code{tbl}: Send a data frame to a work sheet.

\item \code{matrix}: Send a matrix to a work sheet.

\item \code{character}: send a saved image to a work sheet.

\item \code{ggplot}: Send a ggplot object to a work sheet.
}}

\examples{
\dontrun{
book <- createWorkbook()
toSheet(mtcars, book, "mtcars")

library(ggplot2)
p <- ggplot(mtcars, aes(mpg))
p <- p + geom_bar()
toSheet(p)
}
}
