# xlsxSimple

Sometimes we need to save the results of our analysis as multiple tabs in Excel. Perhaps as a first tab with a pretty graph, the second with a summary table and the third with the raw data. The `xlsxSimple` package provides a simple way to do that. 

When the package is loaded, a default workbook is created. Graphics and tabular data can be added as tabs and then the workbook saved. Alternatively, one or more workbooks can be created explicitly, written to and saved.

## Installation

The `xlsxSimple` package rests atop the `xlsx` package which in turn rests atop the Apache POI project, which is written in Java. So you will need to first install Java. And you will only be able to install the version of `xlsxSimple` that corresponds to your version of Java, 32 or 64 bit.

```r
# To install for both the i386 and x64 versions of R:
devtools::install_github("But-I-Digress/xlsxSimple")

# To install for just the version of R that is running:
devtools::install_github("But-I-Digress/xlsxSimple", args = "--no-multiarch")
```

## Examples

```r
# To create a workbook explicitly and use it:
createWorkbook() -> book
toSheet(book, mtcars, "mtcars") -> sheet
addHyperLinks(sheet, "Google", "http://google.com")
saveWorkbook(book, start = TRUE)

# To use the default workbook:

library(ggplot2)
(mtcars %>% ggplot(aes(mpg, cyl)) + geom_point()) %>% toSheet
saveWokbook(start = TRUE)
```

The example above also shows how to use the `magrittr` pipe `%>%` with the  `ggplot2` package. Parenthesis are required.