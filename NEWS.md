# xlsxSImple 2.2.0

* `toExcel()` added. This is a convenience function can be used in place of calling `toSheet()` 
  and then `saveWorkbook()` when you want to look at one data frame or one image. 
  
# xlsxSimple 2.1.0

* The "file.path" parameter has been changed to "filePath" to follow the naming convention 
used elsewhere. The xlsx package uses camel case.  

* An error in the documentation was corrected. The generic for `toSheet()` was incorrectly 
showing a "title" parameter where it should have said "sheetName"

# xlsxSimple 2.0.0

* The “book” parameter has been renamed “workbook” to more closely follow the usage in the
  xlsx package and has been made optional. When xlsxSimple is loaded a default workbook is
  created. This workbook is then used wherever the “workbook” parameter is omitted.
  
* Error trapping has been added for the `toSheet` method. If there is no defined method for the
  object that you are trying to send to a sheet then an error is now thrown. Id est, `toSheet(NULL)`.
  
* `saveWorkbook()` no longer requires “file.path”, “title” or “subject” parameters. If “file.path” is
  omitted then the workbook will be saved with the same path as the script but with an XLSX
  extension.