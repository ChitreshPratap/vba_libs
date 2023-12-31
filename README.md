# vba_libs
 VBA Library to work with vba projects. It contains various utility modules, classes and user forms.

## Available Modules are:
1. [ExcelUtil.bas](#excelutilbas)
1. [DateUtil.bas](#dateutilbas)
1. [FileUtil.bas](#fileutilbas)
<hr>

## ExcelUtil.bas

1. **function toColName (columnNumber as Integer) as String**
	```
	It returns the alphabetical column name of the corresponding integral column number
	Integer columnNumber : The integral column number.
	Returns String : The alphabetical column name of the corresponding integral column number.
	```

1. **Function getExcelLink1(fso As Object, excelFileName As String, sheetName As String, cellRange As String) As String**
	```
	It returns the excel link of given workbook, sheetName and cellRange. It does not open workbook file.
	fso : FileSystemObject : object of the FileSystemObject.
	String excelFileName : Full file name of excel workbook
	String sheetName : Name of the worksheet
	cellRange : String = Address of the cell
	Returns String : the link to the cellRange of the given excelFileName and sheetName
	```

1. **Function getExcelLink2(wb As Workbook, sheetName As String, cellRange As String) As String**
	```
	It returns the excel link of given workbook, sheetName and cellRange. Workbook must be open.
	Workbook wb : Workbook object to create a link with.
	String sheetName : Name of the worksheet
	cellRange : String = Address of the cell
	Returns String : the link to the cellRange of the given excelFileName and sheetName
	```

1. **Function getExcelLink3(cellRange As Range) As String**
	```
	It returns the excel link of cellRange.
	Range cellRange : range object to create a link with.
	Returns String : the link to the cellRange
	```

1. **Function worksheetExists(wb As Workbook, sheetName As String, ifNotExistsRaiseError As Boolean) As Boolean**
	```
	It check whether provided Sheet name exists inside given workbook or not.
	Workbook wb : The workbook in which to check sheet existance
 	String sheetName : Name of the worksheet
 	Boolean ifNotExistsRaiseError : If it is True, Then Error is raised stating worksheet not found in given workbook, if sheet is found then True.
 	If it is False, Then returns False if sheet is not found and True if sheet is found.
 
 	Returns Boolean : If ifNotExistsRaiseError is True, Then Error is raised stating worksheet not found in given workbook
	If it is True, Then return Boolean (True/False) whether sheet exists or not.
	```

## DateUtil.bas

1. **Function getLastDateOfQuarter(iYear As Integer, iQuarterNumber As Integer) As Date**
	```
	It returns last date of the given year and quarter number
 	Integer iYear : Year number
	Integer iQuarterNumber : Quarter Number
 	Returns Date : It returns the last date of the given quarter number and year
	```
1. **Function getQuarterNumber(iDate As Date) As Integer**
	```
	It returns quarter number of the given date
 	Date iDate : Date to find quarter number
 	Returns Integer : It returns the quarter number of the provided date
	```
1. **Function getFormattedString(fDate As Date, stringToFormat As String) As String**
	```
	It returns the formatted string of the given date and date formatted string. It put the Date/Time parts of given Date/Time in the formatted string
 	Date parts symbols must be enclosed inside % %. Example: "I was born in year %YYYY%".
 	Date fDate : Date to formatted string
 	String stringToFormat : Formatted String with date parts enclosed inside %%. Date parts symbols must be enclosed inside % %. Example: "I was born in year %YYYY%".
 	Returns String : It returns the formatted string with the resulted date/time value inside the string
	```

## FileUtil.bas
	Required Dependency : Library FileSystemObject
1. **Sub createFolderPath(pathS As String)**
	```
	It creates the given path if given path not exists. If provided path exists then do nothing. 
 	String pathS : The path to create.
 	Returns : Nothing
	```

1. **Function getFullFilePathByPattern(fullFilePathPattern As String, Optional ifNotExistsRaiseError As Boolean = True) As String**
	```
	It returns the complete path of the provided file path pattern. The pattern is allowed in the file name only.
	In the folder path patterns are not allowed. The folder in fullFilePathPattern must be without pattern.

	String fullFilePathPattern : The file name pattern to get the full file path. The folder name must not include patterns otherwise : Error : Bad file name
	Boolean ifNotExistsRaiseError :
	Returns String : It returns the complete possible existing path of the given.
	```













 












 
