# vba_libs
 VBA Library to work with vba projects. It contains various utility modules, classes and user forms.

## Available Modules are: </h3>
1. ExcelUtil.bas</li>
1. DateUtil.bas</li>
1. FileUtil.bas</li>
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














 












 
