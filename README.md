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

1. **getExcelLink1(fso As Object, excelFileName As String, sheetName As String, cellRange As String) As String**
	```
	It returns the excel link of given workbook, sheetName and cellRange.
	fso : FileSystemObject : object of the FileSystemObject.
	String excelFileName : Full file name of excel workbook
	String sheetName : Name of the worksheet
	cellRange : String = Address of the cell
	Returns String : the link to the cellRange of the given excelFileName and sheetName
	```

