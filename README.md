# vba_libs
 VBA Library to work with vba projects. It contains various utility modules, classes and user forms.

<h3> Available Modules are: </h3>
<ol>
	<li>ExcelUtil.bas</li>
	<li>DateUtil.bas</li>
	<li>FileUtil.bas</li>
</ol>
<hr>
<h3>Modules : </h3>
<h4>ExcelUtil.bas</h4>
<h5>Functions</h5>		
<h6>toColName</h6><h7>(columnNumber as Integer) as String</h7>
		<p>It returns the alphabetical column name of the corresponding integral column number</p>
		<p>columnNumber : Integer = The integral column number</p>
		<p>Returns : String = The alphabetical column name of the corresponding integral column number</p>
	
		<h6>getExcelLink1</h6><h7>(fso As Object, excelFileName As String, sheetName As String, cellRange As String) As String</h7>
		<p>It returns the excel link of given workbook, sheetName and cellRange</p>
		<p>fso : FileSystemObject = object of the FileSystemObject</p>
		<p>excelFileName : String = Full file name of excel workbook</p>
		<p>sheetName : String = Name of the worksheet</p>
		<p>cellRange : String = Address of the cell</p>
		<p>Returns String : the link to the cellRange of the given excelFileName and sheetName</p>
