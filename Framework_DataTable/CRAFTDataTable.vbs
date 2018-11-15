'General Header
'#######################################################################################################################
'Script Description		: Excel data manipulation for CRAFT framework
'Test Tool/Version		: HP Quick Test Professional 10+
'Test Tool Settings		: N.A.
'Application Automated		: N.A.
'Author				: Cognizant
'Date Created			: 04/07/2011
'#######################################################################################################################
Option Explicit	'Forcing Variable declarations

Dim gobjDataTable: Set gobjDataTable = New CraftDataTable

'#######################################################################################################################
'Class Description   	: Class to encapsulate the datatable handling functions
'Author			: Cognizant
'Date Created		: 23/07/2012
'#######################################################################################################################
Class CraftDataTable	
	Private m_strDataTablePath, m_strCommonDataTablePath, m_strDataTableName
	Private m_strDataReferenceIdentifier
	Private m_strCurrentTestcase
	Private m_intCurrentIteration, m_intCurrentSubIteration
	
	
	'###################################################################################################################
	Public Property Get DataTablePath()
		DataTablePath = m_strDataTablePath
	End Property
	
	Public Property Let DataTablePath(strDataTablePath)
		m_strDataTablePath = strDataTablePath
	End Property
	'###################################################################################################################
	
	'###################################################################################################################
	Public Property Get CommonDataTablePath()
		CommonDataTablePath = m_strCommonDataTablePath
	End Property
	
	Public Property Let CommonDataTablePath(strCommonDataTablePath)
		m_strCommonDataTablePath = strCommonDataTablePath
	End Property
	'###################################################################################################################
	
	'###################################################################################################################
	Public Property Get DataTableName()
		DataTableName = m_strDataTableName
	End Property
	
	Public Property Let DataTableName(strDataTableName)
		m_strDataTableName = strDataTableName
	End Property
	'###################################################################################################################
	
	'###################################################################################################################
	Public Property Let DataReferenceIdentifier(strDataReferenceIdentifier)	
		m_strDataReferenceIdentifier = strDataReferenceIdentifier	
	End Property
	'###################################################################################################################
	
	
	'###################################################################################################################
   	Private Sub Class_Initialize()
		m_intCurrentIteration = 0
		m_intCurrentSubIteration = 0
		m_strDataReferenceIdentifier = "#"
	End Sub
	'###################################################################################################################
	
	'###################################################################################################################
	'Function Description   : Function to set the variables required to uniquely identify the exact row of data under consideration
	'Input Parameters	: strCurrentTestcase, intCurrentIteration, intCurrentSubIteration
	'Return Value    	: None
	'Author			: Cognizant
	'Date Created		: 11/10/2012
	'###################################################################################################################
	Public Sub SetCurrentRow(strCurrentTestcase, intCurrentIteration, intCurrentSubIteration)
		m_strCurrentTestCase = strCurrentTestcase
		m_intCurrentIteration = intCurrentIteration
		m_intCurrentSubIteration = intCurrentSubIteration
	End Sub
	'###################################################################################################################
	
	'###################################################################################################################
	'Function Description   : Function to return the test data value corresponding to the field name passed
	'Input Parameters	: strTestDataSheet, strFieldName
	'Return Value    	: strDataValue
	'Author			: Cognizant
	'Date Created		: 11/10/2012
	'###################################################################################################################
	Public Function GetData(strTestDataSheet, strFieldName)
		CheckPreRequisites()
		gobjExcelDataAccess.DatabasePath = m_strDataTablePath
		gobjExcelDataAccess.DatabaseName = m_strDataTableName
		gobjExcelDataAccess.Connect()
		
		Dim strQuery, objTestData
		Set objTestData = CreateObject("ADODB.Recordset")
		strQuery = "Select " & strFieldName & " from [" & strTestDataSheet & "$]" &_
												" where TC_ID = '" & m_strCurrentTestCase &_
												"' and Iteration = '" & m_intCurrentIteration &_
												"' and SubIteration = '" & m_intCurrentSubIteration & "'"
		Set objTestData = gobjExcelDataAccess.ExecuteQuery(strQuery)
		gobjExcelDataAccess.Disconnect()
		
		If objTestData.RecordCount = 0 Then
			Err.Raise 3004, "DataTable Library", "No test data found for the current row: " &_
													"TC_ID = " & m_strCurrentTestCase & ", " &_
													"Iteration = " & m_intCurrentIteration & ", " &_
													"SubIteration = " & m_intCurrentSubIteration
		End If
		
		Dim strDataValue, strFirstChar
		strDataValue = Trim(objTestData(0).Value)
		strFirstChar = Left(strDataValue, 1)
		
		'Release all objects
		objTestData.Close
		Set objTestData = Nothing
		
		If strFirstChar = m_strDataReferenceIdentifier Then
			strDataValue = GetCommonData(strFieldName, strDataValue)
		End If
		
		'Avoid returning Null value
		If IsNull(strDataValue) Then
			strDataValue = ""
		End If
		GetData = strDataValue
	End Function
	'###################################################################################################################
	
	'###################################################################################################################
	Private Sub CheckPreRequisites()
	   If m_strCurrentTestCase = "" Then
		   Err.Raise 3001, "DataTable Library", "Datatable class: Current TestCase is not set"
	   End If
	   If m_intCurrentIteration = 0 Then
		   Err.Raise 3002, "DataTable Library", "Datatable class: Current Iteration is not set"
	   End If
	   If m_intCurrentSubIteration = 0 Then
		   Err.Raise 3003, "DataTable Library", "Datatable class: Current SubIteration is not set"
	   End If
	End Sub
	'###################################################################################################################
	
	'###################################################################################################################
	Private Function GetCommonData(strFieldName, strDataValue)
		gobjExcelDataAccess.DatabasePath = m_strCommonDataTablePath
		gobjExcelDataAccess.DatabaseName = "Common Testdata"
		gobjExcelDataAccess.Connect()
		
		Dim strQuery, objTestData
		Set objTestData = CreateObject("ADODB.Recordset")
		strDataValue = Split(strDataValue, m_strDataReferenceIdentifier)(1)
		strQuery = "Select " & strFieldName & " from [Common_Testdata$] where TD_ID = '" & strDataValue & "'"
		Set objTestData = gobjExcelDataAccess.ExecuteQuery(strQuery)
		gobjExcelDataAccess.Disconnect()
		
		If objTestData.RecordCount = 0 Then
			Err.Raise 3005, "DataTable Library", "No common test data found for the current row: TD_ID = " & strDataValue
		End If
		strDataValue = Trim(objTestData(0).Value)
		
		'Release all objects
		objTestData.Close
		Set objTestData = Nothing
		
		'Avoid returning Null value
		If IsNull(strDataValue) Then
			strDataValue = ""
		End If
		GetCommonData = strDataValue
	End Function
	'###################################################################################################################
	
	'###################################################################################################################
	'Function Description   : Function to output intermediate data (output values)  into the Test data sheet
	'Input Parameters	: strTestDataSheet, strFieldName, strDataValue
	'Return Value    	: None
	'Author			: Cognizant
	'Date Created		: 11/10/2012
	'###################################################################################################################
	Public Sub PutData(strTestDataSheet, strFieldName, strDataValue)
		CheckPreRequisites()
		set excelobject=createobject("excel.application")
		'print m_strDataTablePath
		'Set myexcel=excelobject.workbooks.open("C:\Trackwise_Factory_ExtensionRequest\Datatables\Scenario1.xls")
		Set myexcel=excelobject.workbooks.open(m_strDataTablePath&"\Scenario1.xls")
		Set mysheet=myexcel.Worksheets("General_Data")
		rowcount=mysheet.usedrange.rows.count
		colcount=mysheet.usedrange.columns.count
		for i=1 to rowcount
		
			If mysheet.cells(i,1).value=m_strCurrentTestCase then
				row=i
				exit for
			end if
		next
		for j=1 to colcount
		
			If mysheet.cells(1,j).value=strFieldName then
				col=j
				exit for
			end if
		next
		for i=row to rowcount
			mysheet.cells(i,col)=strDataValue
		next
		
	
		excelobject.DisplayAlerts=False
		excelobject.ActiveWorkbook.save
		myexcel.close
		excelobject.quit
		Set excelobject=nothing
		Set myexcel=nothing
		Set mysheet=nothing
'		gobjExcelDataAccess.DatabasePath = m_strDataTablePath
'		gobjExcelDataAccess.DatabaseName = m_strDataTableName
'		gobjExcelDataAccess.Connect()
'		
'		Dim strNonQuery, objTestData
'		Set objTestData = CreateObject("ADODB.Recordset")
'		strNonQuery = "Update [" & strTestDataSheet & "$] Set " & strFieldName & " = '" & strDataValue & "'" &_
'																" where TC_ID = '" & m_strCurrentTestCase &_
'																"' and Iteration = '" & m_intCurrentIteration &_
'																"' and SubIteration = '" & m_intCurrentSubIteration & "'"
'		gobjExcelDataAccess.ExecuteNonQuery(strNonQuery)
'		gobjExcelDataAccess.Disconnect()
		
		'Report the output value to the results	
		'gobjReport.UpdateTestLog "Output value", _
		'						"Output value " & strDataValue & " written into the " & strFieldName & " column", "Done"
	End Sub
	'###################################################################################################################
	
	'###################################################################################################################
	'Function Description   : Function to return the expected result data (from the Parameterized Checkpoints sheet) corresponding to the field name passed
	'Input Parameters	: strFieldName
	'Return Value    	: strDataValue
	'Author			: Cognizant
	'Date Created		: 11/10/2012
	'###################################################################################################################
	Public Function GetExpectedResult(strFieldName)
		CheckPreRequisites()
		
		gobjExcelDataAccess.DatabasePath = m_strDataTablePath
		gobjExcelDataAccess.DatabaseName = m_strDataTableName
		gobjExcelDataAccess.Connect()
		
		Dim strQuery, objTestData, strCheckPointSheet
		Set objTestData = CreateObject("ADODB.Recordset")
		strCheckPointSheet = "Parametrized_Checkpoints"
		strQuery = "Select " & strFieldName & " from [" & strCheckPointSheet & "$]" &_
												" where TC_ID = '" & m_strCurrentTestCase &_
												"' and Iteration = '" & m_intCurrentIteration &_
												"' and SubIteration = '" & m_intCurrentSubIteration & "'"
		Set objTestData = gobjExcelDataAccess.ExecuteQuery(strQuery)
		gobjExcelDataAccess.Disconnect()
		
		If objTestData.RecordCount = 0 Then
			Err.Raise 3006, "DataTable Library", "No expected results found for the current row: " &_
														"TC_ID = " & m_strCurrentTestCase & ", " &_
														"Iteration = " & m_intCurrentIteration & ", " &_
														"SubIteration = " & m_intCurrentSubIteration
		End If
		
		Dim strDataValue
		strDataValue = Trim(objTestData(0).Value)
		
		'Release all objects
		objTestData.Close
		objConn.Close
		Set objConn = Nothing
		Set objTestData = Nothing
		
		'Avoid returning Null value
		If IsNull(strDataValue) Then
			strDataValue = ""
		End If
		GetExpectedResult = strDataValue
	End Function
	'###################################################################################################################
	
End Class
'#######################################################################################################################
