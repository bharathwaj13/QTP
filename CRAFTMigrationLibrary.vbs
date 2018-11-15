'General Header
'#####################################################################################################################
'Script Description		: CRAFT Migration Library
'Test Tool/Version		: HP Quick Test Professional 10+
'Test Tool Settings		: N.A.
'Application Automated		: Flight Application
'Author				: Cognizant
'Date Created			: 15/04/2012
'#####################################################################################################################
Option Explicit	'Forcing Variable declarations

'#####################################################################################################################
'Function Description   	: Function to return the test data value corresponding to the field name passed
'Input Parameters		: strTestDataSheet, strFieldName
'Return Value    		: strDataValue
'Author				: Cognizant
'Date Created			: 15/04/2012
'#####################################################################################################################
Function CRAFT_GetData(strTestDataSheet, strFieldName)
	CRAFT_GetData = gobjDataTable.GetData(strTestDataSheet, strFieldName)
End Function
'#####################################################################################################################

'#####################################################################################################################
'Function Description   	: Function to report any event related to the current test case
'Input Parameters 		: strStepName, strDescription, strStatus
'Return Value    		: None
'Author				: Cognizant
'Date Created			: 15/04/2012
'#####################################################################################################################
Sub CRAFT_ReportEvent(strStepName, strDescription, strStatus)
	gobjReport.UpdateTestLog strStepName, strDescription, strStatus
End Sub
'#####################################################################################################################
