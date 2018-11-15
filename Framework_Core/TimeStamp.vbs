'General Header
'#######################################################################################################################
'Script Description		: Timestamp class for report
'Test Tool/Version		: HP Quick Test Professional 10+
'Test Tool Settings		: N.A.
'Application Automated		: N.A.
'Author				: Cognizant
'Date Created			: 04/07/2011
'#######################################################################################################################
Option Explicit	'Forcing Variable declarations

Dim gobjTimeStamp : Set gobjTimeStamp = New TimeStamp

'#######################################################################################################################
'Class Description   	: Class to encapsulate utility functions of the framework
'Author			: Cognizant
'Date Created		: 23/07/2012
'#######################################################################################################################
Class TimeStamp
	
	Private m_strPath
	
	'###################################################################################################################
	Public Property Get Path
		Path = m_strPath
	End Property
	
	Public Property Let Path(strPath)
		m_strPath = Trim(strPath)
	End Property
	'###################################################################################################################
	
	
	'###################################################################################################################
	'Function Description   	: Function to calculate the execution time for the current iteration	
	'Input Parameters 		: None
	'Return Value    		: None	
	'Author				: Cognizant	
	'Date Created			: 07/11/2012
	'###################################################################################################################
	Public Sub Initialize()
		Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
		
		Dim strRunConfigurationPath
		strRunConfigurationPath = gobjFrameworkParameters.RelativePath & "\Results\" & gobjFrameworkParameters.RunConfiguration
		
		If Not objFso.FolderExists(strRunConfigurationPath) Then
			objFso.CreateFolder(strRunConfigurationPath)
		End If
		
		If m_strPath = "" Then
			m_strPath = "Run" & "_" & Replace(Date(),"/","-") & "_" & Replace(Time(),":","-")
		End If
		
		Dim strReportPathWithTimeStamp
		strReportPathWithTimeStamp = strRunConfigurationPath & "\" & m_strPath
		
		If Not objFso.FolderExists(strReportPathWithTimeStamp) Then
			objFso.CreateFolder(strReportPathWithTimeStamp)
		End If
		Set objFso = Nothing
	End Sub
	'###################################################################################################################
	
End Class
'#######################################################################################################################
