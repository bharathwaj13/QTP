'General Header
'#######################################################################################################################
'Script Description		: Library to generate different types of reports
'Test Tool/Version		: HP Quick Test Professional 10+
'Test Tool Settings		: N.A.
'Application Automated		: Flight Application
'Author				: Cognizant
'Date Created			: 04/07/2011
'#######################################################################################################################
Option Explicit	'Forcing Variable declarations

Dim gobjReport: Set gobjReport = New Report
Dim gobjReportSettings : Set gobjReportSettings = New ReportSettings
Dim strRelativePath
Environment.Value("TestFailureCheck")="Passed"
'Dim gobjReportTheme: Set gobjReportTheme = New ReportTheme

'#######################################################################################################################
'Class Description   	: Class to handle Reporting
'Author			: Cognizant
'Date Created		: 23/07/2012
'#######################################################################################################################
Class Report
	Private m_intStepNumber
	Private m_intStepsPassed
	Private m_intStepsFailed
	Private m_intTestsPassed
	Private m_intTestsFailed
	Private m_objReportTypes
	
	'###################################################################################################################
	Private Sub Class_Initialize()
		m_intStepNumber = 1
		m_intStepsPassed = 0
		m_intStepsFailed = 0
		m_intTestsPassed = 0
		m_intTestsFailed = 0
	End Sub
	'###################################################################################################################
	
	
	'###################################################################################################################
	'Function Description     	: Function to initialize report
	'Input Parameters	 	: None
	'Return Value    		: None
	'Author				: Cognizant	
	'Date Created			: 11/10/2012
	'###################################################################################################################
'	Public Function InitializeReport()
'		ValidateReportPath()
'		InitializeReportTypes()
'		InitializeReportTheme()
'	End Function
'###################################################################################################################
	
	'###################################################################################################################
	'Function Description   	: Function to report any event related to the current test case
	'Input Parameters 		: strStepName, strStepDescription, strStepStatus
	'Return Value    		: None
	'Author				: HCL
	'Date Created			: 31/07/2017
	'###################################################################################################################
Sub UpdateTestLog(strStepName,strStepDescription,strStepStatus,strStepExpected,strStepActual)
Class_Initialize()
'gobjReportSettings.Class_Initialize()
Dim intStatus,objFso,strScreenshotPath,strCurrentTime,strScreenshotName
intStatus = GetLogLevel(strStepStatus)
If (intStatus < gobjReportSettings.LogLevel) Then
	'Update the QTP results
	'Msgbox "BEFORE QTP"
'	Reporter.ReportEvent GetQtpStatus(strStepStatus), strStepName, strStepDescription
	
	Set objFso = CreateObject("Scripting.FileSystemObject")
	strRelativePath = objFso.GetParentFolderName(PathFinder.Locate("Datatables"))
	strScreenshotPath=strRelativePath&"\"&Environment("TemporaryReportScreenShotFolder")
	If objFso.FolderExists(strScreenshotPath)=false Then
		objFso.CreateFolder(strScreenshotPath)
	End If
	strCurrentTime = Now()
	strScreenshotName = Environment("TestName") & "_" &Replace(Replace(Replace(strCurrentTime, " ", "_"), ":", "-"), "/","-") &".png"
	
	strScreenshotPath=strScreenshotPath&"\"&strScreenshotName
	
	If((strStepStatus = "Failed" Or strStepStatus = "Warning") _
	And gobjReportSettings.TakeScreenshotFailedStep) _
	And objFso.FileExists(strScreenshotPath) = False Then	'check if another screenshot was taken already in the very same second
		Desktop.CaptureBitmap(strScreenshotPath)
	End If
	
	If((strStepStatus = "Passed" Or strStepStatus = "Done") _
	And gobjReportSettings.TakeScreenshotPassedStep) _
	And objFso.FileExists(strScreenshotPath) = False Then	'check if another screenshot was taken already in the very same second
		Desktop.CaptureBitmap(strScreenshotPath)
	End If
	
	If(strStepStatus = "Screenshot") _
	And objFso.FileExists(strScreenshotPath) = False Then	'check if another screenshot was taken already in the very same second
		Desktop.CaptureBitmap(strScreenshotPath)
	End If
	
	Set objFso = Nothing
	'Reporter.ReportEvent GetQtpStatus(strStepStatus), strStepName, strStepDescription			
	Set objTeststep = QCUTIL.CurrentRun.StepFactory.AddItem(Null)
	objTeststep.Name = strStepName
	objTeststep.Field("ST_DESCRIPTION") = strStepDescription
	objTeststep.Field("ST_EXPECTED") = strStepExpected
	objTeststep.Field("ST_ACTUAL") = strStepActual
	objTeststep.Field("ST_STATUS") = strStepStatus
	'objTeststep.Status = strStepStatus
	objTeststep.Post()
	Set objAttachment=objTeststep.Attachments
	Set objAttachItem = objAttachment.AddItem(Null)
	objAttachItem.FileName = strScreenshotPath
	objAttachItem.Type = 1
	objAttachItem.Post()
	If strStepStatus="Failed" Then
		Environment.Value("TestFailureCheck")="Failed"
	End If
	Set objTeststep=Nothing
	Set objAttachment=Nothing
	Set objAttachItem=Nothing
	intStatus=null
	strScreenshotPath=null
	strCurrentTime=null
	strScreenshotName=null
End If
End Sub

Sub ALM_AfterAttachIntoALM_DeleteLocalScreenShotFiles()
Set objFile=CreateObject("Scripting.FileSystemObject")
Dim strScreenshotPath:strScreenshotPath=strRelativePath&"\"&Environment("TemporaryReportScreenShotFolder")
Set objFold=objFile.GetFolder(strScreenshotPath)
For each objItem in objFold.Files
  objFile.DeleteFile objItem.Path,true
Next
Set objFile=Nothing
Set objFold=Nothing
strScreenshotPath=null
End Sub
				
	'###################################################################################################################
	
	'###################################################################################################################
	Private Function GetLogLevel(strStepStatus)
		Dim intStatus
		Select Case strStepStatus
			Case "Failed"
				intStatus = 0
			Case "Warning"
				intStatus = 1
			Case "Passed"
				intStatus = 2
			Case "Screenshot"
				intStatus = 3
			Case "Done"
				intStatus = 4
			Case "Debug"
				intStatus = 5
		End Select
		
		GetLogLevel = intStatus	
	End Function
	'###################################################################################################################
	
	'###################################################################################################################
	Private Function GetQtpStatus(strStepStatus)
		Dim intQtpStatus
		Select Case strStepStatus
			Case "Pass"
				intQtpStatus = micPass
			Case "Fail"
				intQtpStatus = micFail
			Case "Done"
				intQtpStatus = micDone
			Case "Warning"
				intQtpStatus = micWarning
			Case "Screenshot"
				intQtpStatus = micDone
		End Select
		
		GetQtpStatus = intQtpStatus
	End Function
	
	'###################################################################################################################
	
End Class
'#######################################################################################################################
'#######################################################################################################################


'#######################################################################################################################
'Class Description   		: Class to get/set Report settings
'Author				: Cognizant
'Date Created			: 23/07/2012
'#######################################################################################################################
Class ReportSettings
	Private m_strReportPath, m_strReportName
	Private m_strReportTheme
	Private m_strProjectName
	Private m_intLogLevel
	Private m_blnExcelReport, m_blnHtmlReport
	Private m_blnTakeScreenshotPassedStep, m_blnTakeScreenshotFailedStep
	Private m_blnLinkScreenshotsToTestLog
	Private m_blnLinkTestLogsToSummary
	
	'###################################################################################################################

	Public Property Get LogLevel
		LogLevel = m_intLogLevel
	End Property
	
	Public Property Let LogLevel(intLogLevel)
		If intLogLevel < 0 Then
			intLogLevel = 0
		ElseIF intLogLevel > 5 Then
			intLogLevel = 5
		End If
		m_intLogLevel = intLogLevel
	End Property
	'###################################################################################################################
	
	'###################################################################################################################
	Public Property Get TakeScreenshotPassedStep
		TakeScreenshotPassedStep = m_blnTakeScreenshotPassedStep
	End Property
	
	Public Property Let TakeScreenshotPassedStep(blnTakeScreenshotPassedStep)
		m_blnTakeScreenshotPassedStep = blnTakeScreenshotPassedStep
	End Property
	'###################################################################################################################
	
	'###################################################################################################################
	Public Property Get TakeScreenshotFailedStep
		TakeScreenshotFailedStep = m_blnTakeScreenshotFailedStep
	End Property
	
	Public Property Let TakeScreenshotFailedStep(blnTakeScreenshotFailedStep)
		m_blnTakeScreenshotFailedStep = blnTakeScreenshotFailedStep
	End Property
	'###################################################################################################################
	
	'###################################################################################################################
	Public Property Get LinkScreenshotsToTestLog()
		LinkScreenshotsToTestLog = m_blnLinkScreenshotsToTestLog
	End Property
	
	Public Property Let LinkScreenshotsToTestLog(blnLinkScreenshotsToTestLog)
		m_blnLinkScreenshotsToTestLog = blnLinkScreenshotsToTestLog
	End Property
	'###################################################################################################################
	
	'###################################################################################################################
	Public Property Get LinkTestLogsToSummary()
		LinkTestLogsToSummary = m_blnLinkTestLogsToSummary
	End Property
	
	Public Property Let LinkTestLogsToSummary(blnLinkTestLogsToSummary)
		m_blnLinkTestLogsToSummary = blnLinkTestLogsToSummary
	End Property
	'###################################################################################################################
	
	
	'###################################################################################################################
	Private Sub Class_Initialize()
		m_strProjectName = ""
		m_intLogLevel = 4
		m_blnExcelReport = False
		m_blnHtmlReport = True
		m_blnTakeScreenshotFailedStep = True
		m_blnTakeScreenshotPassedStep = False
		m_blnLinkScreenshotsToTestLog = True
		m_blnLinkTestLogsToSummary = True
	End Sub
	'###################################################################################################################
	
End Class
'#######################################################################################################################
