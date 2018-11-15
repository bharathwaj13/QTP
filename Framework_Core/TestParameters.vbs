'General Header
'#####################################################################################################################
'Script Description		: Test Parameters Library
'Test Tool/Version		: HP Quick Test Professional 10+
'Test Tool Settings		: N.A.
'Application Automated		: Flight Application
'Author				: Cognizant
'Date Created			: 04/07/2011
'#######################################################################################################################
Option Explicit	'Forcing Variable declarations

Dim gobjTestParameters: Set gobjTestParameters = New TestParameters

'#######################################################################################################################
'Class Description   	: Class to get/set TestParameters
'Author			: Cognizant
'Date Created		: 23/07/2012
'#######################################################################################################################
Class TestParameters
	Private m_strCurrentScenario
	Private m_strCurrentTestcase
	Private m_strCurrentTestDescription
   	Private m_strIterationMode
	Private m_intStartIteration
   	Private m_intEndIteration
	
	'###################################################################################################################
	Public Property Get CurrentScenario
		CurrentScenario = m_strCurrentScenario
	End Property
	
	Public Property Let CurrentScenario(strCurrentScenario)
		m_strCurrentScenario = strCurrentScenario
	End Property
	'###################################################################################################################
	
	'###################################################################################################################
	Public Property Get CurrentTestcase
		CurrentTestcase = m_strCurrentTestcase
	End Property
	
	Public Property Let CurrentTestcase(strCurrentTestcase)
		m_strCurrentTestcase = strCurrentTestcase
	End Property
	'###################################################################################################################
	
	'###################################################################################################################
	Public Property Get CurrentTestDescription
		CurrentTestDescription = m_strCurrentTestDescription
	End Property
	
	Public Property Let CurrentTestDescription(strCurrentTestDescription)
		m_strCurrentTestDescription = strCurrentTestDescription
	End Property
	'###################################################################################################################
	
	'###################################################################################################################
	Public Property Get IterationMode
		IterationMode = m_strIterationMode
	End Property
	
	Public Property Let IterationMode(strIterationMode)
		m_strIterationMode = strIterationMode
	End Property
	'###################################################################################################################
	
	'###################################################################################################################
	Public Property Get StartIteration
		StartIteration = m_intStartIteration
	End Property
	
	Public Property Let StartIteration(intStartIteration)
		m_intStartIteration = intStartIteration
	End Property
	'###################################################################################################################
	
	'###################################################################################################################
	Public Property Get EndIteration
		EndIteration = m_intEndIteration
	End Property
	
	Public Property Let EndIteration(intEndIteration)
		m_intEndIteration = intEndIteration
	End Property
	'###################################################################################################################
	
End Class
'#######################################################################################################################
