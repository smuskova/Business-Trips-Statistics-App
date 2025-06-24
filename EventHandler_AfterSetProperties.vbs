Const ForReading = 1, ForWriting = 2, ForAppending = 8

LogFolderPath = "D:\M-Files\Temp\log\" & Vault.Name & "\"
BuildFullPath LogFolderPath

Set oPropVals = CreateObject("MFilesApi.PropertyValues")
Set oPropVals = Vault.ObjectPropertyOperations.GetProperties(ObjVer)
LogName = oPropVals.SearchForProperty(100).TypedValue.GetValueAsLookup.DisplayValue
LogState = "EH - Set EH Approved Employee"

' Main Class properties
iPDComplete = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.Complete") ' Boolean
iPDEscalateBusinessTrip = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EscalateBusinessTrip") ' Boolean
iPDEhActionEmployee = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EhActionEmployee") ' Single Lookup

iWFSDraftBusinessTrip = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.DraftBusinessTrip")
iWFSSelectVehicles = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.SelectVehicles")
iWFSRejectedBusinessTrip = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.RejectedBusinessTrip")
iWFSDirectManagerApproval = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.DirectManagerApproval")
iWFS1stLevelApproval = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.1stLevelApproval")
iWFS2ndLevelApproval = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.2ndLevelApproval")
iWFS3rdLevelApproval = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.3rdLevelApproval")
iWFS4thLevelApproval = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.4thLevelApproval")
iWFS5thLevelApproval = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.5thLevelApproval")
iWFS6thLevelApproval = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.6thLevelApproval")
iWFS1stInitialOrganization = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.1stInitialOrganization")
iWFS2ndInitialOrganization = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.2ndInitialOrganization")
iWFSTravelArrangements = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.TravelArrangements")
iWFSReturnedInitialOrganization = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.ReturnedInitialOrganization")
iWFSAwaitingStartTime = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.AwaitingStartTime")
iWFSInProgress = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.InProgress")
iWFSAwaitingReport = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.AwaitingReport")
iWFS1stFinalOrganization = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.1stFinalOrganization")
iWFS2ndFinalOrganization = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.2ndFinalOrganization")
iWFS3rdFinalOrganization = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.3rdFinalOrganization")
iWFSAccountingProcessing = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.AccountingProcessing")
iWFSReturnedFinalOrganization = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.ReturnedFinalOrganization")
iWFSEscaletedBusinessTrip = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.BusinessTrip.EscaletedBusinessTrip")

' Employee properties
iCLEmployee = Vault.ClassOperations.GetObjectClassIDbyAlias("CL.Employee")
iPDUserAccount = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.UserAccount") ' Single Lookup

On Error Resume Next
	boolPDComplete = oPropVals.SearchForProperty(iPDComplete).TypedValue.Value
	szPDComplete = oPropVals.SearchForProperty(iPDComplete).TypedValue.DisplayValue

	boolPDEscalateBusinessTrip = oPropVals.SearchForProperty(iPDEscalateBusinessTrip).TypedValue.Value
	szPDEscalateBusinessTrip = oPropVals.SearchForProperty(iPDEscalateBusinessTrip).TypedValue.DisplayValue

	intCurrentState = oPropVals.SearchForProperty(39).TypedValue.GetValueAsLookup.Item

	szPDEhActionEmployee= oPropVals.SearchForProperty(iPDEhActionEmployee).TypedValue.DisplayValue
On Error Goto 0 

If (Len(szPDEhActionEmployee) = 0 And Len(szPDComplete) > 0 And _
	( _
	intCurrentState = iWFSDraftBusinessTrip Or _
	intCurrentState = iWFSSelectVehicles Or _
	intCurrentState = iWFSDirectManagerApproval Or _
	intCurrentState = iWFSRejectedBusinessTrip Or _
	intCurrentState = iWFS1stLevelApproval Or _
	intCurrentState = iWFS2ndLevelApproval Or _
	intCurrentState = iWFS3rdLevelApproval Or _
	intCurrentState = iWFS4thLevelApproval Or _
	intCurrentState = iWFS5thLevelApproval Or _
	intCurrentState = iWFS6thLevelApproval Or _
	intCurrentState = iWFS1stInitialOrganization Or _
	intCurrentState = iWFS2ndInitialOrganization Or _
	intCurrentState = iWFSTravelArrangements Or _
	intCurrentState = iWFSReturnedInitialOrganization Or _
	intCurrentState = iWFS1stFinalOrganization Or _
	intCurrentState = iWFS2ndFinalOrganization Or _
	intCurrentState = iWFS3rdFinalOrganization Or _
	intCurrentState = iWFSAccountingProcessing Or _
	intCurrentState = iWFSEscaletedBusinessTrip _
	)) _
	Or (Len(szPDEhActionEmployee) = 0 And boolPDComplete And _
	( _
	intCurrentState = iWFSAwaitingStartTime Or _
	intCurrentState = iWFSInProgress Or _
	intCurrentState = iWFSAwaitingReport Or _
	intCurrentState = iWFSReturnedFinalOrganization _
	)) _
	Or (Len(szPDEhActionEmployee) = 0 And boolPDEscalateBusinessTrip And intCurrentState = iWFSAccountingProcessing) _
	Then
		
	WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; CurrentUserID: " & CurrentUserID
	WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; CurrentUserSessionInfo.AccountName: " & CurrentUserSessionInfo.AccountName

	Set oSCs = CreateObject("MFilesAPI.SearchConditions") 

	Set oSCClass = CreateObject("MFilesAPI.SearchCondition") 
	oSCClass.ConditionType = MFConditionTypeEqual
	oSCClass.Expression.DataPropertyValuePropertyDef = 100
	oSCClass.TypedValue.SetValue MFDatatypeLookup, iCLEmployee
	oSCs.Add -1, oSCClass

	Set oSearchCondition = CreateObject("MFilesAPI.SearchCondition")
	oSearchCondition.ConditionType = MFConditionTypeEqual
	oSearchCondition.Expression.DataPropertyValuePropertyDef = iPDUserAccount
	oSearchCondition.TypedValue.SetValue MFDatatypeLookup, CurrentUserID
	oSCs.Add -1, oSearchCondition

	' Search for just the non deleted element
	Set oSearchNonDeleted = CreateObject("MFilesAPI.SearchCondition") 
	oSearchNonDeleted.ConditionType = MFConditionTypeEqual
	oSearchNonDeleted.Expression.DataStatusValueType = MFStatusTypeDeleted
	oSearchNonDeleted.TypedValue.SetValue MFDatatypeBoolean, False
	oSCs.Add -1, oSearchNonDeleted

	' Execute Search	
	Set oSearchResults = Vault.ObjectSearchOperations.SearchForObjectsByConditionsEx(oSCs,MFSearchFlagNone, False) 

	If oSearchResults.Count = 1 Then
		On Error Resume Next
			For Each oSearchResult In oSearchResults
				oPropVals.SearchForProperty(iPDEhActionEmployee).TypedValue.SetValue MFDatatypeLookup, oSearchResult.ObjVer.ID
			Next
		On Error Goto 0

		Vault.ObjectPropertyOperations.SetAllProperties ObjVer, True, oPropVals
		
	ElseIf oSearchResults.Count = 0 Then
		Err.Raise MFScriptCancel, "Missing Employee with ""User Account"" id " & CurrentUserID & " and ""Account Name"" " & CurrentUserSessionInfo.AccountName & "!"
	Else 
		Err.Raise MFScriptCancel, "More than 1 Employee with ""User Account"" id " & CurrentUserID & " and ""Account Name"" " & CurrentUserSessionInfo.AccountName & "!"
	End If 
End If 

'--------------------------------------------------------
'Functions and Procedures 
'--------------------------------------------------------
Sub BuildFullPath(ByVal FullPath)
	Set oFSO = CreateObject("Scripting.FileSystemObject") 
	If Not oFSO.FolderExists(FullPath) Then
		BuildFullPath oFSO.GetParentFolderName(FullPath)
		oFSO.CreateFolder FullPath
	End If
	Set oFSO = Nothing
End Sub

Function TimeStamp
	TimeStamp = PadDigits(Year(Date),4) & "-" & PadDigits(Month(Date),2) & "-" & PadDigits(Day(Date),2) & " " _
		& PadDigits(Hour(Time),2) & ":" & PadDigits(Minute(Time),2) & ":" & PadDigits(Second(Time),2)
End Function 

Sub WriteLog (LogType, LogName, LogMessage)
	On Error Resume Next 
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		LogFileNameWithPath = LogName & "_" & PadDigits(Year(Date),4) & "-" & PadDigits(Month(Date),2) & "-" & PadDigits(Day(Date),2) & ".log"
		Set LogFile = oFSO.OpenTextFile(LogFileNameWithPath , ForAppending, True)
		LogFile.WriteLine(TimeStamp & "; "& LogType & "; " & LogMessage)
		LogFile.Close
		Set oFSO = Nothing
	On Error Goto 0
End Sub

Function PadDigits(val, digits)
	PadDigits = Right(String(digits,"0") & val, digits)
End Function