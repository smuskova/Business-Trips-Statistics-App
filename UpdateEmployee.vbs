Const ForReading = 1, ForWriting = 2, ForAppending = 8

LogFolderPath = "D:\M-Files\Temp\log\" & Vault.Name & "\"
BuildFullPath LogFolderPath

Set oPropVals = CreateObject("MFilesApi.PropertyValues")
Set oPropVals = Vault.ObjectPropertyOperations.GetProperties(ObjVer)
LogName = oPropVals.SearchForProperty(100).TypedValue.GetValueAsLookup.DisplayValue
LogState = oPropVals.SearchForProperty(39).TypedValue.GetValueAsLookup.DisplayValue

' Employee properties
iCLEmployee = Vault.ClassOperations.GetObjectClassIDbyAlias("CL.Employee")

iUserName = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ExtUserName")
iPDExtDistinguishedName = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ExtDistinguishedName")
iPDExtDirectManagerGuid = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ExtDirectManagerGuid")
iPDExtEmployeeGuid = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ExtEmployeeGuid")
iPDExtDirectManager = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ExtDirectManager")
iPDExtDepartmentName = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ExtDepartmentName")

iPDDirectManager = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.DirectManager") ' Single Lookup
iPDDepartment = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.Department") ' Single Lookup

iPDUserAccount = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.UserAccount")
iPDUserGroups = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.UserGroups")

' Department properties
iOTDepartment = Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.Department")

iCLDepartment = Vault.ClassOperations.GetObjectClassIDbyAlias("CL.Department")

iPDDepartmentEn = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.DepartmentEn")
iPDIsActiveFlag = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.IsActiveFlag")

iWFDepartment = Vault.WorkflowOperations.GetWorkflowIDByAlias("WF.Department")
iWFSDepartmentInit = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.Department.Init")

'---------------------------------------------------------------------------------
'----- Read Employee Information -------------------------------------------------
'---------------------------------------------------------------------------------
szUserName = oPropVals.SearchForProperty(iUserName).Value.DisplayValue
szPDExtDirectManagerGuid = oPropVals.SearchForProperty(iPDExtDirectManagerGuid).Value.DisplayValue
szPDExtDirectManager = oPropVals.SearchForProperty(iPDExtDirectManager).Value.DisplayValue
szPDExtDepartmentName = oPropVals.SearchForProperty(iPDExtDepartmentName).Value.DisplayValue

WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szUserName: " & szUserName
WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDExtDirectManagerGuid: " & szPDExtDirectManagerGuid
WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDExtDirectManager: " & szPDExtDirectManager
WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDExtDepartmentName: " & szPDExtDepartmentName

szUserNameArray = Split(szUserName,"@")
If UBound(szUserNameArray) > 0 Then 
	szUserName = szUserNameArray(1) & "\" & szUserNameArray(0)
End If 
WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szUserName: " & szUserName

'---------------------------------------------------------------------------------
'----- Set User Groups -----------------------------------------------------------
'---------------------------------------------------------------------------------
Set VaultUserAccounts = Vault.UserOperations.GetUserAccounts()

For Each VaultUserAccount In VaultUserAccounts
	If LCase(szUserName) = LCase(VaultUserAccount.LoginName) Then
		oPropVals.SearchForProperty(iPDUserAccount).TypedValue.SetValue MFDataTypeLookup, VaultUserAccount.ID
		Set oUserGroups = Vault.UserGroupOperations.GetGroupsOfUserOrGroup(VaultUserAccount.ID,False)
		Set oLookups = CreateObject("MFilesAPI.Lookups")
		Set oLookup = CreateObject("MFilesAPI.Lookup")
		
		For Each oUserGroup In oUserGroups
			If oUserGroup.ID > 2 Then
				oLookup.Item = oUserGroup.ID
				oLookups.Add -1, oLookup
			End If
		Next

		oPropVals.SearchForProperty(iPDUserGroups).TypedValue.SetValueToMultiSelectLookup oLookups
	End If
Next

'---------------------------------------------------------------------------------
'----- Set Direct Manager --------------------------------------------------------
'---------------------------------------------------------------------------------
' Azure/Entra
If Len(szPDExtDirectManagerGuid) > 0 Then 

	Set oSCs = CreateObject("MFilesAPI.SearchConditions") 

	Set oSCClass = CreateObject("MFilesAPI.SearchCondition") 
	oSCClass.ConditionType = MFConditionTypeEqual
	oSCClass.Expression.DataPropertyValuePropertyDef = 100
	oSCClass.TypedValue.SetValue MFDatatypeLookup, iCLEmployee
	oSCs.Add -1, oSCClass

	Set oSearchCondition = CreateObject("MFilesAPI.SearchCondition")
	oSearchCondition.ConditionType = MFConditionTypeEqual
	oSearchCondition.Expression.DataPropertyValuePropertyDef = iPDExtEmployeeGuid
	oSearchCondition.TypedValue.SetValue MFDatatypeText, szPDExtDirectManagerGuid
	oSCs.Add -1, oSearchCondition

	' Search for just the non deleted element
	Set oSearchNonDeleted = CreateObject("MFilesAPI.SearchCondition") 
	oSearchNonDeleted.ConditionType = MFConditionTypeEqual
	oSearchNonDeleted.Expression.DataStatusValueType = MFStatusTypeDeleted
	oSearchNonDeleted.TypedValue.SetValue MFDatatypeBoolean, False
	oSCs.Add -1, oSearchNonDeleted

	' Execute Search	
	Set oSearchResults = Vault.ObjectSearchOperations.SearchForObjectsByConditions(oSCs,MFSearchFlagNone, False) 

	WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; Found Direct Manager count: " & oSearchResults.Count

	If oSearchResults.Count = 1 Then
		
		oPropVals.SearchForProperty(iPDDirectManager).TypedValue.SetValue MFDataTypeLookup, oSearchResults.Item(1).ObjVer.ID
		
	End If

End If

' MS Active Directory (LDAP)
If Len(szPDExtDirectManager) > 0 Then 

	Set oSCs = CreateObject("MFilesAPI.SearchConditions") 

	Set oSCClass = CreateObject("MFilesAPI.SearchCondition") 
	oSCClass.ConditionType = MFConditionTypeEqual
	oSCClass.Expression.DataPropertyValuePropertyDef = 100
	oSCClass.TypedValue.SetValue MFDatatypeLookup, iCLEmployee
	oSCs.Add -1, oSCClass

	Set oSearchCondition = CreateObject("MFilesAPI.SearchCondition")
	oSearchCondition.ConditionType = MFConditionTypeEqual
	oSearchCondition.Expression.DataPropertyValuePropertyDef = iPDExtDistinguishedName
	oSearchCondition.TypedValue.SetValue MFDatatypeText, szPDExtDirectManager
	oSCs.Add -1, oSearchCondition

	' Search for just the non deleted element
	Set oSearchNonDeleted = CreateObject("MFilesAPI.SearchCondition") 
	oSearchNonDeleted.ConditionType = MFConditionTypeEqual
	oSearchNonDeleted.Expression.DataStatusValueType = MFStatusTypeDeleted
	oSearchNonDeleted.TypedValue.SetValue MFDatatypeBoolean, False
	oSCs.Add -1, oSearchNonDeleted

	' Execute Search	
	Set oSearchResults = Vault.ObjectSearchOperations.SearchForObjectsByConditions(oSCs,MFSearchFlagNone, False) 

	WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; Found Direct Manager count: " & oSearchResults.Count

	If oSearchResults.Count = 1 Then
		
		oPropVals.SearchForProperty(iPDDirectManager).TypedValue.SetValue MFDataTypeLookup, oSearchResults.Item(1).ObjVer.ID
		
	End If

End If

'---------------------------------------------------------------------------------
'----- Set Department ------------------------------------------------------------
'---------------------------------------------------------------------------------
If Len(szPDExtDepartmentName) > 0 Then 

	Set oSCs = CreateObject("MFilesAPI.SearchConditions") 

	Set oSCClass = CreateObject("MFilesAPI.SearchCondition") 
	oSCClass.ConditionType = MFConditionTypeEqual
	oSCClass.Expression.DataPropertyValuePropertyDef = 100
	oSCClass.TypedValue.SetValue MFDatatypeLookup, iCLDepartment
	oSCs.Add -1, oSCClass

	Set oSearchCondition = CreateObject("MFilesAPI.SearchCondition")
	oSearchCondition.ConditionType = MFConditionTypeEqual
	oSearchCondition.Expression.DataPropertyValuePropertyDef = iPDDepartmentEn
	oSearchCondition.TypedValue.SetValue MFDatatypeText, szPDExtDepartmentName
	oSCs.Add -1, oSearchCondition

	' Search for just the non deleted element
	Set oSearchNonDeleted = CreateObject("MFilesAPI.SearchCondition") 
	oSearchNonDeleted.ConditionType = MFConditionTypeEqual
	oSearchNonDeleted.Expression.DataStatusValueType = MFStatusTypeDeleted
	oSearchNonDeleted.TypedValue.SetValue MFDatatypeBoolean, False
	oSCs.Add -1, oSearchNonDeleted

	' Execute Search	
	Set oSearchResults = Vault.ObjectSearchOperations.SearchForObjectsByConditions(oSCs,MFSearchFlagNone, False) 

	WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; Found Department count: " & oSearchResults.Count

	If oSearchResults.Count = 0 Then
		
		Set oPropertyValues = CreateObject("MFilesAPI.PropertyValues")
		Set oPropertyValue = CreateObject("MFilesAPI.PropertyValue")
		Set oFiles = CreateObject("MFilesAPI.SourceObjectFiles")
		Set oACL = CreateObject("MFilesAPI.AccessControlList")

		oPropertyValue.PropertyDef = 100
		oPropertyValue.TypedValue.SetValue MFDatatypeLookup, iCLDepartment
		oPropertyValues.Add -1, oPropertyValue

		oPropertyValue.PropertyDef = 38
		oPropertyValue.TypedValue.SetValue MFDatatypeLookup, iWFDepartment
		oPropertyValues.Add -1, oPropertyValue	

		oPropertyValue.PropertyDef = 39
		oPropertyValue.TypedValue.SetValue MFDatatypeLookup, iWFSDepartmentInit
		oPropertyValues.Add -1, oPropertyValue	
		
		oPropertyValue.PropertyDef = iPDDepartmentEn
		oPropertyValue.TypedValue.SetValue MFDatatypeText, szPDExtDepartmentName
		oPropertyValues.Add -1, oPropertyValue	

		oPropertyValue.PropertyDef = iPDIsActiveFlag
		oPropertyValue.TypedValue.SetValue MFDatatypeBoolean, True
		oPropertyValues.Add -1, oPropertyValue	

		Set oObjectVersionAndProperties = CreateObject("MFilesAPI.ObjectVersionAndProperties")
		Set oObjectVersionAndProperties = Vault.ObjectOperations.CreateNewObjectEx(iOTDepartment, oPropertyValues, oFiles, False, True, oACL)

		oPropVals.SearchForProperty(iPDDepartment).TypedValue.SetValue MFDataTypeLookup, oObjectVersionAndProperties.ObjVer.ID
		
	ElseIf oSearchResults.Count = 1 Then

		oPropVals.SearchForProperty(iPDDepartment).TypedValue.SetValue MFDataTypeLookup, oSearchResults.Item(1).ObjVer.ID
		
	End If

Else 

	oPropVals.SearchForProperty(iPDDepartment).TypedValue.SetValue MFDataTypeLookup, ""

End If

Vault.ObjectPropertyOperations.SetAllProperties ObjVer, True, oPropVals

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