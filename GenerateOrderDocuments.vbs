Const ForReading = 1, ForWriting = 2, ForAppending = 8

LogFolderPath = "D:\M-Files\Temp\log\" & Vault.Name & "\"
BuildFullPath LogFolderPath

Set oPropVals = CreateObject("MFilesApi.PropertyValues")
Set oPropVals = Vault.ObjectPropertyOperations.GetProperties(ObjVer)
LogName = oPropVals.SearchForProperty(100).TypedValue.GetValueAsLookup.DisplayValue
LogState = oPropVals.SearchForProperty(39).TypedValue.GetValueAsLookup.DisplayValue

' Business Trip properties 
iPDRequestorEmployee = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.RequestorEmployee") ' Single Lookup
iPDBusinessTripType = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.BusinessTripType") ' Single Lookup
iPDTypeOfBusinessTrip = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.TypeOfBusinessTrip") ' Single Lookup
iPDStartDate = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.StartDate") ' Date
iPDEndDate = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EndDate") ' Date
iPDBusinessTripDuration = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.BusinessTripDuration") ' Integer
iPDRelatedDocument = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.RelatedOrderDocument")' Single Lookup

' Business Trip Document Type properties
iOTBusinessTripDocumentType = Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.BusinessTripDocumentType")

iCLBusinessTripDocumentType = Vault.ClassOperations.GetObjectClassIDbyAlias("CL.BusinessTripDocumentType")

iPDIsOrderDocument = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.IsOrderDocument") ' Single Lookup
iPDIsActiveFlag = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.IsActiveFlag") ' Boolean
iPDDelayReminderInDays = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.DelayReminderInDays")

' Document properties
iCLDocument = Vault.ClassOperations.GetObjectClassIDbyAlias("CL.OrderDocument")

iPDBusinessTripDocumentType = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.BusinessTripDocumentType") ' Single Lookup
iPDRelatedEmployee = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.RelatedEmployee") ' Single Lookup
iPDRelatedBisinessTrip = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.RelatedBusinessTrip") ' Single Lookup

iWFBusinessTripDocument = Vault.WorkflowOperations.GetWorkflowIDByAlias("WF.OrderDocument")
iWFSBusinessTripDocumentInit = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.OrderDocument.Init")
iWFSIsTemplate = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.OrderDocument.IsTemplate")

'----------------------------------------------------------------------
'----- Read Business Trip information ---------------------------------
'----------------------------------------------------------------------
On Error Resume Next 
	intPDRequestorEmployee = oPropVals.SearchForProperty(iPDRequestorEmployee).TypedValue.GetValueAsLookup.Item
	szPDRequestorEmployee = oPropVals.SearchForProperty(iPDRequestorEmployee).TypedValue.DisplayValue

	intPDBusinessTripType = oPropVals.SearchForProperty(iPDBusinessTripType).TypedValue.GetValueAsLookup.Item
	szPDBusinessTripType = oPropVals.SearchForProperty(iPDBusinessTripType).TypedValue.DisplayValue

	intPDTypeOfBusinessTrip = oPropVals.SearchForProperty(iPDTypeOfBusinessTrip).TypedValue.GetValueAsLookup.Item
	szPDTypeOfBusinessTrip = oPropVals.SearchForProperty(iPDTypeOfBusinessTrip).TypedValue.DisplayValue

	datePDStartDate = oPropVals.SearchForProperty(iPDStartDate).TypedValue.Value
	szPDStartDate = oPropVals.SearchForProperty(iPDStartDate).TypedValue.DisplayValue

	datePDEndDate = oPropVals.SearchForProperty(iPDEndDate).TypedValue.Value
	szPDEndDate = oPropVals.SearchForProperty(iPDEndDate).TypedValue.DisplayValue

	intPDBusinessTripDuration = oPropVals.SearchForProperty(iPDBusinessTripDuration).TypedValue.Value
	szPDBusinessTripDuration = oPropVals.SearchForProperty(iPDBusinessTripDuration).TypedValue.DisplayValue
On Error Goto 0
WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDRequestorEmployee: " & szPDRequestorEmployee
WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDBusinessTripType: " & szPDBusinessTripType
WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDTypeOfBusinessTrip: " & szPDTypeOfBusinessTrip
WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDStartDate: " & szPDStartDate
WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDEndDate: " & szPDEndDate
WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDBusinessTripDuration: " & szPDBusinessTripDuration

'------------------------------------------------------
'----- Create Business Trip Document for Order --------
'------------------------------------------------------
Set oSCs = CreateObject("MFilesAPI.SearchConditions") 

Set oSCClass = CreateObject("MFilesAPI.SearchCondition") 
oSCClass.ConditionType = MFConditionTypeEqual
oSCClass.Expression.DataPropertyValuePropertyDef = 100
oSCClass.TypedValue.SetValue MFDatatypeLookup, iCLBusinessTripDocumentType
oSCs.Add -1, oSCClass

Set oSearchCondition = CreateObject("MFilesAPI.SearchCondition")
oSearchCondition.ConditionType = MFConditionTypeEqual
oSearchCondition.Expression.DataPropertyValuePropertyDef = iPDIsOrderDocument
oSearchCondition.TypedValue.SetValue MFDatatypeLookup, 1 ' Yes
oSCs.Add -1, oSearchCondition

Set oSearchCondition = CreateObject("MFilesAPI.SearchCondition")
oSearchCondition.ConditionType = MFConditionTypeEqual
oSearchCondition.Expression.DataPropertyValuePropertyDef = iPDBusinessTripType
oSearchCondition.TypedValue.SetValue MFDatatypeLookup, intPDBusinessTripType
oSCs.Add -1, oSearchCondition

Set oSearchCondition = CreateObject("MFilesAPI.SearchCondition")
oSearchCondition.ConditionType = MFConditionTypeEqual
oSearchCondition.Expression.DataPropertyValuePropertyDef = iPDIsActiveFlag
oSearchCondition.TypedValue.SetValue MFDatatypeBoolean, True
oSCs.Add -1, oSearchCondition

' Search for just the non deleted element
Set oSearchNonDeleted = CreateObject("MFilesAPI.SearchCondition") 
oSearchNonDeleted.ConditionType = MFConditionTypeEqual
oSearchNonDeleted.Expression.DataStatusValueType = MFStatusTypeDeleted
oSearchNonDeleted.TypedValue.SetValue MFDatatypeBoolean, False
oSCs.Add -1, oSearchNonDeleted

' Execute Search	
Set oSearchResults = Vault.ObjectSearchOperations.SearchForObjectsByConditions(oSCs,MFSearchFlagNone, False) 

WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; Found Is Order Business Trip Document Type count: " & oSearchResults.Count

If oSearchResults.Count = 1 Then

	intPDBusinessTripDocumentType = oSearchResults.Item(1).ObjVer.ID
	WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; intPDBusinessTripDocumentType: " & intPDBusinessTripDocumentType

	Set oLookupObj = CreateObject("MFilesAPI.ObjVer")
	oLookupObj.SetIDs iOTBusinessTripDocumentType, intPDBusinessTripDocumentType, -1

	Set oObjectInfo = CreateObject("MFilesAPI.ObjectVersion")
	Set oObjectInfo = Vault.ObjectOperations.GetObjectInfo(oLookupObj, True, False)

	Set oBindProperties = CreateObject("MFilesAPI.PropertyValues")
	Set oBindProperties = Vault.ObjectPropertyOperations.GetProperties(oObjectInfo.ObjVer)

	On Error Resume Next 
		intPDDelayReminderInDays = oBindProperties.SearchForProperty(iPDDelayReminderInDays).TypedValue.Value
	On Error Goto 0 
	WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; intPDDelayReminderInDays: " & intPDDelayReminderInDays

	' Check if template document exists
	Set oSCs = CreateObject("MFilesAPI.SearchConditions") 

	Set oSCClass = CreateObject("MFilesAPI.SearchCondition") 
	oSCClass.ConditionType = MFConditionTypeEqual
	oSCClass.Expression.DataPropertyValuePropertyDef = 100
	oSCClass.TypedValue.SetValue MFDatatypeLookup, iCLDocument
	oSCs.Add -1, oSCClass

	Set oSearchCondition = CreateObject("MFilesAPI.SearchCondition")
	oSearchCondition.ConditionType = MFConditionTypeEqual
	oSearchCondition.Expression.DataPropertyValuePropertyDef = 37
	oSearchCondition.TypedValue.SetValue MFDatatypeBoolean, True
	oSCs.Add -1, oSearchCondition

	Set oSearchCondition = CreateObject("MFilesAPI.SearchCondition")
	oSearchCondition.ConditionType = MFConditionTypeEqual
	oSearchCondition.Expression.DataPropertyValuePropertyDef = iPDBusinessTripType
	oSearchCondition.TypedValue.SetValue MFDatatypeLookup, intPDBusinessTripType
	oSCs.Add -1, oSearchCondition

	Set oSearchCondition = CreateObject("MFilesAPI.SearchCondition")
	oSearchCondition.ConditionType = MFConditionTypeEqual
	oSearchCondition.Expression.DataPropertyValuePropertyDef = iPDTypeOfBusinessTrip
	oSearchCondition.TypedValue.SetValue MFDatatypeLookup, intPDTypeOfBusinessTrip
	oSCs.Add -1, oSearchCondition

	Set oSearchCondition = CreateObject("MFilesAPI.SearchCondition")
	oSearchCondition.ConditionType = MFConditionTypeEqual
	oSearchCondition.Expression.DataPropertyValuePropertyDef = 39
	oSearchCondition.TypedValue.SetValue MFDatatypeLookup, iWFSIsTemplate
	oSCs.Add -1, oSearchCondition

	' Search for just the non deleted element
	Set oSearchNonDeleted = CreateObject("MFilesAPI.SearchCondition") 
	oSearchNonDeleted.ConditionType = MFConditionTypeEqual
	oSearchNonDeleted.Expression.DataStatusValueType = MFStatusTypeDeleted
	oSearchNonDeleted.TypedValue.SetValue MFDatatypeBoolean, False
	oSCs.Add -1, oSearchNonDeleted

	' Execute Search	
	Set oSearchResults = Vault.ObjectSearchOperations.SearchForObjectsByConditions(oSCs,MFSearchFlagNone, False) 

	WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; Found template Business Trip Document count: " & oSearchResults.Count

	If oSearchResults.Count = 1 Then

		'------------------------------------------------------
		'----- Create BT Document for Order -------------------
		'------------------------------------------------------
		Set oPropertyValues = CreateObject("MFilesAPI.PropertyValues")
		Set oPropertyValue = CreateObject("MFilesAPI.PropertyValue")
		Set oFiles = CreateObject("MFilesAPI.SourceObjectFiles")
		Set oACL = CreateObject("MFilesAPI.AccessControlList")

		oPropertyValue.PropertyDef = 100
		oPropertyValue.TypedValue.SetValue MFDatatypeLookup, iCLDocument
		oPropertyValues.Add -1, oPropertyValue

		oPropertyValue.PropertyDef = iPDRelatedEmployee
		oPropertyValue.TypedValue.SetValue MFDatatypeLookup, intPDRequestorEmployee
		oPropertyValues.Add -1, oPropertyValue

		oPropertyValue.PropertyDef = iPDBusinessTripType
		oPropertyValue.TypedValue.SetValue MFDatatypeLookup, intPDBusinessTripType
		oPropertyValues.Add -1, oPropertyValue			

		oPropertyValue.PropertyDef = iPDTypeOfBusinessTrip
		oPropertyValue.TypedValue.SetValue MFDatatypeLookup, intPDTypeOfBusinessTrip
		oPropertyValues.Add -1, oPropertyValue			
		
		oPropertyValue.PropertyDef = iPDBusinessTripDocumentType
		oPropertyValue.TypedValue.SetValue MFDatatypeLookup, intPDBusinessTripDocumentType
		oPropertyValues.Add -1, oPropertyValue			
		
		If intPDDelayReminderInDays > 0 Then 
			oPropertyValue.PropertyDef = iPDDelayReminderInDays
			oPropertyValue.TypedValue.SetValue MFDatatypeInteger, intPDDelayReminderInDays
			oPropertyValues.Add -1, oPropertyValue
		End If

		If Len(szPDStartDate) > 0 Then 
			oPropertyValue.PropertyDef = iPDStartDate
			oPropertyValue.TypedValue.SetValue MFDatatypeDate, datePDStartDate
			oPropertyValues.Add -1, oPropertyValue
		End If

		If Len(szPDEndDate) > 0 Then 
			oPropertyValue.PropertyDef = iPDEndDate
			oPropertyValue.TypedValue.SetValue MFDatatypeDate, datePDEndDate
			oPropertyValues.Add -1, oPropertyValue
		End If

		If Len(szPDBusinessTripDuration) > 0 Then 
			oPropertyValue.PropertyDef = iPDBusinessTripDuration
			oPropertyValue.TypedValue.SetValue MFDatatypeInteger, intPDBusinessTripDuration
			oPropertyValues.Add -1, oPropertyValue
		End If

		oPropertyValue.PropertyDef = iPDRelatedBisinessTrip
		oPropertyValue.TypedValue.SetValue MFDatatypeLookup, ObjVer.ID
		oPropertyValues.Add -1, oPropertyValue			

		'Set Workflow and Workflow State
		oPropertyValue.PropertyDef = 38
		oPropertyValue.TypedValue.SetValue MFDatatypeLookup, iWFBusinessTripDocument
		oPropertyValues.Add -1, oPropertyValue

		oPropertyValue.PropertyDef = 39
		oPropertyValue.TypedValue.SetValue MFDatatypeLookup, iWFSBusinessTripDocumentInit
		oPropertyValues.Add -1, oPropertyValue

		Set oObjectVersionAndProperties = Vault.ObjectOperations.CreateNewObjectEx(0, oPropertyValues, oFiles, False, True, oACL)	

		WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; Created new Business Trip Document of type Order!"

		oPropVals.SearchForProperty(iPDRelatedDocument).TypedValue.SetValue MFDatatypeLookup, oObjectVersionAndProperties.ObjVer.ID
	
	ElseIf oSearchResults.Count > 1 Then
		
		ErrorMessage = ErrorMessage & "More than 1 template Business Trip Document exist"
		
	Else
		
		ErrorMessage = ErrorMessage & "Missing template Business Trip Document"
		
	End If
	
ElseIf oSearchResults.Count > 1 Then

	ErrorMessage = ErrorMessage & "More than 1 Business Trip Document Type for Order Document"
	
Else

	ErrorMessage = ErrorMessage & "Missing Business Trip Document Type for Order Document"
	
End If

Vault.ObjectPropertyOperations.SetAllProperties ObjVer, True, oPropVals	

'------------------------------------------------------
'----- Show Error -------------------------------------
'------------------------------------------------------
If Len(ErrorMessage) > 0 Then 
	Err.Raise MFScriptCancel, ErrorMessage
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