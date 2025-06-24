Const ForReading = 1, ForWriting = 2, ForAppending = 8

LogFolderPath = "D:\M-Files\Temp\log\" & Vault.Name & "\"
BuildFullPath LogFolderPath

Set oPropVals = CreateObject("MFilesApi.PropertyValues")
Set oPropVals = Vault.ObjectPropertyOperations.GetProperties(ObjVer)
LogName = oPropVals.SearchForProperty(100).TypedValue.GetValueAsLookup.DisplayValue
LogState = oPropVals.SearchForProperty(39).TypedValue.GetValueAsLookup.DisplayValue

DownloadFolderPath = "D:\M-Files\Temp\Download\" & Vault.Name & "\"
BuildFullPath DownloadFolderPath

' Document properties
iCLReportDocument = Vault.ClassOperations.GetObjectClassIDbyAlias("CL.ReportDocument")

iPDRelatedEmployee = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.RelatedEmployee") ' Single Lookup
iPDBusinessTripType = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.BusinessTripType") ' Single Lookup
iPDTypeOfBusinessTrip = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.TypeOfBusinessTrip") ' Single Lookup

iWFSIsTemplate = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.ReportDocument.IsTemplate")

' Employee properties
iOTEmployee = Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.Employee")

iPDExtFullName = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ExtFullName") ' Text

' Business Trip Document Type properties
iOTReportDocumentType = Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.ReportDocumentType")

iPDReportDocumentTypeBg = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ReportDocumentTypeBg") ' Text
iPDReportDocumentTypeEn = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ReportDocumentTypeEn") ' Text

'-----------------------------------------------------------------------------
'----- Business Trip Document properties -------------------------------------
'-----------------------------------------------------------------------------
On Error Resume Next 
	intPDRelatedEmployee = oPropVals.SearchForProperty(iPDRelatedEmployee).TypedValue.GetValueAsLookup.Item
	szPDRelatedEmployee = oPropVals.SearchForProperty(iPDRelatedEmployee).TypedValue.DisplayValue

	intPDBusinessTripType = oPropVals.SearchForProperty(iPDBusinessTripType).TypedValue.GetValueAsLookup.Item
	szPDBusinessTripType = oPropVals.SearchForProperty(iPDBusinessTripType).TypedValue.DisplayValue

	intPDTypeOfBusinessTrip = oPropVals.SearchForProperty(iPDTypeOfBusinessTrip).TypedValue.GetValueAsLookup.Item
	szPDTypeOfBusinessTrip = oPropVals.SearchForProperty(iPDTypeOfBusinessTrip).TypedValue.DisplayValue
On Error Goto 0
WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDRelatedEmployee: " & szPDRelatedEmployee
WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDBusinessTripType: " & szPDBusinessTripType
WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDTypeOfBusinessTrip: " & szPDTypeOfBusinessTrip

If Len(szPDRelatedEmployee) > 0 Then 

	Set oLookupObj = CreateObject("MFilesAPI.ObjVer")
	oLookupObj.SetIDs iOTEmployee, intPDRelatedEmployee, -1

	Set oObjectInfo = CreateObject("MFilesAPI.ObjectVersion")
	Set oObjectInfo = Vault.ObjectOperations.GetObjectInfo(oLookupObj, True, False)

	Set oBindProperties = CreateObject("MFilesAPI.PropertyValues")
	Set oBindProperties = Vault.ObjectPropertyOperations.GetProperties(oObjectInfo.ObjVer)

	On Error Resume Next 
		szPDExtFullName = oBindProperties.SearchForProperty(iPDExtFullName).TypedValue.DisplayValue
	On Error Goto 0 
	WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDExtFullName: " & szPDExtFullName

	If Len(szPDExtFullName) > 0 Then 
		szEmployeeName = szPDExtFullName
	End If

Else 
	
	ErrorMessage = ErrorMessage & VbCrLf & "Missing ""Related Employee""!"

End If
WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szEmployeeName: " & szEmployeeName

'------------------------------------------------------
'----- Set Filename -----------------------------------
'------------------------------------------------------
szFileName = "Report"
If Len(szEmployeeName) > 0 Then 
	If Len(szFileName) > 0 Then 
		szFileName = szFileName & " - "
	End If
	szFileName = szFileName & szEmployeeName
End If
WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szFileName: " & szFileName

'--------------------------------------------------------
'----- Upload Template or Rename File -------------------
'--------------------------------------------------------
Vault.ObjectFileOperations.GetFilesForModificationInEventHandler ObjVer

Set ObjectFileCollection = Vault.ObjectFileOperations.GetFiles(ObjVer)

WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; ObjectFileCollection.Count: " & ObjectFileCollection.Count

If ObjectFileCollection.Count = 0 Then ' Upload Template Document

	Set oSCs = CreateObject("MFilesAPI.SearchConditions") 

	Set oSCClass = CreateObject("MFilesAPI.SearchCondition") 
	oSCClass.ConditionType = MFConditionTypeEqual
	oSCClass.Expression.DataPropertyValuePropertyDef = 100
	oSCClass.TypedValue.SetValue MFDatatypeLookup, iCLReportDocument
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

	WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; oSearchResults.Count: " & oSearchResults.Count

	If oSearchResults.Count = 1 Then

		Set oLookup = CreateObject("MFilesAPI.Lookup")
		For Each oSearchResult In oSearchResults
			oLookup.Item = oSearchResult.ObjVer.ID
		Next

		Set oLookupObj = CreateObject("MFilesAPI.ObjVer")
		oLookupObj.SetIDs 0, oLookup.Item, -1

		Set oObjectInfo = CreateObject("MFilesAPI.ObjectVersion")
		Set oObjectInfo = Vault.ObjectOperations.GetObjectInfo(oLookupObj, True, False)

		Set ObjectFileCollection = Vault.ObjectFileOperations.GetFiles(oObjectInfo.ObjVer)

		Set CurrentObjectFile = ObjectFileCollection.Item(1)

		szExt = CurrentObjectFile.Extension 
		szTitle = CurrentObjectFile.Title

		szFilePath = DownloadFolderPath & szTitle & "." & szExt
		WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szFilePath: " & szFilePath

		Call Vault.ObjectFileOperations.DownloadFile(CurrentObjectFile.ID, CurrentObjectFile.Version, szFilePath)

		Vault.ObjectFileOperations.GetFilesForModificationInEventHandler ObjVer
		Vault.ObjectFileOperations.AddFile ObjVer, szFileName, szExt, szFilePath

		Vault.ObjectPropertyOperations.SetAllProperties ObjVer, True, oPropVals

		'Delete File
		On Error Resume Next
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			objFSO.DeleteFile(szFilePath)
			Set objFSO = Nothing
		On Error Goto 0
		
	Else 
		' Missing Template Document. Do Nothing
		WriteLog "WARNING", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; Missing template document for document type: " & szPDDocumentType
	End If 
	
End If 

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