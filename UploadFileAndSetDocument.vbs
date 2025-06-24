Const ForReading = 1, ForWriting = 2, ForAppending = 8

LogFolderPath = "D:\M-Files\Temp\log\" & Vault.Name & "\"
BuildFullPath LogFolderPath

Set oPropVals = CreateObject("MFilesApi.PropertyValues")
Set oPropVals = Vault.ObjectPropertyOperations.GetProperties(ObjVer)
LogName = oPropVals.SearchForProperty(100).TypedValue.GetValueAsLookup.DisplayValue
LogState = oPropVals.SearchForProperty(39).TypedValue.GetValueAsLookup.DisplayValue

DownloadFolderPath = "D:\M-Files\Temp\Download\" & Vault.Name & "\"
BuildFullPath DownloadFolderPath

' Upload Signed Document properties
iPDSelectDocument = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.SelectOrderDocument") ' Single Lookup

' Business Trip Document properties
iPDRelatedEmployee = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.RelatedEmployee") ' Single Lookup
iPDBusinessTripDocumentType = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.BusinessTripDocumentType") ' Single Lookup
iPDRelatedBisinessTrip = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.RelatedBusinessTrip") ' Single Lookup
iPDStartDate = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.StartDate") ' Date
iPDEndDate = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EndDate") ' Date
iPDRegisterYear = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.RegisterYear") ' Text
iPDRegisterNumber = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.RegisterNumber") ' Text

iWFS = Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.OrderDocument.SignedDocument")

' Employee properties
iOTEmployee = Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.Employee")

iPDExtFullName = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ExtFullName") ' Text

' Business Trip Document Type properties
iOTBusinessTripDocumentType = Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.BusinessTripDocumentType")

iPDBusinessTripDocumentTypeBg = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.BusinessTripDocumentTypeBg") ' Text
iPDBusinessTripDocumentTypeEn = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.BusinessTripDocumentTypeEn") ' Text

' Business Trip properties
iOTBusinessTrip = Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.BusinessTrip")

'-----------------------------------------------------------------------------
'----- Set path to downloaded files ------------------------------------------
'-----------------------------------------------------------------------------
Set ObjectFileCollection = Vault.ObjectFileOperations.GetFiles(ObjVer)
Set CurrentObjectFile = ObjectFileCollection.Item(1)

szExt = CurrentObjectFile.Extension 
szTitle = CurrentObjectFile.Title
szFilePath = DownloadFolderPath & szTitle & "." & szExt

WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szExt: " & szExt
WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szTitle: " & szTitle
WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szFilePath: " & szFilePath

If UCase(szExt) = "PDF" Or UCase(szExt) = "P7M" Then 

	'-----------------------------------------------------------------------------
	'----- Read properties of related objects used for setting filename ----------
	'-----------------------------------------------------------------------------
	On Error Resume Next
		intPDSelectDocument = oPropVals.SearchForProperty(iPDSelectDocument).TypedValue.GetValueAsLookup.Item
		szPDSelectDocument = oPropVals.SearchForProperty(iPDSelectDocument).TypedValue.DisplayValue
	On Error Goto 0
	WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDSelectDocument: " & szPDSelectDocument

	If Len(szPDSelectDocument) > 0 Then 
	
		Set oLookupObj = CreateObject("MFilesAPI.ObjVer")
		oLookupObj.SetIDs 0, intPDSelectDocument, -1

		Set oObjectInfo = CreateObject("MFilesAPI.ObjectVersion")
		Set oObjectInfo = Vault.ObjectOperations.GetObjectInfo(oLookupObj, True, False)

		Set oBindProperties = CreateObject("MFilesAPI.PropertyValues")
		Set oBindProperties = Vault.ObjectPropertyOperations.GetProperties(oObjectInfo.ObjVer)

		On Error Resume Next 
			intPDRelatedEmployee = oBindProperties.SearchForProperty(iPDRelatedEmployee).TypedValue.GetValueAsLookup.Item
			szPDRelatedEmployee = oBindProperties.SearchForProperty(iPDRelatedEmployee).TypedValue.DisplayValue

			intPDBusinessTripDocumentType = oBindProperties.SearchForProperty(iPDBusinessTripDocumentType).TypedValue.GetValueAsLookup.Item
			szPDBusinessTripDocumentType = oBindProperties.SearchForProperty(iPDBusinessTripDocumentType).TypedValue.DisplayValue

			intPDRelatedBisinessTrip = oBindProperties.SearchForProperty(iPDRelatedBisinessTrip).TypedValue.GetValueAsLookup.Item
			szPDRelatedBisinessTrip = oBindProperties.SearchForProperty(iPDRelatedBisinessTrip).TypedValue.DisplayValue

			datePDStartDate = oBindProperties.SearchForProperty(iPDStartDate).TypedValue.Value
			szPDStartDate = oBindProperties.SearchForProperty(iPDStartDate).TypedValue.DisplayValue

			datePDEndDate = oBindProperties.SearchForProperty(iPDEndDate).TypedValue.Value
			szPDEndDate = oBindProperties.SearchForProperty(iPDEndDate).TypedValue.DisplayValue

			szPDRegisterYear = oBindProperties.SearchForProperty(iPDRegisterYear).TypedValue.DisplayValue

			szPDRegisterNumber = oBindProperties.SearchForProperty(iPDRegisterNumber).TypedValue.DisplayValue
		On Error Goto 0
		WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDRelatedEmployee: " & szPDRelatedEmployee
		WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDBusinessTripDocumentType: " & szPDBusinessTripDocumentType
		WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDRelatedBisinessTrip: " & szPDRelatedBisinessTrip
		WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDStartDate: " & szPDStartDate
		WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDEndDate: " & szPDEndDate
		WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDRegisterYear: " & szPDRegisterYear
		WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDRegisterNumber: " & szPDRegisterNumber

		If Len(szPDStartDate) > 0 Then 
			szStartDate = PadDigits(Year(datePDStartDate),4) & "." & PadDigits(Month(datePDStartDate),2) & "." & PadDigits(Day(datePDStartDate),2)
		End If

		If Len(szPDEndDate) > 0 Then 
			szEndDate = PadDigits(Year(datePDEndDate),4) & "." & PadDigits(Month(datePDEndDate),2) & "." & PadDigits(Day(datePDEndDate),2)
		End If
		WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szStartDate: " & szStartDate
		WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szEndDate: " & szEndDate

		If Len(szPDRelatedEmployee) > 0 Then 

			Set oLookupObj = CreateObject("MFilesAPI.ObjVer")
			oLookupObj.SetIDs iOTEmployee, intPDRelatedEmployee, -1

			Set oObjectInfo = CreateObject("MFilesAPI.ObjectVersion")
			Set oObjectInfo = Vault.ObjectOperations.GetObjectInfo(oLookupObj, True, False)

			Set oBindProperties = CreateObject("MFilesAPI.PropertyValues")
			Set oBindProperties = Vault.ObjectPropertyOperations.GetProperties(oObjectInfo.ObjVer)

			On Error Resume Next 
				szEmployeeName = oBindProperties.SearchForProperty(iPDFullName).TypedValue.DisplayValue
			On Error Goto 0 

		End If
		WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szEmployeeName: " & szEmployeeName

		If Len(szPDBusinessTripDocumentType) > 0 Then 

			Set oLookupObj = CreateObject("MFilesAPI.ObjVer")
			oLookupObj.SetIDs iOTBusinessTripDocumentType, intPDBusinessTripDocumentType, -1

			Set oObjectInfo = CreateObject("MFilesAPI.ObjectVersion")
			Set oObjectInfo = Vault.ObjectOperations.GetObjectInfo(oLookupObj, True, False)

			Set oBindProperties = CreateObject("MFilesAPI.PropertyValues")
			Set oBindProperties = Vault.ObjectPropertyOperations.GetProperties(oObjectInfo.ObjVer)

			On Error Resume Next 
				szPDBusinessTripDocumentTypeBg = oBindProperties.SearchForProperty(iPDBusinessTripDocumentTypeBg).TypedValue.DisplayValue
				szPDBusinessTripDocumentTypeEn = oBindProperties.SearchForProperty(iPDBusinessTripDocumentTypeEn).TypedValue.DisplayValue
			On Error Goto 0 

		End If
		WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDBusinessTripDocumentTypeBg: " & szPDBusinessTripDocumentTypeBg
		WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDBusinessTripDocumentTypeEn: " & szPDBusinessTripDocumentTypeEn

		If Len(szPDBusinessTripDocumentTypeBg) > 0 Then 
			szBusinessTripDocumentType = szPDBusinessTripDocumentTypeBg
		Else 
			szBusinessTripDocumentType = szPDBusinessTripDocumentTypeEn
		End If
		WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szBusinessTripDocumentType: " & szBusinessTripDocumentType

		If Len(szPDRelatedBisinessTrip) > 0 Then 

			Set oLookupObj = CreateObject("MFilesAPI.ObjVer")
			oLookupObj.SetIDs iOTBusinessTrip, intPDRelatedBisinessTrip, -1

			Set oObjectInfo = CreateObject("MFilesAPI.ObjectVersion")
			Set oObjectInfo = Vault.ObjectOperations.GetObjectInfo(oLookupObj, True, False)

			Set oBindProperties = CreateObject("MFilesAPI.PropertyValues")
			Set oBindProperties = Vault.ObjectPropertyOperations.GetProperties(oObjectInfo.ObjVer)

			On Error Resume Next 
				szRORegisterYear = oBindProperties.SearchForProperty(iPDRegisterYear).TypedValue.DisplayValue
				szRORegisterNumber = oBindProperties.SearchForProperty(iPDRegisterNumber).TypedValue.DisplayValue
			On Error Goto 0 

		End If

		'------------------------------------------------------
		'----- Set Filename -----------------------------------
		'------------------------------------------------------
		' If Len(szPDRegisterYear) > 0 Then 
			' szFileName = szFileName & szPDRegisterYear
		' End If
		' If Len(szPDRegisterNumber) > 0 Then 
			' szFileName = szFileName & szPDRegisterNumber
		' End If
		' If Len(szRORegisterYear) > 0 Then 
			' szFileName = szFileName & szRORegisterYear
		' End If
		' If Len(szRORegisterNumber) > 0 Then 
			' szFileName = szFileName & szRORegisterNumber
		' End If
		If Len(szEmployeeName) > 0 Then 
			If Len(szFileName) > 0 Then 
				szFileName = szFileName & " - "
			End If
			szFileName = szFileName & szEmployeeName
		End If
		If Len(szBusinessTripDocumentType) > 0 Then 
			If Len(szFileName) > 0 Then 
				szFileName = szFileName & " - "
			End If
			szFileName = szFileName & szBusinessTripDocumentType
		End If
		If Len(szStartDate) > 0 Then 
			If Len(szFileName) > 0 Then 
				szFileName = szFileName & " - "
			End If
			szFileName = szFileName & szStartDate
		End If
		If Len(szEndDate) > 0 Then 
			If Len(szFileName) > 0 Then 
				szFileName = szFileName & " - "
			End If
			szFileName = szFileName & szEndDate
		End If
		szFileName = szFileName & " - Signed"
		
		WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szFileName: " & szFileName

	Else 

		ErrorMessage = ErrorMessage & "Not selected ""Business Trip""!" & VbCrLf

	End If 

	'--------------------------------------------------------------------------------
	' Upload Document File
	'--------------------------------------------------------------------------------

	Set oLookupObj = CreateObject("MFilesAPI.ObjVer")
	oLookupObj.SetIDs 0, intPDSelectDocument, -1

	Set oObjectInfo = CreateObject("MFilesAPI.ObjectVersion")
	Set oObjectInfo = Vault.ObjectOperations.GetObjectInfo(oLookupObj, True, False)

	Set oBindProperties = CreateObject("MFilesAPI.PropertyValues")
	Set oBindProperties = Vault.ObjectPropertyOperations.GetProperties(oObjectInfo.ObjVer)

	oBindProperties.SearchForProperty(39).TypedValue.SetValue MFDataTypeLookup, iWFS
	Vault.ObjectPropertyOperations.SetAllProperties oObjectInfo.ObjVer, True, oBindProperties

	Set oDestObj = Vault.ObjectOperations.CheckOut(oObjectInfo.ObjVer.ObjID)
	Vault.ObjectFileOperations.GetFilesForModificationInEventHandler oDestObj.ObjVer
	Vault.ObjectFileOperations.AddFile oDestObj.ObjVer, szFileName, szExt, szFilePath
	Call Vault.ObjectOperations.CheckIn(oDestObj.ObjVer)

Else 

	ErrorMessage = ErrorMessage & "Wrong Filename Extension """ & szExt & """!" & vbCrLf & "Only PDF and P7M files can be uploaded!" & VbCrLf

End If 

'------------------------------------------------------
'----- Delete File ------------------------------------
'------------------------------------------------------
On Error Resume Next
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	objFSO.DeleteFile(szFilePath)
	Set objFSO = Nothing
On Error Goto 0 

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