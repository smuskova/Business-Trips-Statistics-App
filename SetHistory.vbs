Const ForReading = 1, ForWriting = 2, ForAppending = 8

LogFolderPath = "D:\M-Files\Temp\log\" & Vault.Name & "\"
BuildFullPath LogFolderPath

Set oPropVals = CreateObject("MFilesApi.PropertyValues")
Set oPropVals = Vault.ObjectPropertyOperations.GetProperties(ObjVer)
LogName = oPropVals.SearchForProperty(100).TypedValue.GetValueAsLookup.DisplayValue
LogState = oPropVals.SearchForProperty(39).TypedValue.GetValueAsLookup.DisplayValue

' Contract properties
iPDApprover = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EhActionEmployee") ' Single Lookup
iPDApprovalComment = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ApprovalComment") ' Single Lookup
iPDHistory = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.History") ' Single Lookup

' Employee properties 
iOTEmployee = Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.Employee")

iPDEmployeeName = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ExtFullName") ' Text

'--------------------------------------------------------
' Set Approval History
'--------------------------------------------------------
On Error Resume Next 
	intPDApprover = oPropVals.SearchForProperty(iPDApprover).TypedValue.GetValueAsLookup.Item
	szPDApprover = oPropVals.SearchForProperty(iPDApprover).TypedValue.DisplayValue
On Error Goto 0 
WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDApprover: " & szPDApprover

If Len(szPDApprover) > 0 Then 
	Set oLookupObj = CreateObject("MFilesAPI.ObjVer")
	oLookupObj.SetIDs iOTEmployee, intPDApprover, -1

	Set oObjectInfo = CreateObject("MFilesAPI.ObjectVersion")
	Set oObjectInfo = Vault.ObjectOperations.GetObjectInfo(oLookupObj, True, False)

	Set oBindProperties = CreateObject("MFilesAPI.PropertyValues")
	Set oBindProperties = Vault.ObjectPropertyOperations.GetProperties(oObjectInfo.ObjVer)
	
	szPDEmployeeName = oBindProperties.SearchForProperty(iPDEmployeeName).TypedValue.DisplayValue
	WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szPDEmployeeName: " & szPDEmployeeName
End If

mlszPDApprovalComment = oPropVals.SearchForProperty(iPDApprovalComment).TypedValue.DisplayValue
mlszPDHistory = oPropVals.SearchForProperty(iPDHistory).TypedValue.DisplayValue

mlszAddHistory = "- Approved on 1st level by """ & szPDEmployeeName & """ - " & TimeStamp

If Len(mlszPDApprovalComment) > 0 Then 	
	mlszAddHistory = mlszAddHistory & VbCrLf & "- Comment:" & VbCrLf & mlszPDApprovalComment
End If 

If Len(mlszPDHistory) > 0 Then 
	mlszPDHistory = mlszAddHistory & VbCrLf & VbCrLf & mlszPDHistory
Else 
	mlszPDHistory = mlszAddHistory
End If

WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; mlszPDHistory: " & VbCrLf & mlszPDHistory

oPropVals.SearchForProperty(iPDHistory).TypedValue.SetValue MFDatatypeMultiLineText, mlszPDHistory

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