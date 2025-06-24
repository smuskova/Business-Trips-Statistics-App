Const ForReading = 1, ForWriting = 2, ForAppending = 8

LogFolderPath = "D:\M-Files\Temp\log\" & Vault.Name & "\"
BuildFullPath LogFolderPath

Set oPropVals = CreateObject("MFilesApi.PropertyValues")
Set oPropVals = Vault.ObjectPropertyOperations.GetProperties(ObjVer)
LogName = oPropVals.SearchForProperty(100).TypedValue.GetValueAsLookup.DisplayValue
LogState = oPropVals.SearchForProperty(39).TypedValue.GetValueAsLookup.DisplayValue

DownloadFolderPath = "D:\M-Files\Temp\Download\" & Vault.Name & "\"
BuildFullPath DownloadFolderPath

'--------------------------------------------------------------------------------
'Download Object File 
'--------------------------------------------------------------------------------
Set ObjectFileCollection = Vault.ObjectFileOperations.GetFiles(ObjVer)

If ObjectFileCollection.Count = 1 Then

	Set CurrentObjectFile = ObjectFileCollection.Item(1)

	szExt = CurrentObjectFile.Extension 
	szTitle = CurrentObjectFile.Title

	'-----------------------------------------------------------------------------
	'----- Check files extension -------------------------------------------------
	'-----------------------------------------------------------------------------
	If Not (UCase(szExt) = "PDF" Or UCase(szExt) = "P7M") Then 
	
		Err.Raise MFScriptCancel, "Wrong Filename Extension """ & szExt & """!" & vbCrLf & vbCrLf _
			& "Only signed PDF and P7M files can be uploaded!"
	
	End If 

	szFilePath = DownloadFolderPath & szTitle & "." & szExt
	WriteLog "INFO", LogFolderPath & LogName, LogState & "; " & ObjVer.ID & "; szFilePath: " & szFilePath

	'-----------------------------------------------------------------------------
	'----- Download --------------------------------------------------------------
	'-----------------------------------------------------------------------------
	Call Vault.ObjectFileOperations.DownloadFile(CurrentObjectFile.ID, CurrentObjectFile.Version, szFilePath)
Else 

	Err.Raise MFScriptCancel, "Missing file!" & vbCrLf & vbCrLf & "Use ""Add to M-Files"" function to upload file!"

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