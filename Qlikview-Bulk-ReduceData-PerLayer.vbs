
'##############  DEFINE THE VALUES  ###############################################
'Specify where will you be running this script from | Possible values are: "PROD" "DEV" "LOCAL"
RunningFrom = "DEV"

'Environment paths for each environment
ProdPath = "D:\QlikviewData\Environments\"
DevPath = "C:\QlikviewData\Environments\"
LocalPath = "C:\TEMP\QlikViewLocalTest"

'Go through below specified environments or go through all environments:
'1 = Go through below specified environments
'2 = Go through all environments
EnvParameter = 2

'The environments to consider in this scan and reduction
EnvironmentsArray = Array("UAT","Development","QA")

'The Layer to consider in this scan and reduction
Layer = "Model" 'Model, UI or Extract

'Specify the size limit in MB. Anything bigger than this will be reduced
SizeLimit = 10 'MB

'Specify the age limit in days. Anything bigger than this will be reduced
AgeLimit = 90 'days  0.08333 = 2 hours

'Specify strings to exclude
ExcludeArray = Array("ZZZ.")
'##################################################################################



Dim FileNamesArray()
recordNumber = 0
SizeLimitBytes = SizeLimit * 1048576
StartTime = time()
path = DeterminePath()
Set FSO = CreateObject("Scripting.FileSystemObject")
Set PathFolder = FSO.GetFolder(path)


'Scan the files in each folder of the environments and save those that excede the specified size limit
Call ScanFolders(EnvParameter)


If IsArrayDimmed(FileNamesArray) then
	'There are files detected
	Call ReduceFiles
	EndTime = time()
	Call WriteLog
Else
	MsgBox("No reports to reduce")
	Wscript.Quit
End If



MsgBox "Finished"







'#########################  SUB PROCEDURES AND FUNCTIONS  #########################





'###### Iterate the environments and call scan files
Sub ScanFolders(modeVal)
	'on error resume next
	If modeVal = 1 Then
		'Go by the specified environments		
		For each Env in EnvironmentsArray
			EnvPath = path & Env & "\" & Layer &"\"
			If FSO.FolderExists(EnvPath) Then
				Call ScanFiles(FSO.GetFolder(EnvPath))
			End If
		Next
	ElseIf modeVal = 2 Then
		'Go by all the Environments
		For each folder in PathFolder.SubFolders
			EnvPath = path & folder.Name & "\" & Layer &"\"
			If FSO.FolderExists(EnvPath) Then
				Call ScanFiles(FSO.GetFolder(EnvPath))
			End If
		Next
	Else
		MsgBox("Incorrect mode selected. 1 = go by specified environment   2 = Go through all environments")
	End If
	
End Sub


'###### Within the folder scan all files and save path to array
Sub ScanFiles (oFolder)
	' Checks files in a folder and outputs any that haven't been modified in x days
	For each file in oFolder.Files
		If Mid(file.Name,len(file.Name)-3,4) = ".qvw" then
			ExclusionCount = 0
			For each exclusion in ExcludeArray
				If InStr(file.Name,exclusion) = 0 Then
					'If it is not excluded do not increment counter
				Else
					'The file matches one of the exceptions strings, increase counter
					ExclusionCount = ExclusionCount + 1
				End If
			Next
			
			' Determine the file's age and only reduce those with a set amount of aging
			Age = DateDiff("d",CDATE( file.DateLastModified),date())
						
			If ExclusionCount = 0 and file.size > SizeLimitBytes and Age > AgeLimit Then
				ReDim Preserve FileNamesArray(recordNumber)
				FileNamesArray(recordNumber) = file.Path
				recordNumber = recordNumber + 1
			End If
		End If
	Next
End Sub


'###### For those paths within array, reduce them and save them
Sub ReduceFiles
	Set MyApp = CreateObject("QlikTech.QlikView")
	
	For each report in FileNamesArray

		Set MyDoc = MyApp.OpenDocEx (report,1,false,,,,true)
		Set ActiveDocument = MyDoc
		
		On Error Resume Next
		CurrentTables = MyDoc.GetTableCount
		
		If Err.Number <> 0 Then
		  'This means the document could not be opened.  Need to find a way to cath the error from QV itself and not the failure of VBS
		  'MsgBox("Unable to open report: " & report)
		  Err.Clear
		Else
			On Error Goto 0
			MyDoc.RemoveAllData()
			MyDoc.SaveAs(report)
			MyDoc.CloseDoc()
		End If
		
		On Error Goto 0
		
		
		CurrentTables = 0
		Set MyDoc = Nothing

	Next
	
	MyApp.Sleep 2000
	MyApp.Quit()
	
End Sub


'###### Save a file with all the paths that were reduced
Sub WriteLog
	fileName = Replace(Replace("reduced-files_" & date() & "_" & time() & ".txt","/","-"),":","-")
	Set record = FSO.CreateTextFile(fileName, True)

	record.WriteLine StartTime & "-" & date() 'start
	For each report in FileNamesArray
		record.WriteLine(report)
	Next
	record.WriteLine Endtime & "-" & date() 'finish

	record.close
End Sub





'###### Function to detect if array is empty or not
Function IsArrayDimmed(arr)
  IsArrayDimmed = False
  If IsArray(arr) Then
    On Error Resume Next
    Dim ub : ub = UBound(arr)
    If (Err.Number = 0) And (ub >= 0) Then IsArrayDimmed = True
  End If  
End Function

'###### Function to define the path based on the RunningFrom variable
Function DeterminePath()
	If RunningFrom = "PROD" then
		DeterminePath = ProdPath
	ElseIf RunningFrom = "DEV" then
		DeterminePath = DevPath
	Else
		DeterminePath = LocalPath
	End If
End Function

