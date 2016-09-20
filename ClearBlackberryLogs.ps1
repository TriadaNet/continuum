'THIS SCRIPT Checks for folders older than 7 days in the path D:\PROGRAM FILES\RESEARCH IN MOTION\BLACKBERRY ENTERPRISE SERVER\LOGS.
'If folders are older than 7 days they are deleted.
'A logfile is written to the D:\PROGRAM FILES\RESEARCH IN MOTION\BLACKBERRY ENTERPRISE SERVER\LOGS path on completion.
'This logfile is overwritten at each run.

On Error Resume Next

'Define the constants
'Number of days to retain the folder
CONST ADJUSTDATE = -7
'Enter your folder path to the BlackBerry logs below
CONST FOLDERPATH = "c:\PROGRAM FILES\RESEARCH IN MOTION\BLACKBERRY ENTERPRISE SERVER\LOGS"
CONST FILENAME = "DeleteLogsScript.log"

'Get the File System Objects
Set fso = CreateObject("Scripting.FileSystemObject") 
Set folder = fso.GetFolder(FOLDERPATH) 
Set subFolders = folder.SubFolders 

'Set the log file
sLogFilePath = FOLDERPATH & "\" & FILENAME
set logfile = fso.CreateTextFile(sLogFilePath, True)
'Write the log file
logfile.WriteLine("BES Delete Logs started on " & Now )
logfile.WriteLine("The following folders were deleted")

'Set the date to delete folders older than
dateold = DateAdd("d", ADJUSTDATE, Date)

'Process all the folders
For Each folderObject in SubFolders
errorvalue=0 'ensure the error value is reset for each folder
datemod = DateValue(folderObject.DateLastModified)
'Delete folders that are old

IF datemod < dateold Then
fso.DeleteFolder folderobject, true 'this line deletes the folders
errorvalue = Err.Number

If ErrorValue = 0 Then 'The folder was deleted ok
logfile.WriteLine( folderobject.Name )'record folders that were successfully deleted
Else
If Errorvalue = 70 Then 
'Error value 70 denotes the folder could not be deleted for access reasons. Most likely the folder (or a file within it) is in use.
logfile.WriteLine (folderObject.path & " could not be deleted. This folder, or a file within it, may be in use.")
Else
logfile.WriteLine (folderObject.path & " could not be deleted. Unknown error")
End If	
End If
End IF
Next 

'Write the log file
logfile.WriteLine("BES Delete Logs Completed on.")
logfile.Close

'Clear the variables
Set subFolders = Nothing 
Set folder = Nothing 
Set fso = Nothing
Set logfile = Nothing