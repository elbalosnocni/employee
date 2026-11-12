Option Explicit
Dim fso, xl, wb, rootPath, bookPath, macroName
Set fso = CreateObject("Scripting.FileSystemObject")
rootPath = fso.GetParentFolderName(WScript.ScriptFullName)
bookPath = fso.BuildPath(fso.BuildPath(rootPath, "Excel"), "EmployeeSync.xlsm")
macroName = "CopySpecificColumnsToGoogleSheets"

On Error Resume Next
Set xl = CreateObject("Excel.Application")
If Err.Number <> 0 Then WScript.Quit 1
xl.Visible = False
xl.DisplayAlerts = False
xl.EnableEvents = True
xl.AskToUpdateLinks = False
Set wb = xl.Workbooks.Open(bookPath, False, False)
If Err.Number <> 0 Then
  xl.Quit
  WScript.Quit 2
End If
Err.Clear
xl.Run "'" & wb.Name & "'!" & macroName
If Err.Number <> 0 Then
  wb.Close False
  xl.Quit
  WScript.Quit 3
End If
wb.Save
wb.Close False
xl.Quit
Set wb = Nothing
Set xl = Nothing
WScript.Quit 0
