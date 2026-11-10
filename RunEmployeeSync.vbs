Option Explicit

Dim xl, wb
On Error Resume Next

Set xl = CreateObject("Excel.Application")
If Err.Number <> 0 Then WScript.Quit 1

xl.Visible = False
xl.DisplayAlerts = False
xl.EnableEvents = True
xl.AskToUpdateLinks = False

' File containing the VBA module.
Set wb = xl.Workbooks.Open("D:\vtc\github\EmployeeSync.xlsm", False, False)

If Err.Number <> 0 Then
    xl.Quit
    WScript.Quit 2
End If

Err.Clear

xl.Run "'" & wb.Name & "'!CopySpecificColumnsToGoogleSheets"

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
