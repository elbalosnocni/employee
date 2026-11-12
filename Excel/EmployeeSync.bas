Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME)
#Else
    Private Declare Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME)
#End If

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

' EmployeeSync v2 - nhieu chi nhanh, cau hinh ngoai bang BranchConfig.xlsx.
' Chi gui 6 truong cong khai: ma NV, ho ten, ma cham cong, phong ban, bo phan, chuc vu.
' CCCD chi dung kiem tra du 12 so tai may nguon; KHONG DUOC DUA VAO JSON.
' Khong doc/gui cot luong, phu cap, ngay sinh.

Private Const CONFIG_FILE As String = "BranchConfig.xlsx"
Private Const CONFIG_SHEET As String = "Branches"
Private Const CONFIG_FIRST_ROW As Long = 2

Public Sub CopySpecificColumnsToGoogleSheets()
    Dim cfgWb As Workbook, cfgWs As Worksheet
    Dim configPath As String, lastCfg As Long, r As Long
    Dim apiUrl As String, branchCode As String, token As String
    Dim folderPath As String, pattern As String, excelPassword As String
    Dim sourceSheet As String, headerRow As Long, firstDataRow As Long
    Dim colName As String, colEmp As String, colClock As String
    Dim colDept As String, colSection As String, colPosition As String, colCCCD As String
    Dim filePath As String, result As String, sentCount As Long, branchErrors As Long

    On Error GoTo FatalError
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.AskToUpdateLinks = False

    configPath = ThisWorkbook.Path & Application.PathSeparator & CONFIG_FILE
    If Dir(configPath) = "" Then Err.Raise vbObjectError + 100, , "Khong tim thay " & configPath

    Set cfgWb = Workbooks.Open(configPath, ReadOnly:=True)
    Set cfgWs = cfgWb.Worksheets(CONFIG_SHEET)
    lastCfg = cfgWs.Cells(cfgWs.Rows.Count, "A").End(xlUp).Row

    For r = CONFIG_FIRST_ROW To lastCfg
        If UCase$(Trim$(CStr(cfgWs.Cells(r, 1).Value))) = "YES" Or _
           UCase$(Trim$(CStr(cfgWs.Cells(r, 1).Value))) = "TRUE" Or _
           Trim$(CStr(cfgWs.Cells(r, 1).Value)) = "1" Then

            branchCode = UCase$(Trim$(CStr(cfgWs.Cells(r, 2).Value)))
            folderPath = Trim$(CStr(cfgWs.Cells(r, 3).Value))
            pattern = Trim$(CStr(cfgWs.Cells(r, 4).Value))
            excelPassword = CStr(cfgWs.Cells(r, 5).Value)
            sourceSheet = Trim$(CStr(cfgWs.Cells(r, 6).Value))
            headerRow = Val(cfgWs.Cells(r, 7).Value)
            firstDataRow = Val(cfgWs.Cells(r, 8).Value)
            colName = Trim$(CStr(cfgWs.Cells(r, 9).Value))
            colEmp = Trim$(CStr(cfgWs.Cells(r, 10).Value))
            colClock = Trim$(CStr(cfgWs.Cells(r, 11).Value))
            colDept = Trim$(CStr(cfgWs.Cells(r, 12).Value))
            colSection = Trim$(CStr(cfgWs.Cells(r, 13).Value))
            colPosition = Trim$(CStr(cfgWs.Cells(r, 14).Value))
            colCCCD = Trim$(CStr(cfgWs.Cells(r, 15).Value))
            apiUrl = Trim$(CStr(cfgWs.Cells(r, 16).Value))
            token = Trim$(CStr(cfgWs.Cells(r, 17).Value))

            If branchCode = "" Or folderPath = "" Or pattern = "" Or sourceSheet = "" Or _
               firstDataRow < 1 Or colName = "" Or colEmp = "" Or apiUrl = "" Or token = "" Then
                Err.Raise vbObjectError + 101, , "Thieu cau hinh bat buoc tai dong " & r & " trong BranchConfig.xlsx"
            End If

            filePath = FindNewestFile(folderPath, pattern)
            If filePath = "" Then
                LogLocal "SKIP", branchCode, "Khong tim thay file: " & folderPath & pattern
            Else
                result = SyncOneBranch(filePath, excelPassword, sourceSheet, firstDataRow, _
                    branchCode, apiUrl, token, colName, colEmp, colClock, colDept, _
                    colSection, colPosition, colCCCD, sentCount)
                LogLocal "RESULT", branchCode, result & " | File: " & filePath
            End If
        End If
    Next r

    SafeCloseWorkbook cfgWb
    MsgBox "Da chay xong cau hinh cac chi nhanh. Xem log tai sheet SyncRunLog trong file dieu khien.", vbInformation
CleanExit:
    On Error Resume Next
    SafeCloseWorkbook cfgWb
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
    Exit Sub
FatalError:
    LogLocal "ERROR", "", Err.Description
    MsgBox "Dong bo dung: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Function SyncOneBranch(ByVal filePath As String, ByVal pwd As String, _
    ByVal sheetName As String, ByVal firstRow As Long, ByVal branch As String, _
    ByVal apiUrl As String, ByVal token As String, ByVal cName As String, _
    ByVal cEmp As String, ByVal cClock As String, ByVal cDept As String, _
    ByVal cSection As String, ByVal cPosition As String, ByVal cCCCD As String, _
    ByRef sentCount As Long) As String

    Dim srcWb As Workbook, ws As Worksheet, lastRow As Long, i As Long
    Dim arr As Collection, obj As String, cccd As String, emp As String, nm As String
    Dim payload As String, http As Object, response As String

    On Error GoTo BranchError
    Set srcWb = Workbooks.Open(fileName:=filePath, UpdateLinks:=0, ReadOnly:=True, _
        Password:=pwd, WriteResPassword:=pwd, IgnoreReadOnlyRecommended:=True)
    Set ws = srcWb.Worksheets(sheetName)
    lastRow = ws.Cells(ws.Rows.Count, ColumnNumber(cEmp)).End(xlUp).Row
    Set arr = New Collection

    For i = firstRow To lastRow
        emp = CellText(ws.Cells(i, ColumnNumber(cEmp)))
        nm = CellText(ws.Cells(i, ColumnNumber(cName)))
        If emp <> "" And nm <> "" Then
            ' CCCD chi la dieu kien kiem tra tai may nguon; tuyet doi khong gui len API.
            If cCCCD <> "" Then
                cccd = DigitsOnly(CellText(ws.Cells(i, ColumnNumber(cCCCD))))
                If Len(cccd) <> 12 Then GoTo ContinueLoop
            End If
            obj = "{""employee_code"":""" & JsonEscape(emp) & """," & _
                  """full_name"":""" & JsonEscape(nm) & """," & _
                  """time_clock_code"":""" & JsonEscape(CellTextByCol(ws, i, cClock)) & """," & _
                  """department"":""" & JsonEscape(CellTextByCol(ws, i, cDept)) & """," & _
                  """section"":""" & JsonEscape(CellTextByCol(ws, i, cSection)) & """," & _
                  """position"":""" & JsonEscape(CellTextByCol(ws, i, cPosition)) & """}"
            arr.Add obj
        End If
ContinueLoop:
    Next i
    SafeCloseWorkbook srcWb
    If arr.Count = 0 Then
        SyncOneBranch = "Khong co nhan vien hop le"
        Exit Function
    End If

    payload = "{""version"":2,""branch_code"":""" & JsonEscape(branch) & """," & _
        """token"":""" & JsonEscape(token) & """,""timestamp"":" & UtcUnixTimestamp() & ",""employees"":["
    For i = 1 To arr.Count
        If i > 1 Then payload = payload & ","
        payload = payload & arr(i)
    Next i
    payload = payload & "]}"

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 15000, 15000, 30000, 120000
    http.Open "POST", apiUrl, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.send Utf8Bytes(payload)
    response = http.responseText
    If http.Status <> 200 Then Err.Raise vbObjectError + 102, , "HTTP " & http.Status & ": " & response
    If InStr(1, response, """status"":""success""", vbTextCompare) = 0 Then _
        Err.Raise vbObjectError + 103, , "API tá»« chá»‘i: " & response

    sentCount = arr.Count
    SyncOneBranch = "Thanh cong, gui " & arr.Count & " nhan vien. " & response
    Exit Function

BranchError:
    On Error Resume Next
    SafeCloseWorkbook srcWb
    SyncOneBranch = "Lá»–I: " & Err.Description
    LogLocal "ERROR", branch, SyncOneBranch
End Function

Private Function FindNewestFile(ByVal folder As String, ByVal mask As String) As String
    Dim f As String, full As String, best As String, bestDate As Date
    If Right$(folder, 1) <> "\" And Right$(folder, 1) <> "/" Then folder = folder & "\"
    f = Dir(folder & mask)
    Do While f <> ""
        full = folder & f
        If best = "" Or FileDateTime(full) > bestDate Then
            best = full: bestDate = FileDateTime(full)
        End If
        f = Dir()
    Loop
    FindNewestFile = best
End Function

Private Function ColumnNumber(ByVal col As String) As Long
    Dim i As Long, n As Long, ch As Integer
    col = UCase$(Trim$(col))
    If col = "" Then ColumnNumber = 0: Exit Function
    If IsNumeric(col) Then ColumnNumber = CLng(col): Exit Function
    For i = 1 To Len(col)
        ch = Asc(Mid$(col, i, 1))
        If ch < 65 Or ch > 90 Then Err.Raise vbObjectError + 104, , "Cot khong hop le: " & col
        n = n * 26 + ch - 64
    Next i
    ColumnNumber = n
End Function

Private Function CellTextByCol(ws As Worksheet, ByVal rowNo As Long, ByVal col As String) As String
    If Trim$(col) = "" Then CellTextByCol = "" Else CellTextByCol = CellText(ws.Cells(rowNo, ColumnNumber(col)))
End Function

Private Function CellText(rng As Range) As String
    If IsError(rng.Value) Or IsEmpty(rng.Value) Then
        CellText = ""
    Else
        CellText = Trim$(CStr(rng.text))
        If CellText = "####" Then CellText = Trim$(CStr(rng.Value2))
    End If
End Function

Private Function DigitsOnly(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    DigitsOnly = out
End Function

Private Function UtcUnixTimestamp() As Long
    Dim st As SYSTEMTIME, dt As Date
    GetSystemTime st
    dt = DateSerial(st.wYear, st.wMonth, st.wDay) + _
         TimeSerial(st.wHour, st.wMinute, st.wSecond)
    UtcUnixTimestamp = DateDiff("s", #1/1/1970#, dt)
End Function

Private Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEscape = s
End Function

Private Function Utf8Bytes(ByVal text As String) As Variant
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")

    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text

    ' Chuyen sang Binary và bo UTF-8 BOM (3 byte)
    stream.Position = 0
    stream.Type = 1
    stream.Position = 3

    Utf8Bytes = stream.Read

    stream.Close
    Set stream = Nothing
End Function

Private Sub LogLocal(ByVal level As String, ByVal branch As String, ByVal message As String)
    Dim ws As Worksheet, nr As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("SyncRunLog")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "SyncRunLog"
        ws.Range("A1:D1").Value = Array("Time", "Level", "Branch", "Message")
    End If
    nr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nr, 1).Value = Now
    ws.Cells(nr, 2).Value = level
    ws.Cells(nr, 3).Value = branch
    ws.Cells(nr, 4).Value = message
End Sub
Private Sub SafeCloseWorkbook(ByRef wb As Workbook)

    On Error Resume Next

    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
        Set wb = Nothing
    End If

    Err.Clear
    On Error GoTo 0

End Sub
