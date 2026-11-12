Sub CopySpecificColumnsToGoogleSheets()
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim folderPath As String, fileName As String
    Dim passwordExcel As String
    Dim lastRow As Long, i As Long
    Dim http As Object, url As String, payload As String
    Dim jsonRows As String
    
    ' === CAU HINH THONG TIN TAI DAY ===
    ' Hãy thay URL bên du?i b?ng URL b?n tri?n khai m?i nh?t c?a b?n n?u có thay d?i
    url = "https://script.google.com/macros/s/AKfycbxtfVxsHIdtX6Z5nSBbGSM52aD7oTiIYc8Yt_jv1o8OihVVE2VURsKJJmDjAzLxH7X7/exec"
    passwordExcel = "DSCNV"
    folderPath = "\\192.168.0.253\vn hr\DS + PN + TP - 2014\"
    
    Dim googleSpreadsheetId As String
    Dim googleSheetName As String
    googleSpreadsheetId = "1caxiuh1jyzZi8rkz1EqDkzwcr_gf5sIk4XzCh-XBjjQ"
    googleSheetName = "DSCNV"
    ' ==================================
    
    fileName = Dir(folderPath & "DSCNV-*.xlsb")
    If fileName = "" Then
        MsgBox "Khong tim thay file nao bat dau bang 'DSCNV' trong thu muc!", vbCritical, "Loi"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    
    On Error Resume Next
    Set wbSource = Workbooks.Open( _
        fileName:=folderPath & fileName, _
        UpdateLinks:=0, _
        ReadOnly:=True, _
        password:=passwordExcel, _
        WriteResPassword:=passwordExcel, _
        IgnoreReadOnlyRecommended:=True)
    On Error GoTo 0
    
    If wbSource Is Nothing Then
        MsgBox "Khong the mo duoc file. Vui long kiem tra lai thu muc hoac mat khau!", vbCritical, "Loi"
        GoTo CleanExit
    End If
    
    On Error Resume Next
    Set wsSource = wbSource.Sheets("DSCNV")
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        MsgBox "Khong tim thay sheet 'DSCNV' trong file " & fileName, vbCritical, "Loi"
        wbSource.Close SaveChanges:=False
        GoTo CleanExit
    End If
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    lastRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).row
    
    jsonRows = ""
    Dim colH_Value As String
    Dim validRowCount As Long
    validRowCount = 0
    
    For i = 8 To lastRow
        colH_Value = SafeCell(wsSource.Cells(i, 8))
        
        colH_Value = Trim(colH_Value)
        If Left(colH_Value, 1) = "'" Then
            colH_Value = Mid(colH_Value, 2)
        End If
        colH_Value = Trim(colH_Value)
        
        If Len(colH_Value) = 12 Then
            validRowCount = validRowCount + 1
            
            Dim colB As String, colD As String, colE As String, colF As String, colG As String
            Dim colAF As String, colAG As String, colAH As String, colAM As String
            Dim colAO As String, colAP As String, colAQ As String, colAR As String, colAT As String, colAW As String
            
            colB = SafeCell(wsSource.Cells(i, 2))
            colD = SafeCell(wsSource.Cells(i, 4))
            colE = SafeCell(wsSource.Cells(i, 5))
            ' Ép dinh dang Text bat buoc dd/mm/yyyy tu Excel d? d?y sang làm s?ch
            colF = wsSource.Cells(i, 6).Text ' Ngày sinh Nam
            colG = wsSource.Cells(i, 7).Text ' Ngày sinh Nu
            
            
            colH_Value = "'" & colH_Value
            
            colAF = SafeCell(wsSource.Cells(i, 32))
            colAG = SafeCell(wsSource.Cells(i, 33))
            colAH = SafeCell(wsSource.Cells(i, 34))
            colAM = SafeCell(wsSource.Cells(i, 38))
            colAO = SafeCell(wsSource.Cells(i, 41))
            colAP = SafeCell(wsSource.Cells(i, 42))
            colAQ = SafeCell(wsSource.Cells(i, 43))
            colAR = SafeCell(wsSource.Cells(i, 44))
            colAT = SafeCell(wsSource.Cells(i, 46))
            colAW = SafeCell(wsSource.Cells(i, 49))
            
            Dim rowJson As String
            rowJson = "[""" & colB & """,""" & colD & """,""" & colE & """,""" & colF & """,""" & colG & """,""" & colH_Value & """,""" & colAF & """,""" & colAG & """,""" & colAH & """,""" & colAM & """,""" & colAO & """,""" & colAP & """,""" & colAQ & """,""" & colAR & """,""" & colAT & """,""" & colAW & """]"
            
            If jsonRows = "" Then
                jsonRows = rowJson
            Else
                jsonRows = jsonRows & "," & rowJson
            End If
        End If
    Next i
    
    wbSource.Close SaveChanges:=False
    
    If validRowCount = 0 Then
        MsgBox "Khong tim thay dong nao co chua so can cuoc hop le (du 12 ky tu) tai cot H!", vbExclamation, "Thong bao"
        GoTo CleanExit
    End If
    
    payload = "{""spreadsheet_id"":""" & googleSpreadsheetId & """,""sheet_name"":""" & googleSheetName & """,""values"":[" & jsonRows & "]}"

    On Error Resume Next
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.send payload
    
    If Err.Number <> 0 Then
        MsgBox "Loi ket noi Internet: " & Err.Description, vbCritical, "That Bai"
        GoTo CleanExit
    End If
    On Error GoTo 0
    
    If http.Status = 200 Then
        Dim responseText As String
        responseText = http.responseText
        
        ' Hi?n th? chu?i thông báo k?t qu? tr? v? t? Google
        MsgBox responseText, vbInformation, "Hoan Tat"
    Else
        MsgBox "Loi ket noi den may chu Google: " & http.Status & " - " & http.statusText, vbExclamation, "Loi Dong Bo"
    End If

CleanExit:
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True
End Sub

Function SafeCell(rng As Range) As String
    If IsError(rng.Value) Then
        SafeCell = ""
    ElseIf IsEmpty(rng.Value) Then
        SafeCell = ""
    Else
        Dim txt As String
        txt = CStr(rng.Value)
        txt = Replace(txt, "\", "\\")
        txt = Replace(txt, """", "\""")
        txt = Replace(txt, vbCrLf, "\n")
        txt = Replace(txt, vbCr, "\n")
        txt = Replace(txt, vbLf, "\n")
        txt = Replace(txt, vbTab, "\t")
        SafeCell = txt
    End If
End Function


