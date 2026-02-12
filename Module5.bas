Attribute VB_Name = "Module5"

Option Explicit

Sub newsalarydetail()
'
' newsalarydetail 巨集
' 產生新年度薪資明細
'

'
End Sub
Sub 巨集3()
'
' 巨集3 巨集
' 產生新年度明細
'

    Dim file1i As String
    Dim file2i As String
    Dim filePath As String
    Dim salNum As Long
    Dim rowNum As Long
    Dim rowNum1 As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim sheetName As String
    Dim inputYear As String
    Dim yearNum As Long
    Dim nyearText As String
    Dim oyearText As String
    Dim deptName As String
    Dim wb As Workbook
    Dim keepNames As Variant
    Dim sSheet As Worksheet

    sheetName = ActiveSheet.Name
    Set sSheet = ActiveSheet

    inputYear = InputBox(sheetName & " - 請輸入新薪資明細基本檔的年份(ex.115年):", "製作新年度薪資明細基本檔")
    If StrPtr(inputYear) = 0 Then Exit Sub '使用者按取消
    inputYear = Trim$(inputYear)
    If inputYear = vbNullString Then Exit Sub

    yearNum = Val(inputYear) '允許輸入「115年」或「115」
    If yearNum <= 0 Then
        MsgBox sheetName & " - 年份格式不正確，請輸入如 115 或 115年。", vbExclamation, "新年度薪資明細基本檔"
        Exit Sub
    End If

    nyearText = CStr(yearNum) & "年"
    oyearText = CStr(yearNum - 1) & "年"

    If MsgBox(sheetName & " - 確定產生 " & nyearText & " 薪資明細？", vbYesNo, "新年度薪資明細基本檔") = vbNo Then
        Exit Sub
    End If

    filePath = ThisWorkbook.Path
    If Len(filePath) = 0 Then
        MsgBox "無法取得檔案路徑，請先儲存此活頁簿。", vbExclamation, "新年度薪資明細基本檔"
        Exit Sub
    End If
    If Right$(filePath, 1) <> "\" Then filePath = filePath & "\"

    salNum = sSheet.Cells(sSheet.Rows.Count, 1).End(xlUp).Row
    If salNum < 6 Then
        MsgBox sheetName & " - 第 6 列開始沒有可處理資料。", vbInformation, "新年度薪資明細基本檔"
        Exit Sub
    End If

    keepNames = Array( _
        "format", "Mformat", "行政總表", "總表", "拆帳表", _
        oyearText & "12月行政", oyearText & "12月(2)行政", _
        oyearText & "12月", "A碼清冊" _
    )

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo CleanFail

    For i = 6 To salNum
        deptName = Trim$(CStr(sSheet.Cells(i, 6).Value))
        If deptName <> vbNullString Then
            file1i = oyearText & deptName & "薪資明細.xlsx"
            If FileExists(filePath & file1i) Then
                file2i = nyearText & deptName & "薪資明細.xlsx"
                ' 若目標檔已存在，先移除避免 FileCopy 直接報錯
                If FileExists(filePath & file2i) Then
                    On Error Resume Next
                    Kill filePath & file2i
                    On Error GoTo CleanFail
                End If

                FileCopy filePath & file1i, filePath & file2i

                Set wb = Workbooks.Open(filePath & file2i)

                DeleteSheetsExcept wb, keepNames
                FilterSheetKeepTwoMonths wb, "行政總表", oyearText & "12月", oyearText & "12月(2)"
                FilterSheetKeepTwoMonths wb, "總表", oyearText & "12月", oyearText & "12月(2)"

                wb.Save
                wb.Close SaveChanges:=False
                Set wb = Nothing
            End If
        End If
    Next i

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "處理過程發生錯誤：" & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation, "新年度薪資明細基本檔"
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set wb = Nothing
    Resume CleanExit
End Sub

Private Function FileExists(ByVal fullPath As String) As Boolean
    On Error GoTo EH
    FileExists = (Len(Dir$(fullPath, vbNormal)) > 0)
    Exit Function
EH:
    FileExists = False
End Function

Private Function WorksheetExists(ByVal wb As Workbook, ByVal wsName As String) As Boolean
    Dim ws As Worksheet
    On Error GoTo EH
    Set ws = wb.Worksheets(wsName)
    WorksheetExists = True
    Exit Function
EH:
    WorksheetExists = False
End Function

Private Function NameInList(ByVal sheetName As String, ByVal names As Variant) As Boolean
    Dim i As Long
    For i = LBound(names) To UBound(names)
        If LCase$(CStr(names(i))) = LCase$(sheetName) Then
            NameInList = True
            Exit Function
        End If
    Next i
    NameInList = False
End Function

Private Sub DeleteSheetsExcept(ByVal wb As Workbook, ByVal keepNames As Variant)
    Dim idx As Long
    Dim ws As Worksheet

    ' 由後往前刪除，避免集合在迴圈中被修改
    For idx = wb.Worksheets.Count To 1 Step -1
        Set ws = wb.Worksheets(idx)
        If Not NameInList(ws.Name, keepNames) Then
            If wb.Worksheets.Count > 1 Then ws.Delete
        End If
    Next idx
End Sub

Private Sub FilterSheetKeepTwoMonths(ByVal wb As Workbook, ByVal wsName As String, ByVal keepValue1 As String, ByVal keepValue2 As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim v As String

    If Not WorksheetExists(wb, wsName) Then Exit Sub

    Set ws = wb.Worksheets(wsName)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 6 Then Exit Sub

    For r = lastRow To 6 Step -1
        v = CStr(ws.Cells(r, 1).Value)
        If v <> keepValue1 And v <> keepValue2 Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub
