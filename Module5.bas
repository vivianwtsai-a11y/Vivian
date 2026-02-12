Attribute VB_Name = "Module5"

Option Explicit

Sub newsalarydetail()
'
' newsalarydetail 巨集
' 產生新年度薪資明細
'

'
    巨集3
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
    Dim userData As String
    Dim nyearNum As Long
    Dim oyearNum As Long
    Dim nyear As String
    Dim oyear As String
    Dim missingCount As Long
    Dim keep1 As String
    Dim keep2 As String
    Dim v As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim idx As Long

    sheetName = ActiveSheet.Name
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "請先儲存本活頁簿，才能取得檔案路徑並進行複製。", vbExclamation, "新年度薪資明細基本檔"
        Exit Sub
    End If
    filePath = ThisWorkbook.Path & Application.PathSeparator

    userData = InputBox(sheetName & " - 請輸入新薪資明細基本檔的年份(ex.115年):", "製作新年度薪資明細基本檔")
    If StrPtr(userData) = 0 Then Exit Sub

    nyearNum = CLng(Val(userData))
    If nyearNum <= 0 Then
        MsgBox "年份輸入無效: " & userData, vbExclamation, "新年度薪資明細基本檔"
        Exit Sub
    End If

    nyear = CStr(nyearNum) & "年"
    oyearNum = nyearNum - 1
    oyear = CStr(oyearNum) & "年"
    keep1 = oyear & "12月"
    keep2 = oyear & "12月(2)"

    If MsgBox(sheetName & " - 確定產生" & nyear & "薪資明細", vbYesNo, "新年度薪資明細基本檔") = vbNo Then Exit Sub

    salNum = Cells(Rows.Count, 1).End(xlUp).Row
    missingCount = 0

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo CleanFail

    For i = 6 To salNum
        file1i = oyear & CStr(Cells(i, 6).Value) & "薪資明細.xlsx"
        If FileExists(filePath & file1i) Then
            file2i = nyear & CStr(Cells(i, 6).Value) & "薪資明細.xlsx"

            If FileExists(filePath & file2i) Then
                On Error Resume Next
                Kill filePath & file2i
                On Error GoTo CleanFail
            End If

            FileCopy filePath & file1i, filePath & file2i

            Set wb = Workbooks.Open(filePath & file2i)

            ' 刪除不需要的工作表（倒序避免跳過）
            For idx = wb.Worksheets.Count To 1 Step -1
                Set ws = wb.Worksheets(idx)
                If Not ShouldKeepSheet(LCase$(ws.Name), LCase$(oyear)) Then
                    ws.Delete
                End If
            Next idx

            ' 行政總表：僅保留上一年度 12月 / 12月(2) 的資料列
            If SheetExists(wb, "行政總表") Then
                With wb.Worksheets("行政總表")
                    rowNum = .Cells(.Rows.Count, 1).End(xlUp).Row
                    For j = rowNum To 6 Step -1
                        v = CStr(.Cells(j, 1).Value)
                        If v <> keep1 And v <> keep2 Then
                            .Rows(j).Delete
                        End If
                    Next j
                End With
            End If

            ' 總表：僅保留上一年度 12月 / 12月(2) 的資料列
            If SheetExists(wb, "總表") Then
                With wb.Worksheets("總表")
                    rowNum1 = .Cells(.Rows.Count, 1).End(xlUp).Row
                    For k = rowNum1 To 6 Step -1
                        v = CStr(.Cells(k, 1).Value)
                        If v <> keep1 And v <> keep2 Then
                            .Rows(k).Delete
                        End If
                    Next k
                End With
            End If

            wb.Save
            wb.Close SaveChanges:=True
        Else
            missingCount = missingCount + 1
        End If
    Next i

CleanExit:
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    On Error GoTo 0

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "執行中發生錯誤: " & Err.Number & vbCrLf & Err.Description, vbExclamation, "新年度薪資明細基本檔"
    Else
        MsgBox "完成。" & vbCrLf & "找不到舊檔案筆數: " & missingCount, vbInformation, "新年度薪資明細基本檔"
    End If
    Exit Sub

CleanFail:
    Resume CleanExit
End Sub

Private Function FileExists(ByVal fullPath As String) As Boolean
    FileExists = (Len(Dir$(fullPath, vbNormal)) > 0)
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function ShouldKeepSheet(ByVal lowerName As String, ByVal lowerOyear As String) As Boolean
    Select Case lowerName
        Case "format", "mformat", "行政總表", "總表", "拆帳表", "a碼清冊"
            ShouldKeepSheet = True
        Case LCase$(lowerOyear & "12月行政"), LCase$(lowerOyear & "12月(2)行政"), LCase$(lowerOyear & "12月")
            ShouldKeepSheet = True
        Case Else
            ShouldKeepSheet = False
    End Select
End Function
