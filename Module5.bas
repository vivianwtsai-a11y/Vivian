Attribute VB_Name = "Module5"
Option Explicit

Public Sub newsalarydetail()
    ' 維持舊巨集入口，直接轉呼叫主流程。
    巨集3
End Sub

Public Sub 巨集3()
    Dim sourceFileName As String
    Dim newFileName As String
    Dim filePath As String
    Dim employeeName As String
    Dim missingFiles As String
    Dim oldYearLabel As String
    Dim newYearInput As String
    Dim sheetName As String
    Dim newYearNumber As Long
    Dim lastRow As Long
    Dim i As Long
    Dim baseSheet As Worksheet
    Dim targetWorkbook As Workbook

    Set baseSheet = ActiveSheet
    sheetName = baseSheet.Name

    newYearInput = InputBox(sheetName & " - 請輸入新薪資明細基本檔的年份(ex.115年):", "製作新年度薪資明細基本檔")
    If StrPtr(newYearInput) = 0 Then
        Exit Sub
    End If

    newYearInput = Trim$(newYearInput)
    newYearNumber = ParseYearNumber(newYearInput)
    If newYearNumber <= 0 Then
        MsgBox "輸入的年份格式錯誤，請輸入例如 115年。", vbExclamation, "新年度薪資明細基本檔"
        Exit Sub
    End If

    If MsgBox(sheetName & " - 確定產生 " & CStr(newYearNumber) & "年 薪資明細？", vbYesNo + vbQuestion, "新年度薪資明細基本檔") = vbNo Then
        Exit Sub
    End If

    filePath = ThisWorkbook.Path
    If Len(filePath) = 0 Then
        MsgBox "目前活頁簿尚未儲存，無法判斷檔案路徑。", vbExclamation, "新年度薪資明細基本檔"
        Exit Sub
    End If
    If Right$(filePath, 1) <> Application.PathSeparator Then
        filePath = filePath & Application.PathSeparator
    End If

    oldYearLabel = CStr(newYearNumber - 1) & "年"
    lastRow = baseSheet.Cells(baseSheet.Rows.Count, 1).End(xlUp).Row
    If lastRow < 6 Then
        MsgBox "沒有可處理的人員資料。", vbInformation, "新年度薪資明細基本檔"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo CleanFail

    For i = 6 To lastRow
        employeeName = Trim$(CStr(baseSheet.Cells(i, 6).Value))
        If Len(employeeName) = 0 Then
            GoTo NextEmployee
        End If

        sourceFileName = oldYearLabel & employeeName & "薪資明細.xlsx"
        newFileName = CStr(newYearNumber) & "年" & employeeName & "薪資明細.xlsx"

        If FileExists(filePath & sourceFileName) Then
            If FileExists(filePath & newFileName) Then
                Kill filePath & newFileName
            End If

            FileCopy filePath & sourceFileName, filePath & newFileName
            Set targetWorkbook = Workbooks.Open(filePath & newFileName)

            CleanupSalaryWorkbook targetWorkbook, oldYearLabel

            targetWorkbook.Close SaveChanges:=True
            Set targetWorkbook = Nothing
        Else
            missingFiles = missingFiles & vbCrLf & sourceFileName
        End If

NextEmployee:
    Next i

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If Len(missingFiles) > 0 Then
        MsgBox "處理完成，但找不到以下舊年度檔案：" & vbCrLf & missingFiles, vbExclamation, "新年度薪資明細基本檔"
    Else
        MsgBox "處理完成。", vbInformation, "新年度薪資明細基本檔"
    End If
    Exit Sub

CleanFail:
    On Error Resume Next
    If Not targetWorkbook Is Nothing Then
        targetWorkbook.Close SaveChanges:=False
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "處理時發生錯誤：" & Err.Description, vbCritical, "新年度薪資明細基本檔"
End Sub

Private Sub CleanupSalaryWorkbook(ByVal wb As Workbook, ByVal oldYearLabel As String)
    Dim idx As Long
    Dim ws As Worksheet

    For idx = wb.Worksheets.Count To 1 Step -1
        Set ws = wb.Worksheets(idx)
        If Not ShouldKeepSheet(ws.Name, oldYearLabel) Then
            ws.Delete
        End If
    Next idx

    DeleteRowsExcept wb, "行政總表", oldYearLabel & "12月", oldYearLabel & "12月(2)"
    DeleteRowsExcept wb, "總表", oldYearLabel & "12月", oldYearLabel & "12月(2)"
End Sub

Private Sub DeleteRowsExcept(ByVal wb As Workbook, ByVal wsName As String, ByVal keepValue1 As String, ByVal keepValue2 As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim cellValue As String

    On Error Resume Next
    Set ws = wb.Worksheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = lastRow To 6 Step -1
        cellValue = Trim$(CStr(ws.Cells(r, 1).Value))
        If StrComp(cellValue, keepValue1, vbTextCompare) <> 0 And _
           StrComp(cellValue, keepValue2, vbTextCompare) <> 0 Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub

Private Function ShouldKeepSheet(ByVal sheetName As String, ByVal oldYearLabel As String) As Boolean
    Dim normalizedName As String

    normalizedName = LCase$(sheetName)
    Select Case normalizedName
        Case "format", "mformat", LCase$("行政總表"), LCase$("總表"), LCase$("拆帳表"), LCase$("a碼清冊")
            ShouldKeepSheet = True
        Case LCase$(oldYearLabel & "12月行政"), LCase$(oldYearLabel & "12月(2)行政"), _
             LCase$(oldYearLabel & "12月"), LCase$(oldYearLabel & "12月(2)")
            ShouldKeepSheet = True
        Case Else
            ShouldKeepSheet = False
    End Select
End Function

Private Function ParseYearNumber(ByVal yearText As String) As Long
    Dim rawYearText As String

    rawYearText = Trim$(Replace$(yearText, "年", vbNullString))
    If IsNumeric(rawYearText) Then
        ParseYearNumber = CLng(rawYearText)
    End If
End Function

Private Function FileExists(ByVal fileName As String) As Boolean
    FileExists = (Len(Dir$(fileName, vbNormal)) > 0)
End Function
