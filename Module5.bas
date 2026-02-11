Attribute VB_Name = "Module5"

Option Explicit

Public Sub newsalarydetail()
    巨集3
End Sub

Public Sub 巨集3()
    Dim sourceFileName As String
    Dim newFileName As String
    Dim filePath As String
    Dim targetName As String
    Dim missingFiles As String
    Dim oldYearLabel As String
    Dim newYearInput As String
    Dim sheetName As String
    Dim newYearNumber As Long
    Dim oldYearNumber As Long
    Dim newADYear As Long
    Dim oldADYear As Long
    Dim lastRow As Long
    Dim i As Long
    Dim baseSheet As Worksheet
    Dim targetWorkbook As Workbook
    Dim hadError As Boolean
    Dim errorText As String

    Set baseSheet = ActiveSheet
    sheetName = baseSheet.Name

    newYearInput = InputBox(sheetName & " - 請輸入新薪資明細基本檔的年份(ex.115年):", "製作新年度薪資明細基本檔")
    If StrPtr(newYearInput) = 0 Then
        Exit Sub
    End If

    newYearInput = Trim$(newYearInput)
    newYearNumber = ParseYearNumber(newYearInput)
    If newYearNumber <= 0 Then
        MsgBox "年份格式錯誤，請輸入像 115年 或 115。", vbExclamation, "製作新年度薪資明細基本檔"
        Exit Sub
    End If

    If MsgBox(sheetName & " - 確定產生 " & CStr(newYearNumber) & "年 薪資明細？", vbYesNo + vbQuestion, "新年度薪資明細基本檔") = vbNo Then
        Exit Sub
    End If

    oldYearNumber = newYearNumber - 1
    oldYearLabel = CStr(oldYearNumber) & "年"
    newADYear = ToGregorianYear(newYearNumber)
    oldADYear = newADYear - 1

    filePath = ThisWorkbook.Path
    If Len(filePath) = 0 Then
        MsgBox "目前活頁簿尚未儲存，無法判斷檔案路徑。", vbExclamation, "新年度薪資明細基本檔"
        Exit Sub
    End If
    If Right$(filePath, 1) <> Application.PathSeparator Then
        filePath = filePath & Application.PathSeparator
    End If

    lastRow = baseSheet.Cells(baseSheet.Rows.Count, 1).End(xlUp).Row
    If lastRow < 6 Then
        MsgBox "第 6 列之後沒有可處理資料。", vbInformation, "新年度薪資明細基本檔"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo CleanFail

    For i = 6 To lastRow
        targetName = Trim$(CStr(baseSheet.Cells(i, 6).Value))
        If Len(targetName) = 0 Then
            GoTo NextTarget
        End If

        sourceFileName = oldYearLabel & targetName & "薪資明細.xlsx"
        newFileName = CStr(newYearNumber) & "年" & targetName & "薪資明細.xlsx"

        If FileExists(filePath & sourceFileName) Then
            If FileExists(filePath & newFileName) Then
                Kill filePath & newFileName
            End If

            FileCopy filePath & sourceFileName, filePath & newFileName
            Set targetWorkbook = Workbooks.Open(filePath & newFileName)

            CleanupSalaryWorkbook targetWorkbook, oldYearNumber, newYearNumber, oldADYear, newADYear

            targetWorkbook.Close SaveChanges:=True
            Set targetWorkbook = Nothing
        Else
            missingFiles = missingFiles & vbCrLf & sourceFileName
        End If

NextTarget:
    Next i

    GoTo CleanExit

CleanFail:
    hadError = True
    errorText = "處理時發生錯誤：" & Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not targetWorkbook Is Nothing Then
        targetWorkbook.Close SaveChanges:=False
        Set targetWorkbook = Nothing
    End If
    On Error GoTo 0

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If hadError Then
        MsgBox errorText, vbCritical, "新年度薪資明細基本檔"
        Exit Sub
    End If

    If Len(missingFiles) > 0 Then
        MsgBox "處理完成，但找不到以下來源檔：" & vbCrLf & missingFiles, vbExclamation, "新年度薪資明細基本檔"
    Else
        MsgBox "處理完成。", vbInformation, "新年度薪資明細基本檔"
    End If
End Sub

Private Sub CleanupSalaryWorkbook(ByVal wb As Workbook, ByVal oldYearNumber As Long, ByVal newYearNumber As Long, ByVal oldADYear As Long, ByVal newADYear As Long)
    Dim oldYearLabel As String
    Dim newYearLabel As String
    Dim keepSummaryValues As Variant
    Dim keepDateTokens As Variant

    oldYearLabel = CStr(oldYearNumber) & "年"
    newYearLabel = CStr(newYearNumber) & "年"

    DeleteUnneededSheets wb, oldYearLabel

    keepSummaryValues = Array( _
        oldYearLabel & "12月", oldYearLabel & "12月(2)", _
        newYearLabel & "1月", newYearLabel & "1月(2)", _
        newYearLabel & "01月", newYearLabel & "01月(2)" _
    )
    DeleteRowsByColumnAValues wb, "行政總表", keepSummaryValues, 6
    DeleteRowsByColumnAValues wb, "總表", keepSummaryValues, 6

    keepDateTokens = BuildDateTokens(oldYearNumber, newYearNumber, oldADYear, newADYear)
    DeleteRowsByRowTokenMatch wb, "拆帳表", keepDateTokens, 6
    DeleteRowsByRowTokenMatch wb, "AA碼獎金", keepDateTokens, 6
    DeleteRowsByRowTokenMatch wb, "A碼獎金", keepDateTokens, 6
End Sub

Private Sub DeleteUnneededSheets(ByVal wb As Workbook, ByVal oldYearLabel As String)
    Dim idx As Long
    Dim ws As Worksheet

    For idx = wb.Worksheets.Count To 1 Step -1
        Set ws = wb.Worksheets(idx)
        If Not ShouldKeepSheet(ws.Name, oldYearLabel) Then
            If wb.Worksheets.Count > 1 Then
                ws.Delete
            End If
        End If
    Next idx
End Sub

Private Function ShouldKeepSheet(ByVal sheetName As String, ByVal oldYearLabel As String) As Boolean
    Dim normalizedName As String

    normalizedName = LCase$(Trim$(sheetName))
    Select Case normalizedName
        Case "format", "mformat", LCase$("行政總表"), LCase$("總表"), LCase$("拆帳表"), LCase$("a碼清冊"), LCase$("aa碼獎金"), LCase$("a碼獎金")
            ShouldKeepSheet = True
        Case LCase$(oldYearLabel & "12月行政"), LCase$(oldYearLabel & "12月(2)行政"), _
             LCase$(oldYearLabel & "12月"), LCase$(oldYearLabel & "12月(2)")
            ShouldKeepSheet = True
        Case Else
            ShouldKeepSheet = False
    End Select
End Function

Private Sub DeleteRowsByColumnAValues(ByVal wb As Workbook, ByVal wsName As String, ByVal keepValues As Variant, ByVal startRow As Long)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim cellValue As String

    If Not WorksheetExists(wb, wsName) Then
        Exit Sub
    End If

    Set ws = wb.Worksheets(wsName)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < startRow Then
        Exit Sub
    End If

    For r = lastRow To startRow Step -1
        cellValue = Trim$(CStr(ws.Cells(r, 1).Value))
        If Not ValueInArray(cellValue, keepValues) Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub

Private Sub DeleteRowsByRowTokenMatch(ByVal wb As Workbook, ByVal wsName As String, ByVal keepTokens As Variant, ByVal startRow As Long)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long

    If Not WorksheetExists(wb, wsName) Then
        Exit Sub
    End If

    Set ws = wb.Worksheets(wsName)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < startRow Then
        Exit Sub
    End If

    lastCol = LastUsedColumn(ws)
    If lastCol < 1 Then
        lastCol = 1
    End If

    For r = lastRow To startRow Step -1
        If Not RowContainsAnyToken(ws, r, lastCol, keepTokens) Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub

Private Function ValueInArray(ByVal valueText As String, ByVal values As Variant) As Boolean
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If StrComp(valueText, CStr(values(i)), vbTextCompare) = 0 Then
            ValueInArray = True
            Exit Function
        End If
    Next i
    ValueInArray = False
End Function

Private Function RowContainsAnyToken(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal lastCol As Long, ByVal tokens As Variant) As Boolean
    Dim c As Long
    Dim searchText As String

    For c = 1 To lastCol
        searchText = BuildCellSearchText(ws.Cells(rowNumber, c))
        If TextContainsAnyToken(searchText, tokens) Then
            RowContainsAnyToken = True
            Exit Function
        End If
    Next c
End Function

Private Function BuildCellSearchText(ByVal targetCell As Range) As String
    Dim rawValue As Variant

    rawValue = targetCell.Value
    If IsError(rawValue) Then
        BuildCellSearchText = vbNullString
        Exit Function
    End If

    If IsDate(rawValue) Then
        BuildCellSearchText = CStr(targetCell.Text) & "|" & Format$(CDate(rawValue), "yyyy/m/d") & "|" & Format$(CDate(rawValue), "yyyy/m/")
    Else
        BuildCellSearchText = CStr(rawValue)
    End If
End Function

Private Function TextContainsAnyToken(ByVal sourceText As String, ByVal tokens As Variant) As Boolean
    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
        If InStr(1, sourceText, CStr(tokens(i)), vbTextCompare) > 0 Then
            TextContainsAnyToken = True
            Exit Function
        End If
    Next i
    TextContainsAnyToken = False
End Function

Private Function BuildDateTokens(ByVal oldYearNumber As Long, ByVal newYearNumber As Long, ByVal oldADYear As Long, ByVal newADYear As Long) As Variant
    BuildDateTokens = Array( _
        CStr(oldYearNumber) & "年12月", _
        CStr(newYearNumber) & "年1月", _
        CStr(newYearNumber) & "年01月", _
        CStr(oldADYear) & "/12/", _
        CStr(oldADYear) & "/12", _
        CStr(oldADYear) & "-12-", _
        CStr(newADYear) & "/1/", _
        CStr(newADYear) & "/01/", _
        CStr(newADYear) & "-1-", _
        CStr(newADYear) & "-01-" _
    )
End Function

Private Function LastUsedColumn(ByVal ws As Worksheet) As Long
    Dim lastCell As Range

    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0

    If lastCell Is Nothing Then
        LastUsedColumn = 0
    Else
        LastUsedColumn = lastCell.Column
    End If
End Function

Private Function ParseYearNumber(ByVal yearText As String) As Long
    Dim normalizedText As String
    Dim parsedValue As Long

    normalizedText = Trim$(yearText)
    normalizedText = Replace$(normalizedText, "年", vbNullString)

    parsedValue = CLng(Val(normalizedText))
    If parsedValue > 0 Then
        ParseYearNumber = parsedValue
    End If
End Function

Private Function ToGregorianYear(ByVal yearNumber As Long) As Long
    If yearNumber >= 1911 Then
        ToGregorianYear = yearNumber
    Else
        ToGregorianYear = yearNumber + 1911
    End If
End Function

Private Function WorksheetExists(ByVal wb As Workbook, ByVal wsName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(wsName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function FileExists(ByVal fullPath As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(fullPath, vbNormal)) > 0)
    On Error GoTo 0
End Function
