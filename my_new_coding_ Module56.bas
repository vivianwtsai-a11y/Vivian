Attribute VB_Name = "Module521"

Option Explicit

Public Sub 產生新年度明細基本檔()
    Dim file1i As String
    Dim file2i As String
    Dim filePath As String
    Dim srcFullPath As String
    Dim dstFullPath As String
    Dim salNum As Long
    Dim i As Long
    Dim nyear As Long
    Dim oyear As Long
    Dim wyear As Long
    Dim nyearLabel As String
    Dim oyearLabel As String
    Dim fileNotExist As String
    Dim userData As String
    Dim wb As Workbook
    Dim criteria1 As String
    Dim criteria2 As String
    Dim iCnt As Long
    Dim keepCreateTokens As Variant
    Dim baseSheet As Worksheet
    Dim hadError As Boolean
    Dim errMsg As String

    Set baseSheet = ActiveSheet
    iCnt = 0

    userData = InputBox("請輸入新薪資明細基本檔的年份(ex.115年):", "製作新年度薪資明細基本檔")
    If StrPtr(userData) = 0 Then Exit Sub
    If Len(Trim$(userData)) = 0 Then Exit Sub

    If MsgBox("確定產生 " & userData & " 薪資明細基本檔？", vbYesNo + vbQuestion, "新年度薪資明細基本檔") = vbNo Then
        Exit Sub
    End If

    nyear = CLng(Val(Replace$(Trim$(userData), "年", vbNullString)))
    If nyear <= 0 Then
        MsgBox "年份格式錯誤，請輸入例如 115 或 115年。", vbExclamation, "新年度薪資明細基本檔"
        Exit Sub
    End If

    oyear = nyear - 1
    wyear = nyear + 1911
    nyearLabel = CStr(nyear) & "年"
    oyearLabel = CStr(oyear) & "年"

    If Len(ThisWorkbook.Path) > 0 Then
        filePath = ThisWorkbook.Path & Application.PathSeparator
    Else
        filePath = CurDir$ & Application.PathSeparator
    End If

    salNum = baseSheet.Cells(baseSheet.Rows.Count, 1).End(xlUp).Row
    criteria1 = oyearLabel & "12月"
    criteria2 = oyearLabel & "12月(2)"
    keepCreateTokens = Array(CStr(wyear) & "/1/", CStr(wyear) & "/01/")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo CleanFail

    For i = 6 To salNum
        file1i = oyearLabel & CStr(baseSheet.Cells(i, 6).Value) & "薪資明細.xlsx"
        file2i = nyearLabel & CStr(baseSheet.Cells(i, 6).Value) & "薪資明細.xlsx"

        srcFullPath = filePath & file1i
        dstFullPath = filePath & file2i

        If FileExists(srcFullPath) Then
            If FileExists(dstFullPath) Then
                Kill dstFullPath
            End If

            FileCopy srcFullPath, dstFullPath
            Set wb = Workbooks.Open(dstFullPath)

            DeleteUnneededSheets wb, oyearLabel
            FilterRowsByCreateTimeBlock wb, "拆帳表", keepCreateTokens
            FilterRowsByCreateTimeBlock wb, "AA碼季獎金", keepCreateTokens
            FilterRowsByCreateTimeBlock wb, "AA碼獎金", keepCreateTokens
            FilterRowsByMonth wb, "總表", criteria1, criteria2
            FilterRowsByMonth wb, "行政總表", criteria1, criteria2

            If WorksheetExists(wb, "總表") Then
                With wb.Worksheets("總表")
                    .Rows("9:16").Delete
                    With .Columns("A:AO").Font
                        .Name = "Microsoft JhengHei UI"
                        .Size = 10
                        .Underline = xlUnderlineStyleNone
                    End With
                End With
            End If

            wb.Save
            wb.Close SaveChanges:=False
            Set wb = Nothing

            iCnt = iCnt + 1
        Else
            fileNotExist = fileNotExist & baseSheet.Cells(i, 6).Value & vbCrLf
        End If
    Next i

    GoTo CleanExit

CleanFail:
    hadError = True
    errMsg = "處理時發生錯誤：" & Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
        Set wb = Nothing
    End If
    On Error GoTo 0

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If hadError Then
        MsgBox errMsg, vbCritical, "新年度薪資明細基本檔"
    Else
        MsgBox "製作薪資明細基本檔案 " & nyearLabel & " 共 " & iCnt & " 筆。", vbInformation, "新年度薪資明細基本檔"
        If Len(fileNotExist) > 0 Then
            MsgBox "無法製作薪資明細基本檔案名單:" & vbCrLf & fileNotExist, vbExclamation, "新年度薪資明細基本檔"
        End If
    End If
End Sub

Private Sub DeleteUnneededSheets(ByVal wb As Workbook, ByVal oyearLabel As String)
    Dim idx As Long
    Dim sh As Worksheet

    For idx = wb.Worksheets.Count To 1 Step -1
        Set sh = wb.Worksheets(idx)
        If Not ShouldKeepSheet(sh.Name, oyearLabel) Then
            If wb.Worksheets.Count > 1 Then
                sh.Delete
            End If
        End If
    Next idx
End Sub

Private Sub FilterRowsByCreateTimeBlock(ByVal wb As Workbook, ByVal targetSheetName As String, ByVal keepTokens As Variant)
    Dim ws As Worksheet
    Dim starts As Collection
    Dim keepRows() As Boolean
    Dim r As Long
    Dim idx As Long
    Dim blockStart As Long
    Dim blockEnd As Long
    Dim markStart As Long
    Dim firstStart As Long
    Dim lastRow As Long
    Dim headerText As String

    On Error Resume Next
    Set ws = wb.Worksheets(targetSheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then Exit Sub

    Set starts = New Collection
    For r = 1 To lastRow
        headerText = CStr(ws.Cells(r, 1).Value)
        If InStr(1, headerText, "建立時間", vbTextCompare) > 0 Then
            starts.Add r
        End If
    Next r

    If starts.Count = 0 Then
        Exit Sub
    End If

    firstStart = CLng(starts(1))
    ReDim keepRows(1 To lastRow)

    ' Keep static header area above the first "建立時間" block.
    If firstStart > 1 Then
        For r = 1 To firstStart - 1
            keepRows(r) = True
        Next r
    End If

    For idx = 1 To starts.Count
        blockStart = CLng(starts(idx))
        If idx < starts.Count Then
            blockEnd = CLng(starts(idx + 1)) - 1
        Else
            blockEnd = lastRow
        End If

        headerText = CStr(ws.Cells(blockStart, 1).Value)
        If TextContainsAnyToken(headerText, keepTokens) Then
            markStart = blockStart
            ' Keep one blank separator row before target block if present.
            If markStart > firstStart Then
                If Len(CStr(ws.Cells(markStart - 1, 1).Value)) = 0 Then
                    markStart = markStart - 1
                End If
            End If

            For r = markStart To blockEnd
                keepRows(r) = True
            Next r
        End If
    Next idx

    For r = lastRow To firstStart Step -1
        If Not keepRows(r) Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub

Private Sub FilterRowsByMonth(ByVal wb As Workbook, ByVal targetSheetName As String, ByVal criteria1 As String, ByVal criteria2 As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim v As String

    On Error Resume Next
    Set ws = wb.Worksheets(targetSheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = lastRow To 6 Step -1
        v = CStr(ws.Cells(r, 1).Value)
        If (v <> criteria1) And (v <> criteria2) Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub

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

Public Function FileExists(ByVal fullPath As String) As Boolean
    On Error GoTo EH
    FileExists = (Len(Dir$(fullPath, vbNormal)) > 0)
    Exit Function
EH:
    FileExists = False
End Function

Private Function WorksheetExists(ByVal wb As Workbook, ByVal wsName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(wsName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function ShouldKeepSheet(ByVal sheetName As String, ByVal oyearLabel As String) As Boolean
    Dim nameLower As String

    nameLower = LCase$(Trim$(sheetName))
    ShouldKeepSheet = (nameLower = "format") _
        Or (nameLower = "mformat") _
        Or (sheetName = "行政總表") _
        Or (sheetName = "總表") _
        Or (sheetName = (oyearLabel & "12月")) _
        Or (sheetName = (oyearLabel & "12月(2)")) _
        Or (sheetName = (oyearLabel & "12月行政")) _
        Or (sheetName = (oyearLabel & "12月(2)行政")) _
        Or (sheetName = "拆帳表") _
        Or (sheetName = "A碼清冊") _
        Or (sheetName = "AA碼季獎金") _
        Or (sheetName = "AA碼獎金")
End Function
