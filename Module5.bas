Attribute VB_Name = "Module5"

Sub newsalarydetail()
'
' newsalarydetail 巨集
' 產生新年度薪資明細
'

'
End Sub

Function LastRowOfLastTable(ws As Worksheet) As Long
    Dim lo As ListObject
    Dim lastTable As ListObject
    Dim lastTableTop As Long
    Dim lastTableLeft As Long
    Dim currentTop As Long
    Dim currentLeft As Long

    If ws Is Nothing Then
        LastRowOfLastTable = 0
        Exit Function
    End If

    If ws.ListObjects.Count = 0 Then
        LastRowOfLastTable = 0
        Exit Function
    End If

    For Each lo In ws.ListObjects
        currentTop = lo.Range.Row
        currentLeft = lo.Range.Column

        ' Pick the "last" table by position: lower rows first, then rightmost column.
        If lastTable Is Nothing _
            Or currentTop > lastTableTop _
            Or (currentTop = lastTableTop And currentLeft > lastTableLeft) Then
            Set lastTable = lo
            lastTableTop = currentTop
            lastTableLeft = currentLeft
        End If
    Next lo

    If lastTable.DataBodyRange Is Nothing Then
        ' Table has headers only.
        LastRowOfLastTable = lastTable.HeaderRowRange.Row
    Else
        LastRowOfLastTable = lastTable.DataBodyRange.Row + lastTable.DataBodyRange.Rows.Count - 1
    End If
End Function

Sub 巨集3()
'
' 巨集3 巨集
' 產生新年度明細
'

'   Dim file1i As String
    Dim file2i As String
    Dim filePath As String
    Dim salNum As Integer
    Dim rowNum As Long
    Dim rowNum1 As Long
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim sh As Worksheet
    Dim r As Long
    Dim nyear As Integer
    Dim oyear As Integer
    Dim fileNotExist As String
    Dim iCnt As Integer
    Dim jCnt As Integer
    Dim kCnt As Integer
    Dim sheetName As String
    Dim row_range As String
    Dim newsheetName As String
    Dim insurance1 As String
    Dim insurance2 As String
    Dim persentage As String
    Dim activeIndex As Integer
    Dim sSheet As Worksheet
    Dim dSheet As Worksheet
    Dim copyValue As Variant
    sheetName = ActiveSheet.Name
    iCnt = 0
    Userdata = InputBox(sheetName & " - 請輸入新薪資明細基本檔的年份(ex.115年):", "製作新年度薪資明細基本檔")
    If StrPtr(Userdata) = 0 Then
        Exit Sub
    End If
    If MsgBox(sheetName & " - 確定產生" & Userdata & "薪資明細", vbYesNo, "新年度薪資明細基本檔") = vbNo Then
        Exit Sub
    End If
    oyear = Mid(Userdata, 1, 3) - 1 & "年"
    salNum = LastRowOfLastTable(ActiveSheet)
    If salNum = 0 Then
        MsgBox "目前工作表找不到任何表格 (Table)。", vbExclamation
        Exit Sub
    End If
    MsgBox "salNum=" & salNum
    For i = 6 To salNum
        Rows(i).Select
        file1i = "oyear" + "年" + Cells(i, 6) + "薪資明細.xlsx"
        MsgBox "File ：" & file1i & " - " & FileExists(filePath & file1i)
        If FileExists(filePath & file1i) Then
           file2i = "nyear" + "年" + Cells(i, 6) + "薪資明細.xlsx"
           FileCopy file1i, file2i
           Workbooks.Open (filePath & file2i)
           Windows(file2i).Activate
           For Each sh In Worksheets
            If LCase(sh.Name) <> LCase("format" Or "Mformat" Or "行政總表" Or "總表" Or "拆帳表" Or "oyear" + "年" + "12月行政" Or "oyear" + "年" + "12月(2)行政" Or "oyear" + "年12月" Or "A碼清冊") Then
               Sheets.Delete
            End If
            Next sh
            Sheets("行政總表").Activate
            rowNum = Cells(Rows.Count, 1).End(xlUp).Row  '最後一列
            For j = 6 To rowNum
                If Cells(j, 1) <> "oyear" + "年12月" Or "oyear" + "年" + "12月(2)" Then
                Rows(j).Delete
                End If
                Next j
            Sheets("總表").Activate
            rowNum1 = Cells(Rows.Count, 1).End(xlUp).Row   '最後一列
            For k = 6 To rowNum
                If Cells(j, 1) <> "oyear" + "年12月" Or "oyear" + "年" + "12月(2)" Then
                Rows(k).Delete
                End If
                Next k
   Next i
End Sub
