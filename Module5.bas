Attribute VB_Name = "Module5"

Sub newsalarydetail()
'
' newsalarydetail 巨集
' 產生新年度薪資明細
'

'
End Sub

Public Sub DeleteRowsByCriteria(ByVal wb As Workbook, ByVal targetSheetName As String, ByVal criteria3 As String)
    Dim ws1 As Worksheet
    Dim lastRow1 As Long
    Dim s As Long
    Dim rowText As String
    Dim v1 As String

    On Error Resume Next
    Set ws1 = wb.Worksheets(targetSheetName)
    On Error GoTo 0
    If ws1 Is Nothing Then Exit Sub

    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row

    ' Delete from bottom to top to avoid skipping rows.
    For s = lastRow1 To 6 Step -1
        rowText = CStr(ws1.Cells(s, 1).Value) & CStr(ws1.Cells(s, 2).Value)
        If Len(rowText) >= 13 Then
            v1 = Mid$(rowText, 7, 7)
        Else
            v1 = vbNullString
        End If

        If v1 <> criteria3 Then
            ws1.Rows(s).Delete
        End If
    Next s
End Sub
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
    salNum = Cells(Rows.Count, 1).End(p).Row
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
