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
    Dim srcFullPath As String
    Dim dstFullPath As String
    Dim salNum As Long
    Dim rowNum As Long
    Dim rowNum1 As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim nyear As Long
    Dim oyear As Long
    Dim nyearLabel As String
    Dim oyearLabel As String
    Dim sheetName As String
    Dim userData As String
    Dim wb As Workbook
    Dim criteria1 As String
    Dim criteria2 As String

    sheetName = ActiveSheet.Name
    
    userData = InputBox(sheetName & " - 請輸入新薪資明細基本檔的年份(ex.115年):", "製作新年度薪資明細基本檔")
    If Len(Trim$(userData)) = 0 Then Exit Sub
    
    If MsgBox(sheetName & " - 確定產生 " & userData & " 薪資明細", vbYesNo, "新年度薪資明細基本檔") = vbNo Then
        Exit Sub
    End If
    
    nyear = CLng(Val(userData))
    If nyear <= 0 Then Exit Sub
    
    oyear = nyear - 1
    nyearLabel = CStr(nyear) & "年"
    oyearLabel = CStr(oyear) & "年"
    
    If Len(ThisWorkbook.Path) > 0 Then
        filePath = ThisWorkbook.Path & Application.PathSeparator
    Else
        filePath = CurDir$ & Application.PathSeparator
    End If
    
    salNum = Cells(Rows.Count, 1).End(xlUp).Row
    
    criteria1 = oyearLabel & "12月"
    criteria2 = oyearLabel & "12月(2)"
    
    For i = 6 To salNum
        file1i = oyearLabel & CStr(Cells(i, 6).Value) & "薪資明細.xlsx"
        file2i = nyearLabel & CStr(Cells(i, 6).Value) & "薪資明細.xlsx"
        
        srcFullPath = filePath & file1i
        dstFullPath = filePath & file2i
        
        If FileExists(srcFullPath) Then
            FileCopy srcFullPath, dstFullPath
            
            Set wb = Workbooks.Open(dstFullPath)
            
            Application.DisplayAlerts = False
            DeleteUnneededSheets wb, oyearLabel
            Application.DisplayAlerts = True
            
            FilterRowsByMonth wb, "行政總表", criteria1, criteria2
            FilterRowsByMonth wb, "總表", criteria1, criteria2
            
            wb.Save
            wb.Close SaveChanges:=False
        End If
    Next i
End Sub

Private Sub DeleteUnneededSheets(ByVal wb As Workbook, ByVal oyearLabel As String)
    Dim idx As Long
    Dim sh As Worksheet
    
    For idx = wb.Worksheets.Count To 1 Step -1
        Set sh = wb.Worksheets(idx)
        If Not ShouldKeepSheet(sh.Name, oyearLabel) Then
            sh.Delete
        End If
    Next idx
End Sub

Private Function ShouldKeepSheet(ByVal sheetName As String, ByVal oyearLabel As String) As Boolean
    Dim nameLower As String
    nameLower = LCase$(sheetName)
    
    ShouldKeepSheet = (nameLower = "format") _
        Or (nameLower = "mformat") _
        Or (sheetName = "行政總表") _
        Or (sheetName = "總表") _
        Or (sheetName = "拆帳表") _
        Or (sheetName = (oyearLabel & "12月行政")) _
        Or (sheetName = (oyearLabel & "12月(2)行政")) _
        Or (sheetName = (oyearLabel & "12月")) _
        Or (sheetName = "A碼清冊")
End Function

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

Public Function FileExists(ByVal fullPath As String) As Boolean
    On Error GoTo EH
    FileExists = (Len(Dir$(fullPath, vbNormal)) > 0)
    Exit Function
EH:
    FileExists = False
End Function
