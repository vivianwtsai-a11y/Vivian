Attribute VB_Name = "Module521"

Sub 產生新年度明細基本檔()
'
' 巨集3 巨集
' 產生新年度明細
'

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
    Dim wyear As Long
    Dim nyearLabel As String
    Dim oyearLabel As String
    Dim wyearLabel As String
    Dim fileNotExist As String
    Dim sheetName As String
    Dim userData As String
    Dim wb As Workbook
    Dim wb1 As Workbook
    Dim wb2 As Workbook
    Dim criteria1 As String
    Dim criteria2 As String
    Dim criteria3 As String
    Dim iCnt As Integer


    sheetName = ActiveSheet.Name
    
    iCnt = 0
    
    userData = InputBox(" - 請輸入新薪資明細基本檔的年份(ex.115年):", "製作新年度薪資明細基本檔")
    If Len(Trim$(userData)) = 0 Then
    Exit Sub
    End If
    
    If MsgBox(" - 確定產生" & userData & "薪資明細基本檔", vbYesNo, "新年度薪資明細基本檔") = vbNo Then
    Exit Sub
    End If
    
    nyear = CLng(Val(userData))
    If nyear <= 0 Then
    Exit Sub
    End If
    
    
    oyear = nyear - 1
    wyear = nyear + 1911
    nyearLabel = CStr(nyear) & "年"
    oyearLabel = CStr(oyear) & "年"
    wyearLabel = CStr(wyear)
    
    If Len(ThisWorkbook.Path) > 0 Then
        filePath = ThisWorkbook.Path & Application.PathSeparator
    Else
        filePath = CurDir$ & Application.PathSeparator
    End If
    
    salNum = Cells(Rows.Count, 1).End(xlUp).Row
    
    criteria1 = oyearLabel & "12月"
    criteria2 = oyearLabel & "12月(2)"
    criteria3 = wyearLabel & "/1/"
1
    
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
    
            
         FilterRowsByMonth1 wb, "拆帳表", criteria3
         FilterRowsByMonth1 wb, "AA碼季獎金", criteria3
         FilterRowsByMonth wb, "總表", criteria1, criteria2
         FilterRowsByMonth wb, "行政總表", criteria1, criteria2
         
         
         Sheets("總表").Select
         Rows("9:16").Select
         Selection.Delete
         Columns("A:AO").Select
         With Selection.Font
          .Name = "Microsoft JhengHei UI"
          .Size = 10
          .Strikethrough = False
          .Superscript = False
          .Subscript = False
          .OutlineFont = False
          .Shadow = False
          .Underline = xlUnderlineStyleNone
          .TintAndShade = 0
          .ThemeFont = xlThemeFontNone
         End With
         
         wb.Save
         
         wb.Close SaveChanges:=False
        
         iCnt = iCnt + 1
          
       Else
         fileNotExist = fileNotExist & Cells(i, 6) & vbCrLf
          
       End If
    
    Next i
    
    MsgBox " - 製作薪資明細基本檔案" & userData & " 共 " & iCnt & "筆"
    MsgBox " - 無法製作薪資明細基本檔案名單:" & vbCrLf & fileNotExist
    
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
Private Sub FilterRowsByMonth1(ByVal wb As Workbook, ByVal targetSheetName As String, ByVal criteria3 As String)
    Dim ws As Worksheet
    Dim s As Long
    Dim s1 As Long
    Dim t As Long
    Dim text As String
    Dim v1 As String
    Dim lo As ListObject
    Dim lastTable As ListObject
    Dim currenrTop As Long
    Dim lastTableTop As Long
    
    On Error Resume Next
    Set ws = wb.Worksheets(targetSheetName)
    On Error GoTo 0
    If ws Is Nothing Then
       Exit Sub
    End If
     
    If ws.ListObjects.Count = 0 Then
       Exit Sub
    End If
    
    For Each lo In ws.ListObjects
        currentTop = lo.Range.Row
        s = currentTop - 2
        
        s1 = currentTop + lo.DataBodyRange.Rows.Count - 1
        text = CStr(ws.Cells(s, 1).Value)
        If Len(text) >= 13 Then
           v1 = Mid(text, 7, 7)
           Else
           v1 = vbNullString
        End If
        If v1 <> criteria3 Then
           For t = s To s1 Step 1
              ws.Rows(t).Delete
           Next t
        End If
     
     Next lo
    
    
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
Public Function FileExists(ByVal fullPath As String) As Boolean
    On Error GoTo EH
    FileExists = (Len(Dir$(fullPath, vbNormal)) > 0)
    Exit Function
EH:
    FileExists = False
End Function
Private Function ShouldKeepSheet(ByVal sheetName As String, ByVal oyearLabel As String) As Boolean
    Dim nameLower As String
    
    nameLower = LCase$(sheetName)
    ShouldKeepSheet = (nameLower = "format") Or (nameLower = "mformat") Or (sheetName = "行政總表") Or (sheetName = "總表") Or (sheetName = (oyearLabel & "12月")) Or (sheetName = (oyearLabel & "12月(2)")) Or (sheetName = (oyearLabel & "12月行政")) Or (sheetName = (oyearLabel & "12月(2)行政")) Or (sheetName = "拆帳表") Or (sheetName = "A碼清冊") Or (sheetName = "AA碼季獎金")
    
End Function
