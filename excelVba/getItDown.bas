
Sub Marcro1()
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "My Sheet"
End Sub


Sub DateMarcro()
    Dim x As Long, lastrow As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    exampleDate = DateValue("2016/7/28 21:44:34")
    MsgBox exampleDate
    
    For x = lastrow To 1 Step -1
        If DateValue(Cells(x, 2).Value) <> exampleDate Then
            Rows(x).Delete
        End If
       Next x
End Sub


Sub DateMarcro2()
    Dim x As Long, lastrow As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    exampleDate = DateValue("2016/7/28 21:44:34")
    MsgBox exampleDate
    
    With Sheet1
           .AutoFilterMode = False
            .Range("A1:P8").AutoFilter
           .Range("A1:P8").AutoFilter Field:=8, Criteria1:=2
    End With
End Sub

Sub DeleteColumn()
    Columns("C:F").Select
    Selection.Delete Shift:=xlToLeft
End Sub

Sub shiftColumn()
    With ActiveSheet
        .Columns("L:L").Cut
        .Columns("D:D").Insert Shift:=xlToRight
    End With
Application.CutCopyMode = False
End Sub

Sub copyToNewSheet()
    Dim x As Long, lastrow As Long, startrow As Long
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox lastrow
    
 
    
    Set sh = ThisWorkbook.Sheets("Sheet1")
    'Set sh2 = ThisWorkbook.Sheets("Sheet2")
    
    
    'get the latest value
    exampleDate = DateValue(Range("B" & lastrow).Value)
    For x = lastrow To 1 Step -1
        If DateValue(Cells(x, 2).Value) <> exampleDate Then
            Exit For
        End If
    Next x
    
    startrow = x + 1
    'MsgBox startrow
    
        
        
    'Range("D" & lastrow + 1).Formula = "=Sum(D" & startrow & ":D" & lastrow & ")"
    Range("D" & lastrow + 1).Formula = "=Sum(D" & startrow & ":D" & lastrow & ")"
    Range("E" & lastrow + 1).Formula = "=Sum(E" & startrow & ":E" & lastrow & ")"
    Range("F" & lastrow + 1).Formula = "=Sum(F" & startrow & ":F" & lastrow & ")"
    Range("G" & lastrow + 1).Formula = "=Sum(G" & startrow & ":G" & lastrow & ")"
    Range("H" & lastrow + 1).Formula = "=Sum(H" & startrow & ":H" & lastrow & ")"
    
    
    
    ' Add a new worksheet.
    Sheets.Add After:=Sheets(Sheets.Count)
    Set sh2 = ThisWorkbook.Sheets("Sheet" & Sheets.Count)
    
    'sh.Range("A1:L1").Copy _
    'Destination:=sh2.Range("A1")
    sh2.Range("A1").Value = "序号"
    sh2.Range("B1").Value = "提交答卷时间"
    sh2.Range("C1").Value = "姓名"
    sh2.Range("D1").Value = "面谈增员人数"
    sh2.Range("E1").Value = "拜访客户数"
    sh2.Range("F1").Value = "计划书数"
    sh2.Range("G1").Value = "预收件数"
    sh2.Range("H1").Value = "保费（万）"
    sh2.Range("I1").Value = "出单人员"
    sh2.Range("J1").Value = "辅导面谈"
    sh2.Range("K1").Value = "陪访"
    sh2.Range("L1").Value = "重点工作完成情况"
    
    sh.Range("A" & startrow & ":L" & lastrow + 1).Copy _
    Destination:=sh2.Range("A2")
    
End Sub


Sub allTogether()
    Call DeleteColumn
    Call shiftColumn
    Call copyToNewSheet
End Sub




