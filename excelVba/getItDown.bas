Sub copyToNewSheet()
    Dim x As Long, lastrow As Long, startrow As Long, activeLastrow As Long
    Dim sumLineNum As Long

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

    'get the start row here
    startrow = x + 1
    'MsgBox startrow

    ' Add a new worksheet.
    Sheets.Add After:=Sheets(Sheets.Count)
    Set sh2 = ThisWorkbook.Sheets("Sheet" & Sheets.Count)

    'sh.Range("A1:L1").Copy _
    'Destination:=sh2.Range("A1")
    sh2.Range("A1").Value = "序号"
    sh2.Range("B1").Value = "提交答卷时间"
    sh2.Range("G1").Value = "姓名"
    sh2.Range("H1").Value = "拜访客户数"
    sh2.Range("I1").Value = "计划书数"
    sh2.Range("J1").Value = "预收件数"
    sh2.Range("K1").Value = "保费（万）"
    sh2.Range("L1").Value = "出单人员"
    sh2.Range("M1").Value = "辅导面谈"
    sh2.Range("N1").Value = "陪访"
    sh2.Range("O1").Value = "重点工作完成情况"
    sh2.Range("P1").Value = "面谈增员人数"

    sh.Range("A" & startrow & ":P" & lastrow + 1).Copy _
    Destination:=sh2.Range("A2")

    'delete selected column
    sh2.Columns("C:F").Select
    Selection.Delete Shift:=xlToLeft

    'shift the column
    sh2.Columns("L:L").Cut
    sh2.Columns("D:D").Insert Shift:=xlToRight
    Application.CutCopyMode = False

    activeLastrow = 2 + lastrow - startrow
    sumLineNum = activeLastrow + 1
    'sum the selected row
    sh2.Range("D" & sumLineNum).Formula = "=Sum(D2" & ":D" & activeLastrow & ")"
    sh2.Range("E" & sumLineNum).Formula = "=Sum(E2" & ":E" & activeLastrow & ")"
    sh2.Range("F" & sumLineNum).Formula = "=Sum(F2" & ":F" & activeLastrow & ")"
    sh2.Range("G" & sumLineNum).Formula = "=Sum(G2" & ":G" & activeLastrow & ")"
    sh2.Range("H" & sumLineNum).Formula = "=Sum(H2" & ":H" & activeLastrow & ")"

End Sub





