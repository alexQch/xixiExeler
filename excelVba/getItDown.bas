Sub copyToNewSheet()
    Dim x As Long, lastrow As Long, startrow As Long, activeLastrow As Long
    Dim sumLineNum As Long

    Set sh = ThisWorkbook.Sheets("Sheet1")
    'Set sh2 = ThisWorkbook.Sheets("Sheet2")

    lastrow = sh.Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox lastrow

    'get the latest value
    exampleDate = DateValue(sh.Range("B" & lastrow).Value)
    For x = lastrow To 1 Step -1
        If DateValue(sh.Cells(x, 2).Value) <> exampleDate Then
            Exit For
        End If
    Next x

    'get the start row here
    startrow = x + 1
    'MsgBox startrow

    ' Add a new worksheet.
    Sheets.Add After:=Sheets(Sheets.Count)
    Set sh2 = ThisWorkbook.Sheets("Sheet" & Sheets.Count)


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
    Call validateData(sh2.name, activeLastrow)
    Call colorizeIt(sh2.name, sumLineNum)
    Call screenShot(sh2.name, sumLineNum)
    Call checkIfSubmitted(sh2.name, activeLastrow)
    MsgBox ("DONE")
End Sub

'ultimately should pass a parameter of the sheet name
Sub colorizeIt(shName As String, endrow As Long)
    Dim x As Long

     Set sh = ThisWorkbook.Sheets(shName)
     'make the D column green
     sh.Range("D1:D" & endrow).Interior.ColorIndex = 43

    'set WrapText and AutoFit for this range
     sh.Range("A1:L" & endrow).WrapText = True
     sh.Columns("A:L").AutoFit

     'colorize G H I column
     For x = 2 To endrow - 1 Step 1
        If IsNumeric(sh.Cells(x, 8)) Then
            If sh.Cells(x, 8).Value > 0 Then
                If IsNumeric(sh.Cells(x, 7)) And sh.Cells(x, 7) > 0 Then
                    sh.Range("G" & x & ":I" & x).Interior.ColorIndex = 6
                End If
            End If
        End If
     Next x
End Sub


Sub validateData(shName As String, lastrow As Long)

    'If the contents of the cell looks like a numeric value, convert
    ' then cell to be numeirc
    For Each c In Worksheets(shName).Range("D2:H" & lastrow).Cells
        If IsNumeric(c) Then c.Value = Val(c.Value)
    Next

End Sub


'take the screen shot of the target area
Sub screenShot(shName As String, lastrow As Long)
    Set sh = ThisWorkbook.Sheets(shName)
    sh.Range("A1:L" & lastrow).Copy
    sh.Range("O1").Select
    sh.Pictures.Paste Link:=True
    Application.CutCopyMode = False
End Sub

Sub checkIfSubmitted(shName As String, lastrow As Long)
    Dim x As Long, y As Long, found As Boolean


    Dim names(1 To 8) As String, arrayLen As Long
    arrayLen = 8

    names(1) = "尹美林"
    names(2) = "张金娣"
    names(3) = "黄秀莲"
    names(4) = "亓庆涛"
    names(5) = "伍斯斯"
    names(6) = "余治伟"
    names(7) = "谢瑞琴"
    names(8) = "Alex"

    Set sh = ThisWorkbook.Sheets(shName)

    'insert the not found statement
    sh.Cells(lastrow + 2, 1).Value = "没有提交的人有： "
    Dim totalNumNotFound As Long
    totalNumNotFound = 0

    For x = 2 To arrayLen
        For y = 2 To lastrow
            'the Value2 used here can ignore the formatting
            If sh.Cells(y, 3).Value2 = names(x) Then
                'MsgBox ("found match!")
                GoTo FoundMatch
            End If
            If y = lastrow And sh.Cells(y, 3).Value2 <> names(x) Then
                MsgBox ("Not FOUND " + names(x))
                totalNumNotFound = totalNumNotFound + 1
                sh.Cells(lastrow + 2 + totalNumNotFound, 1).Value = names(x)
            End If
        Next y
FoundMatch:
    Next x

    If arrayLen <> lastrow - 1 + totalNumNotFound Then
        sh.Cells(lastrow + 2 + totalNumNotFound + 1, 1).Value = "人数对不上， 请复检。"
    End If


End Sub

