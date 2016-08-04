VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim g1People As Collection, g2People As Collection

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
    'TODO should not need to do the x + 1 here, need to confirm
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
    
        
    'change the index num
    For x = 2 To activeLastrow Step 1
        sh2.Cells(x, 1).Value = x - 1
    Next x
    

    'validate data
    Call validateData(sh2.Name, activeLastrow)

    'sum the selected row
    sh2.Range("D" & sumLineNum).Formula = "=Sum(D2" & ":D" & activeLastrow & ")"
    sh2.Range("E" & sumLineNum).Formula = "=Sum(E2" & ":E" & activeLastrow & ")"
    sh2.Range("F" & sumLineNum).Formula = "=Sum(F2" & ":F" & activeLastrow & ")"
    sh2.Range("G" & sumLineNum).Formula = "=Sum(G2" & ":G" & activeLastrow & ")"
    sh2.Range("H" & sumLineNum).Formula = "=Sum(H2" & ":H" & activeLastrow & ")"

    Call colorizeIt(sh2.Name, sumLineNum)
    Call addFormatting(sh2.Name, sumLineNum)
    Call screenShot(sh2.Name, sumLineNum)
    Call checkIfSubmitted(sh2.Name, activeLastrow)
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

    Set sh = ThisWorkbook.Sheets(shName)

    'insert the not found statement
    sh.Cells(lastrow + 2, 1).Value = "没有提交的人有： "
    Dim totalNumNotFound As Long
    totalNumNotFound = 0
    
    Dim selectedCole As Collection
    If OptionButton1.Value Then
        Set selectedCole = g1People
        'MsgBox ("set g1")
    Else
        Set selectedCole = g2People
    End If
    

    For x = 1 To selectedCole.Count
        For y = 2 To lastrow
            'the Value2 used here can ignore the formatting
            If sh.Cells(y, 3).Value2 = selectedCole.Item(x) Then
                'MsgBox ("found match!")
                GoTo FoundMatch
            End If
            If y = lastrow And sh.Cells(y, 3).Value2 <> selectedCole.Item(x) Then
                totalNumNotFound = totalNumNotFound + 1
                sh.Cells(lastrow + 2 + totalNumNotFound, 1).Value = selectedCole.Item(x)
            End If
        Next y
FoundMatch:
    Next x

    If selectedCole.Count <> lastrow - 1 + totalNumNotFound Then
        sh.Cells(lastrow + 2 + totalNumNotFound + 1, 1).Value = "人数对不上， 请复检。"
    End If


End Sub

Sub addFormatting(shName As String, endrow As Long)
    Set sh = ThisWorkbook.Sheets(shName)

    sh.Range("A1:L1").Font.Bold = True
    sh.Range("A1:L1").Font.Size = 12
    sh.Range("A:L").EntireColumn.AutoFit

    Dim rng As Range

    Set rng = sh.Range("A1:L" & endrow)

    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

End Sub


Private Sub ListBox1_Click()

End Sub

Private Sub CommandButton1_Click()
    
    MsgBox ("Button clicked")
    'Call copyToNewSheet
    
    If (Not OptionButton1.Value) And (Not OptionButton2.Value) Then
        MsgBox ("Please select ONE group")
    Else
        Call copyToNewSheet
    End If
    
        
    
End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub TextBox1_Change()

End Sub



Private Sub UserForm_Initialize()
    Dim x As Long
 
    Set g1People = New Collection
    Set g2People = New Collection
    
    'MsgBox ("For my dearest girl :P")
    
    With g1People
        .Add "尹美林"
        .Add "张金娣"
        .Add "黄秀莲"
        .Add "亓庆涛"
        .Add "伍斯斯"
        .Add "余治伟"
        .Add "谢瑞琴"
    End With
    
        With g2People
        .Add "a"
        .Add "b"
        .Add "d"
        .Add "c"
        .Add "f"
        .Add "e"
        .Add "g"
    End With
    
    For x = 1 To g1People.Count Step 1
        ListBox1.AddItem (g1People.Item(x))
    Next x
    
    For x = 1 To g2People.Count Step 1
        ListBox2.AddItem (g2People.Item(x))
    Next x
    
End Sub
