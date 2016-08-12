VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} XixiExceler 
   Caption         =   "XixiExceler"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   OleObjectBlob   =   "XixiExceler.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "XixiExceler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Dim g1People As Collection, g2People As Collection

Sub copyToNewSheet()
    Dim x As Long, LastRow As Long, startrow As Long, activeLastrow As Long
    Dim sumLineNum As Long

    Set sh = ThisWorkbook.Sheets("Sheet1")
    'Set sh2 = ThisWorkbook.Sheets("Sheet2")

    LastRow = sh.Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox lastrow

    'get the latest value
    exampleDate = DateValue(sh.Range("B" & LastRow).Value)
    For x = LastRow To 1 Step -1
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

    sh.Range("A" & startrow & ":P" & LastRow + 1).Copy _
    Destination:=sh2.Range("A2")
    
    'remove the duplicates
    Call DeleteDups(sh2.Name)
    
    'update the last row after deletion
    activeLastrow = sh2.Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox ("activelastrow: " & activeLastrow)
    sumLineNum = activeLastrow + 1

    'delete selected column
    sh2.Columns("C:F").Select
    Selection.Delete Shift:=xlToLeft

    'shift the column
    sh2.Columns("L:L").Cut
    sh2.Columns("D:D").Insert Shift:=xlToRight
    Application.CutCopyMode = False
        
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
    Call summarize(sh2.Name, activeLastrow)
    MsgBox ("For my dearest girl :)")
End Sub


'Delete duplicates
Sub DeleteDups(shName As String)
     
    Dim x               As Long
    Dim LastRow         As Long
    Dim rowsToDelete    As Collection
    
    Set rowsToDelete = New Collection
    
    Set sh = ThisWorkbook.Sheets(shName)
     
    LastRow = sh.Range("A65536").End(xlUp).Row
    For x = 1 To LastRow Step 1
        'MsgBox ("x: " & x)
        If Application.WorksheetFunction.CountIf(sh.Range("G1:G" & LastRow), sh.Range("G" & x).Value) > 1 Then
            'MsgBox ("remove row: " & x)
            sh.Range("G" & x).EntireRow.Delete
            x = x - 1
        End If
    Next x
    
    'MsgBox ("Remove Duplicates Done!")
    
     
End Sub



'ultimately should pass a parameter of the sheet name
Sub colorizeIt(shName As String, endrow As Long)
    Dim x As Long

     Set sh = ThisWorkbook.Sheets(shName)
     'make header blue
     sh.Range("A1:L1").Interior.ColorIndex = 37
     
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
                    sh.Range("G" & x & ":I" & x).Font.Bold = True
                End If
            End If
        End If
     Next x
End Sub


Sub validateData(shName As String, LastRow As Long)

    'If the contents of the cell looks like a numeric value, convert
    ' then cell to be numeirc
    For Each c In Worksheets(shName).Range("D2:H" & LastRow).Cells
        If IsNumeric(c) Then
            c.Value = Val(c.Value)
        Else
            c.Font.Color = vbRed
        End If
    Next
    
    '检查保费是不是以万为单位
    For Each c In Worksheets(shName).Range("H2:H" & LastRow).Cells
        If IsNumeric(c) Then
            c.Value = Val(c.Value)
            If c.Value > 10 Then
                c.Font.ColorIndex = 46
            End If
        End If
    Next

End Sub


'take the screen shot of the target area
Sub screenShot(shName As String, LastRow As Long)
    Set sh = ThisWorkbook.Sheets(shName)
    sh.Range("A1:L" & LastRow).Copy
    sh.Range("O1").Select
    sh.Pictures.Paste Link:=True
    Application.CutCopyMode = False
End Sub

Sub summarize(shName As String, LastRow As Long)
    Dim x As Long, y As Long, found As Boolean
    Dim startIndex As Long
    Dim line1String As String, notFoundPeople As String
    
    startIndex = LastRow + 2

    Set sh = ThisWorkbook.Sheets(shName)

    'insert the not found statement
    sh.Cells(startIndex, 1).Value = "今日总结： "
    sh.Cells(startIndex, 1).Font.Bold = True
    startIndex = startIndex + 1
    
    If OptionButton1.Value Then
        line1String = "越秀一二区工作达成报告"
    Else
        line1String = "越秀三区工作达成报告"
    End If
    
    sh.Cells(startIndex, 1).Value = Date & line1String
    startIndex = startIndex + 1
    
    sh.Cells(startIndex, 1).Value = "截止至" & Time() & "共 " & LastRow - 1 & " 位主管提交"
    startIndex = startIndex + 1
    
    sh.Cells(startIndex, 1).Value = "总计面谈增员数："
    sh.Cells(startIndex, 2).Formula = "=D" & LastRow + 1
    sh.Cells(startIndex, 3).Value = "人"
    startIndex = startIndex + 1

    sh.Cells(startIndex, 1).Value = "总计拜访客户："
    sh.Cells(startIndex, 2).Formula = "=E" & LastRow + 1
    sh.Cells(startIndex, 3).Value = "人"
    startIndex = startIndex + 1
    
    sh.Cells(startIndex, 1).Value = "总计送计划书："
    sh.Cells(startIndex, 2).Formula = "=F" & LastRow + 1
    sh.Cells(startIndex, 3).Value = "人"
    startIndex = startIndex + 1
    
'     sh.Cells(startIndex, 1).Value = "未提交主管名单如下："
    notFoundPeople = "未提交主管名单如下："

    
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
        For y = 2 To LastRow
            'the Value2 used here can ignore the formatting
            If sh.Cells(y, 3).Value2 = selectedCole.Item(x) Then
                'MsgBox ("found match!")
                GoTo FoundMatch
            End If
            If y = LastRow And sh.Cells(y, 3).Value2 <> selectedCole.Item(x) Then
                totalNumNotFound = totalNumNotFound + 1
                'sh.Cells(startIndex, totalNumNotFound + 1).Value = selectedCole.Item(x)
                notFoundPeople = notFoundPeople & selectedCole.Item(x) & " "
            End If
        Next y
FoundMatch:
    Next x
    
    sh.Cells(startIndex, 1).Value = notFoundPeople
    startIndex = startIndex + 1

    If selectedCole.Count <> LastRow - 1 + totalNumNotFound Then
        sh.Cells(startIndex + 1, 1).Value = "人数对不上， 请检查是否有人把自己名字写错"
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
    
    rng.BorderAround _
        Weight:=xlThick
        
    sh.Range("A1:L1").BorderAround _
        Weight:=xlThick
        
    
    
    rng.VerticalAlignment = xlCenter
    rng.HorizontalAlignment = xlCenter

End Sub


Private Sub ListBox1_Click()

End Sub

Private Sub CommandButton1_Click()
    
    'MsgBox (Date)
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
    
    
   ' Group 1 And 2
With g1People
    .Add "陈文玉"
    .Add "张金娣"
    .Add "雷己兰"
    .Add "谭晓云"
    .Add "谢渠成"
    .Add "杨璐璐"
    .Add "胡彬"
    .Add "罗志刚"
    .Add "黄秀莲"
    .Add "谢瑞琴"
    .Add "成取英"
    .Add "成兰英"
    .Add "余治伟"
    .Add "梁艳红"
    .Add "袁嘉能"
    .Add "郭剑章"
    .Add "伍斯斯"
    .Add "亓庆涛"
    .Add "姚健"
    .Add "叶俊明"
    .Add "肖忠颖"
    .Add "戴詩緯"
    .Add "尹美林"
End With


'Group3
With g2People
    .Add "陈和贵"
    .Add "刘定希"
    .Add "陈美滨"
    .Add "匡书光"
    .Add "梁嘉雯"
    .Add "刘振宇"
    .Add "李敏"
    .Add "朱强"
    .Add "李小敏"
    .Add "谭荣彬"
    .Add "任宇"
    .Add "汤利红"
    .Add "张海燕"
    .Add "曾艳芬"
    .Add "易畅"
    .Add "张玉荣"
    .Add "赵金凤"
End With
    
    For x = 1 To g1People.Count Step 1
        ListBox1.AddItem (g1People.Item(x))
    Next x
    
    For x = 1 To g2People.Count Step 1
        ListBox2.AddItem (g2People.Item(x))
    Next x
    
    'MsgBox (Time())
    
    
End Sub
