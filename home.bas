Attribute VB_Name = "Home"
Sub qy_merge() '合并单元格，并把所有单元格的值用空格分隔后以文本方式联结，作为合并后单元格的值并居中显示
    Dim rng As Range
    Dim str() As String
    Dim sum As Integer
    Dim i As Integer
   
    On Error Resume Next
    Set rng = Selection
    sum = 0

    For Each r In rng
        sum = sum + 1
    Next r
    
    ReDim str(sum - 1)
    
    For Each r In rng
        str(i) = r.Value
        i = i + 1
    Next r

    rng.Merge

    rng.Value = Join(str, " ")

    With rng
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    
End Sub

Sub qy_merge_new() '合并单元格，并把所有单元格的值用输入框输入分隔后以文本方式联结，作为合并后单元格的值并居中显示
    
    Dim rng As Range
    Dim str() As String
    Dim sum As Integer
    Dim i As Integer
    Dim a
   
    On Error Resume Next
    Set rng = Selection
    sum = 0
    a = InputBox("请输入一个分隔符，用于分隔各单元格，按取消为不分隔")
    
    For Each r In rng
        sum = sum + 1
    Next r
    
    ReDim str(sum - 1)
    
    For Each r In rng
        str(i) = r.Value
        i = i + 1
    Next r

    rng.Merge

    rng.Value = Join(str, a)


    With rng
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Public Sub copyIM() '多表数据合并到新表
    Dim sht As Worksheet, xrow As Integer, rng As Range
    Dim lastCount As String
    Dim rngS As Range
    
    Worksheets.Add(before:=Worksheets(1)).Name = "新建汇总表"
    
    '复制各表数据到新表
    For Each sht In Worksheets
        If sht.Name <> ActiveSheet.Name Then     '如果工作表不是当前激活的表
            Set rng = Range("A65536").End(xlUp).Offset(1, 0)  '取得当前表A列第一个非空单元格
            xrow = sht.Range("A1").CurrentRegion.Rows.Count - 1  '取得要复制的工作表的总行数（减去首行）
            sht.Range("A2").Resize(xrow, 7).Copy rng  '把此工作表从A2单元格开始的，xrow行，7列这块区域，复制到活动工作表第一个非空单元格
        End If
    Next
    
    '删除除了第一张表外的各表行头
    xrow = Range("B1").CurrentRegion.Rows.Count
    lastCount = "B5:" & "B" & xrow
    Set rngS = Range(lastCount)
    
    For Each rng In rngS
        
        If InStr(rng, "页码") > 0 Then   '如果该单元格的值包含“页码”这个词
           rng.EntireRow.Delete   '删除此行
           'ActiveSheet.Rows(rng.Row).Delete
        End If
    
    Next rng
    
    Range("B2:F2").Select
    Call qy_merge
    
End Sub


Sub copyWorkbook()   '把同一文件夹内的所有工作簿合并成一个新的工作簿
    
    
    '获取要复制文件所在的目录名--------------***
    Dim copyFilePath As String
    With Application.FileDialog(msoFileDialogFolderPicker)     '打开文件对话框，选择要复制的文件所在的文件夹
        .InitialFileName = "C:\"                                '初始文件夹为C盘根目录
        .Title = "请选择要复制的文件所在的文件夹"
        .Show
        If .SelectedItems.Count > 0 Then                        '如果选择了文件夹
           copyFilePath = .SelectedItems(1)
           'MsgBox copyFilePath
           Else
            MsgBox "没有选择任何目录，退出程序"
            Exit Sub
        End If
    End With
    '***----------------------------------****
    
    
    '选择是全部复制，还是不复制表头----------------***
    Dim fullCopy As Boolean
    If MsgBox("整表复制请选“是”，只保留一份表头请选“否”", vbYesNo, "请选择复制方式") = vbYes Then
        fullCopy = True
    Else
        fullCopy = False
    End If
    '***-------------------------------------------****

    Dim wb As Workbook, Erow As Long, fn As String, FileName As String, sht As Worksheet
    Application.Workbooks.Add  '新建一个工作簿
    
   
    Application.ScreenUpdating = False
    FileName = Dir(copyFilePath & "\*.*") '要复制的文件名
    Dim bt As Boolean      '设定表头是否已复制
    bt = False
    Do While FileName <> ""
        fn = copyFilePath & "\" & FileName '要复制的文件全路径名
        Set wb = GetObject(fn)  '将fn代表的工作簿对象赋给变量
        Set sht = wb.Worksheets(1) '汇总第1张工作表
        
        '取得汇总表第一个非空行---------------------
        If ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Value = "" Then
                Erow = 1
            Else
                Erow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
        End If
       
        If fullCopy = True Then                                '全表复制
            sht.UsedRange.Copy ActiveSheet.Cells(Erow, 1)
        Else                                                  '只复制一次表头
            If bt = False Then
                sht.UsedRange.Copy ActiveSheet.Cells(Erow, 1)
                bt = True
            Else
                sht.Cells(2, 1).Resize(sht.UsedRange.Rows.Count - 1, sht.Cells(sht.Rows.Count, 1).End(xlUp).End(xlToRight).Column).Copy ActiveSheet.Cells(Erow, 1)
            End If
        End If
        
        wb.Close False    '关闭已复制的工作簿
        FileName = Dir    '定位下一个要复制的工作簿
    Loop
   Application.ScreenUpdating = True
      
End Sub

Sub delRow()   '删除某个特定单元格所在的行
    Dim specialWord As String
    specialWord = Application.InputBox("请输入要删除所在行的单词")
    
    If specialWord = "" Then Exit Sub
    
    Application.ScreenUpdating = False
    Dim allRange As Range
    Dim s As Range
    For Each s In ActiveSheet.UsedRange
        If s.Value = specialWord Then
            s.EntireRow.Delete
        End If
    Next s
        
    Application.ScreenUpdating = True

End Sub

Sub 提取单元格中数字并求和()
    Dim total As Single
    Dim txt As String
    txt = ActiveCell.Offset(0, -1).Value
    total = 0
    Dim txtLen As Integer
    txtLen = Len(txt)
    Dim start As Boolean
    start = False
    Dim tmpNum
    'tmpNum = ""
    Dim printIfo As String
    For i = 1 To txtLen
        Dim tmp
        tmp = Mid(txt, i, 1)
        If VBA.IsNumeric(tmp) = False And tmp <> "." Then  '判断是不是数字和小数点
              If start Then  '不是数字和小数点，检测是否开始计数
              '已开始计数
              total = total + Val(tmpNum)
              printIfo = printIfo & tmpNum & "+"
              tmpNum = ""
              start = False
              
              End If
        Else        '是数字和小数点
            If start = False Then start = True
            If Left(tmpNum, 1) = "." Then tmpNum = ""
            tmpNum = tmpNum & tmp

        End If
        
        
     Next i
     
     If Val(tmpNum) <> 0 Then
        total = total + Val(tmpNum)
        printIfo = printIfo & tmpNum
     End If
     
     If Right(printIfo, 1) = "+" Then printIfo = Mid(printIfo, 1, Len(printIfo) - 1)
     
     MsgBox printIfo & "=" & total
     ActiveCell.Value = total
     
    
End Sub

Public Sub copyAllSheets() '同工作簿多表合并
    Dim sht As Worksheet, xrow As Integer, rng As Range
    Dim lastCount As String
    Dim rngS As Range
    
    Worksheets.Add(before:=Worksheets(1)).Name = "合并表" & WorksheetFunction.Substitute(Now(), ":", "-")
    
    '复制各表数据到新表
    For Each sht In Worksheets
        If sht.Name <> ActiveSheet.Name Then     '如果工作表不是当前激活的表
            If Range("A1") = "" And Range("B1") = "" Then
                Set rng = Range("A1")
            Else
                Set rng = Range("A65536").End(xlUp).Offset(1, 0)  '取得当前表A列第一个非空单元格
            End If
                    
            'xrow = sht.Range("A1").CurrentRegion.Rows.Count - 1  '取得要复制的工作表的总行数（减去首行）
            'sht.Range("A2").Resize(xrow, 7).Copy rng  '把此工作表从A2单元格开始的，xrow行，7列这块区域，复制到活动工作表第一个非空单元格
            sht.UsedRange.Copy rng
        End If
    Next
    
    '删除除了第一张表外的各表行头
'    xrow = Range("B1").CurrentRegion.Rows.Count
'    lastCount = "B5:" & "B" & xrow
'    Set rngS = Range(lastCount)
    
   
'    For Each rng In rngS
'
'        If InStr(rng, "页码") > 0 Then   '如果该单元格的值包含“页码”这个词
'           rng.EntireRow.Delete   '删除此行
'           'ActiveSheet.Rows(rng.Row).Delete
'        End If
'
'    Next rng
  
    
'    Range("B2:F2").Select
'    Call qy_merge
    
End Sub

Public Sub test()
'Range("E1").EntireRow.Delete
Application.ScreenUpdating = False
total = 0
For i = 1 To 450
    If Cells(i, "S") = "" Then
        Cells(i, "S") = total
        Cells(i, "S").Offset(0, -1) = "合计"
        total = 0
    Else
        total = total + Cells(i, "S").Value
    End If
Next i
Application.ScreenUpdating = True
End Sub

Sub old竞品数据预处理()
    '获取总行数
    Dim num As Long
    'Range("A1").Select
    'Range("A1").End(xlDown).Select
    'num = ActiveCell.Row
    num = Cells(Rows.Count, 1).End(xlUp).Row
    
    '关闭EXCEL刷新
    Application.ScreenUpdating = False
    
    Columns("B:B").Delete Shift:=xlToLeft
    'Selection.Delete Shift:=xlToLeft
    Columns("C:C").Delete Shift:=xlToLeft
    'Selection.Delete Shift:=xlToLeft
    Range("D:D,E:E,G:G,O:O,P:P,Q:Q,R:R").Delete Shift:=xlToLeft
    Columns("K:K").TextToColumns Destination:=Range("K1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("J:J").TextToColumns Destination:=Range("J1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("I:I").TextToColumns Destination:=Range("I1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    With Columns("H:H")
        .NumberFormatLocal = "yyyy/m/d"
        .TextToColumns Destination:=Range("H1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
            :=Array(Array(1, 5), Array(2, 9)), TrailingMinusNumbers:=True
        .EntireColumn.AutoFit
    End With
    

    
    Columns("I:I").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("I:I").NumberFormatLocal = "G/通用格式"
    
    With Range("I1")
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Value = "当月日期"
    End With
    'ActiveCell.FormulaR1C1 = "当月日期"
    Range("I2").Formula = "=DAY(H2)"
    Range("I2").AutoFill Destination:=Range("I2:I" & num)
    Columns("H:H").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
 

    Range("H1").Formula = "配送企业简称"
    Range("H2").Formula = "=IFERROR(VLOOKUP(G2,'F:\青云\(synology)\竞品\竞品对照表.xlsx'!表1[#All],2,0),"""")"
    'Range("H2").Formula = "=IFERROR(G2,'F:\青云\(synology)\竞品\竞品对照表.xlsx'!表1[#All],2,0),"""")"
    Range("H2").AutoFill Destination:=Range("H2:H" & num)
    Columns("H:H").EntireColumn.AutoFit
    Columns("G:G").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Range("G1").Formula = "疾控简称"
    Range("G2").Formula = "=IFERROR(VLOOKUP(F2,'F:\青云\(synology)\竞品\竞品对照表.xlsx'!表2[#All],2,0),"""")"
    'Range("G2").Formula = "=IFERROR(H2,'F:\青云\(synology)\竞品\竞品对照表.xlsx'!表1[#All],2,0),"""")"
    Range("G2").AutoFill Destination:=Range("G2:G" & num)
    Columns("F:F").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("F1").Formula = "申报企业简称"
    Range("F2").Formula = "=IFERROR(VLOOKUP(E2,'F:\青云\(synology)\竞品\竞品对照表.xlsx'!表3[#All],2,0),"""")"
    Range("F2").AutoFill Destination:=Range("F2:F" & num)
    Columns("F:F").EntireColumn.AutoFit
    'Columns("E:E").Select
    Columns("D:D").Delete Shift:=xlToLeft
    Columns("D:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Formula = "规格"
    Range("D2").FormulaR1C1 = "=IFERROR(IF(FIND(""西林"",RC[-1])>0,""西林""),IFERROR(IF(FIND(""预充"",RC[-1])>0,""预充""),""""))"
    Range("D2").AutoFill Destination:=Range("D2:D" & num)
    Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C1").Formula = "疫苗简称"
    Range("C2").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],[竞品对照表.xlsx]Sheet1!C8:C9,2,0),"""")"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C" & num)
    Columns("C:C").EntireColumn.AutoFit
    Range("A1:P" & num).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    
    '恢复屏幕刷新
    Application.ScreenUpdating = True
    
    MsgBox "处理完成！"
End Sub



Function CaiQiang(ByVal rng As Range) '计算菜钱的公式
    'Application.Volatile True '易失性函数
    Dim total As Single
    Dim txt As String
    txt = rng.Value
    total = 0
    Dim txtLen As Integer
    txtLen = Len(txt)
    Dim start As Boolean
    start = False
    Dim tmpNum
    'tmpNum = ""
    Dim printIfo As String
    For i = 1 To txtLen
        Dim tmp
        tmp = Mid(txt, i, 1)
        If VBA.IsNumeric(tmp) = False And tmp <> "." Then  '判断是不是数字和小数点
              If start Then  '不是数字和小数点，检测是否开始计数
              '已开始计数
              total = total + Val(tmpNum)
              printIfo = printIfo & tmpNum & "+"
              tmpNum = ""
              start = False
              
              End If
        Else        '是数字和小数点
            If start = False Then start = True
            If Left(tmpNum, 1) = "." Then tmpNum = ""
            tmpNum = tmpNum & tmp

        End If
        
        
     Next i
     
     If Val(tmpNum) <> 0 Then
        total = total + Val(tmpNum)
        printIfo = printIfo & tmpNum
     End If
     
     If Right(printIfo, 1) = "+" Then printIfo = Mid(printIfo, 1, Len(printIfo) - 1)
     
     CaiQiang = total
End Function

Sub AddDay()
    For Each r In Selection
        r.Value = r.Value & "日"
    Next r
End Sub
Sub 日报表修整()
   '获取总行数
    Dim num As Long
    Range("A1").Select
    Range("A1").End(xlDown).Select
    num = ActiveCell.Row
    
    
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]&RC[-1]"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C" & num)
    Range("C2:C" & num).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:B").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "商品+规格"
   ' Range("A2").Select
   ' Application.SendKeys ("^a")
    

    
    '所有单元格加全框线,居中，设行高
    Range("B2").CurrentRegion.Select
   ' Application.SendKeys ("^a")
    'Application.Wait (Now + TimeValue("00:00:01"))
    'Application.Wait (Now + TimeValue("00:00:01"))
    
    Selection.RowHeight = 18
    Selection.Columns.AutoFit
    With Selection
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = False
      .ReadingOrder = xlContext
      .MergeCells = False
    End With

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A1").Select
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B2").CurrentRegion.Select

End Sub
Sub 设置色阶()
    Dim colorArr(2) As Integer
    colorArr(0) = 35
    colorArr(1) = 19
    colorArr(2) = 40
    
    Dim colorValue As Integer
    colorValue = 0
    Dim rng As Range
    Range(Cells(Selection.Row, Selection.Column), Cells(Selection.Row, Selection.Column + Selection.Columns.Count - 1)).Interior.ColorIndex = 35
    
    For Each rng In Range(Cells(Selection.Row + 1, Selection.Column), Cells(Selection.Row + Selection.Rows.Count - 1, Selection.Column))
        If rng.Value = rng.Offset(-1, 0).Value Then
            Range(Cells(rng.Row, Selection.Column), Cells(rng.Row, Selection.Column + Selection.Columns.Count - 1)).Interior.ColorIndex = rng.Offset(-1, 0).Interior.ColorIndex
        Else
            colorValue = colorValue + 1
            Range(Cells(rng.Row, Selection.Column), Cells(rng.Row, Selection.Column + Selection.Columns.Count - 1)).Interior.ColorIndex = colorArr(colorValue Mod 2)
        End If
    Next
    
   '设置最后一行色阶
   With Range(Cells(Selection.Row + Selection.Rows.Count - 1, Selection.Column), Cells(Selection.Row + Selection.Rows.Count - 1, Selection.Column + Selection.Columns.Count - 1))
        .Interior.Color = RGB(91, 155, 213)
        .Font.Color = RGB(255, 255, 255)
   End With
   
    '设置第一行色阶
    With Range(Cells(Selection.Row, Selection.Column), Cells(Selection.Row, Selection.Column + Selection.Columns.Count - 1))
        .Interior.Color = RGB(91, 155, 213)
        .Font.Color = RGB(255, 255, 255)
   End With
   
  
End Sub

Sub 色阶()

    Dim rng As Range
    Dim hadColor As Boolean
    Range(Cells(Selection.Row, Selection.Column), Cells(Selection.Row, Selection.Column + Selection.Columns.Count - 1)).Interior.Color = 14277081
    hadColor = True
    For Each rng In Range(Cells(Selection.Row + 1, Selection.Column), Cells(Selection.Row + Selection.Rows.Count - 1, Selection.Column))
        If rng.Value = rng.Offset(-1, 0).Value Then
            If hadColor = True Then
                
                Range(Cells(rng.Row, Selection.Column), Cells(rng.Row, Selection.Column + Selection.Columns.Count - 1)).Interior.Color = 14277081
                hadColor = True
            Else
                Range(Cells(rng.Row, Selection.Column), Cells(rng.Row, Selection.Column + Selection.Columns.Count - 1)).Interior.Color = 16777215
                hadColor = False
            End If
         Else
                If hadColor = True Then
                    Range(Cells(rng.Row, Selection.Column), Cells(rng.Row, Selection.Column + Selection.Columns.Count - 1)).Interior.Color = 16777215
                    hadColor = False
                Else
                    Range(Cells(rng.Row, Selection.Column), Cells(rng.Row, Selection.Column + Selection.Columns.Count - 1)).Interior.Color = 14277081
                    hadColor = True
                End If
        End If
    Next rng
End Sub

Function timeSection(ByVal rng1 As Range, rng2 As Range)
    If rng1 >= 8 And rng1 < 12 Then
        timeSection = "8-12点"
    ElseIf (rng1 >= 12 And rng1 < 14) Or (rng1 = 14 And rng2 < 30) Then
            timeSection = "12点-14:30"
    ElseIf (rng1 = 14 And rng2 >= 30) Or (rng1 >= 15 And rng1 < 18) Then
            timeSection = "14:30 - 18:00"
    Else
            timeSection = "18:00以后"
    End If
        
End Function

Sub 特定词所在行填色()
    Dim num As Long
    num = Range("A2").CurrentRegion.Columns.Count
    
    Dim specialWord As String
    specialWord = Application.InputBox("请输入需要高亮行包含的单词")

    If specialWord = "" Then Exit Sub
    Dim rng As Range
    For Each rng In ActiveCell.CurrentRegion
        If InStr(specialWord, rng) > 0 Then
            'Rows(rng.Row).Interior.ColorIndex = 24
            Range(Cells(rng.Row, rng.CurrentRegion.Column), Cells(rng.Row, num)).Interior.ColorIndex = 22
        End If
    Next
End Sub

Function QyHeBinWenBen(ParamArray inp())  '合并单元格文本
    Dim i, j
'    Dim score()
    For j = 0 To UBound(inp)
        For Each cl In inp(j)
'            i = i + 1
'            ReDim Preserve score(i)
'            score(i) = cl
'            QyHeBinWenBen = QyHeBinWenBen & score(i)
            QyHeBinWenBen = QyHeBinWenBen & cl
        Next
    Next
End Function

Function funTest(ParamArray inp())

End Function

Sub 公式转数值()
    Dim ing As Range
    For Each ing In Selection
        If ing.HasFormula Then ing.Value = ing
    Next
End Sub

Public Sub 保存日报()
   '加日---------------
    Dim dayRow As Integer, dayColumn As Integer
    Dim dayColumnBegin As Integer
    Dim dayColumnEnd As Integer
    
    dayRow = Selection.Row
    dayColumnBegin = Selection.Column + 3
    dayColumnEnd = Cells(Selection.Row, Selection.Column).End(xlToRight).Column - 1
    For i = dayColumnBegin To dayColumnEnd
        Cells(dayRow, i) = Cells(dayRow, i) & "日"
    Next
  '.-------------------END
    dayColumn = Cells(dayRow, Selection.Column).End(xlToRight).Column - 1
    


    '确定文件名
    Dim sFileName As String
    Dim sdate As String

    If Val(Left(Cells(2, dayColumnEnd), Len(Cells(2, dayColumnEnd)) - 1)) < Day(Now()) Then
        sdate = Year(Now()) & "-" & Month(Now()) & "-" & Left(Cells(dayRow, dayColumn), Len(Cells(dayRow, dayColumn)) - 1)
        'sFileName = "C:\Users\QingYun\Desktop\竞品日报" & "(" & sdate & ")"
        sFileName = Environ("userprofile") & "\Desktop\" & "竞品日报" & "(" & sdate & ")"
    Else
        If Month(Now()) = 1 Then
           sdate = Year(Now()) - 1 & "-" & 12 & "-" & Left(Cells(dayRow, dayColumn), Len(Cells(dayRow, dayColumn)) - 1)
        Else
            sdate = Year(Now()) & "-" & Month(Now()) - 1 & "-" & Left(Cells(dayRow, dayColumn), Len(Cells(dayRow, dayColumn)) - 1)
        End If
        sFileName = Environ("userprofile") & "\Desktop\" & "竞品日报" & "(" & sdate & ")"
    End If

   '保存文件
   ActiveWorkbook.SaveAs (sFileName)
   MsgBox "文件" & vbCrLf & sFileName & vbCrLf & "已保存到桌面"
End Sub

Public Sub 保存简报()
    '加日---------------------
    Dim dayRow As Integer, dayColumn As Integer
    Dim dayColumnBegin As Integer
    Dim dayColumnEnd As Integer
    
    dayRow = Selection.Row
    dayColumnBegin = Selection.Column + 3
    dayColumnEnd = Cells(Selection.Row, Selection.Column).End(xlToRight).Column
    For i = dayColumnBegin To dayColumnEnd
        Cells(dayRow, i) = Cells(dayRow, i) & "日"
    Next
   '-------------------------END
   
    dayColumn = Cells(dayRow, Selection.Column).End(xlToRight).Column
    
   Dim sFileName As String
   Dim sdate As String
   'sFileName = "C:\Users\QingYun\Desktop\竞品简报" & "(" & Year(Now()) & "-" & Month(Now()) & "-" & Left(Cells(dayRow, dayColumn), Len(Cells(dayRow, dayColumn)) - 1) & ")"
    If Val(Left(Cells(2, dayColumnEnd), Len(Cells(2, dayColumnEnd)) - 1)) < Day(Now()) Then
        sdate = Year(Now()) & "-" & Month(Now()) & "-" & Left(Cells(dayRow, dayColumn), Len(Cells(dayRow, dayColumn)) - 1)
        'sFileName = "C:\Users\QingYun\Desktop\竞品日报" & "(" & sdate & ")"
        sFileName = Environ("userprofile") & "\Desktop\" & "竞品简报" & "(" & sdate & ")"
    Else
        If Month(Now()) = 1 Then
           sdate = Year(Now()) - 1 & "-" & 12 & "-" & Left(Cells(dayRow, dayColumn), Len(Cells(dayRow, dayColumn)) - 1)
        Else
            sdate = Year(Now()) & "-" & Month(Now()) - 1 & "-" & Left(Cells(dayRow, dayColumn), Len(Cells(dayRow, dayColumn)) - 1)
        End If
        sFileName = Environ("userprofile") & "\Desktop\" & "竞品简报" & "(" & sdate & ")"
    End If
   
   
   
   ActiveWorkbook.SaveAs (sFileName)
   MsgBox "文件" & vbCrLf & sFileName & vbCrLf & "已保存到桌面"
End Sub
Sub 除零()
    Dim rng As Range
    For Each rng In Selection
        If rng = 0 Then rng.ClearContents
    Next
End Sub

Sub 平台合并去重()
    If InStr(ActiveWorkbook.Name, "已确认未配送") = 0 Then
        MsgBox "此脚本不适用于本文件"
        Exit Sub
    End If
    Dim sStart As Long
    Dim sEnd As Long
    Dim wbk As Workbook
    Application.ScreenUpdating = False
    Range("A1").CurrentRegion.Copy
    Set wbk = Workbooks.Open("F:\青云\(synology)\竞品\平台汇总去重(总表2019-4-1起).xlsx")
    'Workbooks("平台汇总去重(总表).xlsx").Activate
   
    sStart = Range("A1").End(xlDown).Row
   
    Cells(Range("A1").End(xlDown).Row + 1, 1).Select
    ActiveSheet.Paste
    Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(1, 2, 13, 15, 16), _
        Header:=xlNo
    sEnd = Range("A1").End(xlDown).Row
    
    ActiveWindow.ScrollRow = 1
    Application.ScreenUpdating = True
    MsgBox "原记录数： " & sStart & vbCrLf & "总记录数： " & sEnd & "" & vbCrLf & "本次新增记录数： " & sEnd - sStart
    
End Sub

Sub 显示文件路径()
    MsgBox ActiveWorkbook.Path
End Sub

Sub 取消筛选()
    Dim rng As Range
    Dim rngAdress As String
    Dim rngLength As Integer
    Dim rngStart As Integer
    rngAdress = Range("A1").CurrentRegion.Address
    rngStart = InStr(rngAdress, ":$")
    rngLength = Len(rngAdress)
    rngAdress = Right(rngAdress, rngLength - rngStart)
    For i = 1 To Range(rngAdress).Column
        [A1].CurrentRegion.AutoFilter field:=i
    Next i
End Sub

Sub 删除空行()
    Dim rng As Range
    Set rng = Cells(Rows.Count, 1).End(xlUp)
    For i = rng.Row To 1 Step -1
        If Cells(i, 1) = "" Or IsEmpty(Cells(i, 1)) Then
            Cells(i, 1).EntireRow.Delete
        End If
    Next i
End Sub

Sub 加空行()
    Dim vCln
    Dim iCln As Integer
    Dim lStart As Long
    vCln = InputBox("请输入做为加空行依据的列号（数字)")
    If Not (IsNumeric(iCln)) Then
        MsgBox "列号不是数字，程序自动退出"
        Exit Sub
    End If
    iCln = Val(vCln)
    lStart = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lStart To 3 Step -1
        If Cells(i, iCln).Value <> Cells(i - 1, iCln).Value Then
            Rows(i).Insert , CopyOrigin:=xlFormatFromLeftOrAbove
        End If
    Next i
End Sub

Sub 正则菜钱()
    Dim objRegExp As Object
    Dim objMh As Object
    Dim i As Integer
    Dim dTotal As Double
    Dim sStr As String
    
    Set objRegExp = CreateObject("vbscript.regexp")
    
    With objRegExp
        .Global = True
        .Pattern = "[0-9]*\.?[0-9]+"
    End With
    
    Set objMh = objRegExp.Execute(ActiveCell.Offset(0, -1).Value)
    
    If objMh.Count > 0 Then
        For i = 0 To objMh.Count - 1
            sStr = sStr & "+" & objMh(i)
            dTotal = dTotal + Val(objMh(i))
        Next i
    Else
        Exit Sub
    End If
    
    MsgBox Right(sStr, Len(sStr) - 1) & "=" & dTotal
    ActiveCell = dTotal
    
End Sub
Sub 竞品数据预处理()
    '获取总行数
    Dim num As Long
    num = Cells(Rows.Count, 1).End(xlUp).Row
    
    '关闭EXCEL刷新
    Application.ScreenUpdating = False
    
   Range("T:T,S:S,R:R,Q:Q,I:I,H:H,G:G,E:E,D:D,B:B").Delete Shift:=xlToLeft
   
    Columns("H:H").TextToColumns Destination:=Range("H1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("J:J").TextToColumns Destination:=Range("J1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("I:I").TextToColumns Destination:=Range("I1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

    With Columns("G:G")
        .NumberFormatLocal = "yyyy/m/d"
        .TextToColumns Destination:=Range("G1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
            :=Array(Array(1, 5), Array(2, 9)), TrailingMinusNumbers:=True
        .EntireColumn.AutoFit
    End With
    

    
    Columns("H:H").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("H:H").NumberFormatLocal = "G/通用格式"
    
    With Range("H1")
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Value = "当月日期"
    End With
    Range("H2").Formula = "=DAY(G2)"
    Range("H2").AutoFill Destination:=Range("H2:H" & num)
    Columns("G:G").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
 

    Range("G1").Formula = "配送企业简称"
    Range("G2").Formula = "=IFERROR(VLOOKUP(F2,'F:\青云\(synology)\竞品\竞品对照表.xlsx'!$F$2:$G$100,2,0),"""")"
    'Range("G2").Formula = "=IFERROR(VLOOKUP(F2,'F:\青云\(synology)\竞品\竞品对照表.xlsx'!表1[#All],2,0),"""")"
    'Range("G2").Formula2R1C1 = "=IFERROR(VLOOKUP(RC[-1],'F:\青云\(synology)\竞品\[竞品对照表.xlsx]Sheet1'!表1[#All],2,0),"""")"
    Range("G2").AutoFill Destination:=Range("G2:G" & num)
    Columns("G:G").EntireColumn.AutoFit
    Columns("F:F").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Range("F1").Formula = "疾控简称"
    Range("F2").Formula = "=IFERROR(VLOOKUP(E2,'F:\青云\(synology)\竞品\竞品对照表.xlsx'!$B$2:$C$200,2,0),"""")"
    'Range("G2").Formula = "=IFERROR(H2,'F:\青云\(synology)\竞品\竞品对照表.xlsx'!表1[#All],2,0),"""")"
    Range("F2").AutoFill Destination:=Range("F2:F" & num)
    Columns("E:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").Formula = "申报企业简称"
    Range("E2").Formula = "=IFERROR(VLOOKUP(D2,'F:\青云\(synology)\竞品\竞品对照表.xlsx'!$D$2:$E$200,2,0),"""")"
    Range("E2").AutoFill Destination:=Range("E2:E" & num)
    Columns("E:E").EntireColumn.AutoFit
    Columns("D:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Formula = "规格"
    Range("D2").Formula = "=IFS(C2=""瓶"",""西林"",C2=""支"",""预充"",TRUE,"""")"
    Range("D2").AutoFill Destination:=Range("D2:D" & num)
    Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C1").Formula = "疫苗简称"
    Range("C2").Formula = "=IFERROR(VLOOKUP(B2,'F:\青云\(synology)\竞品\竞品对照表.xlsx'!$H$2:$I$200,2,0),"""")"
    Range("C2").AutoFill Destination:=Range("C2:C" & num)
    Columns("C:C").EntireColumn.AutoFit
    Range("A1:P" & num).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("D:D").Delete Shift:=xlToLeft
    Range("A1").Select
    
    '恢复屏幕刷新
    Application.ScreenUpdating = True
    
    MsgBox "处理完成！"
End Sub

Sub 更新商务理事库存表()
    Dim fso As Object
    Dim fsoFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sTokill As String
    sTokill = ActiveWorkbook.FullName
    On Error GoTo handle:
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs "\\YBM-DYJ-PC\share\库存.xls"
    Application.DisplayAlerts = True
    MsgBox "库存已更新!"
    
    If Right(sTokill, 4) = "xlam" Then
        Exit Sub
    Else
        fso.deletefile sTokill
        MsgBox "文件 " & fso.getfilename(sTokill) & " 已删除"
        Set fso = Nothing
    End If
    
    ActiveWorkbook.Close
    Exit Sub
   
handle:
    MsgBox "出错了，请检查！"
       
End Sub

Sub 处理交割单()
    '判断是否是交割单，不是则退出
    If InStr(ActiveWorkbook.Name, "交割单查询") = 0 Then
        MsgBox "当前工作表不是交割单"
        Exit Sub
    End If
    
    Dim wsh As Worksheet
    On Error GoTo handle:
    Set wsh = Workbooks("股票2016.xlsm").Worksheets("2016年申万融资帐户明细")
    
    Dim startRows
    startRows = Range("A" & Cells.Rows.Count).End(xlUp).Row - 1
    
    '把要处理的数据读入数组
    ReDim arr(1 To startRows, 1 To 11)
    For i = 1 To startRows
        arr(i, 1) = CDate(Left(Range("A" & i + 1), 4) & "/" & Mid(Range("A" & i + 1), 5, 2) & "/" & Right(Range("A" & i + 1), 2))
        arr(i, 2) = Range("D" & i + 1)
        arr(i, 3) = Range("E" & i + 1)
        arr(i, 4) = Range("F" & i + 1)
        arr(i, 5) = Range("I" & i + 1)
        arr(i, 6) = Range("G" & i + 1)
        arr(i, 7) = Range("J" & i + 1) '成交金额
        arr(i, 8) = Range("U" & i + 1) '实际发生金额
        arr(i, 9) = Range("L" & i + 1) '手续费
        arr(i, 10) = Range("N" & i + 1) '印花税
        arr(i, 11) = Range("O" & i + 1) '过户费
    Next i
    
    '把数组写入统计表
    wsh.Activate
    endRows = Range("A" & wsh.Cells.Rows.Count).End(xlUp).Row + 1
    Range("A" & endRows).Resize(startRows, 11) = arr
    Range("A" & endRows).Resize(startRows, 11).Select
    
    Exit Sub
handle:
    MsgBox "股票2016.xlsm 未打开"
End Sub

Sub 去重平台数据()

    '获取要复制文件所在的目录名--------------***
    Dim copyFilePath As String
    With Application.FileDialog(msoFileDialogFolderPicker)     '打开文件对话框，选择要复制的文件所在的文件夹
        .InitialFileName = "C:\"                                '初始文件夹为C盘根目录
        .Title = "请选择要复制的文件所在的文件夹"
        .Show
        If .SelectedItems.Count > 0 Then                        '如果选择了文件夹
           copyFilePath = .SelectedItems(1)
           'MsgBox copyFilePath
           Else
            MsgBox "没有选择任何目录，退出程序"
            Exit Sub
        End If
    End With
    '***----------------------------------****



    Dim wb As Workbook, Erow As Long, fn As String, FileName As String, sht As Worksheet
    Application.Workbooks.Add  '新建一个工作簿
    
   
    Application.ScreenUpdating = False
    FileName = Dir(copyFilePath & "\*.*") '要复制的文件名
    Dim bt As Boolean      '设定表头是否已复制
    bt = False
    Do While FileName <> ""
        fn = copyFilePath & "\" & FileName '要复制的文件全路径名
        Set wb = GetObject(fn)  '将fn代表的工作簿对象赋给变量
        Set sht = wb.Worksheets(1) '汇总第1张工作表
        
        '取得汇总表第一个非空行---------------------
        If ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Value = "" Then
                Erow = 1
            Else
                Erow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
        End If
       

        sht.UsedRange.Copy ActiveSheet.Cells(Erow, 1) '复制第一张工作表数据
        wb.Close False    '关闭已复制的工作簿
        
        '去重
        Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(1, 2, 13, 15, 16), _
            Header:=xlNo
        
        
        FileName = Dir    '定位下一个要复制的工作簿
    Loop
   Application.ScreenUpdating = True
      

    
End Sub

Sub 合并单元格并填充数据()
    Dim rng As Range
    Dim str As String
    Dim adds
    For Each rng In Selection
        If rng.MergeCells = True Then
            str = rng.Value
            adds = rng.MergeArea.Address
            rng.UnMerge
            Range(adds) = str
            
        End If
    Next
End Sub
