Attribute VB_Name = "ģ��1"
Sub qy_merge() '�ϲ���Ԫ�񣬲������е�Ԫ���ֵ�ÿո�ָ������ı���ʽ���ᣬ��Ϊ�ϲ���Ԫ���ֵ��������ʾ
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

Sub qy_merge_new() '�ϲ���Ԫ�񣬲������е�Ԫ���ֵ�����������ָ������ı���ʽ���ᣬ��Ϊ�ϲ���Ԫ���ֵ��������ʾ
    
    Dim rng As Range
    Dim str() As String
    Dim sum As Integer
    Dim i As Integer
    Dim a
   
    On Error Resume Next
    Set rng = Selection
    sum = 0
    a = InputBox("������һ���ָ��������ڷָ�����Ԫ�񣬰�ȡ��Ϊ���ָ�")
    
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

Public Sub copyIM() '������ݺϲ����±�
    Dim sht As Worksheet, xrow As Integer, rng As Range
    Dim lastCount As String
    Dim rngS As Range
    
    Worksheets.Add(before:=Worksheets(1)).Name = "�½����ܱ�"
    
    '���Ƹ������ݵ��±�
    For Each sht In Worksheets
        If sht.Name <> ActiveSheet.Name Then     '����������ǵ�ǰ����ı�
            Set rng = Range("A65536").End(xlUp).Offset(1, 0)  'ȡ�õ�ǰ��A�е�һ���ǿյ�Ԫ��
            xrow = sht.Range("A1").CurrentRegion.Rows.Count - 1  'ȡ��Ҫ���ƵĹ����������������ȥ���У�
            sht.Range("A2").Resize(xrow, 7).Copy rng  '�Ѵ˹������A2��Ԫ��ʼ�ģ�xrow�У�7��������򣬸��Ƶ���������һ���ǿյ�Ԫ��
        End If
    Next
    
    'ɾ�����˵�һ�ű���ĸ�����ͷ
    xrow = Range("B1").CurrentRegion.Rows.Count
    lastCount = "B5:" & "B" & xrow
    Set rngS = Range(lastCount)
    
    For Each rng In rngS
        
        If InStr(rng, "ҳ��") > 0 Then   '����õ�Ԫ���ֵ������ҳ�롱�����
           rng.EntireRow.Delete   'ɾ������
           'ActiveSheet.Rows(rng.Row).Delete
        End If
    
    Next rng
    
    Range("B2:F2").Select
    Call qy_merge
    
End Sub


Sub copyWorkbook()   '��ͬһ�ļ����ڵ����й������ϲ���һ���µĹ�����
    
    
    '��ȡҪ�����ļ����ڵ�Ŀ¼��--------------***
    Dim copyFilePath As String
    With Application.FileDialog(msoFileDialogFolderPicker)     '���ļ��Ի���ѡ��Ҫ���Ƶ��ļ����ڵ��ļ���
        .InitialFileName = "C:\"                                '��ʼ�ļ���ΪC�̸�Ŀ¼
        .Title = "��ѡ��Ҫ���Ƶ��ļ����ڵ��ļ���"
        .Show
        If .SelectedItems.Count > 0 Then                        '���ѡ�����ļ���
           copyFilePath = .SelectedItems(1)
           'MsgBox copyFilePath
           Else
            MsgBox "û��ѡ���κ�Ŀ¼���˳�����"
            Exit Sub
        End If
    End With
    '***----------------------------------****
    
    
    'ѡ����ȫ�����ƣ����ǲ����Ʊ�ͷ----------------***
    Dim fullCopy As Boolean
    If MsgBox("��������ѡ���ǡ���ֻ����һ�ݱ�ͷ��ѡ����", vbYesNo, "��ѡ���Ʒ�ʽ") = vbYes Then
        fullCopy = True
    Else
        fullCopy = False
    End If
    '***-------------------------------------------****

    Dim wb As Workbook, Erow As Long, fn As String, FileName As String, sht As Worksheet
    Application.Workbooks.Add  '�½�һ��������
    
   
    Application.ScreenUpdating = False
    FileName = Dir(copyFilePath & "\*.*") 'Ҫ���Ƶ��ļ���
    Dim bt As Boolean      '�趨��ͷ�Ƿ��Ѹ���
    bt = False
    Do While FileName <> ""
        fn = copyFilePath & "\" & FileName 'Ҫ���Ƶ��ļ�ȫ·����
        Set wb = GetObject(fn)  '��fn����Ĺ��������󸳸�����
        Set sht = wb.Worksheets(1) '���ܵ�1�Ź�����
        
        'ȡ�û��ܱ��һ���ǿ���---------------------
        If ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Value = "" Then
                Erow = 1
            Else
                Erow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
        End If
       
        If fullCopy = True Then                                'ȫ����
            sht.UsedRange.Copy ActiveSheet.Cells(Erow, 1)
        Else                                                  'ֻ����һ�α�ͷ
            If bt = False Then
                sht.UsedRange.Copy ActiveSheet.Cells(Erow, 1)
                bt = True
            Else
                sht.Cells(2, 1).Resize(sht.UsedRange.Rows.Count - 1, sht.Cells(sht.Rows.Count, 1).End(xlUp).End(xlToRight).Column).Copy ActiveSheet.Cells(Erow, 1)
            End If
        End If
        
        wb.Close False    '�ر��Ѹ��ƵĹ�����
        FileName = Dir    '��λ��һ��Ҫ���ƵĹ�����
    Loop
   Application.ScreenUpdating = True
      
End Sub

Sub delRow()   'ɾ��ĳ���ض���Ԫ�����ڵ���
    Dim specialWord As String
    specialWord = Application.InputBox("������Ҫɾ�������еĵ���")
    
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

Sub ��ȡ��Ԫ�������ֲ����()
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
        If VBA.IsNumeric(tmp) = False And tmp <> "." Then  '�ж��ǲ������ֺ�С����
              If start Then  '�������ֺ�С���㣬����Ƿ�ʼ����
              '�ѿ�ʼ����
              total = total + Val(tmpNum)
              printIfo = printIfo & tmpNum & "+"
              tmpNum = ""
              start = False
              
              End If
        Else        '�����ֺ�С����
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

Public Sub copyAllSheets() 'ͬ���������ϲ�
    Dim sht As Worksheet, xrow As Integer, rng As Range
    Dim lastCount As String
    Dim rngS As Range
    
    Worksheets.Add(before:=Worksheets(1)).Name = "�ϲ���" & WorksheetFunction.Substitute(Now(), ":", "-")
    
    '���Ƹ������ݵ��±�
    For Each sht In Worksheets
        If sht.Name <> ActiveSheet.Name Then     '����������ǵ�ǰ����ı�
            If Range("A1") = "" And Range("B1") = "" Then
                Set rng = Range("A1")
            Else
                Set rng = Range("A65536").End(xlUp).Offset(1, 0)  'ȡ�õ�ǰ��A�е�һ���ǿյ�Ԫ��
            End If
                    
            'xrow = sht.Range("A1").CurrentRegion.Rows.Count - 1  'ȡ��Ҫ���ƵĹ����������������ȥ���У�
            'sht.Range("A2").Resize(xrow, 7).Copy rng  '�Ѵ˹������A2��Ԫ��ʼ�ģ�xrow�У�7��������򣬸��Ƶ���������һ���ǿյ�Ԫ��
            sht.UsedRange.Copy rng
        End If
    Next
    
    'ɾ�����˵�һ�ű���ĸ�����ͷ
'    xrow = Range("B1").CurrentRegion.Rows.Count
'    lastCount = "B5:" & "B" & xrow
'    Set rngS = Range(lastCount)
    
   
'    For Each rng In rngS
'
'        If InStr(rng, "ҳ��") > 0 Then   '����õ�Ԫ���ֵ������ҳ�롱�����
'           rng.EntireRow.Delete   'ɾ������
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
        Cells(i, "S").Offset(0, -1) = "�ϼ�"
        total = 0
    Else
        total = total + Cells(i, "S").Value
    End If
Next i
Application.ScreenUpdating = True
End Sub

Sub old��Ʒ����Ԥ����()
    '��ȡ������
    Dim num As Long
    'Range("A1").Select
    'Range("A1").End(xlDown).Select
    'num = ActiveCell.Row
    num = Cells(Rows.Count, 1).End(xlUp).Row
    
    '�ر�EXCELˢ��
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
    Columns("I:I").NumberFormatLocal = "G/ͨ�ø�ʽ"
    
    With Range("I1")
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Value = "��������"
    End With
    'ActiveCell.FormulaR1C1 = "��������"
    Range("I2").Formula = "=DAY(H2)"
    Range("I2").AutoFill Destination:=Range("I2:I" & num)
    Columns("H:H").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
 

    Range("H1").Formula = "������ҵ���"
    Range("H2").Formula = "=IFERROR(VLOOKUP(G2,'F:\����\(synology)\��Ʒ\��Ʒ���ձ�.xlsx'!��1[#All],2,0),"""")"
    'Range("H2").Formula = "=IFERROR(G2,'F:\����\(synology)\��Ʒ\��Ʒ���ձ�.xlsx'!��1[#All],2,0),"""")"
    Range("H2").AutoFill Destination:=Range("H2:H" & num)
    Columns("H:H").EntireColumn.AutoFit
    Columns("G:G").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Range("G1").Formula = "���ؼ��"
    Range("G2").Formula = "=IFERROR(VLOOKUP(F2,'F:\����\(synology)\��Ʒ\��Ʒ���ձ�.xlsx'!��2[#All],2,0),"""")"
    'Range("G2").Formula = "=IFERROR(H2,'F:\����\(synology)\��Ʒ\��Ʒ���ձ�.xlsx'!��1[#All],2,0),"""")"
    Range("G2").AutoFill Destination:=Range("G2:G" & num)
    Columns("F:F").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("F1").Formula = "�걨��ҵ���"
    Range("F2").Formula = "=IFERROR(VLOOKUP(E2,'F:\����\(synology)\��Ʒ\��Ʒ���ձ�.xlsx'!��3[#All],2,0),"""")"
    Range("F2").AutoFill Destination:=Range("F2:F" & num)
    Columns("F:F").EntireColumn.AutoFit
    'Columns("E:E").Select
    Columns("D:D").Delete Shift:=xlToLeft
    Columns("D:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Formula = "���"
    Range("D2").FormulaR1C1 = "=IFERROR(IF(FIND(""����"",RC[-1])>0,""����""),IFERROR(IF(FIND(""Ԥ��"",RC[-1])>0,""Ԥ��""),""""))"
    Range("D2").AutoFill Destination:=Range("D2:D" & num)
    Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C1").Formula = "������"
    Range("C2").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],[��Ʒ���ձ�.xlsx]Sheet1!C8:C9,2,0),"""")"
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
    
    '�ָ���Ļˢ��
    Application.ScreenUpdating = True
    
    MsgBox "������ɣ�"
End Sub



Function CaiQiang(ByVal rng As Range) '�����Ǯ�Ĺ�ʽ
    'Application.Volatile True '��ʧ�Ժ���
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
        If VBA.IsNumeric(tmp) = False And tmp <> "." Then  '�ж��ǲ������ֺ�С����
              If start Then  '�������ֺ�С���㣬����Ƿ�ʼ����
              '�ѿ�ʼ����
              total = total + Val(tmpNum)
              printIfo = printIfo & tmpNum & "+"
              tmpNum = ""
              start = False
              
              End If
        Else        '�����ֺ�С����
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
        r.Value = r.Value & "��"
    Next r
End Sub
Sub �ձ�������()
   '��ȡ������
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
    ActiveCell.FormulaR1C1 = "��Ʒ+���"
   ' Range("A2").Select
   ' Application.SendKeys ("^a")
    

    
    '���е�Ԫ���ȫ����,���У����и�
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
Sub ����ɫ��()
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
    
   '�������һ��ɫ��
   With Range(Cells(Selection.Row + Selection.Rows.Count - 1, Selection.Column), Cells(Selection.Row + Selection.Rows.Count - 1, Selection.Column + Selection.Columns.Count - 1))
        .Interior.Color = RGB(91, 155, 213)
        .Font.Color = RGB(255, 255, 255)
   End With
   
    '���õ�һ��ɫ��
    With Range(Cells(Selection.Row, Selection.Column), Cells(Selection.Row, Selection.Column + Selection.Columns.Count - 1))
        .Interior.Color = RGB(91, 155, 213)
        .Font.Color = RGB(255, 255, 255)
   End With
   
  
End Sub

Sub ɫ��()

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
        timeSection = "8-12��"
    ElseIf (rng1 >= 12 And rng1 < 14) Or (rng1 = 14 And rng2 < 30) Then
            timeSection = "12��-14:30"
    ElseIf (rng1 = 14 And rng2 >= 30) Or (rng1 >= 15 And rng1 < 18) Then
            timeSection = "14:30 - 18:00"
    Else
            timeSection = "18:00�Ժ�"
    End If
        
End Function

Sub �ض�����������ɫ()
    Dim num As Long
    num = Range("A2").CurrentRegion.Columns.Count
    
    Dim specialWord As String
    specialWord = Application.InputBox("��������Ҫ�����а����ĵ���")

    If specialWord = "" Then Exit Sub
    Dim rng As Range
    For Each rng In ActiveCell.CurrentRegion
        If InStr(specialWord, rng) > 0 Then
            'Rows(rng.Row).Interior.ColorIndex = 24
            Range(Cells(rng.Row, rng.CurrentRegion.Column), Cells(rng.Row, num)).Interior.ColorIndex = 22
        End If
    Next
End Sub

Function QyHeBinWenBen(ParamArray inp())  '�ϲ���Ԫ���ı�
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

Sub ��ʽת��ֵ()
    Dim ing As Range
    For Each ing In Selection
        If ing.HasFormula Then ing.Value = ing
    Next
End Sub

Public Sub �����ձ�()
   '����---------------
    Dim dayRow As Integer, dayColumn As Integer
    Dim dayColumnBegin As Integer
    Dim dayColumnEnd As Integer
    
    dayRow = Selection.Row
    dayColumnBegin = Selection.Column + 3
    dayColumnEnd = Cells(Selection.Row, Selection.Column).End(xlToRight).Column - 1
    For i = dayColumnBegin To dayColumnEnd
        Cells(dayRow, i) = Cells(dayRow, i) & "��"
    Next
  '.-------------------END
    dayColumn = Cells(dayRow, Selection.Column).End(xlToRight).Column - 1
    


    'ȷ���ļ���
    Dim sFileName As String
    Dim sdate As String

    If Val(Left(Cells(2, dayColumnEnd), Len(Cells(2, dayColumnEnd)) - 1)) < Day(Now()) Then
        sdate = Year(Now()) & "-" & Month(Now()) & "-" & Left(Cells(dayRow, dayColumn), Len(Cells(dayRow, dayColumn)) - 1)
        'sFileName = "C:\Users\QingYun\Desktop\��Ʒ�ձ�" & "(" & sdate & ")"
        sFileName = Environ("userprofile") & "\Desktop\" & "��Ʒ�ձ�" & "(" & sdate & ")"
    Else
        If Month(Now()) = 1 Then
           sdate = Year(Now()) - 1 & "-" & 12 & "-" & Left(Cells(dayRow, dayColumn), Len(Cells(dayRow, dayColumn)) - 1)
        Else
            sdate = Year(Now()) & "-" & Month(Now()) - 1 & "-" & Left(Cells(dayRow, dayColumn), Len(Cells(dayRow, dayColumn)) - 1)
        End If
        sFileName = Environ("userprofile") & "\Desktop\" & "��Ʒ�ձ�" & "(" & sdate & ")"
    End If

   '�����ļ�
   ActiveWorkbook.SaveAs (sFileName)
   MsgBox "�ļ�" & vbCrLf & sFileName & vbCrLf & "�ѱ��浽����"
End Sub

Public Sub �����()
    '����---------------------
    Dim dayRow As Integer, dayColumn As Integer
    Dim dayColumnBegin As Integer
    Dim dayColumnEnd As Integer
    
    dayRow = Selection.Row
    dayColumnBegin = Selection.Column + 3
    dayColumnEnd = Cells(Selection.Row, Selection.Column).End(xlToRight).Column
    For i = dayColumnBegin To dayColumnEnd
        Cells(dayRow, i) = Cells(dayRow, i) & "��"
    Next
   '-------------------------END
   
    dayColumn = Cells(dayRow, Selection.Column).End(xlToRight).Column
    
   Dim sFileName As String
   Dim sdate As String
   'sFileName = "C:\Users\QingYun\Desktop\��Ʒ��" & "(" & Year(Now()) & "-" & Month(Now()) & "-" & Left(Cells(dayRow, dayColumn), Len(Cells(dayRow, dayColumn)) - 1) & ")"
    If Val(Left(Cells(2, dayColumnEnd), Len(Cells(2, dayColumnEnd)) - 1)) < Day(Now()) Then
        sdate = Year(Now()) & "-" & Month(Now()) & "-" & Left(Cells(dayRow, dayColumn), Len(Cells(dayRow, dayColumn)) - 1)
        'sFileName = "C:\Users\QingYun\Desktop\��Ʒ�ձ�" & "(" & sdate & ")"
        sFileName = Environ("userprofile") & "\Desktop\" & "��Ʒ��" & "(" & sdate & ")"
    Else
        If Month(Now()) = 1 Then
           sdate = Year(Now()) - 1 & "-" & 12 & "-" & Left(Cells(dayRow, dayColumn), Len(Cells(dayRow, dayColumn)) - 1)
        Else
            sdate = Year(Now()) & "-" & Month(Now()) - 1 & "-" & Left(Cells(dayRow, dayColumn), Len(Cells(dayRow, dayColumn)) - 1)
        End If
        sFileName = Environ("userprofile") & "\Desktop\" & "��Ʒ��" & "(" & sdate & ")"
    End If
   
   
   
   ActiveWorkbook.SaveAs (sFileName)
   MsgBox "�ļ�" & vbCrLf & sFileName & vbCrLf & "�ѱ��浽����"
End Sub
Sub ����()
    Dim rng As Range
    For Each rng In Selection
        If rng = 0 Then rng.ClearContents
    Next
End Sub

Sub ƽ̨�ϲ�ȥ��()
    If InStr(ActiveWorkbook.Name, "��ȷ��δ����") = 0 Then
        MsgBox "�˽ű��������ڱ��ļ�"
        Exit Sub
    End If
    Dim sStart As Long
    Dim sEnd As Long
    Dim wbk As Workbook
    Application.ScreenUpdating = False
    Range("A1").CurrentRegion.Copy
    Set wbk = Workbooks.Open("F:\����\(synology)\��Ʒ\ƽ̨����ȥ��(�ܱ�2019-4-1��).xlsx")
    'Workbooks("ƽ̨����ȥ��(�ܱ�).xlsx").Activate
   
    sStart = Range("A1").End(xlDown).Row
   
    Cells(Range("A1").End(xlDown).Row + 1, 1).Select
    ActiveSheet.Paste
    Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(1, 2, 13, 15, 16), _
        Header:=xlNo
    sEnd = Range("A1").End(xlDown).Row
    
    ActiveWindow.ScrollRow = 1
    Application.ScreenUpdating = True
    MsgBox "ԭ��¼���� " & sStart & vbCrLf & "�ܼ�¼���� " & sEnd & "" & vbCrLf & "����������¼���� " & sEnd - sStart
    
End Sub

Sub ��ʾ�ļ�·��()
    MsgBox ActiveWorkbook.Path
End Sub

Sub ȡ��ɸѡ()
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

Sub ɾ������()
    Dim rng As Range
    Set rng = Cells(Rows.Count, 1).End(xlUp)
    For i = rng.Row To 1 Step -1
        If Cells(i, 1) = "" Or IsEmpty(Cells(i, 1)) Then
            Cells(i, 1).EntireRow.Delete
        End If
    Next i
End Sub

Sub �ӿ���()
    Dim vCln
    Dim iCln As Integer
    Dim lStart As Long
    vCln = InputBox("��������Ϊ�ӿ������ݵ��кţ�����)")
    If Not (IsNumeric(iCln)) Then
        MsgBox "�кŲ������֣������Զ��˳�"
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

Sub �����Ǯ()
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
Sub ��Ʒ����Ԥ����()
    '��ȡ������
    Dim num As Long
    num = Cells(Rows.Count, 1).End(xlUp).Row
    
    '�ر�EXCELˢ��
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
    Columns("H:H").NumberFormatLocal = "G/ͨ�ø�ʽ"
    
    With Range("H1")
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Value = "��������"
    End With
    Range("H2").Formula = "=DAY(G2)"
    Range("H2").AutoFill Destination:=Range("H2:H" & num)
    Columns("G:G").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
 

    Range("G1").Formula = "������ҵ���"
    Range("G2").Formula = "=IFERROR(VLOOKUP(F2,'F:\����\(synology)\��Ʒ\��Ʒ���ձ�.xlsx'!$F$2:$G$100,2,0),"""")"
    'Range("G2").Formula = "=IFERROR(VLOOKUP(F2,'F:\����\(synology)\��Ʒ\��Ʒ���ձ�.xlsx'!��1[#All],2,0),"""")"
    'Range("G2").Formula2R1C1 = "=IFERROR(VLOOKUP(RC[-1],'F:\����\(synology)\��Ʒ\[��Ʒ���ձ�.xlsx]Sheet1'!��1[#All],2,0),"""")"
    Range("G2").AutoFill Destination:=Range("G2:G" & num)
    Columns("G:G").EntireColumn.AutoFit
    Columns("F:F").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Range("F1").Formula = "���ؼ��"
    Range("F2").Formula = "=IFERROR(VLOOKUP(E2,'F:\����\(synology)\��Ʒ\��Ʒ���ձ�.xlsx'!$B$2:$C$200,2,0),"""")"
    'Range("G2").Formula = "=IFERROR(H2,'F:\����\(synology)\��Ʒ\��Ʒ���ձ�.xlsx'!��1[#All],2,0),"""")"
    Range("F2").AutoFill Destination:=Range("F2:F" & num)
    Columns("E:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").Formula = "�걨��ҵ���"
    Range("E2").Formula = "=IFERROR(VLOOKUP(D2,'F:\����\(synology)\��Ʒ\��Ʒ���ձ�.xlsx'!$D$2:$E$200,2,0),"""")"
    Range("E2").AutoFill Destination:=Range("E2:E" & num)
    Columns("E:E").EntireColumn.AutoFit
    Columns("D:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Formula = "���"
    Range("D2").Formula = "=IFS(C2=""ƿ"",""����"",C2=""֧"",""Ԥ��"",TRUE,"""")"
    Range("D2").AutoFill Destination:=Range("D2:D" & num)
    Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C1").Formula = "������"
    Range("C2").Formula = "=IFERROR(VLOOKUP(B2,'F:\����\(synology)\��Ʒ\��Ʒ���ձ�.xlsx'!$H$2:$I$200,2,0),"""")"
    Range("C2").AutoFill Destination:=Range("C2:C" & num)
    Columns("C:C").EntireColumn.AutoFit
    Range("A1:P" & num).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("D:D").Delete Shift:=xlToLeft
    Range("A1").Select
    
    '�ָ���Ļˢ��
    Application.ScreenUpdating = True
    
    MsgBox "������ɣ�"
End Sub

Sub �����������¿���()
    Dim fso As Object
    Dim fsoFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sTokill As String
    sTokill = ActiveWorkbook.FullName
    On Error GoTo handle:
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs "\\YBM-DYJ-PC\share\���.xls"
    Application.DisplayAlerts = True
    MsgBox "����Ѹ���!"
    
    If Right(sTokill, 4) = "xlam" Then
        Exit Sub
    Else
        fso.deletefile sTokill
        MsgBox "�ļ� " & fso.getfilename(sTokill) & " ��ɾ��"
        Set fso = Nothing
    End If
    
    ActiveWorkbook.Close
    Exit Sub
   
handle:
    MsgBox "�����ˣ����飡"
       
End Sub

Sub �����()
    '�ж��Ƿ��ǽ�����������˳�
    If InStr(ActiveWorkbook.Name, "�����ѯ") = 0 Then
        MsgBox "��ǰ�������ǽ��"
        Exit Sub
    End If
    
    Dim wsh As Worksheet
    On Error GoTo handle:
    Set wsh = Workbooks("��Ʊ2016.xlsm").Worksheets("2016�����������ʻ���ϸ")
    
    Dim startRows
    startRows = Range("A" & Cells.Rows.Count).End(xlUp).Row - 1
    
    '��Ҫ��������ݶ�������
    ReDim arr(1 To startRows, 1 To 11)
    For i = 1 To startRows
        arr(i, 1) = CDate(Left(Range("A" & i + 1), 4) & "/" & Mid(Range("A" & i + 1), 5, 2) & "/" & Right(Range("A" & i + 1), 2))
        arr(i, 2) = Range("D" & i + 1)
        arr(i, 3) = Range("E" & i + 1)
        arr(i, 4) = Range("F" & i + 1)
        arr(i, 5) = Range("I" & i + 1)
        arr(i, 6) = Range("G" & i + 1)
        arr(i, 7) = Range("J" & i + 1) '�ɽ����
        arr(i, 8) = Range("U" & i + 1) 'ʵ�ʷ������
        arr(i, 9) = Range("L" & i + 1) '������
        arr(i, 10) = Range("N" & i + 1) 'ӡ��˰
        arr(i, 11) = Range("O" & i + 1) '������
    Next i
    
    '������д��ͳ�Ʊ�
    wsh.Activate
    endRows = Range("A" & wsh.Cells.Rows.Count).End(xlUp).Row + 1
    Range("A" & endRows).Resize(startRows, 11) = arr
    Range("A" & endRows).Resize(startRows, 11).Select
    
    Exit Sub
handle:
    MsgBox "��Ʊ2016.xlsm δ��"
End Sub

Sub ȥ��ƽ̨����()

    '��ȡҪ�����ļ����ڵ�Ŀ¼��--------------***
    Dim copyFilePath As String
    With Application.FileDialog(msoFileDialogFolderPicker)     '���ļ��Ի���ѡ��Ҫ���Ƶ��ļ����ڵ��ļ���
        .InitialFileName = "C:\"                                '��ʼ�ļ���ΪC�̸�Ŀ¼
        .Title = "��ѡ��Ҫ���Ƶ��ļ����ڵ��ļ���"
        .Show
        If .SelectedItems.Count > 0 Then                        '���ѡ�����ļ���
           copyFilePath = .SelectedItems(1)
           'MsgBox copyFilePath
           Else
            MsgBox "û��ѡ���κ�Ŀ¼���˳�����"
            Exit Sub
        End If
    End With
    '***----------------------------------****



    Dim wb As Workbook, Erow As Long, fn As String, FileName As String, sht As Worksheet
    Application.Workbooks.Add  '�½�һ��������
    
   
    Application.ScreenUpdating = False
    FileName = Dir(copyFilePath & "\*.*") 'Ҫ���Ƶ��ļ���
    Dim bt As Boolean      '�趨��ͷ�Ƿ��Ѹ���
    bt = False
    Do While FileName <> ""
        fn = copyFilePath & "\" & FileName 'Ҫ���Ƶ��ļ�ȫ·����
        Set wb = GetObject(fn)  '��fn����Ĺ��������󸳸�����
        Set sht = wb.Worksheets(1) '���ܵ�1�Ź�����
        
        'ȡ�û��ܱ��һ���ǿ���---------------------
        If ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Value = "" Then
                Erow = 1
            Else
                Erow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
        End If
       

        sht.UsedRange.Copy ActiveSheet.Cells(Erow, 1) '���Ƶ�һ�Ź���������
        wb.Close False    '�ر��Ѹ��ƵĹ�����
        
        'ȥ��
        Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(1, 2, 13, 15, 16), _
            Header:=xlNo
        
        
        FileName = Dir    '��λ��һ��Ҫ���ƵĹ�����
    Loop
   Application.ScreenUpdating = True
      

    
End Sub

Sub �ϲ���Ԫ���������()
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
