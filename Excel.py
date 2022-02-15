Option Private Module 'Excel workbook为1个VBA工程，只在当前模块中使用以下定义的子流程
Dim Variable As String '定义适用于当前模块的变量
Public Variable As String '定义适用于所有workbooks的变量，或者配合Private Module适用于当前workbook

'默认定义Public Sub全局子流程，若在Private Module下或者Private Sub则定义当前模块局部子流程
Sub Content(ByRef AmountVariable As String) '定义参数，可选参数和默认值参数放最后Optional Param As Integer = 10 
'单引号用于注释，而且只能用英文，不能用中文
    Dim StringVariable As String '设置Sub变量数据类型
    StringVariable = "Cities" '字符串必须用双引号
    Sheets("sheet1").Select '选择指定worksheet，或者Worksheets(1)，Sheets(1)，Sheet1
    Range("A1").Select '默认选择当前worksheet的单元格，除非指定sheet
    ActiveCell.FormulaR1C1 = StringVariable '当前选中单元格的第1行第1列公式，其他写法：Selection.Value =, Range("A1").Value = 
    
    AmountVariable = "Amount2" '默认ByRef为浅拷贝，ByValue即深拷贝，决定在子过程中变量的变动是否会影响父过程
    Range("B1").Value = AmountVariable '直接指定值，还可以Cells(1,2).Value 第1行第2列

    Range("B31").Value = "=SUM(R[-29]C:R30C2)" '以B31为标准的相对行索引R[-29]和当前列相对索引C(B2)，以及绝对索引第30行第2列$B$30，从1而非0开始算
End Sub '结束模块



Sub Fill()
    Dim AmountVariable As String
    AmountVariable = "Amount1"
    Exit Sub 'break出当前子进程，不会运行以下代码，End则是结束所有进程
    Content AmountVariable '调用函数，用逗号隔开多个参数
    Content AmountVariable:=AmountVariable '加上参数名赋值
    Call Content(AmountVariable) '或者用call调用
    MsgBox AmountVariable '当ByRef时，从Amount1变为Amount2
End Sub



Sub Format()
    Range("A1:B1").Select '选择单元格范围，选择整列Range("A:A"), Columns("A")，整行rows("1")
    With Selection.Borders(xlEdgeBottom) '指定单元格边框格式，还有xlEdgeTop/Bottom/Right/Left，或者直接Borders默认所有方向
        .LineStyle = xlDouble '单元格边框样式，没有的为xlNone
        .Weight = xlThick '粗细
        .ColorIndex = xlAutomatic '颜色，或者数字
    End With 

    Range("B31").Select 
    With Selection.Font '字体格式
        .Color = -16776961 
        .TintAndShade = 0
        .Bold = True
        .Italic = True
        .Underline = xlUnderlineStyleSingle
    End With
    ExecuteExcel4Macro "PATTERNS(1,0,65535,TRUE,2,3,0,0)" '黄色填充

    Dim i As Integer '设置变量数据类型，还有Long
    For i = 2 To 30 '循环
        Worksheets(1).Range("B"&i).Select '迭代不同单元格
        If Selection.Value > 50000000 Then '通过&来连接字符串
            Selection.Font.Color = -16776961 '字体变红
        Else
            Selection.Font.Color = -11489280 '字体变绿
        End If
    Next i
    MsgBox "Finished" '弹窗
End Sub



Sub Calculation()
    MsgBox Application.WorksheetFunction.CountA(Sheets("sheet1").Columns("A:A")) '计算第A列的非空值行数
    MsgBox Sheet1.Columns("A:A").Count '计算第A列共有多少列
End Sub



Sub CreateWorksheet()
    For i = 1 To 3
        Sheets.Add '创建sheet，删除可以先选择sheet，再ActiveSheet.Delete
        ActiveSheet.Name = "sheet"&i
    Next
    MsgBox Sheets.Count 'sheets数量

    Sheets("sheet1").UsedRange.Copy '自动复制所有有效内容，或者指定Range("A1:B30")
    Sheets("sheet2").Range("A2").Select '指定开始粘贴内容的左上单元格
    Cells(Sheets("sheet2").Range("A65536").End(xlUp).Row+2, 2).Select '选择一个极大数为最后一行，单元格向上选择最后一个非空单元格行数，从而不覆盖原有内容情况下粘贴
    Sheets("sheet2").Paste '默认从A1开始粘贴

    rows("1:2").Insert '第1-2行插入空行，内容往下平移
    Range("A1") = "Workbook: " & ActiveWorkbook.Path & "/" & ActiveWorkbook.Name '获取现有
End Sub



Sub DeleteEmptyRows()
    Application.ScreenUpdating = False '停止屏幕刷新，适用于对内容产生删减时，希望后续判断仍基于原数据而非变动中的数据时
    Sheets(2).Select 
    For i = Sheets(2).UsedRange.rows.Count To 1 Step -1 '当停止刷新，可以从第1行开始，否则从最后1行开始
        If rows(i).Find("*") Is Nothing Then '如果行(对象)为None
            MsgBox "Row " & i & " is blank"
            Range("A" & i).Select '先选中该行
            Selection.EntireRow.Delete '删除整行
        ElseIf Len(Trim(rows(i).Find("*"))) = 0 Then '如果不为None，但是是空格，就通过trim判断，两个条件不能用or连接，因为None无法取trim
            MsgBox "Row " & i & " includes only spaces"
            Range("A" & i).Select
            Selection.EntireRow.Delete '____
        End If
    Next i
    Application.ScreenUpdating = True
End Sub