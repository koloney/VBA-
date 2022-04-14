Attribute VB_Name = "随机奖励名单创建"
'参数count随机数量
Sub CreateRandomItems(count)

    '数据源表名
    source_sheets = "源名单表"
    target_sheets = "奖励名单表"
    If Sheets(source_sheets) Is Nothing Then
        MsgBox (source_sheets & "不存在")
        Exit Sub
    End If
    
    On Error Resume Next
    If Sheets(target_sheets) Is Nothing Then
        Sheets.Add(After:=Worksheets(Worksheets.count)).Name = target_sheets
    End If
    Worksheets(target_sheets).UsedRange.ClearContents
    
    Dim sc As New Collection
     '注意：属性名放在第1行，数据从第2行开始
    i = 2
    str1 = ""
    Do While Worksheets(source_sheets).Cells(i, 1) <> ""
       sc.Add (i)
       i = i + 1
    Loop
    
    '如果数据源数量超过需要随机的数量则进行随机
    
    If sc.count > count Then
        Dim rc As New Collection
        i = 1
        Do While i <= count
            r = Int(Rnd() * sc.count) + 1
            n = sc.Item(r)
            sc.Remove (r)
            rc.Add (n)
            i = i + 1
        Loop
    Else
        Set rc = sc
    End If
    
    '创建标题行
    j = 1
    Do While Worksheets(source_sheets).Cells(1, j) <> ""
         Worksheets(target_sheets).Cells(1, j) = Worksheets(source_sheets).Cells(1, j)
         j = j + 1
    Loop
    maxj = j - 1
    i = 1
    
    Do While i <= rc.count
        j = 1
        Do While j <= maxj
            si = rc.Item(i)
            Worksheets(target_sheets).Cells(i + 1, j) = Worksheets(source_sheets).Cells(si, j)
            j = j + 1
        Loop
    i = i + 1
    Loop
        
End Sub
Sub Sort_Target_Sheets(count)
    target_sheets = "奖励名单表"
    j = 1
    Do While Worksheets(target_sheets).Cells(1, j) <> ""
         j = j + 1
    Loop
    maxj = j - 1
    h = Chr(64 + maxj)

    Worksheets(target_sheets).Sort.SortFields.Clear
    Worksheets(target_sheets).Sort.SortFields.Add Key:=Range("A1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("奖励名单表").Sort
        .SetRange Range("A2:" & h & CStr(count + 1))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


Sub 随机发放奖励名单()    
    '奖励名额设置
    count = 300
    '奖励名额生成并保存到“奖励名单表”
    CreateRandomItems (count)
     '对“奖励名单表”显示排序（可选）
    Sort_Target_Sheets (count)
End Sub
