Sub onekey()

Dim formulate As String
Dim mass_column As Integer
Dim instrument_num_col As Integer
Dim water_rate_col As Integer

'1.指定此元素的仪器读数位于第几列，需要区分不同元素zn
instrument_num_col = 2

'2.指定质量数据位于序列工作表的第几列
mass_column = 6

'3.指定样品含水率数据位于第几列
water_rate_col = 7


'填充质量公式
For index = 15 To 21 Step 1
formulate = "=VLOOKUP($A" & index & ",序列!$A$1:$G$99," & mass_column & ",0)"
Range("d" & index).Value = formulate
Next

'填充仪器读数公式
For index2 = 13 To 21 Step 1
formulate = "=VLOOKUP($A" & index2 & ",序列!$A$1:$G$99," & instrument_num_col & ",0)"
Range("i" & index2).Value = formulate
Next

'填充结果计算公式
For index3 = 17 To 21 Step 1
formulate = "=I" & index3 & "*50*" & "H" & index3 & "/D" & index3 & "/L" & index3
Range("K" & index3).Value = formulate
Next

'填充含水率公式
For index3 = 17 To 21 Step 1
formulate = "=VLOOKUP($A" & index3 & ",序列!$A$1:$G$99," & water_rate_col & ",0)"
Range("l" & index3).Value = formulate
Next
'-------------------------------------------------------
'id为长编号
    Dim id As String
'length为序列的长度
    Const length As Integer = 99
    Dim i As Integer
    Dim i2 As Integer
    Dim i3 As Integer
    Dim i4 As Integer
    Dim id_array(1 To length) As String
    Dim perfix, id_short As String
    Dim ret
    Dim ret2
    Dim short_id_dict As Object
    Set short_id_dict = CreateObject("scripting.dictionary")
    Dim times As Integer
    Dim start_number As Integer
    
    times = 0
    perfix = "EN2021"
    
    '构建字典与数组
    For i = 1 To length Step 1
        id_array(i) = Sheets("序列").Range("A" & i).Value
        ret = InStr(1, id_array(i), perfix, 1)
        '判断是否为样品编号，是则加入字典
        If ret <> 0 Then
            '截取样品四位编号
            id_short = Mid(id_array(i), 9, 4)
            '将四位编号作为字典的键
            short_id_dict(id_short) = ""
        End If
    Next
    
    '逐个建立sheet
    For Each short_id In short_id_dict.keys
    
        For i2 = LBound(id_array) To UBound(id_array)
            ret2 = InStr(1, id_array(i2), short_id, 1)
            If ret2 <> 0 Then
                start_number = i2
                '得到第一个编号开始的位置，跳出循环
                Exit For
            End If
        Next
        
        '复制第一张样表
        Sheets("temp").Copy After:=ActiveSheet
        
        '以短编号命名当前选中的工作表
        ActiveSheet.Name = short_id
        
        '计算当前short_id出现次数
        For i3 = LBound(id_array) To UBound(id_array)
            'ret2作为四位编号是否在长编号的标志，为0则不在
            ret2 = InStr(1, id_array(i3), short_id, 1)
            '判断ret2是否为0，不为零则统计出现的次数
            If ret2 <> 0 Then
                times = times + 1
            End If
        Next
        
        '依据编号出现次数，填充编号
        For i4 = 17 To times + 16 Step 1
            Range("a" & i4).Value = id_array(start_number)
            start_number = 1 + start_number
        Next
        '重置出现次数
        times = 0
    Next

End Sub