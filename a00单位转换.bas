Attribute VB_Name = "a00单位转换"
Function Conversion_of_Unit(Unit As Single, PropertyID As Single) As Single

'这个程序主要用来实现 单位的转换
'其中PropertyID = 10xx为“长度”单位转换
'其中PropertyID = 11xx为“重量”单位转换



'长度单位换算
'PropertyID = 1001 单位为 英尺(ft) 转换为 米(m)
'PropertyID = 1002 单位为 米(m)    转换为 英尺(ft)
'PropertyID = 1011 单位为 英寸(in) 转换为 英尺(ft)
'PropertyID = 1012 单位为 英尺(ft) 转换为 英寸(in)
'PropertyID = 1021 单位为 英寸(in) 转换为 米(m)
'PropertyID = 1022 单位为 米(m)    转换为 英寸(in)

'重量单位换算
'PropertyID = 1101 单位为 磅(lb)      转换为   千克(Kg)
'PropertyID = 1102 单位为 千克(Kg)    转换为   磅(lb)

'功率单位转换
'PropertyID = 1201 单位为 BHP   转换为  kW
'PropertyID = 1202 单位为 kW    转换为  BHP



'————————————————————————————————————————————'
   
'长度单位换算
    If PropertyID = 1001 Then
        Conversion_of_Unit = Unit * 0.3048
    ElseIf PropertyID = 1002 Then
        Conversion_of_Unit = Unit * 3.28084
    ElseIf PropertyID = 1011 Then
        Conversion_of_Unit = Unit * 0.083333
    ElseIf PropertyID = 1012 Then
        Conversion_of_Unit = Unit * 12.00005
    ElseIf PropertyID = 1021 Then
        Conversion_of_Unit = Unit * 0.0254
    ElseIf PropertyID = 1022 Then
        Conversion_of_Unit = Unit * 39.37008
    End If
    
'重量单位换算
    If PropertyID = 1101 Then
        Conversion_of_Unit = Unit * 0.453592
    ElseIf PropertyID = 1102 Then
        Conversion_of_Unit = Unit * 2.204623
    End If
    
'功率单位转换
    If PropertyID = 1201 Then
        Conversion_of_Unit = Unit * 0.746
    ElseIf PropertyID = 1202 Then
        Conversion_of_Unit = Unit * 1.34
    End If
    
    
    
    
    
End Function
