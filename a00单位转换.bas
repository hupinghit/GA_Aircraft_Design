Attribute VB_Name = "a00��λת��"
Function Conversion_of_Unit(Unit As Single, PropertyID As Single) As Single

'���������Ҫ����ʵ�� ��λ��ת��
'����PropertyID = 10xxΪ�����ȡ���λת��
'����PropertyID = 11xxΪ����������λת��



'���ȵ�λ����
'PropertyID = 1001 ��λΪ Ӣ��(ft) ת��Ϊ ��(m)
'PropertyID = 1002 ��λΪ ��(m)    ת��Ϊ Ӣ��(ft)
'PropertyID = 1011 ��λΪ Ӣ��(in) ת��Ϊ Ӣ��(ft)
'PropertyID = 1012 ��λΪ Ӣ��(ft) ת��Ϊ Ӣ��(in)
'PropertyID = 1021 ��λΪ Ӣ��(in) ת��Ϊ ��(m)
'PropertyID = 1022 ��λΪ ��(m)    ת��Ϊ Ӣ��(in)

'������λ����
'PropertyID = 1101 ��λΪ ��(lb)      ת��Ϊ   ǧ��(Kg)
'PropertyID = 1102 ��λΪ ǧ��(Kg)    ת��Ϊ   ��(lb)

'���ʵ�λת��
'PropertyID = 1201 ��λΪ BHP   ת��Ϊ  kW
'PropertyID = 1202 ��λΪ kW    ת��Ϊ  BHP



'����������������������������������������������������������������������������������������'
   
'���ȵ�λ����
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
    
'������λ����
    If PropertyID = 1101 Then
        Conversion_of_Unit = Unit * 0.453592
    ElseIf PropertyID = 1102 Then
        Conversion_of_Unit = Unit * 2.204623
    End If
    
'���ʵ�λת��
    If PropertyID = 1201 Then
        Conversion_of_Unit = Unit * 0.746
    ElseIf PropertyID = 1202 Then
        Conversion_of_Unit = Unit * 1.34
    End If
    
    
    
    
    
End Function
