Attribute VB_Name = "a0928���ں�ʱ���չ�ұ�"
Function AR_e_E(C_L_cru As Single, E As Single, c_t As Single, W_ini As Single, W_fin As Single, C_D_min As Single) As Single
    
'������ƺ���ȱ��չ�ұ�

'C_L_cru    Ѳ��ʱ��ƽ������ϵ��    ��λΪ1
'E          ����                    ��λΪs
'c_t        ��������λ������        ��λΪ1/s
'W_ini      Ѳ����ʼʱ����          ��λΪlbf
'W_fin      Ѳ������ʱ����          ��λΪlbf
'C_D_min    ��С����ϵ��            ��λΪ1

    AR_e_E = C_L_cru ^ 2 / 3.1415926 / (C_L_cru / E / c_t * Log(W_ini / W_fin) - C_D_min)
    
    
End Function

