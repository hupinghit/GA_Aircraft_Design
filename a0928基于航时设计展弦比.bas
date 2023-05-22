Attribute VB_Name = "a0928基于航时设计展弦比"
Function AR_e_E(C_L_cru As Single, E As Single, c_t As Single, W_ini As Single, W_fin As Single, C_D_min As Single) As Single
    
'基于设计航程缺点展弦比

'C_L_cru    巡航时的平均升力系数    单位为1
'E          航程                    单位为s
'c_t        发动机单位耗油率        单位为1/s
'W_ini      巡航开始时重量          单位为lbf
'W_fin      巡航结束时重量          单位为lbf
'C_D_min    最小阻力系数            单位为1

    AR_e_E = C_L_cru ^ 2 / 3.1415926 / (C_L_cru / E / c_t * Log(W_ini / W_fin) - C_D_min)
    
    
End Function

