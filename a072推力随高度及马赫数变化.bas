Attribute VB_Name = "a072推力随高度及马赫数变化"
Function Engine_Thrust(F0 As Single, TR As Single, H As Single, m As Single, deltaOTA_C As Single, Mode As Byte) As Single

'this routine can be used to estimate the change in thrust depending on
'flight conditions for the following engine:
'
' 1.Piston engines(Mode = 0)
' 2.Turboprops(Mode = 1)
' 3.Turbojets, maximum power(Mode = 2)
' 4.Turbojets, military power(Mode = 3)
' 5.Low Bypass Ratio Turfans(Mode = 4)
' 6.Hign Bypass Ratio Turbofans(Mode = 5)
'
' Variables:    F0      = Engine thrust in lbf              推力
'               TR      = Throttle ratio                    节流比
'               H       = Altitude at condition, in ft      高度
'               M       = Mach Number                       马赫数
'               deltaOTA= Deviation from ISA in ℃          从ISA得到的数据
'
'NOTE1: For piston engines, the function only treats the power, so F0 is
'       the rated engine power at S - L and ISA
'NOTE2: The Function calls the AtmosProperty function, so it must be present.

'Initialize
    'Atmospheric properties
    Dim P As Single, OTA As Single, rho As Single
    'Property ratios
    Dim Sigma As Single, delta As Single, theta As Single
    
'presets
    P = AtmosProperty(H, 11)            'Pressure at H in lbf/ft2
    OTA = AtmosProperty(H, 10)          'Standard OTA at H in °R
    OTA = OTA + deltaOTA_C * 1.8        'Include temperature deviation. Note T°R = 18T K
    rho = P / (1716 * OTA)              'Density in slug/ft3
    
    'Pressure Ratio
    delta = P / 2116 * (1 + 0.2 * m ^ 2) ^ 3.5
    'temperature ratio
    theta = OTA / 518.67 * (1 + 0.2 * m ^ 2)
    'Density Ratio
    Sigma = rho / 0.002377427

'Process
    Select Case Mode
    Case 0                  'Piston per Gagg and Ferrar
        Engine_Thrust = F0 * (1.132 * Sigma - 0.132)
        
    Case 1                  'Turboprop per Mattingly, et al.
        If m <= 0.1 Then
            Engine_Thrust = F0 * delta
        Else
            If theta <= TR Then
                Engine_Thrust = F0 * delta * (1 - 0.96 * (m - 0.1) ^ 0.25)
            Else
                Engine_Thrust = F0 * delta * (1 - 0.96 * (m - 0.1) ^ 0.25 - 3 * (theta - TR) / (8.13 * (m - 0.1)))
            End If
        End If
        
    Case 2, 3               'Turbojet per Mattingly, et al.
        If Mode = 2 Then    'Max thrust
            If theta <= TR Then
                Engine_Thrust = F0 * delta * (1 - 0.3 * (theta - 1) - 0.1 * Sqr(m))
            Else
                Engine_Thrust = F0 * delta * (1 - 0.3 * (theta - 1) - 0.1 * Sqr(m) - 1.5 * (theta - TR) / theta)
            End If
        ElseIf Mode = 3 Then    'military thrust(afterburner)
            If theta <= TR Then
                Engine_Thrust = 0.8 * F0 * delta * (1 - 0.16 * Sqr(m))
            Else
                Engine_Thrust = 0.8 * F0 * delta * (1 - 0.16 * Sqr(m) - 24 * (theta - TR) / ((9 + m) * theta))
            End If
        End If
    
    Case 4                  'LBR Turbofan per Mattingly, et al.
        If theta <= TR Then
            Engine_Thrust = F0 * delta
        Else
            Engine_Thrust = F0 * delta * (1 - 3.5 * (theta - TR) / theta)
        End If
    Case 5                  'HBR Turbofan per Mattingly, et al.
        If theta <= TR Then
            Engine_Thrust = F0 * delta * (1 - 0.49 * Sqr(m))
        Else
            Engine_Thrust = F0 * delta * (1 - 0.49 * Sqr(m) - 3 * (theta - TR) / (1.5 + m))
        End If
    End Select
End Function
