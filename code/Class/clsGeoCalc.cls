VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GeoCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Before using most of these functions, call function SetEccent to initialise the
'semi-major axis and eccentricity for the ellipsoid to be used.

Private Const MaxErr As Double = 0.00000000001 'Maximum error
Private Const MaxIter As Integer = 100 'Maximum number of iterations in iterative calculations

Private Const QuartPi As Double = 0.785398163398
Private Const HalfPi As Double = 1.570796326795
Private Const Pi As Double = 3.14159265359
Private Const TwoPi As Double = 6.28318530718

Private ES As Double 'Eccentricity
Private ES2 As Double 'Eccentricity squared
Private ES4 As Double 'Eccentricity ^ 4
Private ES6 As Double 'Eccentricity ^ 6
Private ES8 As Double 'Eccentricity ^ 8
Private a As Double 'Semi-major axis

Public Function SetEccent(dblA As Double, dblES As Double) As Boolean
'Set constants for given ellipsoid
'dblA = semi-major axis
'dblES = eccentricity

    a = dblA
    ES = dblES
    ES2 = ES * ES
    ES4 = ES2 * ES2
    ES6 = ES2 * ES4
    ES8 = ES4 * ES4
    SetEccent = True
End Function

Public Function AdjustLon(Longitude As Double) As Double
'Convert radians to a number in the range -pi to +pi
    
    If Abs(Longitude) < Pi Then
        AdjustLon = Longitude
    Else
        AdjustLon = Longitude - TwoPi * Int((Longitude + Pi) / TwoPi)
    End If
End Function

Public Function MSFNZ(Phi As Double) As Double
' Function to compute the constant small m which is the radius of
' a parallel of latitude, phi, divided by the semimajor axis.

    Dim Temp As Double
    
    Temp = ES * Sin(Phi)
    MSFNZ = Cos(Phi) / (Sqr(1 - Temp * Temp))
End Function

Public Function RadCurv(Latitude As Double) As Double
'radius of curvature on the prime vertical (i.e. perpendicular to the meridian)
'at given latitude

    Dim Temp As Double
    
    Temp = Sin(Latitude) * ES
    RadCurv = a / Sqr(1 - Temp * Temp)
End Function

Public Function RCML(Latitude As Double) As Double
'radius of curvature of the meridian at Latitude
    
    Dim SinPhi As Double
    
    SinPhi = Sin(Latitude)
    RCML = a * (1 - ES2) / (1 - ES2 * SinPhi * SinPhi) ^ 1.5

End Function

Public Function QSFNZ(dblLat As Double) As Double
' Function to compute constant small q which is the radius of a
' parallel of latitude, phi, divided by the semi-major axis.
' Equation 3-12 in Snyder

    Dim Temp As Double
    Dim Temp1 As Double
    
    Temp = Sin(dblLat)
    If (ES > MaxErr) Then
        Temp1 = ES * Temp
        QSFNZ = (1 - ES2) * (Temp / (1 - Temp1 * Temp1) - 1 / (2 * ES) _
                    * Log(((1 - Temp1) / (1 + Temp1))))
    Else
        QSFNZ = 2 * Temp
    End If
End Function

Public Function MLFNZ(Latitude As Double) As Double
'Snyder equation 3-21
'Compute the value of M which is the distance along a meridian
'from the Equator to Latitude

    Const C02 As Double = 0.25
    Const C04 As Double = 0.046875
    Const C06 As Double = 0.01953125
    Const C08 As Double = 0.01068115234375
    Const C22 As Double = 0.75
    Const C44 As Double = 0.46875
    Const C46 As Double = 1.30208333333333E-02
    Const C48 As Double = 7.12076822916667E-03
    Const C66 As Double = 0.364583333333333
    Const C68 As Double = 5.69661458333333E-03
    Const C88 As Double = 0.3076171875
    
    Dim T As Double 'temp variable
    Dim E0 As Double
    Dim E1 As Double
    Dim E2 As Double
    Dim E3 As Double
    Dim E4 As Double
    
    Dim CPhi As Double
    Dim SPhi As Double
    
    SPhi = Sin(Latitude)
    CPhi = Cos(Latitude) * SPhi
    SPhi = SPhi * SPhi

    T = C04 + ES2 * (C06 + ES2 * C08)
    E0 = 1 - ES2 * (C02 + ES2 * T)
    E1 = ES2 * (C22 - ES2 * T)
    E2 = ES4 * (C44 - ES2 * (C46 + ES2 * C48))
    E3 = ES6 * (C66 - ES * C68)
    E4 = ES8 * C88
    T = E2 + SPhi * (E3 + SPhi * E4)
    MLFNZ = E0 * Latitude - CPhi * (E1 + SPhi * T)
End Function

Public Function UI(M As Double) As Double
'Snyder equation 7-19

    Const C1 As Double = 0.046875
    Const C2 As Double = 0.01953125
    
    UI = M / (a * (1 - ES2 / 4 - C1 * ES4 - C2 * ES6))

End Function

Public Function E2fn() As Double
'Snyder equation 3-24

    Dim Temp As Double
 
    Temp = Sqr(1 - ES2)
    E2fn = (1 - Temp) / (1 + Temp)
    
End Function

Public Function CalcM(Northing As Double) As Double
'Snyder equation 14-15
    Dim SinPhi As Double
    
    SinPhi = Sin(Northing)
    CalcM = Cos(Northing) / Sqr(1 - ES2 * SinPhi * SinPhi)
End Function

Public Function CalcG(Northing As Double) As Double
'Snyder equation 3-1a
    Dim Temp As Double
    Dim S1 As Double, S2 As Double
    
    Temp = Sin(Northing)
    If Temp = 1 Then
        CalcG = 1E+99
        Exit Function
    End If
    S1 = (1 + Temp) / (1 - Temp)
    Temp = ES * Temp
    S2 = (1 - Temp) / (1 + Temp)
    CalcG = 2 * Atn(Sqr(S1 * S2 ^ ES)) - HalfPi

End Function

Public Function Phi2z(TS As Double) As Double
'Function to compute the inverse latitude angle, phi2, for the conformal latitude
'for a conformal sphere, in radians
'Snyder equation 3-5

    Const C1 As Double = 0.208333333333333
    Const C2 As Double = 0.036111111111111
    Const C3 As Double = 0.145833333333333
    Const C4 As Double = 0.120833333333333
    Const C5 As Double = 0.070399305555555
    Const C6 As Double = 0.058333333333333
    Const C7 As Double = 7.23214285714286E-02
    Const C8 As Double = 0.026531498015873

    Phi2z = TS + (ES2 / 2 + C1 * ES4 + ES6 / 12 + C2 * ES8) * Sin(2 * TS) _
           + (C3 * ES4 + C4 * ES6 + C5 * ES8) * Sin(4 * TS) _
           + (C6 * ES6 + C7 * ES8) * Sin(6 * TS) _
           + C8 * ES8 * Sin(8 * TS)
    
End Function

Public Function Phi3z(U As Double, E1 As Double) As Double
'Snyder equation 3-26

    Const C1 As Double = 2.25
    Const C2 As Double = 1.3125
    Const C3 As Double = 1.71875
    Const C4 As Double = 1.57291666666666
    Const C5 As Double = 2.142578125
    
    Dim E1_3 As Double
    Dim E1_4 As Double
    
    E1_3 = E1 * E1 * E1
    E1_4 = E1_3 * E1

    Phi3z = U + (1.5 * E1 - C1 * E1_3) * Sin(2 * U) + (C2 * E1 * E1 - C3 * E1_4) * Sin(4 * U) _
            + C4 * E1_3 * Sin(6 * U) + C5 * E1_4 * Sin(8 * U)
End Function

Public Function Phi4Z(A1 As Double, B1 As Double, C1 As Double, Phi As Double) As Double
On Error GoTo Phi4Z_Err

    Dim i As Integer
    Dim SinPhi As Double
    Dim TanPhi As Double
    Dim Sin2PH As Double
    Dim M1 As Double
    Dim M1P As Double
    Dim DPhi As Double
    
    Phi = A1
    For i = 1 To MaxIter
        SinPhi = Sin(Phi)
        TanPhi = Tan(Phi)
        C1 = TanPhi * Sqr(1 - ES2 * SinPhi * SinPhi)
        Sin2PH = Sin(2 * Phi)
        
        M1 = MLFNZ(Phi)
        M1P = 1 - ES2 / 4 - 3 * ES2 * ES2 / 64 - 5 * ES2 ^ 3 / 256 _
              - 2 * (3 * ES2 / 8 + 3 * ES2 * ES2 / 32 + 45 * ES2 ^ 3 / 1024) * Cos(2 * Phi) _
              + 4 * (15 * ES2 * ES2 / 256 + 45 * ES2 ^ 3 / 1024) * Cos(4 * Phi) _
              - 6 * (35 * ES2 ^ 3 / 3072) * Cos(6 * Phi)
        DPhi = (A1 * (C1 * M1 + 1) - M1 - 0.5 * (M1 * M1 + B1) * C1) / (ES2 * Sin2PH * (M1 * M1 + B1 - 2 * A1 * M1) / 4 * C1 + (A1 - M1) * (C1 * M1P - 0.5 * Sin2PH) - M1P)
        Phi = Phi - DPhi
        If Abs(DPhi) <= MaxErr Then
            Exit Function
        End If
    Next i

Phi4Z_Done:
    Exit Function
    
Phi4Z_Err:
    Phi4Z = 0
    Resume Phi4Z_Done
End Function

Public Function Phi5z(Beta As Double) As Double
'Snyder equation 3-18
'Calculate the latitude, given the authalic latitude Beta, for a sphere having
'the same surface area as the ellipsoid

    Const C1 As Double = 0.172222222222222
    Const C2 As Double = 0.102579365079365
    Const C3 As Double = 6.38888888888889E-02
    Const C4 As Double = 6.64021164021164E-02
    Const C5 As Double = 1.67768959435626E-02

    Phi5z = Beta + (ES2 / 3 + C1 * ES4 + C2 * ES6) * Sin(2 * Beta) _
            + (C3 * ES4 + C4 * ES6) * Sin(4 * Beta) _
            + C5 * ES6 * Sin(6 * Beta)
End Function

Public Function TSFNZ(Phi As Double) As Double
'Snyder 15-9 or 9-13
  Dim SinPhi As Double
  Dim Temp As Double
  
  SinPhi = Sin(Phi)
  Temp = ES * SinPhi
  TSFNZ = Tan(QuartPi - Phi / 2) / ((1 - Temp) / (1 + Temp)) ^ (ES / 2)
End Function



