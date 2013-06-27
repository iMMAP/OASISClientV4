VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GeoMercator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Mercator Projection Calculations
'--------------------------------

'Constants used in projection calculations
Private Const HalfPi As Double = 1.570796326795
Private Const QuartPi As Double = 0.785398163398

'Local copies of ellipsoid parameters and calculated constants
Private mfSphere As Boolean
Private a As Double 'Semi-major axis
Private ES2 As Double 'Eccentricity squared
Private ES As Double 'Eccentricity

'Local copies of projection parameters and calculated constants
Private FN As Double
Private FE As Double
Private LatOrig As Double
Private LonOrig As Double
Private Ko As Double

'Local reference to Geocalc class module
Private objGeoCalc As GeoCalc

Public Function SetEllipsoid(Radius As Double, Flattening As Double, _
                             fSphere As Boolean) As String
On Error GoTo SetEllipsoid_Err
'Function initialises ellipsoid parameters. This method must be called
'before SetProjection or the Forward or Inverse functions, else run-time errors
'will occur.
'In: Radius - semi major axis of ellipsoid or radius of sphere, in metres
'In: Flattening - ellipsoidal flattening ratio, for ellipsoids, or 0, for sphere
'In: fSphere - boolean flag, true if spherical formulae are to be used for
'    the calculations, false for ellipsoidal calcs
'Returns: String value of error message, if errors encounted during initialisation.
'         Returns zero-length string if successful
    
    Dim b As Double  'Semi-minor axis of ellipsoid, meters
    
    mfSphere = fSphere
    If Radius <= 0 Then
        SetEllipsoid = "Radius must be a positive number!"
    ElseIf Flattening <= 0 And Not mfSphere Then
        SetEllipsoid = "Inverse flattening must be a positive number!"
    Else
        a = Radius
        If Not mfSphere Then
            b = a * (Flattening - 1) / Flattening
            ES2 = (a * a - b * b) / (a * a)
            ES = Sqr(ES2)
            objGeoCalc.SetEccent a, ES
        End If
        SetEllipsoid = ""
    End If
    
SetEllipsoid_Done:
    Exit Function
    
SetEllipsoid_Err:
    SetEllipsoid = "Unexpected error in SetEllipsoid!" & vbCr & "Error Code: " & Err & vbCr & Err.Description
    Resume SetEllipsoid_Done
    
End Function

Public Function SetProjection(FalseNorthing As Double, FalseEasting As Double, _
                              LatitudeOrigin As Double, LongitudeOrigin As Double) As String
On Error GoTo SetProjection_Err
'Initialise the projection parameters.
'In: FalseNorthing - projection system false northing value, in metres
'In: FalseEasting - projection system false easting value, in metres
'In: LatitudeOrigin - latitude of origin/true scale, in radians
'In: LongitudeOrigin - longitude of origin, in radians
'Returns: String value of error message, if errors encounted during initialisation.
'         Returns zero-length string if successful

'Save local copies of projection configuration parameters
    FN = FalseNorthing
    FE = FalseEasting
    LatOrig = LatitudeOrigin
    LonOrig = LongitudeOrigin
    
'Verify that latitude of origin is within the range -90 to 90 degrees
    If Abs(LatOrig) > HalfPi Then
        SetProjection = "Latitude of Origin must be less than 90 degrees!"
        Exit Function
    End If
    
'Set the projection constants. For the Mercator projection, the only constant is
'Ko, which is only applicable to the ellipsoidal transformations
    If Not mfSphere Then
        Ko = a * Cos(Abs(LatOrig)) / Sqr(1 - ES2 * (Sin(LatOrig)) ^ 2)
    End If
    SetProjection = ""
    
SetProjection_Done:
    Exit Function
    
SetProjection_Err:
    SetProjection = "Unexpected Error in function SetProjection" & vbCr & "Error Code: " & Err & vbCr & Err.Description
    Resume SetProjection_Done
End Function

Public Function Forward(Northing As Double, Easting As Double) As String
On Error GoTo Forward_Err
'Convert latitude and longitude (in radians) to northing and easting (in metres)
'In: Northing - latitude of coordinate to be converted, in radians
'Out: Northing - northing of converted coordinate, in metres
'In: Easting - longitude of coordinate to be converted, in radians
'Out: Easting - easting of converted coordinate, in metres
'Returns: String value of error message, if errors encounted during conversion.
'         Returns zero-length string if successful

    Dim Temp As Double
                                  
'If projection has been configured for a spherical earth model, then perform
'the calculations in the mfSphere branch.
    If mfSphere Then
        Easting = FE + a * (Easting - LonOrig)
        Northing = FN + a * Log(Tan(QuartPi + Northing / 2))
    Else
'For an ellipsoidal earth model, perform the calculations below.
        Easting = FE + Ko * (Easting - LonOrig)
        Temp = ES * Sin(Northing)
        Northing = FN + Ko * Log(Abs(Tan(QuartPi + Northing / 2)) * ((1 - Temp) / (1 + Temp)) ^ ((ES / 2)))
    End If
    Forward = ""
    
Forward_Done:
    Exit Function
    
Forward_Err:
    Forward = "Unexpected Error in function Forward" & vbCr & "Error Code: " & Err & vbCr & Err.Description
    Resume Forward_Done

End Function

Public Function Inverse(Northing As Double, Easting As Double) As String
On Error GoTo Inverse_Err
'Convert northing and easting (in metres) to latitude and longitude (in radians)
'In: Northing - northing of coordinate to be converted, in metres
'Out: Northing - latitude of converted coordinate, in radians
'In: Easting - easting of coordinate to be converted, in metres
'Out: Easting - longitude of converted coordinate, in degrees
'Returns: String value of error message, if errors encounted during conversion.
'         Returns zero-length string if successful

    Dim X As Double
    
'Perform calculations common to spherical and ellipsoidal models
    Easting = Easting - FE
    Northing = Northing - FN
    
'If projection has been configured for a spherical earth model, then perform
'the calculations in the mfSphere branch.
    If mfSphere Then
        Northing = HalfPi - 2 * Atn(Exp(-Northing / a))
        Easting = Easting / a + LonOrig
    Else
'For an ellipsoidal earth model, perform the calculations below.
        X = HalfPi - 2 * Atn(Exp((-Northing) / Ko))
        Easting = Easting / Ko + LonOrig
        Northing = objGeoCalc.Phi2z(X)
    End If

    Inverse = ""

Inverse_Done:
    Exit Function
    
Inverse_Err:
    Inverse = "Unexpected Error in function Inverse" & vbCr & "Error Code: " & Err & vbCr & Err.Description
    Resume Inverse_Done

End Function

Private Sub Class_Initialize()
'Initialise the GeoCalc object. This provides access to functions that are used
'in a number of different projection transformation calculations
    Set objGeoCalc = New GeoCalc
End Sub

Private Sub Class_Terminate()
'Free up resources used by the GeoCalc object
    Set objGeoCalc = Nothing
End Sub
