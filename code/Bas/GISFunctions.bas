Attribute VB_Name = "GISFunctions"
Option Explicit

Public Function miniRecalc(ByVal ptg As TatukGIS_XDK10.XGIS_Point) As TatukGIS_XDK10.XGIS_Point

        Dim ptgr As TatukGIS_XDK10.XGIS_Point 'result
        Dim deltax, deltay As Double
        Dim P1, p2, p3, p4 As TatukGIS_XDK10.XGIS_Point
        Dim ex As TatukGIS_XDK10.XGIS_Extent

        Set P1 = minishp.GetPoint(0, 0)
        Set p2 = minishp.GetPoint(0, 1)
        Set p3 = minishp.GetPoint(0, 2)
        Set p4 = minishp.GetPoint(0, 3)
        Set ex = frmMain.GIS10.Extent
        deltax = (p2.x - P1.x) / 2 'delta 1/2 of mini rect length
        deltay = (p3.y - p2.y) / 2
        Set ptgr = ptg
        If (ptg.x < (ex.xmin + deltax)) Then ptgr.x = (ex.xmin + deltax)
        If (ptg.y < (ex.ymin + deltay)) Then ptgr.y = (ex.ymin + deltay)
        If (ptg.x > (ex.xmax - deltax)) Then ptgr.x = (ex.xmax - deltax)
        If (ptg.y > (ex.ymax - deltay)) Then ptgr.y = (ex.ymax - deltay)

        Set miniRecalc = ptgr
    End Function

Public Sub Vector2DBConverter(shpType As String, lyrToImport As TatukGIS_XDK10.XGIS_LayerVector, sDBSource As String)
    Dim lL As TatukGIS_XDK10.XGIS_LayerVector
    Dim shape_type As TatukGIS_XDK10.XGIS_ShapeType

    'Console.WriteLine ("TatukGIS Samples - ANY->SQL converter")

'    If args.Length <> 3 Then
'        Console.WriteLine ("Usage: ANY2SQL source_file destination.ttkls shape_type")
'        Console.WriteLine ("Where shape_type:")
'        Console.WriteLine (" A - Arc")
'        Console.WriteLine (" G - polyGon")
'        Console.WriteLine (" P - Point")
'        Console.WriteLine (" M - Multipoint")


    'lyrToImport Set lm = GisUtils.GisCreateLayer("SHAPE NAME", "c:/shp.shp")

    Select Case shpType

        Case "A"
            shape_type = 4 'XGIS_ShapeType.gisShapeTypeArc

        Case "G"
            shape_type = 5 'XGIS_ShapeType.gisShapeTypePolygon

        Case "P"
            shape_type = 2 'XGIS_ShapeType.gisShapeTypePoint

        Case "M"
            shape_type = 3 'XGIS_ShapeType.gisShapeTypeMultiPoint

        Case Else
            shape_type = 4 'XGIS_ShapeType.gisShapeTypeArc
    End Select

    Set lL = GisUtils.GisCreateLayer("", sDBSource)


    lL.ImportLayer lyrToImport, lyrToImport.Extent, shape_type, "", False
End Sub
