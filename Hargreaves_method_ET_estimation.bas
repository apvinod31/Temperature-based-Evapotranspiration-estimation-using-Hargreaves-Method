Attribute VB_Name = "Module1"
Sub Hargreaves_ET_ArrayFast_Tmean()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataArr As Variant
    Dim outArr() As Variant
    
    Dim i As Long
    
    Dim Tmax As Double, Tmin As Double, Tmean As Double
    Dim Lat_deg As Double, Lat_rad As Double
    Dim J As Double, Gsc As Double
    
    Dim dr As Double, delta As Double
    Dim ws_angle As Double
    Dim Ra_MJ As Double, Ra_mm As Double
    Dim ETo As Double
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    'Read input data into array
    dataArr = ws.Range("A2:F" & lastRow).Value
    
    'Prepare output array (G to N = 8 columns)
    ReDim outArr(1 To UBound(dataArr), 1 To 8)
    
    For i = 1 To UBound(dataArr)
        
        Tmax = dataArr(i, 2)
        Tmin = dataArr(i, 3)
        Lat_deg = dataArr(i, 4)
        Gsc = dataArr(i, 5)
        J = dataArr(i, 6)
        
        'Mean temperature
        Tmean = (Tmax + Tmin) / 2
        
        'Latitude radians
        Lat_rad = Lat_deg * WorksheetFunction.Pi() / 180
        
        'Inverse relative Earth-Sun distance
        dr = 1 + 0.033 * Cos(2 * WorksheetFunction.Pi() * J / 365)
        
        'Solar declination
        delta = 0.409 * Sin(2 * WorksheetFunction.Pi() * J / 365 - 1.39)
        
        'Sunset hour angle
        ws_angle = WorksheetFunction.Acos(-Tan(Lat_rad) * Tan(delta))
        
        'Extraterrestrial radiation MJ/m2/day
        Ra_MJ = (24 * 60 / WorksheetFunction.Pi()) * Gsc * dr * _
                (ws_angle * Sin(Lat_rad) * Sin(delta) + _
                Cos(Lat_rad) * Cos(delta) * Sin(ws_angle))
        
        'Convert to mm/day
        Ra_mm = 0.408 * Ra_MJ
        
        'Hargreaves ET
        If Tmax > Tmin Then
            ETo = 0.002 * Ra_mm * Sqr(Tmax - Tmin) * (Tmean + 17.8)
        Else
            ETo = ""
        End If
        
        'Store outputs
        outArr(i, 1) = Tmean
        outArr(i, 2) = dr
        outArr(i, 3) = delta
        outArr(i, 4) = Lat_rad
        outArr(i, 5) = ws_angle
        outArr(i, 6) = Ra_MJ
        outArr(i, 7) = Ra_mm
        outArr(i, 8) = ETo
        
    Next i
    
    'Write output to sheet in one shot
    ws.Range("G2").Resize(UBound(outArr), 8).Value = outArr
    
    MsgBox "Ultra-Fast Hargreaves ET (Auto Tmean) Completed", vbInformation

End Sub

