Sub sanpz()
'—×‚Ìsheet‚©‚çŽæ“¾

For i = 1 To 20
    Dim ch As Chart
    Set ch = ActiveSheet.ChartObjects.Add(10, 10, 300, 150).Chart
    If i Mod 5 = 1 Then
    ActiveSheet.ChartObjects(i).Left = Cells(1 + 1 * (i), 2).Left
    ActiveSheet.ChartObjects(i).Top = Cells(1 + 1 * (i), 2).Top
    End If
    If i Mod 5 = 2 Then
    ActiveSheet.ChartObjects(i).Left = Cells(1 + 1 * (i), 6).Left
    ActiveSheet.ChartObjects(i).Top = Cells(1 + 1 * (i), 6).Top
    End If
    If i Mod 5 = 3 Then
    ActiveSheet.ChartObjects(i).Left = Cells(1 + 1 * (i), 10).Left
    ActiveSheet.ChartObjects(i).Top = Cells(1 + 1 * (i), 10).Top
    End If
    If i Mod 5 = 4 Then
    ActiveSheet.ChartObjects(i).Left = Cells(1 + 1 * (i), 14).Left
    ActiveSheet.ChartObjects(i).Top = Cells(1 + 1 * (i), 14).Top
    End If
    If i Mod 5 = 0 Then
    ActiveSheet.ChartObjects(i).Left = Cells(1 + 1 * (i), 18).Left
    ActiveSheet.ChartObjects(i).Top = Cells(1 + 1 * (i), 18).Top
    End If
    ActiveSheet.ChartObjects(i).Chart.Axes(xlValue).MaximumScale = 10
    ActiveSheet.ChartObjects(i).Chart.Axes(xlValue).MinimumScale = -1
    'xa = ActiveSheet.Name + "!" + "B" + Str(i + 1)
    
    'ActiveSheet.ChartObjects.Top = Range(xa).Top
    'ActiveSheet.ChartObjects.Left = Range(xa).Left
    
    ch.ChartType = xlXYScatter
    x = ActiveSheet.Name + "!" + "B1:AY1"
    y = ActiveSheet.Name + "!" + "B" + Str(i + 1) + ":AY" + Str(i + 1)
    With ch.SeriesCollection.NewSeries
        .XValues = Range(Cells(1, 2), Cells(1, 51))
        .Values = Range(Cells(1 + i, 2), Cells(1 + i, 51))
    End With
    ch.HasTitle = True
    ch.ChartTitle.Text = Cells(i + 1, 1)
    Set sh2 = Worksheets("Sheet4")
    With ch.SeriesCollection.NewSeries
        '.XValues = sh2.Range(sh2.Cells(1, 2), sh2.Cells(1, 51))
        '.Values = sh2.Range(sh2.Cells(1 + i, 2), sh2.Cells(1 + i, 51))
        .XValues = sh2.Range(sh2.Cells(1 + i, 2), sh2.Cells(1 + i, 51))
        .Values = sh2.Range(sh2.Cells(1, 2), sh2.Cells(1, 51))
        .ChartType = xlXYScatterSmoothNoMarkers
    End With
Next
    
End Sub