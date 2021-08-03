Attribute VB_Name = "Module11"
Sub CombineCsvFiles()

    Dim xFilesToOpen As Variant
    Dim I As Integer
    Dim xWb As Workbook
    Dim xTempWb As Workbook
    Dim xDelimiter As String
    Dim xScreen As Boolean
    On Error GoTo ErrHandler
    xScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    xDelimiter = "|"
    xFilesToOpen = Application.GetOpenFilename("Text Files (*.csv), *.csv", , "Kutools for Excel", , True)
    If TypeName(xFilesToOpen) = "Boolean" Then
        MsgBox "No files were selected", , "Kutools for Excel"
        GoTo ExitHandler
    End If
    I = 1
    Set xTempWb = Workbooks.Open(xFilesToOpen(I))
    xTempWb.Sheets(1).Copy
    Set xWb = Application.ActiveWorkbook
    xTempWb.Close False
    Do While I < UBound(xFilesToOpen)
        I = I + 1
        Set xTempWb = Workbooks.Open(xFilesToOpen(I))
        xTempWb.Sheets(1).Move , xWb.Sheets(xWb.Sheets.Count)
    Loop
ExitHandler:
    Application.ScreenUpdating = xScreen
    Set xWb = Nothing
    Set xTempWb = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description, , "Kutools for Excel"
    Resume ExitHandler
    
    
End Sub



Sub graph()

'start with the first sheet
Application.ScreenUpdating = False

For Each sh In Worksheets
    sh.Activate
    
    name = ActiveSheet.name
    
    
    'By selecting a cell, it won't automatically select data within the graph
    Range("L1").Select
    
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterLinesNoMarkers).Select
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).XValues = "=" & "'" & name & "'" & "!$A$2:$A$400"
    ActiveChart.FullSeriesCollection(1).Values = "=" & "'" & name & "'" & "!$B$2:$B$400"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).XValues = "=" & "'" & name & "'" & "!$C$2:$C$400"
    ActiveChart.FullSeriesCollection(2).Values = "=" & "'" & name & "'" & "!$D$2:$D$400"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).XValues = "=" & "'" & name & "'" & "!$E$2:$E$400"
    ActiveChart.FullSeriesCollection(3).Values = "=" & "'" & name & "'" & "!$F$2:$F$400"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(4).XValues = "=" & "'" & name & "'" & "!$G$2:$G$400"
    ActiveChart.FullSeriesCollection(4).Values = "=" & "'" & name & "'" & "!$H$2:$H$400"
    
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    ActiveChart.FullSeriesCollection(3).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(112, 48, 160)
        .Transparency = 0
    End With
    ActiveChart.FullSeriesCollection(4).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .Transparency = 0
    End With
    
    ActiveChart.ChartTitle.Select
    Selection.Characters.Text = name
    
    ActiveChart.Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
    ActiveChart.Axes(xlValue).MaximumScale = 4
    ActiveChart.Axes(xlValue).MinimumScale = -6
    ActiveChart.Axes(xlCategory, xlPrimary).MinimumScale = 0
    ActiveChart.Axes(xlCategory, xlPrimary).MaximumScale = 350
    
Next

End Sub



    

