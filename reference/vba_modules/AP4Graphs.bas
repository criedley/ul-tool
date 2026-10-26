Attribute VB_Name = "AP4Graphs"
Sub createGraphs()
Dim reportmaker As Workbook
Dim rawDataWb As Workbook
Dim TCsheet As Worksheet
Dim rawDataSheet As Worksheet

Set reportmaker = ActiveWorkbook
Set rawDataWb = GetDataV2("Select report file")
Set rawDataSheet = rawDataWb.Sheets("RawData")
Set TCsheet = rawDataWb.Sheets("AP4_Chamber")

'creating graphs

'find row that the channels are on
Dim channelrow As Integer
channelrow = 0
Dim x As Range


rawDataSheet.Activate
Set x = rawDataSheet.Range("A1:C50").Find("Date/Time", LookIn:=xlValues)
If x Is Nothing Then
    MsgBox ("Could not find correct row to look for the TC channels on raw data sheet")
Else
    channelrow = x.Row
End If

Dim datastart As Integer
'need to find DataStart - different per chamber, need to find in col 1 where there is a datatime
'isdate()
For i = 1 To 100
    If IsDate(rawDataSheet.Cells(i, 1)) Then
        datastart = channelrow + i
            Exit For
    End If
Next i

'find last row
lastrow = rawDataSheet.Range("A65000").End(xlUp).Row

'array of channel column titles
'Dim seriesTitle(100) As Variant
'seriesTitle(1) = "PID_PV_ID_DB_TEMP_STS"
'seriesTitle(2) = "PID_PV_ID_WB_TEMP_STS"
'seriesTitle(3) = "PID_PV_OD_DB_TEMP_STS"
'seriesTitle(4) = "PID_PV_OD_WB_TEMP_STS"
'seriesTitle(5) = "UUT_ID_UNIT_AMPS"
'seriesTitle(6) = "UUT_ID_UNIT_VOLTS"
'seriesTitle(7) = "UUT_ID_UNIT_WATTS"
'seriesTitle(8) = "PID_SP_ID_TNL_AIRFLOW_PRESS_STS"
'seriesTitle(9) = "UUT_REFRIG_PRESS_1001"
'seriesTitle(10) = "UUT_REFRIG_PRESS_1002"
'seriesTitle(11) = "UUT_REFRIG_PRESS_1003"
'seriesTitle(12) = "UUT_REFRIG_PRESS_1004"

'Dim graphTitle(100) As Variant
'graphTitle(1) = "ID Temperatures"

'graph 1 ID temps
'graph 2 OD temps
'graph 3 V and amps
'graph 4 watts
'graph 5 ID Pressure
'graph 6 OD pressure
'graph 7 ID TCs
'graph 8 OD TCS
'Dim yaxistemp(100) As Variant

Dim channelCol As Integer

Dim xaxis As Range ' each graph will have x axis of time
Set xaxis = rawDataSheet.Range(Cells(datastart, 2), Cells(lastrow, 2))

Dim NewChart As Chart
rawDataSheet.Cells(1, 1).Activate
'graph 1 ID temps
If FindIfThereisAlreadyNamedSheet("ID Room Temps", rawDataWb) Then
    Set NewChart = Charts.Add
    NewChart.ChartArea.Clear
    NewChart.ChartType = xlXYScatterSmoothNoMarkers
   
    Set chartSeries = NewChart.SeriesCollection.NewSeries
    channelCol = FindColFromName(rawDataSheet, "PID_PV_ID_DB_TEMP_STS", channelrow)
    'Set chartSeries = NewChart.SeriesCollection.NewSeries
            With chartSeries
                .XValues = xaxis
                .Values = rawDataSheet.Range(Cells(datastart, channelCol), Cells(lastrow, channelCol))
                .Name = "ID Db Temps"
                .AxisGroup = xlPrimary
            End With
    Set chartSeries = NewChart.SeriesCollection.NewSeries
        channelCol = FindColFromName(rawDataSheet, "PID_PV_ID_WB_TEMP_STS", channelrow)
            With chartSeries
                .XValues = xaxis
                .Values = rawDataSheet.Range(Cells(datastart, channelCol), Cells(lastrow, channelCol))
                .Name = "ID Wb Temps"
                .AxisGroup = xlSecondary
            End With

        With NewChart
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlSecondary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "Drybulb Temp (F)"
            .Axes(xlValue, xlSecondary).AxisTitle.Text = "Wetbulb Temp (F)"
            .HasTitle = True
            .ChartTitle.Text = "ID Room Temps"
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
        End With
    NewChart.Name = "ID Room Temps"
End If
    
'graph 2 OD temps
If FindIfThereisAlreadyNamedSheet("OD Room Temps", rawDataWb) Then
    Set NewChart = Charts.Add
    NewChart.ChartArea.Clear
    NewChart.ChartType = xlXYScatterSmoothNoMarkers

    Set chartSeries = NewChart.SeriesCollection.NewSeries
    channelCol = FindColFromName(rawDataSheet, "PID_PV_OD_DB_TEMP_STS", channelrow)
    'Set chartSeries = NewChart.SeriesCollection.NewSeries
            With chartSeries
                .XValues = xaxis
                .Values = rawDataSheet.Range(Cells(datastart, channelCol), Cells(lastrow, channelCol))
                .Name = "OD Db Temps"
                .AxisGroup = xlPrimary
            End With
    Set chartSeries = NewChart.SeriesCollection.NewSeries
        channelCol = FindColFromName(rawDataSheet, "PID_PV_OD_WB_TEMP_STS", channelrow)
            With chartSeries
                .XValues = xaxis
                .Values = rawDataSheet.Range(Cells(datastart, channelCol), Cells(lastrow, channelCol))
                .Name = "OD Wb Temps"
                .AxisGroup = xlSecondary
            End With

        With NewChart
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlSecondary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "Drybulb Temp (F)"
            .Axes(xlValue, xlSecondary).AxisTitle.Text = "Wetbulb Temp (F)"
            .HasTitle = True
            .ChartTitle.Text = "ID Room Temps"
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
        End With
   NewChart.Name = "OD Room Temps"
End If

'graph 3 V and amps
If FindIfThereisAlreadyNamedSheet("Amps and Volts", rawDataWb) Then
    Set NewChart = Charts.Add
    NewChart.ChartArea.Clear
   NewChart.ChartType = xlXYScatterSmoothNoMarkers
     
    'graph 1 ID temps
    Set chartSeries = NewChart.SeriesCollection.NewSeries
    channelCol = FindColFromName(rawDataSheet, "UUT_ID_UNIT_AMPS", channelrow)
    'Set chartSeries = NewChart.SeriesCollection.NewSeries
            With chartSeries
                .XValues = xaxis
                .Values = rawDataSheet.Range(Cells(datastart, channelCol), Cells(lastrow, channelCol))
                .Name = "Amps (A)"
                .AxisGroup = xlPrimary
            End With
    Set chartSeries = NewChart.SeriesCollection.NewSeries
        channelCol = FindColFromName(rawDataSheet, "UUT_ID_UNIT_VOLTS", channelrow)
            With chartSeries
                .XValues = xaxis
                .Values = rawDataSheet.Range(Cells(datastart, channelCol), Cells(lastrow, channelCol))
                .Name = "Volts (V)"
                .AxisGroup = xlSecondary
            End With

        With NewChart
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlSecondary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "Amps (A)"
            .Axes(xlValue, xlSecondary).AxisTitle.Text = "Volts (V)"
            .HasTitle = True
            .ChartTitle.Text = "Amps and Volts"
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
        End With
   NewChart.Name = "Amps and Volts"
End If

'graph 4 watts
If FindIfThereisAlreadyNamedSheet("Wattage (W)", rawDataWb) Then
Set NewChart = Charts.Add
    NewChart.ChartArea.Clear
   NewChart.ChartType = xlXYScatterSmoothNoMarkers
     
    'graph 1 ID temps
    Set chartSeries = NewChart.SeriesCollection.NewSeries
    channelCol = FindColFromName(rawDataSheet, "UUT_ID_UNIT_WATTS", channelrow)
    'Set chartSeries = NewChart.SeriesCollection.NewSeries
            With chartSeries
                .XValues = xaxis
                .Values = rawDataSheet.Range(Cells(datastart, channelCol), Cells(lastrow, channelCol))
                .Name = "Wattage (W)"
                .AxisGroup = xlPrimary
            End With

        With NewChart
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "Wattage (W)"
            .HasTitle = True
            .ChartTitle.Text = "Watts"
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
        End With
   NewChart.Name = "Wattage (W)"
End If

'graph 5 ID Pressure ESP
If FindIfThereisAlreadyNamedSheet("ID ESP", rawDataWb) Then
    Set NewChart = Charts.Add
    NewChart.ChartArea.Clear
    NewChart.ChartType = xlXYScatterSmoothNoMarkers
   
    Set chartSeries = NewChart.SeriesCollection.NewSeries
    channelCol = FindColFromName(rawDataSheet, "PID_PV_ID_TNL_AIRFLOW_PRESS_STS", channelrow)
    'Set chartSeries = NewChart.SeriesCollection.NewSeries
            With chartSeries
                .XValues = xaxis
                .Values = rawDataSheet.Range(Cells(datastart, channelCol), Cells(lastrow, channelCol))
                .Name = "ID ESP"
                .AxisGroup = xlPrimary
            End With
            
        With NewChart
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "ID ESP"
            .HasTitle = True
            .ChartTitle.Text = "ID ESP"
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
        End With
    NewChart.Name = "ID ESP"
End If
    
'graph 6 ID Pressure Taps
If FindIfThereisAlreadyNamedSheet("ID Pressure", rawDataWb) Then
    Set NewChart = Charts.Add
    NewChart.ChartArea.Clear
    NewChart.ChartType = xlXYScatterSmoothNoMarkers
   
    Set chartSeries = NewChart.SeriesCollection.NewSeries
    channelCol = FindColFromName(rawDataSheet, "UUT_REFRIG_PRESS_1001", channelrow)
    'Set chartSeries = NewChart.SeriesCollection.NewSeries
            With chartSeries
                .XValues = xaxis
                .Values = rawDataSheet.Range(Cells(datastart, channelCol), Cells(lastrow, channelCol))
                .Name = "UUT_REFRIG_PRESS_1001"
                .AxisGroup = xlPrimary
            End With
    Set chartSeries = NewChart.SeriesCollection.NewSeries
    channelCol = FindColFromName(rawDataSheet, "UUT_REFRIG_PRESS_1002", channelrow)
    'Set chartSeries = NewChart.SeriesCollection.NewSeries
            With chartSeries
                .XValues = xaxis
                .Values = rawDataSheet.Range(Cells(datastart, channelCol), Cells(lastrow, channelCol))
                .Name = "UUT_REFRIG_PRESS_1002"
                .AxisGroup = xlPrimary
            End With
    Set chartSeries = NewChart.SeriesCollection.NewSeries
    channelCol = FindColFromName(rawDataSheet, "UUT_REFRIG_PRESS_1003", channelrow)
    'Set chartSeries = NewChart.SeriesCollection.NewSeries
            With chartSeries
                .XValues = xaxis
                .Values = rawDataSheet.Range(Cells(datastart, channelCol), Cells(lastrow, channelCol))
                .Name = "UUT_REFRIG_PRESS_1003"
                .AxisGroup = xlPrimary
            End With
   ' Set chartSeries = NewChart.SeriesCollection.NewSeries
   ' channelCol = FindColFromName(rawDataSheet, "UUT_REFRIG_PRESS_1004", channelrow)
    'Set chartSeries = NewChart.SeriesCollection.NewSeries
   '         With chartSeries
   '             .XValues = xaxis
   '             .Values = rawDataSheet.Range(Cells(datastart, channelCol), Cells(lastrow, channelCol))
   '             .Name = "UUT_REFRIG_PRESS_1004"
   '             .AxisGroup = xlPrimary
   '         End With

        With NewChart
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "ID Pressure"
            .HasTitle = True
            .ChartTitle.Text = "ID Pressure"
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
        End With
    NewChart.Name = "ID Pressure"
End If

'graph 7 Capacity
If FindIfThereisAlreadyNamedSheet("Capacity", rawDataWb) Then
    Set NewChart = Charts.Add
    NewChart.ChartArea.Clear
    NewChart.ChartType = xlXYScatterSmoothNoMarkers
    Set chartSeries = NewChart.SeriesCollection.NewSeries
    channelCol = FindColFromName(rawDataSheet, "ID_TOTAL_CAPACITY_REPORT", channelrow)
    'Set chartSeries = NewChart.SeriesCollection.NewSeries
            With chartSeries
                .XValues = xaxis
                .Values = rawDataSheet.Range(Cells(datastart, channelCol), Cells(lastrow, channelCol))
                .Name = "Total Capacity Indoor Side"
                .AxisGroup = xlPrimary
            End With
    Set chartSeries = NewChart.SeriesCollection.NewSeries
        channelCol = FindColFromName(rawDataSheet, "ID_SENSIBLE_CAPACITY_REPORT", channelrow)
            With chartSeries
                .XValues = xaxis
                .Values = rawDataSheet.Range(Cells(datastart, channelCol), Cells(lastrow, channelCol))
                .Name = "Sensible Capacity Indoor Side"
                .AxisGroup = xlPrimary
            End With
    Set chartSeries = NewChart.SeriesCollection.NewSeries
        channelCol = FindColFromName(rawDataSheet, "ID_LATENT_CAPACITY_REPORT", channelrow)
            With chartSeries
                .XValues = xaxis
                .Values = rawDataSheet.Range(Cells(datastart, channelCol), Cells(lastrow, channelCol))
                .Name = "Latent Capacity Indoor Side"
                .AxisGroup = xlPrimary
            End With

        With NewChart
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "BTUs"
            .HasTitle = True
            .ChartTitle.Text = "Indoor Capacity"
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
        End With
   NewChart.Name = "Capacity"
End If

End Sub

Function FindColFromName(wkst As Worksheet, columnName As String, channelrow As Integer)
wkst.Activate
Set x = wkst.Range(Cells(channelrow - 1, 1), Cells(channelrow + 1, 700)).Find(columnName, LookIn:=xlValues)
FindColFromName = x.Column
Exit Function

End Function

Function FindIfThereisAlreadyNamedSheet(sheetname As String, wb As Workbook)
'return boolean if the chart should be made, returns True if there is no sheet and you want to make the graph

For i = 1 To wb.Sheets.Count
    If wb.Sheets(i).Name = sheetname Then
        FindIfThereisAlreadyNamedSheet = False
        Exit Function
    End If
    
Next i
FindIfThereisAlreadyNamedSheet = True
End Function
