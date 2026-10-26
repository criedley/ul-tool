Attribute VB_Name = "ReportMaker"
Public reportmaker As Workbook
Public ExitSub As Boolean
'Option Explicit

Sub ReportConsolidator()

'Notes: there are several assumptions made based on the file directory
'structure and filenames of the files chosen for inclusion in the final workbook.
'If these assumptions are no longer valid, some modification of the code is required.

Dim myFile As FileDialog
Dim myNewWorkbook As Workbook
Dim reportWorkbook As Workbook
Dim tempNewSheet As Worksheet
Dim myFileName(100) As String
Dim measuredDataFileName As String
Dim calculatedDataFileName As String
Dim timeTemp As String
Dim timeSecs As Long
Dim timeMins As Long
Dim timeHrs As Long
Dim timeDay As Long
Dim timeMonth As Long
Dim timeYear As Long
Dim timeTol As Long
Dim tempString As String
Dim tempInt As Long
Dim J As Long
Dim dataFileBase As String
Dim myDate As Double
Dim myTime As Variant
Dim myInitialPath As String
Const reportSheetName As String = "Report"
Const measuredDataSheetName As String = "Measured Data"
Const calculatedDataSheetName As String = "Calculated Data"
Dim k As Integer
Dim m As Integer
Dim TCS(40) As String
Dim TCYaxis(40) As Range
Dim TempSeries(100) As Series 'Creat series variable for chart
Set reportmaker = ActiveWorkbook
Dim xaxis As Range
Dim noEER As Boolean

'adjusting the TC summary page on the consolidate reports workbook first before rest of code 10/28/24

AdjustTCSummaryPage

'code for making this a multi report system 5/31/23


    '[1. Report]
                Application.ScreenUpdating = False
                noEER = False 'starting with the assumption that the test is good to use- will trip and skip reports if bad
                'pick a report file to copy to new workbook
                    Set myFile = Application.FileDialog(msoFileDialogFilePicker)
                    
                    'Assume that the file takes the form "*.xlsx", but also include other options as well
                    myInitialPath = Range("Q1").Text
                    
                    With myFile
                        .Title = "Select Report File"
                        'If answer = vbYes Then
                        '    .AllowMultiSelect = False
                        'Else
                            .AllowMultiSelect = True
                        'End If
                        .Filters.Clear
                        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.csv"
                        .InitialFileName = myInitialPath
                        'if user hits cancel, abort
                        If .Show <> -1 Then
                            Exit Sub
                        End If
                        k = myFile.SelectedItems.Count
                        For i = 1 To myFile.SelectedItems.Count
                            myFileName(i) = myFile.SelectedItems(i)
                        Next i
                        
                    End With
                    
                    Dim saveFile As FileDialog
                    'pick a report file to copy to new workbook
                    Set saveFile = Application.FileDialog(msoFileDialogFolderPicker)
                    
                    'Assume that the file takes the form "*.xlsx", but also include other options as well
                    myInitialPath = Range("Q1").Text
                    
                    With myFile
                        .Title = "Select where to save report file/s"
                        'If answer = vbYes Then
                        '    .AllowMultiSelect = False
                        'Else
                            .AllowMultiSelect = True
                        'End If
                        .Filters.Clear
                        
                        .InitialFileName = myInitialPath
                        'if user hits cancel, abort
                        If .Show <> -1 Then
                            Exit Sub
                        End If
                        
                        
                    End With
   
                    Dim mySaveLocation As String
                    
                    mySaveLocation = saveFile.SelectedItems(1)
                    
                    'myFileName = myFile.SelectedItems.Item(1)
    For i = 1 To k
                    noEER = False 'needs to reset for each worksheet
                    Set reportWorkbook = Application.Workbooks.Open(myFileName(i))
                    Dim totalTime As Integer
                    totalTime = reportWorkbook.Sheets(1).Range("rpt_WarmupTime").Value + reportWorkbook.Sheets(1).Range("rpt_StabTime").Value + reportWorkbook.Sheets(1).Range("rpt_MeasTime").Value
                    If reportWorkbook.Sheets(1).Range("rpt_EER") = "" Or reportWorkbook.Sheets(1).Range("rpt_EER").Value = 0 Or totalTime <= 15 Then
                        noEER = True
                        'will ask if you want to skip after pulling up the measurement data sheet too
                        'reportWorkbook.Close
                        'GoTo SkipCode
                    End If
                    
                    Set myNewWorkbook = Application.Workbooks.Add
                    
                    Dim reportWorkbookname As String
                    reportWorkbookname = reportWorkbook.Name
                    
                    reportWorkbook.Sheets(1).Copy before:=myNewWorkbook.Sheets(1)
                    
                    ActiveSheet.Name = reportSheetName
                    reportWorkbook.Close
                    
                ' adding TR number in if inputed
                If reportmaker.Worksheets("Consolidate Reports").Range("Q2").Value <> "" And myNewWorkbook.Worksheets("Report").Range("rpt_TestReqNum").Value = "" Then
                    myNewWorkbook.Worksheets("Report").Range("rpt_TestReqNum").Value = reportmaker.Worksheets("Consolidate Reports").Range("Q2").Value
                End If
            '>>>>ignore this stuff
            
            
            'Get a base path for raw data files. Assume that the data folder is at the same level as the "TestReports" folder
            'dataFileBase = Left(myFileName, InStr(1, myFileName, "\TestReports\"))
            'dataFileBase = dataFileBase & "Data\"
            
            
            'Get a string with "YYYY_MM_DD_HH_mm_SS" , per the assumed filename format above.
            'take the right-most 24 characters (includes ".xlsx" of filename) and select just
            'the desired string, which is 19 characters long, including the underscores (_)
        
            'timeTemp = Mid(Right(myFileName, 24), 1, 19)
            
            'up to this point, we have been dealing with strings... need to cast these as numerals. In this case, integers.
            'Split the string and assign the various parts (#) to unique variables
            
            'timeYear = CInt(Split(timeTemp, "_")(0))
            'timeMonth = CInt(Split(timeTemp, "_")(1))
            'timeDay = CInt(Split(timeTemp, "_")(2))
            'timeHrs = CInt(Split(timeTemp, "_")(3))
            'timeMins = CInt(Split(timeTemp, "_")(4))
            'timeSecs = CInt(Split(timeTemp, "_")(5))
            
            ' how many seconds before or after the main report's timestamp to search for the raw data file's timestamp in the filename?
            'timeTol = 10
            
            'myDate = CDbl(DateSerial(timeYear, timeMonth, timeDay) + TimeSerial(timeHrs, timeMins, timeSecs))
            
            'tempString = ""
            
            'For j = 0 To timeTol
                
                'first add j seconds to the main report's timestamp
                'myDate = DateAdd("s", j, myDate)
                
                'tempString = dataFileBase & _
                             Format(Year(myDate), "0000") & "_" & _
                             Format(Month(myDate), "00") & "_" & _
                             Format(Day(myDate), "00") & "_" & _
                             Format(Hour(myDate), "00") & "_" & _
                             Format(Minute(myDate), "00") & "_" & _
                             Format(Second(myDate), "00") & "_MEASEMU.csv"
            
        '        If Dir(tempString) <> "" And tempString <> "" Then
        '            measuredDataFileName = tempString
        '            Exit For
        '        End If
        '
        '        'then subtract j seconds from the main report's timestamp
        '        myDate = DateAdd("s", -2 * j, myDate)
        '
        '        tempString = dataFileBase & _
        '                     Format(Year(myDate), "0000") & "_" & _
        '                     Format(Month(myDate), "00") & "_" & _
        '                     Format(Day(myDate), "00") & "_" & _
        '                     Format(Hour(myDate), "00") & "_" & _
        '                     Format(Minute(myDate), "00") & "_" & _
        '                     Format(Second(myDate), "00") & "_MEASEMU.csv"
        '
        '        If Dir(tempString) <> "" And tempString <> "" Then
        '            measuredDataFileName = tempString
        '            Exit For
        '        End If
        '
        '    Next j

        
        'dataFileBase = ("E:\2-15-24" & "\" & Right(Range("rpt_DataFilename").Value, Len(Range("rpt_DataFilename").Value) - InStrRev(Range("rpt_DataFilename").Value, "\")))
        'hard coding in so can run on my computer
            
             
            dataFileBase = Range("rpt_DataFilename")
            If Dir(dataFileBase) <> "" And dataFileBase <> "" Then
                measuredDataFileName = dataFileBase
            End If
        
            If measuredDataFileName = "" Then
                    'Couldn't find an automatic match ... ask user to identify the corresponding measurement file
                    
                    Set myFile = Application.FileDialog(msoFileDialogFilePicker)
        
                    With myFile
                        .Title = "Select Measurement File that corresponds to report ending in " & Right(myFileName(i), 35)
                        .AllowMultiSelect = False
                        .Filters.Clear
                        .Filters.Add "CSV Files", "*.csv"
                        
                        'if user hits cancel, abort
                        If .Show <> -1 Then
                            Exit Sub
                        End If
                    
                    End With
                    
                    measuredDataFileName = myFile.SelectedItems.Item(1)
            End If
        
        
        
        
        
        '[2. Pics]
            myNewWorkbook.Worksheets.Add after:=myNewWorkbook.Worksheets("Report")
            ActiveSheet.Name = "Pics"
            
        
        
        '[6. Measured Data]
        
            Set reportWorkbook = Application.Workbooks.Open(measuredDataFileName)
            
            reportWorkbook.Sheets(1).Copy after:=myNewWorkbook.Sheets("Pics")
            
            'Code below depends on this sheet's name being exactly as specified: "Measured Data"
            ActiveSheet.Name = measuredDataSheetName
            
            reportWorkbook.Close
            
            'need to check if there is measurement data
            Worksheets(measuredDataSheetName).Activate
            If noEER Then  'myNewWorkbook.Sheets(measuredDataSheetName).Range("C30").Value = "" And
                Worksheets(reportSheetName).Activate
                ActiveWindow.ScrollRow = 1
                Application.ScreenUpdating = True
                Dim answer As Integer
                answer = MsgBox("Should this test be skipped? This event is triggered when the EER is 0 or the test was scheduled for 15min or less", vbQuestion + vbYesNo + vbDefaultButton2, "Skip Test?")
                If answer = vbYes Then
                    myNewWorkbook.Close savechanges:=False   'close it dont need to save because there was no data
                    GoTo SkipCode
                End If
            End If
            
            Application.ScreenUpdating = False
        '[3. ID Room Temps]
        
            Worksheets(measuredDataSheetName).Activate
            
            'This next line of code selects the columns that contain timestamp (assumed to be column C), as well as "IR Drybulb Inlet Temp" and "IR Wetbulb Inlet Temp". The
            'last two are searched for in row 2 in order to find the correct columns.
            Union(Range("C:C"), Range("2:2").Find("IR Drybulb Inlet Temp").EntireColumn, Range("2:2").Find("IR Wetbulb Inlet Temp").EntireColumn).Select
            
            'now a chart is created based on the selected columns
            Dim Chart1 As Chart
            Set Chart1 = Charts.Add
            With Chart1
                .ChartType = xlXYScatterLinesNoMarkers
                .SeriesCollection(1).Name = "=""IR Drybulb Inlet"""
                .SeriesCollection(2).Name = "=""IR Wetbulb Inlet"""
                .HasAxis(xlValue, xlSecondary) = True
                .SeriesCollection(2).AxisGroup = xlSecondary
                .Axes(xlValue, xlPrimary).HasTitle = True
                .Axes(xlValue, xlSecondary).HasTitle = True
                .Axes(xlValue, xlSecondary).AxisTitle.Text = "Wetbulb Temp (F)"
                .Axes(xlValue, xlPrimary).AxisTitle.Text = "Drybulb Temp (F)"
                .Axes(xlCategory, xlPrimary).HasTitle = True
                .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
                
            End With
            
            Chart1.Move before:=myNewWorkbook.Sheets(measuredDataSheetName)
            Chart1.Name = "ID Room Temps"
        
        '[4. OD Room Temps]
            Worksheets(measuredDataSheetName).Activate
            
            Union(Range("C:C"), Range("2:2").Find("OR Drybulb Inlet Temp").EntireColumn, Range("2:2").Find("OR Wetbulb Inlet Temp").EntireColumn).Select
            
            Set Chart1 = Charts.Add
            With Chart1
                .ChartType = xlXYScatterLinesNoMarkers
                .SeriesCollection(1).Name = "=""OR Drybulb Inlet"""
                .SeriesCollection(2).Name = "=""OR Wetbulb Inlet"""
                .HasAxis(xlValue, xlSecondary) = True
                .SeriesCollection(2).AxisGroup = xlSecondary
                .Axes(xlValue, xlPrimary).HasTitle = True
                .Axes(xlValue, xlSecondary).HasTitle = True
                .Axes(xlValue, xlPrimary).AxisTitle.Text = "Drybulb Temp (F)"
                .Axes(xlValue, xlSecondary).AxisTitle.Text = "Wetbulb Temp (F)"
                .Axes(xlCategory, xlPrimary).HasTitle = True
                .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
            End With
            
            Chart1.Move before:=myNewWorkbook.Sheets(measuredDataSheetName)
            Chart1.Name = "OD Room Temps"
        
        '[5. UUT Power]
        
            Worksheets(measuredDataSheetName).Activate
            Union(Range("C:C"), Range("2:2").Find("UUT Power").EntireColumn).Select
            
            
            Set Chart1 = Charts.Add
            With Chart1
                .ChartType = xlXYScatterLinesNoMarkers
                .SeriesCollection(1).Name = "=""UUT Power"""
                .Axes(xlValue, xlPrimary).HasTitle = True
                .Axes(xlValue, xlPrimary).AxisTitle.Text = "Power (W)"
                .Axes(xlCategory, xlPrimary).HasTitle = True
                .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
            End With
            
            Chart1.Move before:=myNewWorkbook.Sheets(measuredDataSheetName)
            Chart1.ChartTitle.Text = "UUT Power " & Range("rpt_ReportFilename")
            Chart1.Name = "UUT Power"
        
        '[9. Calculated Data]
        
            'assume that the calculated data file name simply replaces the "MEASEMU" from the measurement data file name with "CalcsEMU"
            calculatedDataFileName = Replace(measuredDataFileName, "MEASEMU", "CalcsEMU")
            
            Set reportWorkbook = Application.Workbooks.Open(calculatedDataFileName)
            
            reportWorkbook.Sheets(1).Copy after:=myNewWorkbook.Sheets(measuredDataSheetName)
            
            ActiveSheet.Name = calculatedDataSheetName
            
            reportWorkbook.Close
            
        '[7. ID Total Capacity]
            Worksheets(calculatedDataSheetName).Activate
            
            'This next line of code selects the columns that contain timestamp (assumed to be column C) and "I.R. Total Capacity (ACFM)". The
            'latter is searched for in row 1 in order to find the correct column.
            
            Union(Range("C:C"), Range("1:1").Find("I.R. Total Capacity (ACFM)").EntireColumn).Select
            
            Set Chart1 = Charts.Add
            With Chart1
                .ChartType = xlXYScatterLinesNoMarkers
                .SeriesCollection(1).Name = "=""ID Total Capacity"""
                .Axes(xlCategory, xlPrimary).HasTitle = True
                .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
            End With
            
            Chart1.Move before:=myNewWorkbook.Sheets(calculatedDataSheetName)
            Chart1.Name = "ID Total Capacity"
            
        '[8. ID Airflow]
            Worksheets(calculatedDataSheetName).Activate
            Union(Range("C:C"), Range("1:1").Find("I.R. Measured Airflow").EntireColumn).Select
            
            Set Chart1 = Charts.Add
            With Chart1
                .ChartType = xlXYScatterLinesNoMarkers
                .SeriesCollection(1).Name = "=""ID Airflow"""
                .Axes(xlCategory, xlPrimary).HasTitle = True
                .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
            End With
            
            Chart1.Move before:=myNewWorkbook.Sheets(calculatedDataSheetName)
            Chart1.Name = "ID Airflow"
            
        '[9. ID Static Pressure]
            Worksheets(measuredDataSheetName).Activate
            
            Union(Range("C:C"), Range("2:2").Find("IR UUT Differential Pressure").EntireColumn).Select
            
            Set Chart1 = Charts.Add
            With Chart1
                .ChartType = xlXYScatterLinesNoMarkers
                .SeriesCollection(1).Name = "=""IR ESP"""
                .Axes(xlCategory, xlPrimary).HasTitle = True
                .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
            End With
            
            Chart1.Move before:=myNewWorkbook.Sheets(measuredDataSheetName)
            Chart1.Name = "IR ESP"
            
        '[10. Volts-Amps]
            Worksheets(measuredDataSheetName).Activate
            
            Union(Range("C:C"), Range("2:2").Find("UUT Voltage").EntireColumn, Range("2:2").Find("UUT Current").EntireColumn).Select
            
            Set Chart1 = Charts.Add
            With Chart1
                .ChartType = xlXYScatterLinesNoMarkers
                .SeriesCollection(1).Name = "=""UUT Voltage"""
                .SeriesCollection(2).Name = "=""UUT Current"""
                .HasAxis(xlValue, xlSecondary) = True
                .SeriesCollection(2).AxisGroup = xlSecondary
                .Axes(xlValue, xlPrimary).HasTitle = True
                .Axes(xlValue, xlSecondary).HasTitle = True
                .Axes(xlValue, xlPrimary).AxisTitle.Text = "Voltage (V)"
                .Axes(xlValue, xlSecondary).AxisTitle.Text = "Current (A)"
                .Axes(xlCategory, xlPrimary).HasTitle = True
                .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
                
                .HasTitle = True
            End With
            
            Chart1.Move before:=myNewWorkbook.Sheets(measuredDataSheetName)
            Chart1.ChartTitle.Text = "Volts-Amps " & Range("rpt_ReportFilename")
            Chart1.Name = "Volts-Amps"
            
        '[11. Volts-Amps]
            Worksheets(measuredDataSheetName).Activate
            
            Union(Range("C:C"), Range("1:1").Find("PE101").EntireColumn, Range("1:1").Find("PE102").EntireColumn).Select
            
            Set Chart1 = Charts.Add
            With Chart1
                .ChartType = xlXYScatterLinesNoMarkers
                .SeriesCollection(1).Name = "=""PE101"""
                .SeriesCollection(2).Name = "=""PE102"""

                .Axes(xlValue, xlPrimary).HasTitle = True

                .Axes(xlValue, xlPrimary).AxisTitle.Text = "PSI"
                .Axes(xlCategory, xlPrimary).HasTitle = True
                .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
                
                .HasTitle = True
            End With
            
            Chart1.Move before:=myNewWorkbook.Sheets(measuredDataSheetName)
            Chart1.ChartTitle.Text = "Pressure " & Range("rpt_ReportFilename")
            Chart1.Name = "Pressure"
            
        '[12. TC channels and names]
            
            'needs to put name in array
            
            Dim counter As Boolean
            
            
            For m = 1 To 40
                If reportmaker.Worksheets("Consolidate Reports").Cells(m + 1, 19).Value <> "" Then
                    TCS(m) = reportmaker.Worksheets("Consolidate Reports").Cells(m + 1, 19).Value
                    counter = True
                End If
            Next m
            
            Dim col As Integer
            
            If counter Then
                'need yaxis in array
                'TCYaxis
                
                For m = 1 To 40
                    If TCS(m) <> "" Then
                    myNewWorkbook.Sheets(measuredDataSheetName).Activate
                         With myNewWorkbook.Sheets(measuredDataSheetName)
                            
                            If m < 10 Then
                            col = myNewWorkbook.Sheets(measuredDataSheetName).Rows(1).Find(What:="TC0" & m).Column
                            'myNewWorkbook.Sheets(measuredDataSheetName).Range("A1").Select
                            
                            'MsgBox (Columns(Col).Address)
                                Set TCYaxis(m) = Range(Columns(col).Address).EntireColumn
                              'Set TCYaxis(m) = Columns(Col).Address.EntireColumn.Select
                                'Set TCYaxis(m) = Range(Cells(1, Col), Cells(2, Col)).EntireColumn
                                
                            Else
                                col = myNewWorkbook.Sheets(measuredDataSheetName).Rows(1).Find(What:="TC" & m).Column
                                Set TCYaxis(m) = Range(Columns(col).Address).EntireColumn
                                'Set f = .Range("2:2").Find(What:="TC" & m)
                                'Set TCYaxis(m) = .Range("find:find").EntireColumn
                            End If
                        End With
                    End If
                Next m
                
                Worksheets(measuredDataSheetName).Activate
                
                'Union(Range("C:C"), Range("2:2").Find("UUT Voltage").EntireColumn, Range("2:2").Find("UUT Current").EntireColumn).Select
                
                
                
                Set xaxis = myNewWorkbook.Sheets(measuredDataSheetName).Range(Columns(3).Address).EntireColumn
                Set Chart1 = Charts.Add
                If Chart1.SeriesCollection(1).Name <> "" Then
                    For m = 1 To 2
                    Chart1.SeriesCollection(1).Delete
                    Next m
                End If
                With Chart1
                    
                    .ChartType = xlXYScatterLinesNoMarkers
                    
                    For m = 1 To 40
                        If TCS(m) <> "" Then
                            Set TempSeries(m) = Chart1.SeriesCollection.NewSeries
                                With TempSeries(m)
                                    .Values = TCYaxis(m)
                                    .XValues = xaxis
                                    .Name = TCS(m)
                                End With
                        End If
                    Next m
                    
                    
                    
                End With
                
                Chart1.Axes(xlValue).MinimumScaleIsAuto = True
                Chart1.Axes(xlCategory).MinimumScaleIsAuto = True
                'ActiveChart.Axes(xlValue).MaximunScaleIsAuto = True
               ' ActiveChart.Axes(xlCategory).MaximunScaleIsAuto = True
                Chart1.Axes(xlValue, xlPrimary).HasTitle = True
                Chart1.Axes(xlValue, xlPrimary).AxisTitle.Text = "Temperature (F)"
                Chart1.Move before:=myNewWorkbook.Sheets(measuredDataSheetName)
                Chart1.Name = "TC Channels"
                Chart1.Axes(xlCategory, xlPrimary).HasTitle = True
                Chart1.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
          End If
        
         'Adding agilent data and checking temps if K13 is selected as Yes, otherwise it skips
        
        
        
         Application.ScreenUpdating = True
          Worksheets(reportSheetName).Activate
          Application.ScreenUpdating = False
         If reportmaker.Sheets("Consolidate Reports").Range("N5").Value = "Yes" Then
            answer = MsgBox("Do you need to add Agilent Data?", vbQuestion + vbYesNo + vbDefaultButton2, "Agilent Data?")
            If answer = vbYes Then
                'adding TCs from agilent tab to a new graph no reports made
                'pull up data source
                Dim agDataWB As Workbook
                Set agDataWB = GetData()
                    If ExitSub = True Then Exit Sub
                Dim agDatasheet As Worksheet
                Set agDatasheet = ActiveSheet
                agDatasheet.Name = "AgilentData"
                reportmaker.Sheets("AgilentTCs").Copy after:=myNewWorkbook.Sheets(myNewWorkbook.Sheets.Count)
            End If
         End If
                   
        
        If counter = True Then
            'need to make a different sheet for tescor since may only have that and not aligent
            Sheets.Add after:=myNewWorkbook.Sheets("TC Channels")
            Dim tescorTCs As Worksheet
            Set tescorTCs = ActiveSheet
            tescorTCs.Name = "Tescor TCs"
            tescorTCs.Cells(1, 1).Value = "Channel"
            tescorTCs.Cells(1, 2).Value = "Channel Name"
            tescorTCs.Cells(1, 3).Value = "Max Temps allowed (C)"
            looping = True
            Dim n As Integer
            Dim p As Integer
            n = 2
            p = 2
            While looping 'need to only put in channels that have values
                If reportmaker.Sheets("Consolidate Reports").Cells(p, 19).Value <> "" Then
                   tescorTCs.Cells(n, 1).Value = reportmaker.Sheets("Consolidate Reports").Cells(p, 18).Value 'channel
                   tescorTCs.Cells(n, 2).Value = reportmaker.Sheets("Consolidate Reports").Cells(p, 19).Value 'name
                   tescorTCs.Cells(n, 3).Value = reportmaker.Sheets("Consolidate Reports").Cells(p, 20).Value 'max temp
                   n = n + 1
                End If
                p = p + 1
                If p = 50 Then
                    looping = False
                End If
            Wend
        

        'Add TCSummary
        reportmaker.Sheets("TC Summary").Copy after:=myNewWorkbook.Sheets(myNewWorkbook.Sheets.Count)
        'myNewWorkbook.Sheets("TC Summary").Activate
        myNewWorkbook.Sheets("TC Summary").UsedRange.Value = myNewWorkbook.Sheets("TC Summary").UsedRange.Value  'for some reason works better
        'reportmaker.Sheets("TC Summary").UsedRange.Value = reportmaker.Sheets("TC Summary").UsedRange.Value 'should remove formulas from TC Summary sheet
        
        'checkTescor 'checks the tscor numbers if there are any max temp limits listed
        
        'stability check
        'If reportmaker.Sheets("Consolidate Reports").Range("D15").Value = "Yes" Then
        '    CheckStability
        '    End If
        End If
        
        
Application.DisplayAlerts = False
'Sheets("Sheet1").Delete
Application.DisplayAlerts = True

myNewWorkbook.Sheets("Report").Activate
ActiveWindow.ScrollRow = 1
Dim saveLocation As String
If Right(mySaveLocation, 1) = "\" Then
    saveLocation = mySaveLocation & "Report_" & reportWorkbookname
    Else
    saveLocation = mySaveLocation & "\Report_" & reportWorkbookname
End If
myNewWorkbook.SaveAs Filename:=saveLocation
'mySaveLocation
'reportWorkbook.Name
myNewWorkbook.Close
 
'if the file has no data then skipped
SkipCode:

Next i
       
Application.ScreenUpdating = True
       

End Sub

Sub CombineFiles()

Dim myWorkbook As Workbook
Dim myWorksheet As Worksheet
Dim myDate As String
Dim myTime As String
Dim myReports As FileDialog
Dim configuredSource As Boolean
Dim configuredDest As Boolean
Dim actingWorkbook As Workbook
Dim i As Integer
Dim names() As String
Dim sourceRanges() As String
Dim destRanges() As String
Dim TempRow As Long
Dim tempValue As Double
Dim selectionType As Long

Set myWorkbook = Application.Workbooks.Add
myWorkbook.Activate
Set myWorksheet = myWorkbook.Worksheets(1)

'select folder with reports to be processed

Dim nextCell As Range
Dim headerHeight As Long

headerHeight = 0

Set myReports = MultiFileSelect()

For Each myReport In myReports.SelectedItems
    Set actingWorkbook = Application.Workbooks.Open(myReport)
    actingWorkbook.Activate
    If headerHeight = 0 Then
        headerHeight = Application.inputbox("Enter # of rows in header", , , , , , 1)
        actingWorkbook.Worksheets(1).UsedRange.Copy
        myWorksheet.Range("A1").PasteSpecial xlPasteAll
        Application.CutCopyMode = False
    Else
        actingWorkbook.Worksheets(1).UsedRange.Copy
        myWorkbook.Activate
        TempRow = NextEmptyCellinColumn(myWorksheet.Range("A1")).Row
        NextEmptyCellinColumn(myWorksheet.Range("A1")).PasteSpecial xlPasteAll
        Application.CutCopyMode = False
                
        For i = 1 To headerHeight
            myWorksheet.Cells(TempRow, 1).EntireRow.Delete
        Next i
    
    End If
    
actingWorkbook.Close
Next myReport



End Sub

Function MultiFileSelect(Optional DefaultPath As String) As FileDialog

Set myFile = Application.FileDialog(msoFileDialogFilePicker)

With myFile
    .Title = "Select .XLS File"
    .AllowMultiSelect = True
    .Filters.Clear
    .Filters.Add "Excel Files", "*.xls; *.xlsx; *.csv"
    .InitialFileName = DefaultPath
    If .Show <> -1 Then
        Exit Function
    End If
    'myFiles = .SelectedItems
End With

Set MultiFileSelect = myFile

End Function

Function NextEmptyCellinColumn(TempRange As Range) As Range

    Set NextEmptyCellinColumn = Range(Cells(1048576, TempRange.Column).Address).End(xlUp).Offset(1)

End Function

Sub ClearContents()

ActiveSheet.Range("Q2").Clear
ActiveSheet.Range("S2:T41").Clear
End Sub
Sub checkTescorOld()
Dim reportWB As Workbook
Set reportWB = ActiveWorkbook 'we start on the newly made report

'need to make a different sheet for tescor since may only have that and not aligent
reportWB.Sheets.Add after:=reportWB.Sheets("TC Channels")
Dim tescorTCs As Worksheet
Set tescorTCs = ActiveSheet
tescorTCs.Name = "Tescor TCs"

'take this out and add to checks
'add headings
tescorTCs.Cells(1, 1).Value = "Channel"
tescorTCs.Cells(1, 2).Value = "Channel Name"
tescorTCs.Cells(1, 3).Value = "Max Temp allowed"
tescorTCs.Cells(1, 4).Value = "Max temp from Data (F)"
tescorTCs.Cells(1, 4).Value = "Max temp from Data (C)"
tescorTCs.Cells(1, 6).Value = "Difference"
tescorTCs.Cells(1, 7).Value = "Over?"

'check tescor data
    For i = 1 To 40
        If reportmaker.Sheets("Consolidate Reports").Cells(i + 1, 19).Value <> "" Then
           tescorTCs.Cells(i + 1, 1).Value = reportmaker.Sheets("Consolidate Reports").Cells(i + 1, 18).Value 'channel
           tescorTCs.Cells(i + 1, 2).Value = reportmaker.Sheets("Consolidate Reports").Cells(i + 1, 19).Value 'name
           tescorTCs.Cells(i + 1, 3).Value = reportmaker.Sheets("Consolidate Reports").Cells(i + 1, 20).Value 'max temp
        End If
           channelcolOffset = 19 + i
           If tescorTCs.Cells(i + 1, 3).Value <> "" Then
                tescorTCs.Cells(i + 1, 4).FormulaR1C1 = "=max('Measured Data'!C" & channelcolOffset & ")"
                tescorTCs.Cells(i + 1, 5).FormulaR1C1 = "=convert(R[0]C[-1],""F"",""C"")"
                tescorTCs.Cells(i + 1, 6).Formula = "=(R[0]C[-3]-R[0]C[-1])"
                If tescorTCs.Cells(i + 1, 6).Value < 0 Then
                    tescorTCs.Cells(i + 1, 7).Value = "Over"
                    tescorTCs.Cells(i + 1, 7).Interior.ColorIndex = 3
                End If
                'check numbers
           End If
    Next i


End Sub


Sub addAgilentOld()
'adding TCs from agilent tab to a new graph no reports made
'Check drop down and skip if dont want data added
Dim reportWB As Workbook
Set reportWB = ActiveWorkbook 'we start on the newly made report ' was D14
If reportmaker.Sheets("Consolidate Reports").Range("K13").Value = "No" Then 'reportmaker.Worksheets("Consolidate Reports").
Else
    'pull up data source
    Dim agDataWB As Workbook
    Set agDataWB = GetData()
        If ExitSub = True Then Exit Sub
    Dim agDatasheet As Worksheet
    Set agDatasheet = ActiveSheet
    agDatasheet.Name = "AgilentData"
    reportmaker.Sheets("AgilentTCs").Copy after:=reportWB.Sheets(reportWB.Sheets.Count)
    Dim agTCSheet As Worksheet
    Set asTCsheet = reportmaker
    'add copy of the aglient TCs from report consolidator, can then have the max temps allowed and put down what was the max temps
    
    'take this out and add to checks
    
    'need time stamps to line up with the times on the measured data
    'find time stamps needed on the measured data
    Dim firstRowTime As String
    Dim lastRowTime As String
    reportWB.Sheets("Measured Data").Activate
        firstRowTime = ConvertTimeFormat(reportWB.Sheets("Measured Data").Range("C4").Value) '(Worksheets(measuredDataSheetName).Range("C4").Value) '2023_10_11_19_01_22
        lastrow = FindLastRow
        lastRowTime = ConvertTimeFormat(reportWB.Sheets("Measured Data").Cells(lastrow, 3).Value) '(Worksheets(measuredDataSheetName).Range(FindLastRow)) '2023_10_11_22_01_31
    'find times on agilent data and return the row numbers
    Dim time1row As Integer
    Dim time2row As Integer
    agDatasheet.Activate
    time1row = FindTimeRowfromTimeString(firstRowTime, agDatasheet) '9 'will need to tell it what worksheet to go to '3671
    time2row = FindTimeRowfromTimeString(lastRowTime, agDatasheet) '2168 '5833
    
    'need to pull max temp from between those time rows for each channel that has a number
    looping = True
    Dim channelcolOffset As Integer
    Dim channelcounter As Integer
    Dim overChannelNamescounter As Integer
    Dim overChannelNames(1000) As String
    Dim i As Integer
    Dim TCS(1000) As String
    i = 1
    channelcolOffset = 1
    channelcounter = 2
    overChannelNamescounter = 1
    
    'checking temperatures
    While looping
        If reportWB.Sheets("AgilentTCs").Cells(channelcounter, 3).Value <> "" Then
            'make sure there is a channel name and then check the number
            channelCol = Sheets("AgilentTCs").Cells(channelcounter, 3).Row 'agDatasheet.Name
            reportWB.Sheets("AgilentTCs").Cells(channelcounter, 5).Formula = "=max('" & agDatasheet.Name & "'!R" & time1row & "C" & channelCol & ":R" & time2row & "C" & channelCol & ")"
            reportWB.Sheets("AgilentTCs").Cells(channelcounter, 6).Formula = "=convert(R[0]C[-1],""F"",""C"")"
                
                If reportWB.Sheets("AgilentTCs").Cells(channelcounter, 4).Value <> "" Then
                    reportWB.Sheets("AgilentTCs").Cells(channelcounter, 7).Formula = "=(R[0]C[-3]-R[0]C[-1])"
                End If
            
            If reportWB.Sheets("AgilentTCs").Cells(channelcounter, 7).Value < 0 Then
                reportWB.Sheets("AgilentTCs").Cells(channelcounter, 8).Value = "Over"
                reportWB.Sheets("AgilentTCs").Cells(channelcounter, 8).Interior.ColorIndex = 3
                overChannelNames(overChannelNamescounter) = reportWB.Sheets("AgilentTCs").Cells(channelcounter, 3).Value
                overChannelNamescounter = overChannelNamescounter + 1
            End If
            channelcounter = channelcounter + 1
        Else
            channelcounter = channelcounter + 1
            
            If Sheets("AgilentTCs").Cells(channelcounter, 2).Value = "" Then
                'stop looping and checking if there is no channel number next to it
                looping = False
            End If
        End If
    Wend
    
    
    'Dim TCYaxis(1000) As Range
    'Dim TempSeries(1000) As Series 'Creat series variable for chart
    'looping = True
    'While looping
   '
   '     'create TC channels to graph- Yaxis
   '     If Sheets("AglientTCs").Cells(i + 1, 3) <> "" Then
   '         TCS(i) = Sheets("AglientTCs").Cells(i + 1, 3).Value 'allowing there to be blanks to keep the channels at the righ number in the array
   '         ' i will be the row that is used on the aglient TCs sheet to correspond to the columns on the agDatasheet
   '         Set TCYaxis(i) = agDatasheet.Range(Cells(time1row, i + 1), Cells(time2row, i + 1))
   '     End If
   '     If Sheets("AglientTCs").Cells(i + 1, 1).Value = "" Then
   '         looping = False
   '     End If
   '     i = i + 1
   ' Wend
   '     'set xaxis as the time column
   '     Dim xaxis As Range
   '     Z = agDatasheet.Name
   '     Set xaxis = agDatasheet.Range(Cells(time1row, 1), Cells(time2row, 1))
   '     'Graphing the TCs
   '     Set Chart1 = Charts.Add
   '     Chart1.ChartArea.Clear

    '   With Chart1
    '       .ChartType = xlXYScatterLinesNoMarkers
    '        For m = 1 To i - 1
    '            If TCS(m) <> "" Then
    '                Set TempSeries(m) = Chart1.SeriesCollection.NewSeries
    '                    With TempSeries(m)
    '                        .Values = TCYaxis(m)
    '                        .XValues = xaxis
    '                        .Name = TCS(m)
    '                    End With
     '           End If
     '       Next m
     '   End With
    
    'msgbox to tell about the TCs that hit their limit
    If overChannelNames(overChannelNamescounter - 1) <> "" Then
        'alert if there were some channels over limits
        msg = "Channel names over limits: " & vbNewLine
        msg2 = ""
            While overChannelNamescounter > 0
                msg2 = msg2 & overChannelNames(overChannelNamescounter) & vbNewLine
                overChannelNamescounter = overChannelNamescounter - 1
            Wend
        MsgBox (msg & msg2)
    End If
End If


End Sub

Function GetData()
'function to quickly pull the data from the file, copy and close the original
'at this point will use the actual data file, would be safer than using the test report being generated
Dim rawdatasheetwb As Workbook
Dim datasheetwb As Workbook
Set datasheetwb = ActiveWorkbook
Dim rawdatasheetname As String

With Application.FileDialog(msoFileDialogFilePicker)
    .Title = "Please pick raw Agilent data file to copy out the data"
    .AllowMultiSelect = False
    If .Show = True Then
        rawdatasheetname = .SelectedItems(1)
        If Right(rawdatasheetname, 4) <> ".csv" Then
            Dim answer As Integer
            answer = MsgBox("This does not look like the raw data, would you like to procede?", vbQuestion + vbYesNo + vbDefaultButton2, "Please check data")
                If answer = vbNo Then
                   ExitSub = True
                End If
        End If
    End If
End With

If ExitSub = False Then
    Workbooks.Open (rawdatasheetname)
    Set rawdatasheetwb = ActiveWorkbook
    
    
    y = rawdatasheetwb.Name
    z = datasheetwb.Name
    rawdatasheetwb.Sheets(1).Copy after:=datasheetwb.Sheets(datasheetwb.Sheets.Count)
    'reportWorkbook.Sheets(1).Copy after:=myNewWorkbook.Sheets("Pics")
    rawdatasheetwb.Close
    
    Set GetData = datasheetwb
End If

End Function


Function ConvertTimeFormat(time As String)
'needs to be configured for all data sets
    Sec = CInt(Right(time, 2))
    Min = CInt(Mid(time, 15, 2))
    hr = CInt(Mid(time, 12, 2))
    Days = CInt(Mid(time, 9, 2))
    Mon = CInt(Mid(time, 6, 2))
    Years = CInt(Left(time, 4))
    
    ConvertTimeFormat = DateSerial(Years, Mon, Days) + TimeSerial(hr, Min, Sec)


End Function

Function FindLastRow()

FindLastRow = Cells(Rows.Count, 3).End(xlUp).Row

End Function

Function FindTimeRowfromTimeString(time As String, agDatasheet As Worksheet)
'parameter should be in a basic "mm/dd/yyy hh:mm:ss"
'returning the time row via an integer
Dim rowcounter As Integer
rowcounter = FindLastRow()
looping = True
Do While looping
    T1 = agDatasheet.Cells(rowcounter, 1).Value
    T2 = agDatasheet.Cells(rowcounter - 1, 1).Value
    If DateDiff("s", time, T1) >= 0 And DateDiff("s", time, T2) < 0 Then
        looping = False
    Else
        rowcounter = rowcounter - 1
    End If
    
Loop

FindTimeRowfromTimeString = rowcounter


End Function

Sub AdjustTCSummaryPage()


Dim summaryPage As Worksheet
Dim agilentTCPage As Worksheet
Dim tescorTCPage As Worksheet

Set reportmaker = ActiveWorkbook
Set summaryPage = reportmaker.Sheets("TC Summary")
Set agilentTCPage = reportmaker.Sheets("AgilentTCs")
Set tescorTCPage = reportmaker.Sheets("Consolidate Reports")

'remove past items
summaryPage.Range("1:4").Clear

'set up
summaryPage.Range("A2").Value = "TC#"
summaryPage.Range("A3").Value = "Desciption"
summaryPage.Range("A4").Value = "Requirements-->"

looping = True
i = 1
'starting with aglient TCS, should put any with a channel number in column B on agilent TCs, will stop when reaching the end of the those channels
While agilentTCPage.Cells(i + 1, 2) <> ""
    summaryPage.Cells(1, i + 1).Value = i
    summaryPage.Cells(2, i + 1).Value = "TC" + agilentTCPage.Cells(i + 1, 2).Text
    summaryPage.Cells(3, i + 1).FormulaR1C1 = "=IF(ISBLANK(AgilentTCs!R" & i + 1 & "C3),""(LEFT BLANK)"",AgilentTCs!R" & i + 1 & "C3)"
    
    summaryPage.Cells(4, i + 1).FormulaR1C1 = "=if(isblank(AgilentTCs!R" & i + 1 & "C4),"""",AgilentTCs!R" & i + 1 & "C4)"
    
    'summaryPage.Cells(2, i + 1).Value = "TC" + agilentTCs.Cells(i + 1, 2).Text
    i = i + 1
Wend

'slighlty different, will skip all empty of channel names and check all the way to the end of the channel numbers
J = 1
While tescorTCPage.Cells(J + 1, 18).Value <> ""
    If tescorTCPage.Cells(J + 1, 19).Value <> "" Then
        summaryPage.Cells(1, i + 1).Value = tescorTCPage.Cells(J + 1, 18).Value
        summaryPage.Cells(2, i + 1).Value = tescorTCPage.Cells(i + 1, 19).Value
        summaryPage.Cells(3, i + 1).Value = "=IF(ISBLANK('Consolidate Reports'!R" & J + 1 & "C19),"""",'Consolidate Reports'!R" & J + 1 & "C19)"
        summaryPage.Cells(4, i + 1).Value = "=IF(ISBLANK('Consolidate Reports'!R" & J + 1 & "C20),"""",'Consolidate Reports'!R" & J + 1 & "C20)"
        i = i + 1
    End If
    J = J + 1
    
Wend

Set rng = summaryPage.Range("1:4")
With rng
    .Borders.LineStyle = xlContinuous
    .ColumnWidth = 12
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = True
End With

summaryPage.Range("5:5").Clear
End Sub
