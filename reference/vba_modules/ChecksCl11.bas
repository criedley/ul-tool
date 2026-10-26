Attribute VB_Name = "ChecksCl11"
Public col() As Variant
Public ExitSub As Boolean
Public timeCol As Integer
Public DataAq As String
Public DataSheetOffset As Integer
Public reportmaker As Workbook
Public reportWB As Workbook
Public Cl19 As Boolean
Public tescorTCs As Worksheet
Public agTCSheet  As Worksheet
Function ConvertStringtoDateTime(userinput As String, wb As Workbook) 'taking in a string that is hh:mm miltary time and translating it to a date time to function later
DataAq = "Tescor"
'find date from wb "Measured Data" and then pull together
desiredtimestamp = CDate(userinput)
'10/25/24- adding in extra input- need to know date to use if 2 dates on tecor data
'funciton to test if there are 2 dates
sheetDate = WhichDateDecider(wb)
 ' = ConvertTimeFormat(wb.Sheets("Measured Data").Range("C4").Value, DataAq)
'y = CLng(sheetDate)
'Z = CDate(y)
'a = CDate(Format(sheetDate, "yyyy-mm-dd"))
desiredDate = CDate(Format(sheetDate, "yyyy-mm-dd"))
Dim combinedDateTime As String
combinedDateTime = desiredDate + desiredtimestamp
ConvertStringtoDateTime = combinedDateTime
End Function
Function WhichDateDecider(wb As Workbook)
'should always be Tescor
wb.Sheets("Measured Data").Activate
lastrow = FindLastRow()
FirstDate = CDate(Format(ConvertTimeFormat(Range("C4").Value, DataAq), "yyyy-mm-dd"))
LastDate = CDate(Format(ConvertTimeFormat(Cells(lastrow, 3).Value, DataAq), "yyyy-mm-dd"))

'if thie first date on the tescor data doesn't match the last then it will prompt you to tell it which date you meant
If FirstDate <> LastDate Then
    answer = MsgBox("This tescor data spans overnight to 2 dates. Which date should be used? Click Yes for: " & FirstDate & " or click No for: " & LastDate & "", vbQuestion + vbYesNo + vbDefaultButton2, "Date in question")
        If answer = vbYes Then
            WhichDateDecider = FirstDate
        Else
            WhichDateDecider = LastDate
        End If
Else
    WhichDateDecider = FirstDate

End If

End Function


Sub patch()

d = Range(Cells(1, 1), Cells(2, 100)).Find("Max Temps (C)").Column


MsgBox (Range("E4").Formula)
Dim z As String
z = Range("E4").Formula
Dim y(10) As String
a = InStr(z, "'")
b = InStrRev(z, ")")
c = Mid(z, a, (b - (a + 1)))

Range("R2").Formula2 = "=max(" & c & ")"

End Sub


Sub CleanReportCl11()
Dim reportmaker As Workbook
Set reportmaker = ActiveWorkbook 'we start on the report consolidator
Dim reportWB As Workbook
Set reportWB = GetData("Select Tescor report file to clean into a new workbook")

'handle not having some of the sheets that we need
'Dim reportsheetsArray As Variant
'reportsheetsArray = Array("ID Room Temps", "OD Room Temps", "UUT Power", "Volts-Amps", "TC Channels", "Tescor TCs", "Measured Data", "Agilent Data", "AgilentTCs", "TC Summary")

'k = 1
'Dim sheetsArray(20) As Variant
'    For i = 1 To reportWB.Sheets.Count
'        For j = 0 To 9 'number of array components
'        y = reportWB.Sheets(i).Name
'        z = reportsheetsArray(j)
'            If reportWB.Sheets(i).Name = reportsheetsArray(j) Then
'                sheetsArray(k) = reportWB.Sheets(i).Name
'                k = k + 1
'            End If
'        Next j
'    Next i

'copy tabs to new workbook

'reportWB.Sheets(Array(sheetsArray)).Copy

reportWB.Sheets(Array("ID Room Temps", "OD Room Temps", "UUT Power", "Volts-Amps", "TC Channels", "Tescor TCs", "Measured Data", "AgilentData", "AgilentTCs", "TC Summary")).Copy
', "Measured Data", "TC Channels", "Volts-Amps", "OD Room Temps", "ID Room Temps", "UUT Power", "AgilentTCs", "Agilent Data", "TC Summary"
Set cleanedDatawb = ActiveWorkbook
reportWB.Close (False)
'change formula produced on Max Temps (C)
Dim TCsheet As Worksheet
Dim dataAqArray(10) As String
dataAqArray(1) = "Tescor"
If CheckforAgilentData = True Then
'If reportmaker.Sheets("Consolidate Reports").Range("K13").Value = "Yes" Then
    dataAqArray(2) = "Aglient 34980A" 'can add more types of data later
    End If
Dim datasheet As Worksheet


i = 1 'data aq counter

While dataAqArray(i) <> ""
    If dataAqArray(i) = "Tescor" Then
            Set datasheet = cleanedDatawb.Sheets("Measured Data")
            Set TCsheet = cleanedDatawb.Sheets("Tescor TCs")
            coloffset = 19
            channelNameCol = 2
            maxAllowedTempCol = 3
        Else
            Set datasheet = cleanedDatawb.Sheets("AgilentData")
            Set TCsheet = cleanedDatawb.Sheets("AgilentTCs")
            coloffset = 1
            channelNameCol = 3
            maxAllowedTempCol = 4
        End If
    
    TCsheet.Activate
    Dim formulaAddr As Range
    Set formulaAddr = TCsheet.Range(Cells(1, 1), Cells(2, 100)).Find("Max Temps (F) from data", LookIn:=xlValues, SearchDirection:=xlPrevious)
    
    formulaCol = formulaAddr.Column
    Dim maxTempAddr As Range
    Set maxTempAddr = TCsheet.Range(Cells(1, 1), Cells(2, 100)).Find("Max Temps (C)", LookIn:=xlValues, SearchDirection:=xlPrevious)
    maxTempCol = maxTempAddr.Column
    Dim str As String
    
    
    
        'Next Highest Temp (F) remove this column
        Set nexthighestAddr = TCsheet.Range(Cells(1, 1), Cells(2, 100)).Find("Next Highest Temp (F)", LookIn:=xlValues, SearchDirection:=xlPrevious)
        If IsEmpty(nexthighestAddr) Then
            nexthighestCol = nexthighestAddr.Column
        
            J = 2 'channel row counter
            While TCsheet.Cells(J, channelNameCol - 1).Value <> ""
                 If TCsheet.Cells(J, channelNameCol).Value <> "" Then
                        'Max Temps (F) from data
                        str = TCsheet.Cells(J, formulaCol).Formula
                        StartStr = InStr(str, "'")
                        endStr = InStrRev(str, ")")
                        fullAddr = Mid(str, StartStr, (endStr - (StartStr)))
                        'redoing formula to not reference another cell that is going to be deleted
                        TCsheet.Cells(J, maxTempCol).Formula = "=convert(large(" & fullAddr & ",{1}),""F"",""C"")"
                        
                        'match color format from next highest column
                        If TCsheet.Cells(J, nexthighestCol).Interior.ColorIndex <> xlNone Then
                           TCsheet.Cells(J, maxTempCol).Interior.ColorIndex = TCsheet.Cells(J, nexthighestCol).Interior.ColorIndex
                        End If
                End If
                J = J + 1
            Wend
            
            'deleting that next highest number column because it is no longer needed
            Columns(nexthighestCol).EntireColumn.Delete
        
    End If
    
i = i + 1
Wend

End Sub

Sub GetTescorConditions()
Set reportmaker = ActiveWorkbook 'we start on the report consolidator

Set reportWB = GetData("Select Report to get Tescor conditons of")

Dim reportsheet As Worksheet
Set reportsheet = reportWB.Sheets("Report")

Dim tescorData As Worksheet
Set tescorData = reportWB.Sheets("Measured Data")

DataAq = "Tescor"
timeCol = 3
Dim endtime As String

Dim userinput As String
userinput = Application.inputbox("Input end time of stabilization, if you want it to pull from the end, leave at 0. Input in military time (hh:mm)", "Time input", 0) 'like for examples 14:10
    If userinput = "" Then Exit Sub
    
    
tescorData.Activate

If userinput = "0" Then
    endtime = ConvertTimeFormat(tescorData.Cells(FindLastRow(), 3).Value, DataAq)
Else
    'endtime = ConvertTimeFormat(ConvertStringtoDateTime(userinput, reportWB), DataAq)
    endtime = ConvertStringtoDateTime(userinput, reportWB)
End If




Dim k As Integer
Dim timeStringArray(60) As String
    For k = 0 To 60
        timeStringArray(k) = GetTimefromSubtractedMin(k, endtime)
        'so 1 will be 1 min before the end of the test
    Next k
    
Dim timeRowArray(60) As Integer
    For k = 0 To 60
        timeRowArray(k) = FindTimeRowfromTimeString(timeStringArray(k), FindLastRow())
        'so now 1 will be the row 1min from end of the test
    Next k

colOff = 15
rowOff = 10

Dim colArray(10) As Integer

'Ambient C - ID Db 'E 5
reportsheet.Cells(rowOff + 1, colOff).Value = "Ambient C - ID Db"
colArray(1) = 5
'Ambient C - ID Wb 'F 6
reportsheet.Cells(rowOff + 2, colOff).Value = "Ambient C - ID Wb"
colArray(2) = 6
'Ambient C - OD Db 'I 9
reportsheet.Cells(rowOff + 3, colOff).Value = "Ambient C - OD Db"
colArray(3) = 9
'Ambient C - OD Wb 'J 10
reportsheet.Cells(rowOff + 4, colOff).Value = "Ambient C - OD Wb"
colArray(4) = 10
'V 'AT 46
reportsheet.Cells(rowOff + 5, colOff).Value = "V"
colArray(5) = 46
'A 'AU 47
reportsheet.Cells(rowOff + 6, colOff).Value = "A"
colArray(6) = 47
'W 'AV 48
reportsheet.Cells(rowOff + 7, colOff).Value = "W"
colArray(7) = 48
'Hi pres 'AR 44
reportsheet.Cells(rowOff + 8, colOff).Value = "Hi pres"
colArray(8) = 44
'Low pres 'AS 45
reportsheet.Cells(rowOff + 9, colOff).Value = "Low pres"
colArray(9) = 45

reportsheet.Cells(rowOff, colOff + 1).Value = "0min"
reportsheet.Cells(rowOff, colOff + 2).Value = "15min"
reportsheet.Cells(rowOff, colOff + 3).Value = "30min"
reportsheet.Cells(rowOff, colOff + 4).Value = "45min"
reportsheet.Cells(rowOff, colOff + 5).Value = "60min"
reportsheet.Cells(rowOff, colOff + 6).Value = "Max"

For i = 1 To 4
    reportsheet.Cells(rowOff + i, colOff + 1).FormulaR1C1 = "=convert('" & tescorData.Name & "'!R" & timeRowArray(60) & "C" & colArray(i) & ",""F"",""C"")"
    reportsheet.Cells(rowOff + i, colOff + 2).FormulaR1C1 = "=convert('" & tescorData.Name & "'!R" & timeRowArray(45) & "C" & colArray(i) & ",""F"",""C"")"
    reportsheet.Cells(rowOff + i, colOff + 3).FormulaR1C1 = "=convert('" & tescorData.Name & "'!R" & timeRowArray(30) & "C" & colArray(i) & ",""F"",""C"")"
    reportsheet.Cells(rowOff + i, colOff + 4).FormulaR1C1 = "=convert('" & tescorData.Name & "'!R" & timeRowArray(15) & "C" & colArray(i) & ",""F"",""C"")"
    reportsheet.Cells(rowOff + i, colOff + 5).FormulaR1C1 = "=convert('" & tescorData.Name & "'!R" & timeRowArray(0) & "C" & colArray(i) & ",""F"",""C"")"
Next i

For i = 5 To 9
    reportsheet.Cells(rowOff + i, colOff + 1).FormulaR1C1 = "='" & tescorData.Name & "'!R" & timeRowArray(60) & "C" & colArray(i) & ""
    reportsheet.Cells(rowOff + i, colOff + 2).FormulaR1C1 = "='" & tescorData.Name & "'!R" & timeRowArray(45) & "C" & colArray(i) & ""
    reportsheet.Cells(rowOff + i, colOff + 3).FormulaR1C1 = "='" & tescorData.Name & "'!R" & timeRowArray(30) & "C" & colArray(i) & ""
    reportsheet.Cells(rowOff + i, colOff + 4).FormulaR1C1 = "='" & tescorData.Name & "'!R" & timeRowArray(15) & "C" & colArray(i) & ""
    reportsheet.Cells(rowOff + i, colOff + 5).FormulaR1C1 = "='" & tescorData.Name & "'!R" & timeRowArray(0) & "C" & colArray(i) & ""
    reportsheet.Cells(rowOff + i, colOff + 6).FormulaR1C1 = "=large('" & tescorData.Name & "'!C" & colArray(i) & ":C" & colArray(i) & ", {1})"
    'LARGE('Measured Data'!$AU:AU,{1})
Next i

reportsheet.Activate


End Sub

Sub addingFromOldReportMaker()
'just need to add the summary file, TCs and agilent TCs from this report consolidator
'Adding agilent data and checking temps if K13 is selected as Yes, otherwise it skips
Set reportmaker = ActiveWorkbook 'we start on the report consolidator
Set reportWB = GetData("Select older report to pull data into")


         'Adding agilent data and checking temps if K13 is selected as Yes, otherwise it skips
        'If reportmaker.Sheets("Consolidate Reports").Range("N5").Value = "Yes" Then
        answer = MsgBox("Do you need to add Agilent Data?", vbQuestion + vbYesNo + vbDefaultButton2, "Agilent Data?")
        If answer = vbYes Then
            'adding TCs from agilent tab to a new graph no reports made
            'pull up data source
            Dim agDataWB As Workbook
            Set agDataWB = GetData("Select agilent data to pull")
                If ExitSub = True Then Exit Sub
            Dim agDatasheet As Worksheet
            Set agDatasheet = ActiveSheet
            agDatasheet.Name = "AgilentData"
            agDatasheet.Copy after:=reportWB.Sheets(reportWB.Sheets.Count)
            reportmaker.Sheets("AgilentTCs").Copy after:=reportWB.Sheets(reportWB.Sheets.Count)
            agDataWB.Close savechanges = False
        End If
            
        'need to make a different sheet for tescor since may only have that and not aligent
        Sheets.Add after:=reportWB.Sheets("Measured Data")
        Dim tescorTCs As Worksheet
        Set tescorTCs = ActiveSheet
        tescorTCs.Name = "Tescor TCs"
        tescorTCs.Cells(1, 1).Value = "Channel"
        tescorTCs.Cells(1, 2).Value = "Channel Name"
        tescorTCs.Cells(1, 3).Value = "Max Temps allowed (C)"
        looping = True
        n = 2
        p = 2
        While looping 'need to only put in channels that have values
            If reportmaker.Sheets("Consolidate Reports").Cells(n + 1, 19).Value <> "" Then
               tescorTCs.Cells(p, 1).Value = reportmaker.Sheets("Consolidate Reports").Cells(n + 1, 18).Value 'channel
               tescorTCs.Cells(p, 2).Value = reportmaker.Sheets("Consolidate Reports").Cells(n + 1, 19).Value 'name
               tescorTCs.Cells(p, 3).Value = reportmaker.Sheets("Consolidate Reports").Cells(n + 1, 20).Value 'max temp
               p = p + 1
            End If
            n = n + 1
            
            If n = 41 Then
             looping = False
            End If
        Wend
        
        'Add TCSummary
        reportmaker.Sheets("TC Summary").Copy after:=reportWB.Sheets(reportWB.Sheets.Count)
        'myNewWorkbook.Sheets("TC Summary").Activate
        reportWB.Sheets("TC Summary").UsedRange.Value = reportWB.Sheets("TC Summary").UsedRange.Value  'for some reason works better
        
        reportWB.Save

End Sub


Sub checkMaxTemps()

Set reportmaker = ActiveWorkbook 'we start on the report consolidator
Dim TCsheet As Worksheet
'need to select report to be checked
Set reportWB = GetData("Select Report to check for the max temps on")
Dim datasheet As Worksheet
Dim dataAqArray(10) As String
dataAqArray(1) = "Tescor"
Dim i As Integer
Application.ScreenUpdating = False
Dim maxAllowedTempCol As Integer
If CheckforAgilentData = True Then
'If reportmaker.Sheets("Consolidate Reports").Range("K13").Value = "Yes" Then
    dataAqArray(2) = "Aglient 34980A" 'can add more types of data later
    End If
J = 1

While dataAqArray(J) <> ""
    'checking from a premade report
    DataAq = dataAqArray(J) 'everything is already set up for DataAq
    If DataAq = "Tescor" Then
        Set datasheet = reportWB.Sheets("Measured Data")
        Set TCsheet = reportWB.Sheets("Tescor TCs")
        coloffset = 19
        channelNameCol = 2
        maxAllowedTempCol = 3
        'channelCol = agTCSheet.Cells(channelcounter, 3).Row 'agDatasheet.Name
        
    Else
        Set datasheet = reportWB.Sheets("AgilentData")
        Set TCsheet = reportWB.Sheets("AgilentTCs")
        coloffset = 1
        channelNameCol = 3
        maxAllowedTempCol = 4
        'channelCol = agTCSheet.Cells(channelcounter, 3).Row 'agDatasheet.Name
    End If
    
    Dim sheetOffset As Integer
    sheetOffset = findNextCol(TCsheet) 'needs the Channel Name column filled so added that back to the reportmaker

    Dim firstRowTime As String
    Dim lastRowTime As String
    Dim lastrow As Integer
    'firstRowTime = ConvertTimeFormat(reportWB.Sheets("Measured Data").Range("C4").Value, "Tescor") '(Worksheets(measuredDataSheetName).Range("C4").Value) '2023_10_11_19_01_22
    'lastRow = FindLastRow
    'lastRowTime = ConvertTimeFormat(reportWB.Sheets("Measured Data").Cells(lastRow, 3).Value, "Tescor") '(Worksheets(measuredDataSheetName).Range(FindLastRow)) '2023_10_11_22_01_31
    
    Dim time1row As Integer
    Dim time2row As Integer
    Dim channelrow As Integer
    datasheet.Activate
    If DataAq = "Tescor" Then
        timeCol = 3
        firstRowTime = ConvertTimeFormat(reportWB.Sheets("Measured Data").Range("C4").Value, "Tescor") '(Worksheets(measuredDataSheetName).Range("C4").Value) '2023_10_11_19_01_22
        lastrow = FindLastRow
        lastRowTime = ConvertTimeFormat(reportWB.Sheets("Measured Data").Cells(lastrow, 3).Value, "Tescor") '(Worksheets(measuredDataSheetName).Range(FindLastRow)) '2023_10_11_22_01_31
        time2row = lastrow 'need int
        time1row = 4
        
        'time2row = FindTimeRowfromTimeString(lastRowTime, lastRow)
        'time1row = FindTimeRowfromTimeString(ConvertTimeFormat(firstRowTime, DataAq), time2row) '9 'will need to tell it what worksheet to go to '3671
    Else
        
       
        timeCol = 1
        lastrow = FindLastRow
        time2row = FindTimeRowfromTimeString(lastRowTime, lastrow)
        '11/25/24 adding in code to delete everything outside of the relevant time data on the agilent code
        If time2row <> FindLastRow Then
            datasheet.Range(Cells(time2row + 1, 1), Cells(FindLastRow, 1)).EntireRow.Delete
        End If
        time1row = FindTimeRowfromTimeString(ConvertTimeFormat(firstRowTime, DataAq), time2row)  '9 'will need to tell it what worksheet to go to '3671
        
        For i = 1 To 50
            If datasheet.Cells(i, 1).Value = "Time" Then
             channelrow = i
             End If
        Next i
        If channelrow + 1 <> time1row Then
            datasheet.Range(Cells(channelrow + 1, 1), Cells(time1row - 1, 1)).EntireRow.Delete
        End If
        
        datasheet.Range(Cells(1, 1), Cells(6, 1)).EntireRow.Insert
        time2row = FindTimeRowfromTimeString(lastRowTime, FindLastRow)
        time1row = FindTimeRowfromTimeString(ConvertTimeFormat(firstRowTime, DataAq), time2row)
         
        datasheet.Cells(1, 1) = "Max Temps(F)"
        datasheet.Cells(2, 1) = "Max Temps(C)"
        datasheet.Cells(3, 1) = "Max Temps allowed"
        datasheet.Cells(4, 1) = "Delta"
            
    End If
    
    
    
    
    If dataAqArray(J) = "Aglient 34980A" Then
        'adding new agilent chart
        Dim agilentgraph As Chart
        Set agilentgraph = Charts.Add
                With agilentgraph
                    .ChartType = xlXYScatterLinesNoMarkers
                    .ChartArea.Clear
                    .Axes(xlCategory, xlPrimary).HasTitle = True
                    .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
                End With
                
        agilentgraph.Move before:=datasheet
        agilentgraph.Name = "Agilent Plot"
    End If
            
    datasheet.Activate
            
    Dim agilentDataRange As Range
    Dim agilentTimeRange As Range
    
    Dim ser As Series
    'add headings
    'tescorTCs.Cells(1, 1).Value = "Channel"
    'tescorTCs.Cells(1, 2).Value = "Channel Name"
    'tescorTCs.Cells(1, 3).Value = "Max Temps allowed (C)"
    TCsheet.Cells(1, sheetOffset).Value = "Max Temps (F) from data" '4
    TCsheet.Cells(1, sheetOffset + 1).Value = "Average Top 5-10 temps (F)"
    TCsheet.Cells(1, sheetOffset + 2).Value = "Delta"
    'TCSheet.Cells(1, sheetOffset + 3).Value = "Next Highest Temp (F)"
    TCsheet.Cells(1, sheetOffset + 3).Value = "Max Temps (C)"
    TCsheet.Cells(1, sheetOffset + 4).Value = "Delta"
    looping = True
    looping2 = True
    Dim channelNumber As String
    i = 2
    While looping
    'check raw data
    '    For i = 1 To 40
           'channelcolOffset = 19
           'change channelCol if on agilent system
           If dataAqArray(J) = "Aglient 34980A" Then
                'need to search for the channel number on the TC sheet- make it a function to return the channelCol
                channelCol = FindChannelColFromAgilentSheet(TCsheet, datasheet, i)
           Else
                channelCol = coloffset + TCsheet.Cells(i, 1) 'number from column 1 should be the channel on all TC sheet lists
           End If
           
           looping2 = True
           If TCsheet.Cells(i, 3).Value <> "" Then 'checking to see if it has a name
                TCsheet.Cells(i, sheetOffset).FormulaR1C1 = "=max('" & datasheet.Name & "'!R" & time1row & "C" & channelCol & ":R" & time2row & "C" & channelCol & ")"
                'max temps afrom data
                TCsheet.Cells(i, sheetOffset + 1).FormulaR1C1 = "=average(LARGE('" & datasheet.Name & "'!R" & time1row & "C" & channelCol & ":R" & time2row & "C" & channelCol & ", {5,6,7,8,9,10}))"
                'Average Top 5-10 temps (F)
                TCsheet.Cells(i, sheetOffset + 2).FormulaR1C1 = "=abs(R[0]C[-2]-R[0]C[-1])"
                'Delta
                
                'only need next highest if not within delta
                If TCsheet.Cells(i, sheetOffset + 2).Value > 10 Then
                    TCsheet.Cells(i, sheetOffset).Interior.ColorIndex = 3
                    TCsheet.Cells(i, sheetOffset + 2).Interior.ColorIndex = 3
                End If
                
                TCsheet.Cells(i, sheetOffset + 3).FormulaR1C1 = "=convert(large('" & datasheet.Name & "'!R" & time1row & "C" & channelCol & ":R" & time2row & "C" & channelCol & ",{1}),""F"",""C"")" 'max temps (C)
                TCsheet.Cells(i, sheetOffset + 4).Formula = "=if(isblank(R[0]C" & maxAllowedTempCol & "),"""",R[0]C" & maxAllowedTempCol & "-R[0]C[-1])" 'Delta =(D6-I6) =if(isblank(R[0]C[-6]),"",R[0]C[-6]-R[0]C[-1])
                If TCsheet.Cells(i, sheetOffset + 4).Value < 0 Then
                    TCsheet.Cells(i, sheetOffset + 4).Interior.ColorIndex = 3
                End If
                
                'adding in extra on agilent raw data
                If dataAqArray(J) = "Aglient 34980A" Then
                    k = 1
                    datasheet.Activate
                    Set agilentTimeRange = datasheet.Range(Cells(time1row, 1), Cells(time2row, 1))
                    'max temps F '=MAX(B12:B1000000)
                    datasheet.Cells(k, channelCol).FormulaR1C1 = "=max(R" & time1row & "C" & channelCol & ":R" & time2row & "C" & channelCol & ")"
                    'max temps C =CONVERT(B1,"F","C")
                    datasheet.Cells(k + 1, channelCol).FormulaR1C1 = "=convert(R[-1]C,""F"",""C"")"
                    'temp limit from page
                    datasheet.Cells(k + 2, channelCol).FormulaR1C1 = "=if(isblank('" & TCsheet.Name & "'!R" & i & "C" & maxAllowedTempCol & "),"""",'" & TCsheet.Name & "'!R" & i & "C" & maxAllowedTempCol & ")"
                    'Delta c =B3-B2
                    datasheet.Cells(k + 3, channelCol).FormulaR1C1 = "=if(isnumber(R[-1]C),R[-1]C - R[-2]C,"""")"
                    
                    'adding in graph of only the items I need
                    Set agilentDataRange = datasheet.Range(Cells(time1row, channelCol), Cells(time2row, channelCol))
                    With agilentgraph
                        Set ser = .SeriesCollection.NewSeries
                        '.ChartArea.Clear
                        ser.Values = agilentDataRange
                        ser.XValues = agilentTimeRange
                    End With
                End If
            
            
            
           End If
           If TCsheet.Cells(i, 1).Value = "" Then
                'stop looping and checking if there is no channel number next to it
                looping = False
            End If
           i = i + 1
        'Next i
    Wend 'Looping
    
   J = J + 1
Wend 'dataAqArray(j) <> ""
    
'If reportmaker.Sheets("Consolidate Reports").Range("K13").Value = "Yes" Then
'    checkAgilentMaxTemps 'will check if agilent temps are needed and then check them as well
'End If

'pull all numbers to the TC Summary
'need to match the channel number, agilent system will be "TC1001" while tescor will be "2"
Dim TCSummary As Worksheet
Dim SummaryRow As Integer
Dim pullingSummaryData As Boolean
Dim maxTempCol As Integer

If CheckforAgilentData Then
    Set TCSummary = reportWB.Sheets("TC Summary")
    Set TCsheet = reportWB.Sheets("AgilentTCs")
    'maxTempCol
    SummaryRow = 6
    pullingSummaryData = True
    i = 2
    
    If dataAqArray(2) = "Aglient 34980A" Then 'reportmaker.Sheets("Consolidate Reports").Range("K13").Value = "Yes" Then
        'find max temp column for agilent data
        While TCsheet.Cells(1, i).Value <> ""
            If TCsheet.Cells(1, i).Value = "Max Temps (C)" Then
                maxTempCol = i
            End If
            i = i + 1
        Wend
    
        i = 2
        J = 2
        z = TCsheet.Name
        While pullingSummaryData 'will have new dataaq system soon and will need to be changed
            x = TCsheet.Cells(J, 3).Value
            y = TCSummary.Cells(3, i).Value
    
            If TCsheet.Cells(J, 3).Value = TCSummary.Cells(3, i).Value Then
                TCSummary.Cells(SummaryRow, i).FormulaR1C1 = "=('" & TCsheet.Name & "'!R" & J & "C" & maxTempCol & ")" 'hard wired where max temps will be
                i = i + 1
            End If
            J = J + 1
            a = TCsheet.Cells(J, 2).Value
            If TCsheet.Cells(J, 2).Value = "" Then
                i = i + 1
                J = 2
            End If
    
           If TCsheet.Cells(i, 2).Value = "" Then
                'end of TCs
                pullingSummaryData = False
           End If
           '
            If i > 1000 Then 'need a way to make it break out if it cannot find the channel number
                MsgBox ("Double check temps on TC Summary")
                pullingSummaryData = False
            End If
        Wend
    End If
    
    i = 1
    'find max temp column for Tescor data
    Set TCsheet = reportWB.Sheets("Tescor TCs")
    While TCsheet.Cells(1, i).Value <> ""
        If TCsheet.Cells(1, i).Value = "Max Temps (C)" Then
            maxTempCol = i
        End If
        i = i + 1
    Wend
    
    pullingSummaryData = True
        
    i = 1
    J = 2
    
    While pullingSummaryData 'Tescor Data
        If TCsheet.Cells(J, 2) = TCSummary.Cells(3, i) Then 'these are simple int so no need to convert
        'TCSummary.Cells(SummaryRow, i + Summaryoffset).Formula = "='AglientTCs'!R2C11"
            TCSummary.Cells(SummaryRow, i).FormulaR1C1 = "=('" & TCsheet.Name & "'!R" & J & "C" & maxTempCol & ")" 'hard wired where max temps will be
            J = J + 1
        End If
        
        i = i + 1
        If i > 1000 Then 'need a way to make it break out if it cannot find the channel number
            MsgBox ("Double check temps on TC Summary")
            pullingSummaryData = False
        End If
        
        If TCsheet.Cells(J, 2) = "" Then 'once it reaches the end of the channel numbers - will stop pulling
            pullingSummaryData = False
        End If
    Wend
End If

If CheckforAgilentData Then



End If

'i = 2
'new section 11/25/24 - Kai wants the max temps on the actual data page, at least on the agilent data
'If CheckForAgilentData = True Then
'    'only need if there is agilent data
'    Set TCSheet = reportWB.Sheets("AgilentTCs")
'    datasheet.Range(Cells(1, 1), Cells(6, 1)).EntireRow.Insert
'TCSummary = reportWB.Sheets("TC Summary")
'TCSheet = reportWB.Sheets("AgilentTCs")
'    time1row = time1row + 6
'    time2row = time2row + 6
'    datasheet.Cells(1, 1) = "Max Temps(F)"
'    datasheet.Cells(2, 1) = "Max Temps(C)"
'    datasheet.Cells(3, 1) = "Max Temps allowed"
'    datasheet.Cells(4, 1) = "Delta"
'
'    'while loop to go through all the TCs that need to be done
'    While TCSheet.Cells(i, 2).Value <> ""'''

 '       channelCol = FindChannelColFromAgilentSheet(TCSheet, datasheet, i)
 '       If TCSheet.Cells(i, 3).Value <> "" Then 'checking to see if it has a name
 '               '"=max('" & agDatasheet.Name & "'!R" & time1row & "C" & channelCol & ":R" & time2row & "C" & channelCol & ")" 'max temps from data
 '               TCSheet.Cells(i, sheetOffset).FormulaR1C1 = "=max('" & datasheet.Name & "'!R" & time1row & "C" & channelCol & ":R" & time2row & "C" & channelCol & ")" 'max temps afrom data
 '               '"=average(LARGE('" & agDatasheet.Name & "'!R" & time1row & "C" & channelCol & ":R" & time2row & "C" & channelCol & ", {5,6,7,8,9,10}))" 'Average Top 5-10 temps (F)
 '               TCSheet.Cells(i, sheetOffset + 1).FormulaR1C1 = "=average(LARGE('" & datasheet.Name & "'!R" & time1row & "C" & channelCol & ":R" & time2row & "C" & channelCol & ", {5,6,7,8,9,10}))" 'Average Top 5-10 temps (F)
 '               '=abs(R[0]C[-2]-R[0]C[-1])"
  '              TCSheet.Cells(i, sheetOffset + 2).FormulaR1C1 = "=abs(R[0]C[-2]-R[0]C[-1])" 'Delta
  '
  ''
  '
  '      i = i + 1
  '  Wend'

'End If




'reportWB.Save

End Sub

Function CheckforAgilentData()
'Agilent Data need to check if there is a sheet named this

For i = 1 To Sheets.Count
    If Sheets(i).Name = "AgilentData" Then
        CheckforAgilentData = True
        Exit Function
    End If
Next i

CheckforAgilentData = False

End Function

Function CopyData() 'not being used
'function to quickly pull the data from the file, copy and close the original
'at this point will use the actual data file, would be safer than using the test report being generated
Dim rawdatasheetwb As Workbook
Dim datasheetwb As Workbook

Dim rawdatasheetname As String

With Application.FileDialog(msoFileDialogFilePicker)
    .Title = "Please pick raw data file to copy out the data"
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
    
    rawdatasheetwb.Sheets(1).Copy
    Set datasheetwb = ActiveWorkbook
    
    rawdatasheetwb.Close
    
    Set CopyData = datasheetwb
End If

End Function


Function CheckDataAq() 'not being used
'function to determine the data aq system
If Range("K22").Value = "" Then
    MsgBox ("Data system needs to be inputed")
    ExitSub = True
End If
CheckDataAq = Range("K22").Value
'returning a string

End Function

Function GetTimeColumn(DataAq)
'function to look at the sofware used and give back a tine column number
If DataAq = "Tescor" Then
    GetTimeColumn = 3
    DataSheetOffset = 40
ElseIf DataAq = "Aglient 34980A" Then
    GetTimeColumn = 1
    DataSheetOffset = 1
Else
    TimeColumn = Application.inputbox("What is the Number correlating to the time column for this data?", "Time Column")
    DataSheetOffset = TimeColumn + 1 'can add a way later to get more data but right now only takes columns after the time column
    GetTimeColumn = TimeColumn
End If

End Function
Function findNextCol(wkst As Worksheet) 'hopefuly works with both tescor and aglient
looping = True
looping2 = True
i = 1
J = 2
'need to check the sheet for a channel name and then see how far over that goes
While looping
    If wkst.Cells(1, i).Value = "Channel Name" Then 'i on column
        While looping2
            If wkst.Cells(J, i) <> "" Then 'j on the row
            'found a channel name
            findNextCol = wkst.Cells(J, Columns.Count).End(xlToLeft).Column + 1
                Exit Function
            End If
            
            J = J + 1
            If J = 10000 Then
                looping2 = False 'breaking out at 10000- will never have that many channels
                Exit Function
            End If
        Wend
      
    End If
    
    i = i + 1
    If i = 10000 Then
        looping = False
        Exit Function
    End If
Wend

MsgBox ("could not find next empty column") 'should only trigger if it didn't find anything
'return int for col

End Function
Sub CheckStability()
'checking stability at the end of the data set, at 10min before, and at 20min before

Dim datasheetwb As Workbook
Dim datasheet As Worksheet
Set reportmaker = ActiveWorkbook
Dim k As Integer
Dim channelCol As Integer
'pull up data source
 Set datasheetwb = GetData("Select Report to check for stability on")
Application.ScreenUpdating = False

Dim dataAqArray(10) As String
dataAqArray(1) = "Tescor"

If CheckforAgilentData = True Then 'can add more later and return an array
'If reportmaker.Sheets("Consolidate Reports").Range("K13").Value = "Yes" Then
    dataAqArray(2) = "Aglient 34980A" 'can add more types of data later
    End If
    
Dim i As Integer
Dim coloffset As Integer
i = 1
Dim TCsheet As Worksheet

Dim channelNameCol As Integer
'sheetOffset = 11 ' can move the numbers over to the right as needed


While dataAqArray(i) <> ""

    'checking from a premade report
    DataAq = dataAqArray(i) 'everything is already set up for DataAq
    If DataAq = "Tescor" Then
        Set datasheet = datasheetwb.Sheets("Measured Data")
        Set TCsheet = datasheetwb.Sheets("Tescor TCs")
        coloffset = 19
        channelNameCol = 2
        
    Else
        Set datasheet = datasheetwb.Sheets("AgilentData")
        Set TCsheet = datasheetwb.Sheets("AgilentTCs")
        coloffset = 1
        channelNameCol = 3
    End If
    TCsheet.Activate
    sheetOffset = findNextCol(TCsheet) 'finds next free column for you
    TCsheet.Cells(1, sheetOffset).Value = "0min (C)"
    TCsheet.Cells(1, sheetOffset + 1).Value = "15min (C)"
    TCsheet.Cells(1, sheetOffset + 2).Value = "30min (C)"
    TCsheet.Cells(1, sheetOffset + 3).Value = "Difference"
    'TCSheet.Cells(1, sheetOffset + 4).Value = "Max Temp Rise 0-60min(C)"

    datasheet.Activate
    
    timeCol = GetTimeColumn(DataAq)
    'Dim lastRow As Integer
    'Dim lastTime As String
    Dim endtime As String
    Dim userinput As String
    
    If dataAqArray(i) = "Tescor" Then
        
        userinput = Application.inputbox("Input end time of stabilization, if you want it to pull from the end, leave at 0. Input in military time (hh:mm)", "Time input", 0) 'like for examples 14:10
            If userinput = "" Then Exit Sub
        datasheet.Activate
        
        If userinput = "0" Then
            endtime = ConvertTimeFormat(datasheet.Cells(FindLastRow(), 3).Value, DataAq)
        Else
            endtime = ConvertStringtoDateTime(userinput, datasheetwb)
        End If

        'lastRow = FindLastRow()
        'lastTime = Cells(lastRow, TimeCol).Value
        
        'endtime = ConvertTimeFormat(lastTime, DataAq)
    End If
    'goes back 60min - puts times in array according to min
        'then puts row numbers in array according to the time

    Dim timeStringArray(60) As String
    For k = 0 To 60
        timeStringArray(k) = GetTimefromSubtractedMin(k, endtime)
        'so 1 will be 1 min before the end of the test
    Next k
    
    Dim timeRowArray(60) As Integer
    For k = 0 To 60
        timeRowArray(k) = FindTimeRowfromTimeString(timeStringArray(k), FindLastRow())
        'so now 1 will be the row 1min from end of the test
    Next k
    Dim J As Integer
    J = 2
    
    
    While TCsheet.Cells(J, 1) <> ""
    
    'sheetOffset = 11 ' can move the numbers over to the right as needed
    If DataAq = "Aglient 34980A" Then
        'need to search for the channel number on the TC sheet- make it a function to return the channelCol
        channelCol = FindChannelColFromAgilentSheet(TCsheet, datasheet, J)
    Else
        channelCol = coloffset + TCsheet.Cells(J, 1) 'number from column 1 should be the channel on all TC sheet lists
    End If
    
      If TCsheet.Cells(J, channelNameCol) <> "" Then

          TCsheet.Cells(J, sheetOffset).FormulaR1C1 = "=convert('" & datasheet.Name & "'!R" & timeRowArray(30) & "C" & channelCol & ",""F"",""C"")"
          TCsheet.Cells(J, sheetOffset + 1).FormulaR1C1 = "=convert('" & datasheet.Name & "'!R" & timeRowArray(15) & "C" & channelCol & ",""F"",""C"")"
          TCsheet.Cells(J, sheetOffset + 2).FormulaR1C1 = "=convert('" & datasheet.Name & "'!R" & timeRowArray(0) & "C" & channelCol & ",""F"",""C"")"
          TCsheet.Cells(J, sheetOffset + 3).FormulaR1C1 = "=max(R[0]C[-3]:R[0]C[-1])-min(R[0]C[-3]:R[0]C[-1])"
          If TCsheet.Cells(J, sheetOffset + 3).Value > 3 Then
            TCsheet.Cells(J, sheetOffset + 3).Interior.ColorIndex = 3
          End If

          'TCSheet.Cells(j, sheetOffset + 4).NumberFormat = "0.0000"
          'TCSheet.Cells(j, sheetOffset + 4).FormulaR1C1 = "=convert(large('" & datasheet.Name & "'!R" & timeRowArray(60) & "C" & channelCol & _
          ":R" & timeRowArray(0) & "C" & channelCol & ",{1}),""F"",""C"")-convert('" & datasheet.Name & "'!R" & timeRowArray(60) & "C" & channelCol & ",""F"",""C"")"
            'If TCSheet.Cells(j, sheetOffset + 4).Value > 10 Then
            '   TCSheet.Cells(j, sheetOffset + 4).Interior.ColorIndex = 6
           ' End If
        End If
        J = J + 1
    Wend
    
    ''adding in temps for DB,WB, voltage,ampre, watts,Hi pressure and low pressure which are needed on datasheets
    ''only needed with TESCOR
    'If DataAq = "Aglient 34980A" Then
    
    
    
    i = i + 1
    
    
Wend





Application.ScreenUpdating = True

datasheetwb.Save
End Sub
Function ConvertFtoC(number As Double)
ConvertFtoC = (number - 32) * (5 / 9)
End Function


'End Function
Function GetData(message As String)
'function to quickly pull the data from the file, copy and close the original
'at this point will use the actual data file, would be safer than using the test report being generated

Dim rawdatasheetname As String

With Application.FileDialog(msoFileDialogFilePicker)
    .Title = message
    .AllowMultiSelect = False
    If .Show = True Then
        rawdatasheetname = .SelectedItems(1)
    End If
End With

If ExitSub = False Then
    Workbooks.Open (rawdatasheetname)
    Set GetData = ActiveWorkbook
End If

End Function

Function FindLastRow()

FindLastRow = Cells(Rows.Count, timeCol).End(xlUp).Row

End Function

Function FindTimeRowfromTimeString(time As String, startrow As Integer)
'parameter should be in a basic "mm/dd/yyy hh:mm:ss"
'returning the time row via an integer
rowcounter = startrow
looping = True
Do While looping
    T1 = ConvertTimeFormat(Cells(rowcounter, timeCol).Value, DataAq)
    
    T2 = ConvertTimeFormat(Cells(rowcounter - 1, timeCol).Value, DataAq)
    'need a break for when the tesor data is started before the agilent data
    If T1 = "" Or T2 = "" Then
        'if T1 = "" then it is at the end, T2 at the beginneg, either way just taking that rowcounter as the right row should work
        FindTimeRowfromTimeString = rowcounter
        Exit Function
    Else
        If IsDate(T1) = False Or IsDate(T2) = False Then
        'if either are not dates then next part will not work
            FindTimeRowfromTimeString = rowcounter 'again just take the number that it is on
            Exit Function
        Else
        x = DateDiff("s", time, T1)
        y = DateDiff("s", time, T2)
            If (DateDiff("s", time, T1) >= 0 And DateDiff("s", time, T2) < 0) Or (DateDiff("s", time, T2) < 0 And DateDiff("s", time, T1) < 0) Then
                looping = False
            Else
                rowcounter = rowcounter - 1
            End If
        End If
    End If
Loop

FindTimeRowfromTimeString = rowcounter


End Function


Function ConvertTimeFormat(time As String, DataAq)
'needs to be configured for all data sets
If time <> "" Then
    If DataAq = "Tescor" Then
    
        Sec = CInt(Right(time, 2))
        Min = CInt(Mid(time, 15, 2))
        hr = CInt(Mid(time, 12, 2))
        Days = CInt(Mid(time, 9, 2))
        Mon = CInt(Mid(time, 6, 2))
        Years = CInt(Left(time, 4))
        
        ConvertTimeFormat = DateSerial(Years, Mon, Days) + TimeSerial(hr, Min, Sec)
    Else
        ConvertTimeFormat = time
    End If
End If
End Function

Function GetTimefromSubtractedMin(userinput As Integer, time As String)
'1/24 = 1hr
'1/1440 = min
'1/86400 = sec

GetTimefromSubtractedMin = CDate(time) - (userinput / 1440)
End Function

Function FindChannelColFromAgilentSheet(TCsheet As Worksheet, datasheet As Worksheet, J As Integer)
Dim channelNumber As String
Dim k As Integer
Dim loopingChannel As Boolean
loopingChannel = True
Dim channelrow As Integer
For k = 1 To 50
    If datasheet.Cells(k, 1).Value = "Time" Then
         channelrow = k
    End If
Next k
p = TCsheet.Name
q = datasheet.Name
k = 1
While loopingChannel
    channelNumber = TCsheet.Cells(J, 2).Value
    If channelNumber = TCsheet.Cells(J, 3).Value <> "" Then
        If Left(datasheet.Cells(channelrow, k).Value, 4) = channelNumber Then
            FindChannelColFromAgilentSheet = k
            Exit Function
            loopingChannel = False 'break out of loop
        End If
    End If
    k = k + 1
    If k = 1000 Then
        loopingChannel = False
    End If
Wend

FindChannelColFromAgilentSheet = k 'in case it breaks out of the loop accidently


End Function



