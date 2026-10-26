Attribute VB_Name = "ChecksCl19"
Public col() As Variant
Public ExitSub As Boolean
Public timeCol As Integer
Public DataAq As String
Public DataSheetOffset As Integer
Public reportmaker As Workbook
Sub test()
Dim wb As Workbook
Set wb = ActiveWorkbook
z = ConvertStringtoDateTime("11:15", wb)


End Sub

Function ConvertStringtoDateTime(userinput As String, wb As Workbook)
'taking in a string that is hh:mm miltary time and translating it to a date time to function later
DataAq = "Tescor"
'find date from wb "Measured Data" and then pull together
desiredtimestamp = CDate(userinput)
sheetDate = ConvertTimeFormat(wb.Sheets("Measured Data").Range("C4").Value)
'y = CLng(sheetDate)
'Z = CDate(y)
'a = CDate(Format(sheetDate, "yyyy-mm-dd"))
desiredDate = CDate(Format(sheetDate, "yyyy-mm-dd"))
Dim combinedDateTime As String
combinedDateTime = desiredDate + desiredtimestamp
ConvertStringtoDateTime = combinedDateTime


End Function

Sub CheckCl19104DV3()
'need to have the temperature rise checked- should ask the user how far back should the data set be checked
'no longer asking for the time frame to check back on - require start and end time in restriction
'can then check 30min before the start of restriction to check stability
Set reportmaker = ActiveWorkbook
Set reportWB = GetData("Select Report to check average rise temps on")
Dim datasheet As Worksheet
Dim dataAqArray(10) As String
dataAqArray(1) = "Tescor"
Dim datasheetwb As Workbook
Dim userinput As String
Dim Starttime As String
Dim endtime As String
Dim channelCol As Integer

DataAq = "Tescor"
'userinput = Application.inputbox("Input stabilization start time (hh:mm) miltary time", "Time input") 'like for examples 14:10
'    If userinput = "" Then Exit Sub
'starttime = ConvertStringtoDateTime(userinput, reportWB)
userinput = Application.inputbox("Input stabilization end time (hh:mm) miltary time", "Time input") 'like for examples 14:10
    If userinput = "" Then Exit Sub
endtime = ConvertStringtoDateTime(userinput, reportWB)
If CheckforAgilentData = True Then
'If reportmaker.Sheets("Consolidate Reports").Range("K13").Value = "Yes" Then
    dataAqArray(2) = "Aglient 34980A" 'can add more types of data later
    End If
Dim J As Integer
Dim i As Integer
J = 1

Dim TCsheet As Worksheet

While dataAqArray(J) <> ""
    'checking from a premade report
    DataAq = dataAqArray(J) 'everything is already set up for DataAq
    If DataAq = "Tescor" Then
        Set datasheet = reportWB.Sheets("Measured Data")
        Set TCsheet = reportWB.Sheets("Tescor TCs")
        coloffset = 19
        channelNameCol = 2
        maxAllowedTempCol = 3
        timeCol = 3
    Else
        Set datasheet = reportWB.Sheets("AgilentData")
        Set TCsheet = reportWB.Sheets("AgilentTCs")
        coloffset = 1
        channelNameCol = 3
        maxAllowedTempCol = 4
        timeCol = 1
    End If
    'check for max temp rise- per min
    'should add extra sheet when it finds something that is too large
    
    'first find the time rows we should be checking
    Dim firstRowTime As String
    Dim lastRowTime As String
    Dim lastrow As Integer
    
    Dim time1row As Integer
    Dim time2row As Integer
    Dim ReturnMaxTemp As Double
    datasheet.Activate
    'If DataAq = "Tescor" Then

        'lastRow = FindLastRow
        'lastRowTime = ConvertTimeFormat(reportWB.Sheets("Measured Data").Cells(lastRow, 3).Value) '(Worksheets(measuredDataSheetName).Range(FindLastRow)) '2023_10_11_22_01_31

        'time2row = FindTimeRowTimeString(ConvertTimeFormat(endtime), lastRow) 'lastRow 'need int'1628
        'time1row = FindTimeRowfromTimeString(ConvertTimeFormat(starttime), time2row) 'firstRowTime'1007

    'Else
        'TimeCol = 1
        lastrow = FindLastRow
        time2row = FindTimeRowfromTimeString(ConvertTimeFormat(endtime), lastrow)
        
        
        'End If
    
    'Dim timeStampArray() As Variant
    'timeStampArray = CreateArrayofTimeStamps(30, starttime)
    'Dim timeRowArray() As Variant
    'timeRowArray = ReturnArrayofTimeRows(30, timeStampArray)
    Dim k As Integer
    Dim timeStampArray(60) As String
    For k = 0 To 60
        timeStampArray(k) = GetTimefromSubtractedMin(k, endtime) 'starttime is the user inputed time
        'so 1 will be 1 min before the end of the test
        '30 will 30min before the starttime
    Next k
    
    Dim timeRowArray(60) As Integer
    For k = 0 To 60
        timeRowArray(k) = FindTimeRowfromTimeString(timeStampArray(k), FindLastRow())
        'so now 1 will be the row 1min from end of the test
        '30 will 30min before the starttime
    Next k
    
    time1row = timeRowArray(60) 'FindTimeRowfromTimeString(ConvertTimeFormat(starttime), time2row) '9 'will need to tell it what worksheet to go to '3671
    'Timediff = ConvertTimeFormat(endtime) - ConvertTimeFormat(starttime)
    
    sheetOffset = findNextCol(TCsheet) 'finds next free column for you
    
    TCsheet.Cells(1, sheetOffset).Value = "0min"
    TCsheet.Cells(1, sheetOffset + 1).Value = "15min"
    TCsheet.Cells(1, sheetOffset + 2).Value = "30min"
    TCsheet.Cells(1, sheetOffset + 3).Value = "45min"
    TCsheet.Cells(1, sheetOffset + 4).Value = "60min"
    TCsheet.Cells(1, sheetOffset + 5).Value = "Delta"
    TCsheet.Cells(1, sheetOffset + 6).Value = "Max Temp during temp Rise (F)"
    TCsheet.Cells(1, sheetOffset + 7).Value = "Max Temp during temp Rise (C) 19.104DV3"
    
    'should have the times and rows to check back from
    'now need to check all the channels and go through a full loop
    i = 2
    looping = True
    Dim rng As Range
    
    While looping   'check raw data
        If dataAqArray(J) = "Aglient 34980A" Then
             'need to search for the channel number on the TC sheet- make it a function to return the channelCol
             channelCol = FindChannelColFromAgilentSheet(TCsheet, datasheet, i)
        Else
             channelCol = coloffset + TCsheet.Cells(i, 1) 'number from column 1 should be the channel on all TC sheet lists
        End If
        Set rng = datasheet.Range(datasheet.Cells(time1row, channelCol), datasheet.Cells(time2row, channelCol))
         'timeinMin = (CDate(endtime) - CDate(starttime)) * 1440
        If TCsheet.Cells(i, 3).Value <> "" Then 'should
            
            TCsheet.Cells(i, sheetOffset).FormulaR1C1 = "=convert('" & datasheet.Name & "'!R" & timeRowArray(60) & "C" & channelCol & ",""F"",""C"")"
            TCsheet.Cells(i, sheetOffset + 1).FormulaR1C1 = "=convert('" & datasheet.Name & "'!R" & timeRowArray(45) & "C" & channelCol & ",""F"",""C"")"
            TCsheet.Cells(i, sheetOffset + 2).FormulaR1C1 = "=convert('" & datasheet.Name & "'!R" & timeRowArray(30) & "C" & channelCol & ",""F"",""C"")"
            TCsheet.Cells(i, sheetOffset + 3).FormulaR1C1 = "=convert('" & datasheet.Name & "'!R" & timeRowArray(15) & "C" & channelCol & ",""F"",""C"")"
            TCsheet.Cells(i, sheetOffset + 4).FormulaR1C1 = "=convert('" & datasheet.Name & "'!R" & timeRowArray(0) & "C" & channelCol & ",""F"",""C"")"
            TCsheet.Cells(i, sheetOffset + 5).FormulaR1C1 = "=max(R[0]C[-5]:R[0]C[-1])-min(R[0]C[-5]:R[0]C[-1])"
            If TCsheet.Cells(i, sheetOffset + 5).Value > 3 Then
                TCsheet.Cells(i, sheetOffset + 5).Interior.Color = vbRed
            End If
            
            'TCSheet.Cells(i, sheetOffset + 4).Formula2 = "=max(" & rng.Address & ")"
            'TCSheet.Cells(i, sheetOffset + 4).Formula2 = "=(convert(large('" & datasheet.Name & "'!" & rng.Address & ",{1}),""F"",""C"")-convert(min('" & datasheet.Name & "'!" & rng.Address & "),""F"",""C""))/" & timeinMin & ""
            'TCSheet.Cells(i, sheetOffset).Value = ReturnMaxTempRise(time1row, time2row, channelCol, datasheet) 'find max temp and return it
            TCsheet.Cells(i, sheetOffset + 6).Value = "=large('" & datasheet.Name & "'!" & rng.Address & ",{1})"
            'If TCSheet.Cells(i, sheetOffset + 4).Value > 1 Then 'used to be 1.8F for checking F now switched to C and the proper number is 1K/min
            '    TCSheet.Cells(i, sheetOffset + 4).Interior.Color = vbRed
            '    TCSheet.Cells(i, sheetOffset + 5).Interior.Color = vbRed
            'End If
            TCsheet.Cells(i, sheetOffset + 7).Value = "=convert(RC[-1],""F"",""C"")"
            
        End If
      
        If TCsheet.Cells(i, 1).Value = "" Then
                    'stop looping and checking if there is no channel number next to it
                    looping = False
        End If
        
        i = i + 1
    Wend
    
    J = J + 1
Wend
    
reportWB.Save
    
End Sub

Sub checkCl19104DV2()


'need to have the temperature rise checked- should ask the user how far back should the data set be checked
'no longer asking for the time frame to check back on - require start and end time in restriction
'can then check 30min before the start of restriction to check stability
Set reportmaker = ActiveWorkbook
Set reportWB = GetData("Select Report to check average rise temps on")
Dim datasheet As Worksheet
Dim dataAqArray(10) As String
dataAqArray(1) = "Tescor"
Dim datasheetwb As Workbook
Dim userinput As String
Dim Starttime As String
Dim endtime As String
Dim channelCol As Integer

DataAq = "Tescor"
userinput = Application.inputbox("Input restriction start time (hh:mm) miltary time", "Time input") 'like for examples 14:10
    If userinput = "" Then Exit Sub
Starttime = ConvertStringtoDateTime(userinput, reportWB)
userinput = Application.inputbox("Input restriction end time (hh:mm) miltary time", "Time input") 'like for examples 14:10
    If userinput = "" Then Exit Sub
endtime = ConvertStringtoDateTime(userinput, reportWB)
If CheckforAgilentData = True Then
'If reportmaker.Sheets("Consolidate Reports").Range("K13").Value = "Yes" Then
    dataAqArray(2) = "Aglient 34980A" 'can add more types of data later
    End If
Dim J As Integer
Dim i As Integer
J = 1

Dim TCsheet As Worksheet

While dataAqArray(J) <> ""
    'checking from a premade report
    DataAq = dataAqArray(J) 'everything is already set up for DataAq
    If DataAq = "Tescor" Then
        Set datasheet = reportWB.Sheets("Measured Data")
        Set TCsheet = reportWB.Sheets("Tescor TCs")
        coloffset = 19
        channelNameCol = 2
        maxAllowedTempCol = 3
        timeCol = 3
    Else
        Set datasheet = reportWB.Sheets("AgilentData")
        Set TCsheet = reportWB.Sheets("AgilentTCs")
        coloffset = 1
        channelNameCol = 3
        maxAllowedTempCol = 4
        timeCol = 1
    End If
    'check for max temp rise- per min
    'should add extra sheet when it finds something that is too large
    
    'first find the time rows we should be checking
    Dim firstRowTime As String
    Dim lastRowTime As String
    Dim lastrow As Integer
    
    Dim time1row As Integer
    Dim time2row As Integer
    Dim ReturnMaxTemp As Double
    datasheet.Activate
    'If DataAq = "Tescor" Then

        'lastRow = FindLastRow
        'lastRowTime = ConvertTimeFormat(reportWB.Sheets("Measured Data").Cells(lastRow, 3).Value) '(Worksheets(measuredDataSheetName).Range(FindLastRow)) '2023_10_11_22_01_31

        'time2row = FindTimeRowTimeString(ConvertTimeFormat(endtime), lastRow) 'lastRow 'need int'1628
        'time1row = FindTimeRowfromTimeString(ConvertTimeFormat(starttime), time2row) 'firstRowTime'1007

    'Else
        'TimeCol = 1
        lastrow = FindLastRow
        time2row = FindTimeRowfromTimeString(ConvertTimeFormat(endtime), lastrow)
        time1row = FindTimeRowfromTimeString(ConvertTimeFormat(Starttime), time2row) '9 'will need to tell it what worksheet to go to '3671
        
    'End If
    
    'Dim timeStampArray() As Variant
    'timeStampArray = CreateArrayofTimeStamps(30, starttime)
    'Dim timeRowArray() As Variant
    'timeRowArray = ReturnArrayofTimeRows(30, timeStampArray)
    Dim k As Integer
    Dim timeStampArray(60) As String
    For k = 0 To 30
        timeStampArray(k) = GetTimefromSubtractedMin(k, Starttime) 'starttime is the user inputed time
        'so 1 will be 1 min before the end of the test
        '30 will 30min before the starttime
    Next k
    
    Dim timeRowArray(60) As Integer
    For k = 0 To 30
        timeRowArray(k) = FindTimeRowfromTimeString(timeStampArray(k), FindLastRow())
        'so now 1 will be the row 1min from end of the test
        '30 will 30min before the starttime
    Next k
    
    sheetOffset = findNextCol(TCsheet) 'finds next free column for you
    
    TCsheet.Cells(1, sheetOffset).Value = "0min"
    TCsheet.Cells(1, sheetOffset + 1).Value = "15min"
    TCsheet.Cells(1, sheetOffset + 2).Value = "30min"
    TCsheet.Cells(1, sheetOffset + 3).Value = "Delta"
    TCsheet.Cells(1, sheetOffset + 4).Value = "Max Temp Rise over Time(C/min) Cl 19"
    TCsheet.Cells(1, sheetOffset + 5).Value = "Max Temp during temp Rise (F)"
    TCsheet.Cells(1, sheetOffset + 6).Value = "Max Temp during temp Rise (C) 19.104DV2"
    
    'should have the times and rows to check back from
    'now need to check all the channels and go through a full loop
    i = 2
    looping = True
    Dim rng As Range
    
    While looping   'check raw data
        If dataAqArray(J) = "Aglient 34980A" Then
             'need to search for the channel number on the TC sheet- make it a function to return the channelCol
             channelCol = FindChannelColFromAgilentSheet(TCsheet, datasheet, i)
        Else
             channelCol = coloffset + TCsheet.Cells(i, 1) 'number from column 1 should be the channel on all TC sheet lists
        End If
        Set rng = datasheet.Range(datasheet.Cells(time1row, channelCol), datasheet.Cells(time2row, channelCol))
       
        'Set rng = Range(datasheet.Cells(1007, 21), datasheet.Cells(1628, 21))

         'Set rng = datasheet.Range(R(time1row) C(channelCol)), Cells(time2row, channelCol))
         timeinMin = (CDate(endtime) - CDate(Starttime)) * 1440
        If TCsheet.Cells(i, 3).Value <> "" Then 'should
            
            TCsheet.Cells(i, sheetOffset).FormulaR1C1 = "=convert('" & datasheet.Name & "'!R" & timeRowArray(30) & "C" & channelCol & ",""F"",""C"")"
            TCsheet.Cells(i, sheetOffset + 1).FormulaR1C1 = "=convert('" & datasheet.Name & "'!R" & timeRowArray(15) & "C" & channelCol & ",""F"",""C"")"
            TCsheet.Cells(i, sheetOffset + 2).FormulaR1C1 = "=convert('" & datasheet.Name & "'!R" & timeRowArray(0) & "C" & channelCol & ",""F"",""C"")"
            TCsheet.Cells(i, sheetOffset + 3).FormulaR1C1 = "=max(R[0]C[-3]:R[0]C[-1])-min(R[0]C[-3]:R[0]C[-1])"
            If TCsheet.Cells(i, sheetOffset + 3).Value > 3 Then
                TCsheet.Cells(i, sheetOffset + 3).Interior.Color = vbRed
            End If
            
            'TCSheet.Cells(i, sheetOffset + 4).Formula2 = "=max(" & rng.Address & ")"
            TCsheet.Cells(i, sheetOffset + 4).Formula2 = "=(convert(large('" & datasheet.Name & "'!" & rng.Address & ",{1}),""F"",""C"")-convert(min('" & datasheet.Name & "'!" & rng.Address & "),""F"",""C""))/" & timeinMin & ""
            'TCSheet.Cells(i, sheetOffset).Value = ReturnMaxTempRise(time1row, time2row, channelCol, datasheet) 'find max temp and return it
            TCsheet.Cells(i, sheetOffset + 5).Value = "=large('" & datasheet.Name & "'!" & rng.Address & ",{1})"
            If TCsheet.Cells(i, sheetOffset + 4).Value > 1 Then 'used to be 1.8F for checking F now switched to C and the proper number is 1K/min
                TCsheet.Cells(i, sheetOffset + 4).Interior.Color = vbRed
                TCsheet.Cells(i, sheetOffset + 5).Interior.Color = vbRed
            End If
            TCsheet.Cells(i, sheetOffset + 6).Value = "=convert(RC[-1],""F"",""C"")"
            
        End If
      
        If TCsheet.Cells(i, 1).Value = "" Then
                    'stop looping and checking if there is no channel number next to it
                    looping = False
        End If
        
        i = i + 1
    Wend
    
    J = J + 1
Wend

reportWB.Save

End Sub

Sub checkingarrays()

Dim testsarray() As String
'testsarray = CreateArrayofTimeStamps(5, "2/15/2024  8:39:48 AM")
If IsDate("2/15/2024  8:39:48 AM") Then
    MsgBox ("")
End If
Dim testRowArray() As Variant
testRowArray = ReturnArrayofTimeRows(5, testsarray)



End Sub


Public Function CreateArrayofTimeStamps(i As Integer, endtime As String)
Dim k As Integer
Dim timeStampArray() As String
Dim looping As Boolean
k = 0
While k <= i
        timeStampArray(k) = GetTimefromSubtractedMin(k, endtime)
        'so 1 will be 1 min before the end of the test
        k = k + 1
Wend
CreateArrayofTimeStamps = timeStampArray
    
End Function
    
Public Function ReturnArrayofTimeRows(i As Integer, ByRef timeStringArray As Variant)
Dim k As Integer
Dim timeRowArray() As Variant
    For k = 0 To i
        timeRowArray(k) = FindTimeRowfromTimeString(timeStringArray(k), FindLastRow())
        'so now 1 will be the row 1min from end of the test
    Next k
ReturnArrayofTimeRows = timeRowArray
End Function


Function GetDataCopy()
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
    
    Set GetData = datasheetwb
End If

End Function

Function CheckDataAq()
'function to determine the data aq system
If Range("B5").Value = "" Then
    MsgBox ("Data system needs to be inputed")
    ExitSub = True
End If
CheckDataAq = Range("B5").Value
'returning a string

End Function

Function GetTimeColumn(DataAq)
'function to look at the sofware used and give back a tine column number
If DataAq = "Tescor" Then
    GetTimeColumn = 3
    DataSheetOffset = 19
ElseIf DataAq = "Aglient 34980A" Then
    GetTimeColumn = 1
    DataSheetOffset = 1
Else
    TimeColumn = Application.inputbox("What is the Number correlating to the time column for this data?", "Time Column")
    DataSheetOffset = TimeColumn + 1 'can add a way later to get more data but right now only takes columns after the time column
    GetTimeColumn = TimeColumn
End If

End Function

Function ReturnMaxTempRise(time1row As Integer, time2row As Integer, channelCol As Integer, datasheet As Worksheet)
MaxRiseAmount = 0
'Time1Row will be the starting time in the data
'Time2Row will be the ending time

Do While time2row > time1row 'start with the last row to be checked, and work back until it is at Time1Row - updating the Time2Row as you go
    'check temps between every min of testing
    
    time2 = ConvertTimeFormat(Cells(time2row, timeCol))
    Dim timeMinus1min As String
    timeMinus1min = GetTimefromSubtractedMin(1, ConvertTimeFormat(datasheet.Cells(time2row, timeCol).Value))
    timeRowMinus1min = FindTimeRowfromTimeString(timeMinus1min, time2row)
    time1 = ConvertTimeFormat(Cells(timeRowMinus1min, timeCol))
    
    Timediff = DateDiff("s", time1, time2) / 60 'should give updated amount of min if not 60sec
    
        tempdiff = Cells(time2row, channelCol).Value - Cells(timeRowMinus1min, channelCol).Value
        
        RiseAmount = (tempdiff / Timediff)

            If RiseAmount > MaxRiseAmount Then
                MaxRiseAmount = RiseAmount
            End If
        If RiseAmount > 1.8 Then '/ Sec
            Cells(time2row, 1).EntireRow.Interior.Color = vbYellow
        End If
    'Loop
    
    time2row = time2row - 1
Loop
ReturnMaxTempRise = MaxRiseAmount
End Function

Sub CheckTempRiseFullReport()
'need to have the temperature rise checked- should ask the user how far back should the data set be checked
'no longer asking for the time frame to check back on - require start and end time in restriction
'can then check 30min before the start of restriction to check stability
Set reportmaker = ActiveWorkbook
Set reportWB = GetData("Select Report to check max rise temps on")
'DataAq = CheckDataAq()
'    If ExitSub = True Then Exit Sub
'Dim TimeCol As Integer
'TimeCol = GetTimeColumn(DataAq)
Dim datasheet As Worksheet
Dim dataAqArray(10) As String
dataAqArray(1) = "Tescor"
Dim datasheetwb As Workbook


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
         timeCol = 3
        'channelCol = agTCSheet.Cells(channelcounter, 3).Row 'agDatasheet.Name
        
    Else
        Set datasheet = reportWB.Sheets("AgilentData")
        Set TCsheet = reportWB.Sheets("AgilentTCs")
        coloffset = 1
        channelNameCol = 3
        maxAllowedTempCol = 4
        timeCol = 1
        'channelCol = agTCSheet.Cells(channelcounter, 3).Row 'agDatasheet.Name
    End If
    'check for max temp rise- per min
    'should add extra sheet when it finds something that is too large
    
    'first find the time rows we should be checking
    Dim firstRowTime As String
    Dim lastRowTime As String
    Dim lastrow As Integer
    'firstRowTime = ConvertTimeFormat(reportWB.Sheets("Measured Data").Range("C4").Value, "Tescor") '(Worksheets(measuredDataSheetName).Range("C4").Value) '2023_10_11_19_01_22
    'lastRow = FindLastRow
    'lastRowTime = ConvertTimeFormat(reportWB.Sheets("Measured Data").Cells(lastRow, 3).Value, "Tescor") '(Worksheets(measuredDataSheetName).Range(FindLastRow)) '2023_10_11_22_01_31
    
    Dim time1row As Integer
    Dim time2row As Integer
    Dim ReturnMaxTemp As Double
    datasheet.Activate
    'If DataAq = "Tescor" Then
    '    TimeCol = 3
        'firstRowTime = ConvertTimeFormat(reportWB.Sheets("Measured Data").Range("C4").Value, "Tescor") '(Worksheets(measuredDataSheetName).Range("C4").Value) '2023_10_11_19_01_22
    
        lastrow = FindLastRow
        lastRowTime = ConvertTimeFormat(reportWB.Sheets("Measured Data").Cells(lastrow, 3).Value, "Tescor") '(Worksheets(measuredDataSheetName).Range(FindLastRow)) '2023_10_11_22_01_31
        'need to find the firstRowTime from subtracting the amount from the lastRow Time
        Dim userinput As Integer
        userinput = CInt(Application.inputbox("How many minutes do you want to check back from the end? Whole numbers only please"))
        firstRowTime = FindTimeRowfromTimeString(GetTimefromSubtractedMin(userinput, lastRowTime), lastrow)
        time2row = lastrow 'need int
        time1row = firstRowTime
        'GetTimefromSubtractedMin(userinput As Integer, time As String)
        'time2row = FindTimeRowfromTimeString(lastRowTime, lastRow)
        'time1row = FindTimeRowfromTimeString(ConvertTimeFormat(firstRowTime, DataAq), time2row) '9 'will need to tell it what worksheet to go to '3671
    'Else
        timeCol = 1
        lastrow = FindLastRow
        time2row = FindTimeRowfromTimeString(lastRowTime, lastrow)
        time1row = FindTimeRowfromTimeString(ConvertTimeFormat(firstRowTime, DataAq), time2row) '9 'will need to tell it what worksheet to go to '3671
        
    End If
    
    sheetOffset = findNextCol(TCsheet) 'finds next free column for you
    TCsheet.Cells(1, sheetOffset).Value = "Max Temp Rise (C/min)"
    
    'should have the times and rows to check back from
    'now need to check all the channels and go through a full loop
    i = 2
    While looping   'check raw data
        If dataAqArray(J) = "Aglient 34980A" Then
             'need to search for the channel number on the TC sheet- make it a function to return the channelCol
             channelCol = FindChannelColFromAgilentSheet(TCsheet, datasheet, i)
        Else
             channelCol = coloffset + TCsheet.Cells(i, 1) 'number from column 1 should be the channel on all TC sheet lists
        End If
    If TCsheet.Cells(i, 3).Value <> "" Then 'should
        TCsheet.Cells(i, sheetOffset).Value = ReturnMaxTempRise(time1row, time2row, channelCol, datasheet) 'find max temp and return it
        
    End If
    
    If TCsheet.Cells(i, 1).Value = "" Then
                'stop looping and checking if there is no channel number next to it
                looping = False
            End If
    
    i = i + 1
    Wend
    
    J = J + 1
Wend

'quickly getting the data from the file bring made- or could be a file that is already done
'Set datasheetwb = GetData()
'    If ExitSub = True Then Exit Sub
'Dim datasheet As Worksheet
'Set datasheet = ActiveSheet'

'Dim lastRow As Integer
'lastRow = FindLastRow()

'Dim lastTime As String
'lastTime = Cells(lastRow, TimeCol).Value
'lastTime = ConvertTimeFormat(lastTime, DataAq)

'Dim timeframe As Integer
'timeframe = CInt(Application.inputbox("How far back do you wish to check (whole numbers only and in min please)?"))

'Dim newtime As String
'newtime = GetTimefromSubtractedMin(timeframe, lastTime)
'userinputrow = FindTimeRowfromTimeString(newtime)

''need to check that rise temp betweem points is not more than 1K/min = 1.8F/min = 0.03F/sec
'col = inputbox() 'getting columns for testing
'Dim tempdiff As Double
'Dim Timediff As Double
'nopass = False
'Dim i As Integer
'i = lastRow
'Dim time2 As String
'Dim time1 As String
'Dim RiseAmount As Double
'Dim MaxRiseAmount As Double
'Dim MaxColRise() As Double
'Dim MaxColTemp() As Double

'MaxRiseAmount = 0
'Do While i >= userinputrow 'cant do the last row since nothing after it
    'check temps on all columns between every row of data point - will need time diff between rows
    
'    time2 = ConvertTimeFormat(Cells(i, TimeCol).Value, DataAq)
'    time1 = GetTimefromSubtractedMin(1, time2)

'    Dim time1row As Integer
'    time1row = FindTimeRowfromTimeString(time1)
'    Z = DateDiff("s", time1, time2)
'    Timediff = DateDiff("s", time1, time2) / 60 'should give updated amount of min if not 60sec
'
'    countercol = 1
'    Do While col(countercol) <> 0
'        tempdiff = Cells(i, col(countercol)).Value - Cells(time1row, col(countercol)).Value
'        RiseAmount = (tempdiff / Timediff)
 ''       If countercol = 1 Then
 '           MaxColRise(countercol) = RiseAmount
 '           MaxColTemp(countercol) = Cells(i, col(countercol)).Value
 '       Else
 '           If MaxColRise(countercol) < RiseAmount Then
 '               MaxColRise(countercol) = RiseAmount
  '          End If
  '          If MaxColTemp(countercol) < Cells(i, col(countercol)).Value Then
  '              MaxColTemp(countercol) = Cells(i, col(countercol)).Value
  ''          End If
  '      End If
  '          If RiseAmount > MaxRiseAmount Then
  '              MaxRiseAmount = RiseAmount
  '          End If
  '      If RiseAmount > 1.8 Then '/ Sec
  '          Cells(i, 1).EntireRow.Interior.Color = vbYellow
   '         If nopass = False Then
   '
   ''             Workbooks.Add
   '             Dim checksheetwb As Workbook
   '             Set checksheetwb = ActiveWorkbook
   '             Dim checksheet As Worksheet
   '             Set checksheet = ActiveSheet
   '
   '             lastCheckSheetrow = FindLastRow()
    '            datasheetwb.Activate 'go back to data
    '        End If
    ''        checksheet.Rows(lastCheckSheetrow).Value = datasheet.Rows(i).Value
    '        lastCheckSheetrow = lastCheckSheetrow + 1
    '        nopass = True
    '    End If
    '    countercol = countercol + 1
    'Loop
    'i = i - 1
'Loop

'If nopass = True Then
'    MsgBox ("Larger rise than 1K/min (1.8F/min) with max rise of " & MaxRiseAmount & "F/min")
'Else
'    MsgBox ("Passed with Max Rise of " & MaxRiseAmount & "F/min")
'End If



End Sub

Sub CheckTempRise()
'need to have the temperature rise checked- should ask the user how far back should the data set be checked

DataAq = CheckDataAq()
    If ExitSub = True Then Exit Sub
'Dim TimeCol As Integer
timeCol = GetTimeColumn(DataAq)

Dim datasheetwb As Workbook
'quickly getting the data from the file bring made- or could be a file that is already done
Set datasheetwb = GetDataCopy()
    If ExitSub = True Then Exit Sub
Dim datasheet As Worksheet
Set datasheet = ActiveSheet

Dim lastrow As Integer
lastrow = FindLastRow()

Dim lastTime As String
lastTime = Cells(lastrow, timeCol).Value
lastTime = ConvertTimeFormat(lastTime, DataAq)

Dim timeframe As Integer
timeframe = CInt(Application.inputbox("How far back do you wish to check (whole numbers only and in min please)?"))

Dim newtime As String
newtime = GetTimefromSubtractedMin(timeframe, lastTime)
userinputrow = FindTimeRowfromTimeString(newtime)

'need to check that rise temp betweem points is not more than 1K/min = 1.8F/min = 0.03F/sec
col = inputbox() 'getting columns for testing
Dim tempdiff As Double
Dim Timediff As Double
nopass = False
Dim i As Integer
i = lastrow
Dim time2 As String
Dim time1 As String
Dim RiseAmount As Double
Dim MaxRiseAmount As Double
Dim MaxColRise() As Double
Dim MaxColTemp() As Double

MaxRiseAmount = 0
Do While i >= userinputrow 'cant do the last row since nothing after it
    'check temps on all columns between every row of data point - will need time diff between rows
    
    time2 = ConvertTimeFormat(Cells(i, timeCol).Value, DataAq)
    time1 = GetTimefromSubtractedMin(1, time2)

    Dim time1row As Integer
    time1row = FindTimeRowfromTimeString(time1)
    z = DateDiff("s", time1, time2)
    Timediff = DateDiff("s", time1, time2) / 60 'should give updated amount of min if not 60sec
    
    countercol = 1
    Do While col(countercol) <> 0
        tempdiff = Cells(i, col(countercol)).Value - Cells(time1row, col(countercol)).Value
        RiseAmount = (tempdiff / Timediff)
        If countercol = 1 Then
            MaxColRise(countercol) = RiseAmount
            MaxColTemp(countercol) = Cells(i, col(countercol)).Value
        Else
            If MaxColRise(countercol) < RiseAmount Then
                MaxColRise(countercol) = RiseAmount
            End If
            If MaxColTemp(countercol) < Cells(i, col(countercol)).Value Then
                MaxColTemp(countercol) = Cells(i, col(countercol)).Value
            End If
        End If
            If RiseAmount > MaxRiseAmount Then
                MaxRiseAmount = RiseAmount
            End If
        If RiseAmount > 1.8 Then '/ Sec
            Cells(i, 1).EntireRow.Interior.Color = vbYellow
            If nopass = False Then
                
                Workbooks.Add
                Dim checksheetwb As Workbook
                Set checksheetwb = ActiveWorkbook
                Dim checksheet As Worksheet
                Set checksheet = ActiveSheet
                
                lastCheckSheetrow = FindLastRow()
                datasheetwb.Activate 'go back to data
            End If
            checksheet.Rows(lastCheckSheetrow).Value = datasheet.Rows(i).Value
            lastCheckSheetrow = lastCheckSheetrow + 1
            nopass = True
        End If
        countercol = countercol + 1
    Loop
    i = i - 1
Loop

If nopass = True Then
    MsgBox ("Larger rise than 1K/min (1.8F/min) with max rise of " & MaxRiseAmount & "F/min")
Else
    MsgBox ("Passed with Max Rise of " & MaxRiseAmount & "F/min")
End If



End Sub

Sub CheckStabilityQuick()
'checking stability at the end of the data set, at 10min before, and at 20min before

If ActiveSheet.Range("F5").Value = "Raw Data" Then
    'this pathway if checking the data on the fly and need to pull the data quickly
    DataAq = CheckDataAq()
        If ExitSub = True Then Exit Sub
    'Dim TimeCol As Integer
    timeCol = GetTimeColumn(DataAq)
    
    Dim datasheetwb As Workbook
    'quickly getting the data from the file bring made- or could be a file that is already done
    Set datasheetwb = GetData()
        If ExitSub = True Then Exit Sub
    
    Dim datasheet As Worksheet
Else
    'this pathway if pulling from the report section- how will you select the data?
    'Aglient Data '=tab name
    
    
End If


Dim lastrow As Integer
lastrow = FindLastRow()

Dim lastTime As String
lastTime = Cells(lastrow, timeCol).Value

Dim lastSec As Integer
Dim lastMin As Integer

lastTime = ConvertTimeFormat(lastTime, DataAq)

Dim TenMinTime As String
TenMinTime = GetTimefromSubtractedMin(10, lastTime)
Dim TwentyMinTime As String
TwentyMinTime = GetTimefromSubtractedMin(20, lastTime)

Dim TenMinRow As Integer
TenMinRow = FindTimeRowfromTimeString(TenMinTime)

Dim twentyMinRow  As Integer
twentyMinRow = FindTimeRowfromTimeString(TwentyMinTime)

If MaxDifference(TenMinRow, twentyMinRow, lastrow) < 5.4 Then
    MsgBox ("passed")
Else
    MsgBox ("Did not pass with max difference of: " & MaxDiff)
End If

'MsgBox (counter)


End Sub


Function inputbox()

'input box that can take in channels grouped like 3-7, 8,9,10- 21
'needs take in a string and ignore spaces
'use commas as indicators for new channels
'and dashes for groups of channels
Dim Colstring As Variant
retryinput:
Colstring = Application.inputbox("Input channel numbers. Need to be inputed as #-# with commas in between or singular numbers. Ex; 4-5, 6, 8,9-12")

'take out spaces from string
Colstring = Replace(Colstring, " ", "")
i = 1
Dim midarray(500) As String
'deconstruct the rest into smaller bites- probably need another array to handle
Do While Len(Colstring) > 0
    
    midarray(i) = Mid(Colstring, InStrRev(Colstring, ",") + 1, Abs(InStrRev(Colstring, ",") - Len(Colstring)))
    Colstring = Replace(Colstring, midarray(i), "", , 1)
    'remove , if there is one
    If Right(Colstring, 1) = "," Then
        Colstring = Left(Colstring, Len(Colstring) - 1)
    End If
    i = i + 1
Loop

Dim colstr(500) As Variant
'each array will need to be processed into the final array

J = 1
k = 1
ArrayBol = True
Dim w As String
Dim Startj As Integer
Dim Endj As Integer

Do While ArrayBol
    If InStrRev(midarray(k), "-") = 0 Then
        If IsNumeric(midarray(k)) = False Then
            MsgBox ("Non-channel NUMBERS detected, please retry input")
            GoTo retryinput
        Else
            colstr(J) = CInt(midarray(k)) + DataSheetOffset
            k = k + 1
            J = J + 1
        End If
    
    Else
        'need to dissect the ##-##
        x = Mid(midarray(k), 1, InStrRev(midarray(k), "-") - 1)
        y = Mid(midarray(k), InStrRev(midarray(k), "-") + 1, Abs(InStrRev(midarray(k), "-") - Len(midarray(k))))
        If IsNumeric(x) = False Or IsNumeric(y) = False Then
            MsgBox ("Non-channel NUMBERS detected, please retry input")
            GoTo retryinput
        Else
            Startj = CInt(x)
            Endj = CInt(y)
            
                For i = Startj To Endj
                   colstr(J) = CInt(i) + DataSheetOffset
                   J = J + 1
                Next i
            k = k + 1
        End If
        
    End If
    
    If midarray(k) = "" Then
        ArrayBol = False
    End If
    
Loop

'need to

inputbox = colstr
'output will be an array will take in numbers and output col(0) = 6, col(1)= 7
'output will need to be integers to use in cells()
'these channel numbers will have to be translated over to the correct column

End Function

Function MaxDifference(TenMinRow As Integer, twentyMinRow As Integer, lastrow As Integer)
Dim col() As Variant

col = inputbox()
counter = 0
For i = 1 To 500
    If col(i) = 0 Or col(i) = "" Then
    Else
        counter = counter + 1
    End If
Next i

Dim Max As Double
Dim Maxi As Double
Dim Mini As Double
Dim Min As Double
Dim Maxdif As Double
Maxdif = 0
Max = 0


For i = 1 To counter
    Maxi = Application.WorksheetFunction.Max(Cells(TenMinRow, col(i)), Cells(twentyMinRow, col(i)), Cells(lastrow, col(i)))
    Mini = Application.WorksheetFunction.Min(Cells(TenMinRow, col(i)), Cells(twentyMinRow, col(i)), Cells(lastrow, col(i)))
    Max = Maxi - Mini
    If Maxdif < Max Then
        Maxdif = Max
    End If
    
Next i

'passes or doesn't pass - need to be able to see the data fully on a new data sheet
Dim datasheet As Workbook
Set datasheet = ActiveWorkbook
Workbooks.Add
Dim checksheet As Workbook
Set checksheet = ActiveWorkbook
Range("A1").Value = "Channels"
Range("A2").Value = "Time Frame"
Range("A3").Value = "20min check"
Range("A4").Value = "10min check"
Range("A5").Value = "0min check"
Range("A6").Value = "Max/channel"
Range("A7").Value = "Difference/channel"
Range("A10").Value = "Max Difference"

Range("A10").Value = "Line Numbers"
Range("A11").Value = twentyMinRow & "= 20min row"
Range("A12").Value = TenMinRow & "= 10min row"
Range("A13").Value = lastrow & "= 0min row"

For i = 1 To counter
    checksheet.ActiveSheet.Cells(1, col(i) - (DataSheetOffset - 1)) = col(i) - DataSheetOffset
    checksheet.ActiveSheet.Cells(3, col(i) - (DataSheetOffset - 1)) = datasheet.ActiveSheet.Cells(twentyMinRow, col(i))
    checksheet.ActiveSheet.Cells(4, col(i) - (DataSheetOffset - 1)) = datasheet.ActiveSheet.Cells(TenMinRow, col(i))
    checksheet.ActiveSheet.Cells(5, col(i) - (DataSheetOffset - 1)) = datasheet.ActiveSheet.Cells(lastrow, col(i))
    checksheet.ActiveSheet.Cells(6, col(i) - (DataSheetOffset - 1)).FormulaR1C1 = "=Max(R[-1]C:R[-3]C)"
    checksheet.ActiveSheet.Cells(7, col(i) - (DataSheetOffset - 1)).FormulaR1C1 = "=Max(R[-2]C:R[-4]C)-Min(R[-2]C:R[-4]C)"
    
Next i
Range("B10").Value = "Max Difference across channels"
Range("C10").Value = Maxdif
MaxDifference = Maxdif

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
    T1 = ConvertTimeFormat(Cells(rowcounter, timeCol).Value)
    
    T2 = ConvertTimeFormat(Cells(rowcounter - 1, timeCol).Value)
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


Function ConvertTimeFormat(time As String)
'needs to be configured for all data sets
If DataAq = "Tescor" And Not IsDate(time) Then

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

End Function

Function GetTimefromSubtractedMin(userinput As Integer, time As String)
'1/24 = 1hr
'1/1440 = min
'1/86400 = sec

GetTimefromSubtractedMin = CDate(time) - (userinput / 1440)


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
x = datasheet.Cells(channelrow, k).Value
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
        MsgBox ("could not find channel or over 1000 channels")
    End If
Wend

FindChannelColFromAgilentSheet = k 'in case it breaks out of the loop accidently


End Function

Function GetData(message As String)
'function to quickly pull the data from the file, copy and close the original
'at this point will use the actual data file, would be safer than using the test report being generated

Dim rawdatasheetname As String

With Application.FileDialog(msoFileDialogFilePicker)
    .Title = message
    .AllowMultiSelect = False
    .InitialFileName = "\\techlabs.appl.ge.com\labdata\AP2\AirConditioning\Zoneline\Reports"
    .Filters.Clear
    If .Show = True Then
        rawdatasheetname = .SelectedItems(1)
    End If
End With

If ExitSub = False Then
    Workbooks.Open (rawdatasheetname)
    Set GetData = ActiveWorkbook
End If

End Function
