Attribute VB_Name = "AP4Checks"
Public ExitSub
Sub ConsolidateReportAP4()
Dim reportmaker As Workbook
Dim rawDataWb As Workbook
Dim TCsheet As Worksheet
Dim rawDataSheet As Worksheet

Set reportmaker = ActiveWorkbook
Set rawDataWb = GetDataV2("Select raw data file")
Set rawDataSheet = rawDataWb.Sheets(1)

'add TC channel sheet to the raw data
reportmaker.Sheets("AP4_Chamber").Copy after:=rawDataWb.Sheets(rawDataWb.Sheets.Count)
Set TCsheet = rawDataWb.Sheets("AP4_Chamber")
rawDataSheet.Name = "RawData"


'add compact graph excel file
Set compactWb = GetDataV2("Select compact report excel file")
Set compactWksht = compactWb.Sheets(1)

compactWksht.Copy before:=rawDataWb.Sheets(1)
compactWb.Close
Set compactWksht = rawDataWb.Sheets(1)
compactWksht.Name = "Report"

If reportmaker.Sheets("AP4_Instructions").Range("H6") = "Yes" Then
    answer = MsgBox("Would you like to add Agilent Data?", vbYesNo)
    If answer = vbYes Then
        'add agilent data
        Set compactWb = GetDataV2("Select agilent data file")
        Set compactWksht = compactWb.Sheets(1)
        
        compactWksht.Copy after:=rawDataWb.Sheets(rawDataWb.Sheets.Count)
        compactWb.Close
        rawDataWb.Sheets(rawDataWb.Sheets.Count).Name = "AgilentData"
        
        reportmaker.Sheets("AgilentTCs").Copy after:=rawDataSheet
    End If
End If
x = rawDataWb.Path + "Report_" + rawDataWb.Name
rawDataWb.SaveAs Filename:=CStr(rawDataWb.Path + "Report_" + rawDataWb.Name), FileFormat:=xlWorkbookDefault

End Sub

Sub Check19_104DV2TestAP4()
Dim reportmaker As Workbook
Dim rawDataWb As Workbook
Dim TCsheet As Worksheet
Dim rawDataSheet As Worksheet

Set reportmaker = ActiveWorkbook
Set rawDataWb = GetDataV2("Select report file")
Set rawDataSheet = rawDataWb.Sheets("RawData")
Set TCsheet = rawDataWb.Sheets("AP4_Chamber")

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
For i = 0 To 100
    If IsDate(rawDataSheet.Cells(channelrow + i, 1)) Then
        datastart = channelrow + i
            Exit For
    End If
Next i

'find last row
lastrow = rawDataSheet.Range("A1000000").End(xlUp).Row

Dim endtime As Date
Dim Starttime As Date

NeedStartTime:

'setting up time stamps to use
userinput = Application.inputbox("Input start of rise time. Input in military time (hh:mm)", "Time input") 'like for examples 14:10
    If userinput = "" Then Exit Sub
    If userinput = "0" Then
        MsgBox ("You cannot use the end of the data set for a start time, Retry")
        GoTo NeedStartTime
    Else
        Starttime = ConvertStringtoDateTime(CStr(userinput), rawDataWb)
    End If
    
userinput = Application.inputbox("Input end time of rise time, if you want it to pull from the end, leave at 0. Input in military time (hh:mm)", "Time input", 0) 'like for examples 14:10
    If userinput = "" Then Exit Sub
    If userinput = "0" Then
        endtime = CDate(rawDataSheet.Cells(lastrow, 1).Value)
    Else
     
        endtime = ConvertStringtoDateTime(CStr(userinput), rawDataWb)
    End If
timeinMin = (CDate(endtime) - CDate(Starttime)) * 1440

Dim sub15min As Date
Dim sub30min As Date

dateSelect = CDate(Format(endtime, "MM-DD-YYYY"))

sub15min = dateSelect + CDate(TimeValue(Starttime) - TimeValue("0:15:00"))
sub30min = dateSelect + CDate(TimeValue(Starttime) - TimeValue("0:30:00"))

'Finding stability before the rise time
    'Cells(Application.Match(CLng(Date), Range("A:A"), 1), "A").Select
    sub15minRow = FindClosestRowtoTimeString(rawDataSheet, sub15min, datastart)
    sub30minRow = FindClosestRowtoTimeString(rawDataSheet, sub30min, datastart)
    sub0minRow = FindClosestRowtoTimeString(rawDataSheet, Starttime, datastart)

    endrow = FindClosestRowtoTimeString(rawDataSheet, endtime, datastart)
    
    'setting up columns
    nextCol = TCsheet.Cells(1, Columns.Count).End(xlToLeft).Column + 1
    TCsheet.Cells(1, nextCol).Value = "0min (C)"
    TCsheet.Cells(1, nextCol + 1).Value = "15min (C)"
    TCsheet.Cells(1, nextCol + 2).Value = "30min (C)"
    TCsheet.Cells(1, nextCol + 3).Value = "Difference"

'inputing rise time column names on sheet
TCsheet.Cells(1, nextCol + 4).Value = "Max Temp Rise over Time(C/min) Cl 19"
TCsheet.Cells(1, nextCol + 5).Value = "Max Temp during temp Rise (F)"
TCsheet.Cells(1, nextCol + 6).Value = "Max Temp during temp Rise (C) 19.104DV2"

i = 2
Dim rng As Range
'find channel col
'have about 700 col to search
rawDataSheet.Activate
While TCsheet.Cells(i, 2).Value <> ""
    'ex Z_UUT_ID_TC_035
    If TCsheet.Cells(i, 3).Value <> "" Then
        IDOD = TCsheet.Cells(i, 1).Value 'should just be ID or OD
        channelNum = Format(TCsheet.Cells(i, 2).Value, "000")  'needs to be in 3 digit format
        columnName = "Z_UUT_" + IDOD + "_TC_" + channelNum
        Set x = rawDataSheet.Range(Cells(channelrow - 1, 1), Cells(channelrow + 1, 700)).Find(columnName, LookIn:=xlValues)
        channelCol = x.Column
        'get stability temps
        TCsheet.Cells(i, nextCol).FormulaR1C1 = "=convert('" & rawDataSheet.Name & "'!R" & sub30minRow & "C" & channelCol & ",""F"",""C"")"
        TCsheet.Cells(i, nextCol + 1).FormulaR1C1 = "=convert('" & rawDataSheet.Name & "'!R" & sub15minRow & "C" & channelCol & ",""F"",""C"")"
        TCsheet.Cells(i, nextCol + 2).FormulaR1C1 = "=convert('" & rawDataSheet.Name & "'!R" & sub0minRow & "C" & channelCol & ",""F"",""C"")"
        TCsheet.Cells(i, nextCol + 3).FormulaR1C1 = "=max(R[0]C[-3]:R[0]C[-1])-min(R[0]C[-3]:R[0]C[-1])"
        If TCsheet.Cells(i, nextCol + 3).Value > 3 Then
            TCsheet.Cells(i, nextCol + 3).Interior.ColorIndex = 3
        End If
        'get rise time and max temp during rise
        'name range
        Set rng = rawDataSheet.Range(rawDataSheet.Cells(sub0minRow, channelCol), rawDataSheet.Cells(endrow, channelCol))
        TCsheet.Cells(i, nextCol + 4).Formula2 = "=(convert(large('" & rawDataSheet.Name & "'!" & rng.Address & ",{1}),""F"",""C"")-convert(min('" & rawDataSheet.Name & "'!" & rng.Address & "),""F"",""C""))/" & timeinMin & ""
        TCsheet.Cells(i, nextCol + 5).Value = "=large('" & rawDataSheet.Name & "'!" & rng.Address & ",{1})"
            If TCsheet.Cells(i, nextCol + 4).Value > 1 Then 'used to be 1.8F for checking F now switched to C and the proper number is 1K/min
                TCsheet.Cells(i, nextCol + 4).Interior.Color = vbRed
                TCsheet.Cells(i, nextCol + 5).Interior.Color = vbRed
            End If
        TCsheet.Cells(i, nextCol + 6).Value = "=convert(RC[-1],""F"",""C"")"
        
    End If
    i = i + 1
Wend

ActiveWorkbook.Save


End Sub

Sub Check19_104DV3TestAP4()

Dim reportmaker As Workbook
Dim rawDataWb As Workbook
Dim TCsheet As Worksheet
Dim rawDataSheet As Worksheet

Set reportmaker = ActiveWorkbook
Set rawDataWb = GetDataV2("Select report file")
Set rawDataSheet = rawDataWb.Sheets("RawData")
Set TCsheet = rawDataWb.Sheets("AP4_Chamber")

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
For i = 0 To 100
    If IsDate(rawDataSheet.Cells(channelrow + i, 1)) Then
        datastart = channelrow + i
            Exit For
    End If
Next i

'find last row
lastrow = rawDataSheet.Range("A1000000").End(xlUp).Row

Dim endtime As Date
Dim Starttime As Date

userinput = Application.inputbox("Input end time of stablization time, if you want it to pull from the end, leave at 0. Input in military time (hh:mm)", "Time input", 0) 'like for examples 14:10
    If userinput = "" Then Exit Sub
    If userinput = "0" Then
        endtime = CDate(rawDataSheet.Cells(lastrow, 1).Value)
    Else
        endtime = ConvertStringtoDateTime(CStr(userinput), rawDataWb)
    End If
        
dateSelect = CDate(Format(endtime, "MM-DD-YYYY"))
        
Dim sub15min As Date
Dim sub30min As Date
Dim sub45min As Date
Dim sub60min As Date

sub15min = dateSelect + CDate(TimeValue(endtime) - TimeValue("0:15:00"))
sub30min = dateSelect + CDate(TimeValue(endtime) - TimeValue("0:30:00"))
sub45min = dateSelect + CDate(TimeValue(endtime) - TimeValue("0:45:00"))
sub60min = dateSelect + CDate(TimeValue(endtime) - TimeValue("1:00:00"))


'Finding stability before the rise time
    'Cells(Application.Match(CLng(Date), Range("A:A"), 1), "A").Select
    sub15minRow = FindClosestRowtoTimeString(rawDataSheet, sub15min, datastart)
    sub30minRow = FindClosestRowtoTimeString(rawDataSheet, sub30min, datastart)
    sub45minRow = FindClosestRowtoTimeString(rawDataSheet, sub45min, datastart)
    sub60minRow = FindClosestRowtoTimeString(rawDataSheet, sub60min, datastart)
    sub0minRow = FindClosestRowtoTimeString(rawDataSheet, endtime, datastart)

'setting up columns
    nextCol = TCsheet.Cells(1, Columns.Count).End(xlToLeft).Column + 1
    TCsheet.Cells(1, nextCol).Value = "0min (C)"
    TCsheet.Cells(1, nextCol + 1).Value = "15min (C)"
    TCsheet.Cells(1, nextCol + 2).Value = "30min (C)"
    TCsheet.Cells(1, nextCol + 3).Value = "45min (C)"
    TCsheet.Cells(1, nextCol + 4).Value = "60min (C)"
    TCsheet.Cells(1, nextCol + 5).Value = "Delta- technically does not need to be within 3K"
    TCsheet.Cells(1, nextCol + 6).Value = "Max Temp during temp Rise (F)"
    TCsheet.Cells(1, nextCol + 7).Value = "Max Temp during temp Rise (C) 19.104DV3"
    
i = 2
Dim rng As Range
'find channel col
'have about 700 col to search
rawDataSheet.Activate
While TCsheet.Cells(i, 2).Value <> ""
    'ex Z_UUT_ID_TC_035
    If TCsheet.Cells(i, 3).Value <> "" Then
        IDOD = TCsheet.Cells(i, 1).Value 'should just be ID or OD
        channelNum = Format(TCsheet.Cells(i, 2).Value, "000")  'needs to be in 3 digit format
        columnName = "Z_UUT_" + IDOD + "_TC_" + channelNum
        Set x = rawDataSheet.Range(Cells(channelrow - 1, 1), Cells(channelrow + 1, 700)).Find(columnName, LookIn:=xlValues)
        channelCol = x.Column
        'get stability temps
        TCsheet.Cells(i, nextCol).FormulaR1C1 = "=convert('" & rawDataSheet.Name & "'!R" & sub60minRow & "C" & channelCol & ",""F"",""C"")"
        TCsheet.Cells(i, nextCol + 1).FormulaR1C1 = "=convert('" & rawDataSheet.Name & "'!R" & sub45minRow & "C" & channelCol & ",""F"",""C"")"
        TCsheet.Cells(i, nextCol + 2).FormulaR1C1 = "=convert('" & rawDataSheet.Name & "'!R" & sub30minRow & "C" & channelCol & ",""F"",""C"")"
        TCsheet.Cells(i, nextCol + 3).FormulaR1C1 = "=convert('" & rawDataSheet.Name & "'!R" & sub15minRow & "C" & channelCol & ",""F"",""C"")"
        TCsheet.Cells(i, nextCol + 4).FormulaR1C1 = "=convert('" & rawDataSheet.Name & "'!R" & sub0minRow & "C" & channelCol & ",""F"",""C"")"
        TCsheet.Cells(i, nextCol + 5).FormulaR1C1 = "=max(R[0]C[-5]:R[0]C[-1])-min(R[0]C[-5]:R[0]C[-1])"
        If TCsheet.Cells(i, nextCol + 5).Value > 3 Then
            TCsheet.Cells(i, nextCol + 5).Interior.ColorIndex = 3
        End If
        'get rise time and max temp during rise
        'name range
        Set rng = rawDataSheet.Range(rawDataSheet.Cells(sub0minRow, channelCol), rawDataSheet.Cells(sub60minRow, channelCol))
        TCsheet.Cells(i, nextCol + 6).Formula2 = "=large('" & rawDataSheet.Name & "'!" & rng.Address & ",{1})"
        TCsheet.Cells(i, nextCol + 7).Value = "=convert(RC[-1],""F"",""C"")"
        
    End If
    i = i + 1
Wend

ActiveWorkbook.Save


End Sub


Sub CheckStablilityAP4()
Dim reportmaker As Workbook
Dim rawDataWb As Workbook
Dim TCsheet As Worksheet
Dim rawDataSheet As Worksheet
Dim AddAgilent As Boolean 'flag if we are adding agilent data
Dim timeCol As Integer 'time col on raw data sheet
Dim TCsheetCol As Integer 'col with channel number on TC sheet
Dim channelCol As Integer 'col for channel on actual raw data
Dim datastart As Integer
Set reportmaker = ActiveWorkbook
Set rawDataWb = GetDataV2("Select report file")

TCsheetCol = 2
AddAgilent = CheckforAgilentData()
If AddAgilent Then
    looping = 2
  Else
    looping = 1
End If
  
timeCol = 1
Dim str As String

userinput = Application.inputbox("Input end time of stabilization, if you want it to pull from the end, leave at 0. Input in military time (hh:mm)", "Time input", 0) 'like for examples 14:10
        If userinput = "" Then Exit Sub

For J = 1 To looping
    
    'set TC sheet and raw data sheet
    If J = 1 Then
        Set rawDataSheet = rawDataWb.Sheets("RawData")
        Set TCsheet = rawDataWb.Sheets("AP4_Chamber")
        str = "Date/Time"
    Else
        Set rawDataSheet = rawDataWb.Sheets("AgilentData")
        Set TCsheet = rawDataWb.Sheets("AgilentTCs")
        str = "Time"
    End If
    
    'find row that the channels are on
    Dim channelrow As Integer
    channelrow = 0
    Dim x As Range
    
    rawDataSheet.Activate
    'find where the TC shannel names are
    Set x = rawDataSheet.Range("A1:C50").Find(str, LookIn:=xlValues)
    If x Is Nothing Then
        MsgBox ("Could not find correct row to look for the TC channels on raw data sheet")
    Else
        channelrow = x.Row
        
    End If
    
    'need to find DataStart - different per chamber, need to find in col 1 where there is a datatime
    'isdate()
    For i = 0 To 100
        If IsDate(rawDataSheet.Cells(channelrow + i, timeCol)) Then
            datastart = channelrow + i
                Exit For
        End If
    Next i
     
    Dim timeinterval As Integer
    'find time interval to determine a good point to find the corresponding times on agilent system
    timeinterval = CInt(DateDiff("s", rawDataSheet.Cells(datastart, timeCol), rawDataSheet.Cells(datastart + 1, timeCol)))
     
    'find Max temps allowed column
    TCsheet.Activate
    Set x = TCsheet.Range(Cells(1, 1), Cells(1, 100)).Find("Max Temps Allowed", LookIn:=xlValues)
    maxTempCol = x.Column
    
    'find last row
    lastrow = rawDataSheet.Range("A1000000").End(xlUp).Row
    
    'need to match up the agilent data to the CTS data
    'datastart
    'lastrow
    k = 1
    If J = 2 Then
    'need to change these for agilent data to start and end at times that match CTS
        k = datastart
        'skip loop if the agilent data was started after the CTS sw
        If DateDiff("s", rawDataSheet.Cells(k, timeCol), Agstart) >= 0 Then
            Do While k < lastrow
                If DateDiff("s", rawDataSheet.Cells(k + 1, timeCol), Agstart) < timeinterval + 5 _
                And DateDiff("s", rawDataSheet.Cells(k, timeCol), Agstart) >= timeinterval - 5 Then
                    datastart = k
                    Exit Do
                End If
                k = k + 1
            Loop
        End If 'should keep datastart as the initial starting row
        k = lastrow
        'skip loop if agilent data was turned off before CTS sw
        If DateDiff("s", rawDataSheet.Cells(k, timeCol), Aglast) <= 0 Then
            Do While k > datastart
                If DateDiff("s", rawDataSheet.Cells(k - 1, timeCol), Aglast) <= timeinterval + 5 _
                And DateDiff("s", rawDataSheet.Cells(k, timeCol), Aglast) >= timeinterval - 5 Then
                    lastrow = k
                    Exit Do
                End If
                k = k - 1
                If k = datastart Then
                    MsgBox ("didn't find a time stamp on agilent data that matches the last row of CTS data")
                End If
            Loop
        End If 'should keep last row as the ending row
    Else
        'find this out on CTS loop, will need on next loop with agilent system
        Agstart = rawDataSheet.Cells(datastart, timeCol) 'time stamp we need to look for on agilent system
        Aglast = rawDataSheet.Cells(lastrow, timeCol)
    End If

    Dim endtime As Date
    'setting up time stamps to use
    'userinput = Application.inputbox("Input end time of stabilization, if you want it to pull from the end, leave at 0. Input in military time (hh:mm)", "Time input", 0) 'like for examples 14:10
    '    If userinput = "" Then Exit Sub
        If userinput = "0" Then
            endtime = CDate(rawDataSheet.Cells(lastrow, 1).Value)
        Else
            
            endtime = ConvertStringtoDateTime(CStr(userinput), rawDataWb)
        End If
        
    Dim sub15min As Date
    Dim sub30min As Date
    
    dateSelect = CDate(Format(endtime, "MM-DD-YYYY"))
    
    sub15min = dateSelect + CDate(TimeValue(endtime) - TimeValue("0:15:00"))
    sub30min = dateSelect + CDate(TimeValue(endtime) - TimeValue("0:30:00"))
    
    rawDataSheet.Activate
    'finding rows to use, start from bottom for end time
    'Cells(Application.Match(CLng(Date), Range("A:A"), 1), "A").Select
    sub15minRow = FindClosestRowtoTimeString(rawDataSheet, sub15min, datastart)
    sub30minRow = FindClosestRowtoTimeString(rawDataSheet, sub30min, datastart)
    sub0minRow = FindClosestRowtoTimeString(rawDataSheet, endtime, datastart)
    'setting up columns
    nextCol = TCsheet.Cells(1, Columns.Count).End(xlToLeft).Column + 1
    TCsheet.Cells(1, nextCol).Value = "0min (C)"
    TCsheet.Cells(1, nextCol + 1).Value = "15min (C)"
    TCsheet.Cells(1, nextCol + 2).Value = "30min (C)"
    TCsheet.Cells(1, nextCol + 3).Value = "Difference"

    i = 2
    'find channel col
    'have about 700 col to search
    rawDataSheet.Activate
    While TCsheet.Cells(i, 2).Value <> ""
        'ex Z_UUT_ID_TC_035
        If TCsheet.Cells(i, 3).Value <> "" Then
            IDOD = TCsheet.Cells(i, 1).Value 'should just be ID or OD
            If J = 1 Then
                channelNum = Format(TCsheet.Cells(i, TCsheetCol).Value, "000")  'needs to be in 3 digit format
                columnName = "Z_UUT_" + IDOD + "_TC_" + channelNum
            Else
                columnName = CStr(TCsheet.Cells(i, TCsheetCol).Value) + " (F)" 'agilent system just takes the # in 1001 format
            End If
            'set range of full data
            Set x = rawDataSheet.Range(Cells(channelrow - 1, 1), Cells(channelrow + 1, 700)).Find(columnName, LookIn:=xlValues)
            If Not x Is Nothing Then
            channelCol = x.Column
                TCsheet.Cells(i, nextCol).FormulaR1C1 = "=convert('" & rawDataSheet.Name & "'!R" & sub30minRow & "C" & channelCol & ",""F"",""C"")"
                TCsheet.Cells(i, nextCol + 1).FormulaR1C1 = "=convert('" & rawDataSheet.Name & "'!R" & sub15minRow & "C" & channelCol & ",""F"",""C"")"
                TCsheet.Cells(i, nextCol + 2).FormulaR1C1 = "=convert('" & rawDataSheet.Name & "'!R" & sub0minRow & "C" & channelCol & ",""F"",""C"")"
                TCsheet.Cells(i, nextCol + 3).FormulaR1C1 = "=max(R[0]C[-3]:R[0]C[-1])-min(R[0]C[-3]:R[0]C[-1])"
                If TCsheet.Cells(i, nextCol + 3).Value > 3 Then
                    TCsheet.Cells(i, nextCol + 3).Interior.ColorIndex = 3
                End If
            End If
        End If
        i = i + 1
    Wend
Next J

End Sub


Sub MaxTempsAP4()

Dim reportmaker As Workbook
Dim rawDataWb As Workbook
Dim TCsheet As Worksheet
Dim rawDataSheet As Worksheet
Dim AddAgilent As Boolean 'flag if we are adding agilent data
Dim timeCol As Integer 'time col on raw data sheet
Dim TCsheetCol As Integer 'col with channel number on TC sheet
Dim channelCol As Integer 'col for channel on actual raw data
Dim datastart As Integer
Dim lastrow As Long
Dim channelrow As Integer
Set reportmaker = ActiveWorkbook
Set rawDataWb = GetDataV2("Select report file")
Dim x As Range
TCsheetCol = 2
AddAgilent = CheckforAgilentData()
If AddAgilent Then
    looping = 2
  Else
    looping = 1
End If
  
timeCol = 1
Dim str As String

For J = 1 To looping
    
    'set TC sheet and raw data sheet
    If J = 1 Then
        Set rawDataSheet = rawDataWb.Sheets("RawData")
        Set TCsheet = rawDataWb.Sheets("AP4_Chamber")
        str = "Date/Time"
    Else
        Set rawDataSheet = rawDataWb.Sheets("AgilentData")
        Set TCsheet = rawDataWb.Sheets("AgilentTCs")
        str = "Time"
    End If
    
    'find row that the channels are on
    channelrow = 0
    
    rawDataSheet.Activate
    'find where the TC shannel names are
    Set x = rawDataSheet.Range("A1:C50").Find(str, LookIn:=xlValues)
    If x Is Nothing Then
        MsgBox ("Could not find correct row to look for the TC channels on raw data sheet")
    Else
        channelrow = x.Row
        
    End If
    
    'need to find DataStart - different per chamber, need to find in col 1 where there is a datatime
    'isdate()
    For i = 0 To 100
        If IsDate(rawDataSheet.Cells(channelrow + i, timeCol)) Then
            datastart = channelrow + i
                Exit For
        End If
    Next i
     
    Dim timeinterval As Integer
    'find time interval to determine a good point to find the corresponding times on agilent system
    timeinterval = CInt(DateDiff("s", rawDataSheet.Cells(datastart, timeCol), rawDataSheet.Cells(datastart + 1, timeCol)))
     
    'find Max temps allowed column
    TCsheet.Activate
    Set x = TCsheet.Range(Cells(1, 1), Cells(1, 100)).Find("Max Temps Allowed", LookIn:=xlValues)
    maxTempCol = x.Column
    
    'find last row
    
    lastrow = rawDataSheet.Range("A1000000").End(xlUp).Row
    
    'need to match up the agilent data to the CTS data
    'datastart
    'lastrow
    k = 1
    If J = 2 Then
    rawDataSheet.Activate
    'need to change these for agilent data to start and end at times that match CTS
        k = datastart
        'skip loop if the agilent data was started after the CTS sw
        If DateDiff("s", rawDataSheet.Cells(k, timeCol), Agstart) >= 0 Then
            Do While k < lastrow
                If DateDiff("s", rawDataSheet.Cells(k + 1, timeCol), Agstart) < timeinterval + 5 _
                And DateDiff("s", rawDataSheet.Cells(k, timeCol), Agstart) >= timeinterval - 5 Then
                    datastart = k
                    Exit Do
                End If
                k = k + 1
            Loop
        End If 'should keep datastart as the initial starting row

        'rawDataSheet.Range(Cells(channelrow + 1, 1), Cells(datastart - 1, 1)).EntireRow.Delete
        k = lastrow
        'skip loop if agilent data was turned off before CTS sw
        If DateDiff("s", rawDataSheet.Cells(k, timeCol), Aglast) <= 0 Then
            Do While k > datastart
                If DateDiff("s", rawDataSheet.Cells(k - 1, timeCol), Aglast) <= timeinterval + 5 _
                And DateDiff("s", rawDataSheet.Cells(k, timeCol), Aglast) >= timeinterval - 5 Then
                    lastrow = k
                    Exit Do
                End If
                k = k - 1
                If k = datastart Then
                    MsgBox ("didn't find a time stamp on agilent data that matches the last row of CTS data")
                End If
            Loop
        End If 'should keep last row as the ending row
       'Remove excess data from agilent raw data sheet
       Call trimexcessAgdata(rawDataSheet, lastrow, datastart, channelrow)
        'rawDataSheet.Range(Cells(lastrow + 1, 1), Cells(wkst.Range("A1000000").End(xlUp).Row + 1, 1)).EntireRow.Delete
        'have to reestablish that the first row and last row of the Ag data is now good to use
        datastart = channelrow + 1
        lastrow = rawDataSheet.Range("A1000000").End(xlUp).Row
    Else
        'find this out on CTS loop, will need on next loop with agilent system
        Agstart = rawDataSheet.Cells(datastart, timeCol) 'time stamp we need to look for on agilent system
        Aglast = rawDataSheet.Cells(lastrow, timeCol)
    End If
    
    
    'find next open column and add the column names
    nextCol = TCsheet.Cells(1, Columns.Count).End(xlToLeft).Column + 1
    TCsheet.Cells(1, nextCol).Value = "Max Temps (F) from data"
    TCsheet.Cells(1, nextCol + 1).Value = "Average Top 5-10 temps (F)"
    TCsheet.Cells(1, nextCol + 2).Value = "Delta"
    TCsheet.Cells(1, nextCol + 3).Value = "Max Temps (C)"
    TCsheet.Cells(1, nextCol + 4).Value = "Delta"
    
    Dim channelNum As String
    Dim columnName As String
    Dim IDOD As String
    Dim fullData As Range
    
    i = 2
    'find channel col
    'have about 700 col to search
    rawDataSheet.Activate
    While TCsheet.Cells(i, TCsheetCol).Value <> ""
        'ex Z_UUT_ID_TC_035
        If TCsheet.Cells(i, 3).Value <> "" Then
            IDOD = TCsheet.Cells(i, 1).Value 'should just be ID or OD
            'need to change locations to look for channel numbers for agilent system
            If J = 1 Then
                channelNum = Format(TCsheet.Cells(i, TCsheetCol).Value, "000")  'needs to be in 3 digit format
                columnName = "Z_UUT_" + IDOD + "_TC_" + channelNum
            Else
                columnName = CStr(TCsheet.Cells(i, TCsheetCol).Value) + " (F)" 'agilent system just takes the # in 1001 format
            End If
                      
            Set x = rawDataSheet.Range(Cells(channelrow - 1, 1), Cells(channelrow + 1, 700)).Find(columnName, LookIn:=xlValues)
            If Not x Is Nothing Then
                channelCol = x.Column
                    
                'set range of full data
                Set fullData = rawDataSheet.Range(Cells(datastart, channelCol), Cells(lastrow, channelCol))
                a = fullData.Address
                TCsheet.Cells(i, nextCol).FormulaR1C1 = "=max(R1C1)"
                'max temps from rawdatasheet
                TCsheet.Cells(i, nextCol).Formula = "=max('" & rawDataSheet.Name & "'!" & fullData.Address & ")"
                'average of 5-10 largest numbers
                TCsheet.Cells(i, nextCol + 1).Formula2 = "=average(large('" & rawDataSheet.Name & "'!" & fullData.Address & ",{5,6,7,8,9,10}))"
                'delta of that and max, if over 10 need to make a differnt color
                TCsheet.Cells(i, nextCol + 2).Formula = "=abs(R[0]C[-2] - R[0]C[-1])"
                If TCsheet.Cells(i, nextCol + 2) > 10 Then
                    TCsheet.Cells(i, nextCol + 2).Interior.Color = vbRed
                End If
                'convert the max temp to C
                TCsheet.Cells(i, nextCol + 3).Formula2 = "=convert(large('" & rawDataSheet.Name & "'!" & fullData.Address & ",{1}),""F"",""C"")"
                'find delta from our max temps allowed
                TCsheet.Cells(i, nextCol + 4).Formula = "=if(isblank(R[0]C" & maxTempCol & "),"""",R[0]C" & maxTempCol & " - R[0]C[-1])"
                 If TCsheet.Cells(i, nextCol + 4) < 0 Then
                    TCsheet.Cells(i, nextCol + 4).Interior.Color = vbRed
                End If
            End If
        End If
        
        i = i + 1
    Wend

Next J

End Sub



Function CheckforAgilentData()
i = 1
While i <= ActiveWorkbook.Sheets.Count
    If ActiveWorkbook.Sheets(i).Name = "AgilentData" Then
        CheckforAgilentData = True
    End If
    i = i + 1
Wend
End Function


Public Function GetDataV2(message As String)
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
    Set GetDataV2 = ActiveWorkbook
End If

End Function

Function ConvertStringtoDateTime(userinput As String, wb As Workbook) 'taking in a string that is hh:mm miltary time and translating it to a date time to function later
desiredtimestamp = CDate(userinput)
'10/25/24- adding in extra input- need to know date to use if 2 dates on tecor data
'funciton to test if there are 2 dates
sheetDate = WhichDateDecider(wb)
desiredDate = CDate(Format(sheetDate, "yyyy-mm-dd"))
Dim combinedDateTime As String
combinedDateTime = desiredDate + desiredtimestamp
ConvertStringtoDateTime = combinedDateTime
End Function
Function FindClosestRowtoTimeString(wkst As Worksheet, timetoFind As Date, datastart As Integer)
lastrow = wkst.Range("A1000000").End(xlUp).Row

i = lastrow
While i > datastart
    time1 = CDate(wkst.Cells(i, 1).Value)
    time2 = CDate(wkst.Cells(i - 1, 1).Value)
    x = DateDiff("s", timetoFind, time1)
    y = DateDiff("s", timetoFind, time2)
    If (DateDiff("s", timetoFind, time1) <= 5 And DateDiff("s", timetoFind, time2) >= -5) Or (time1 = timetoFind) Then
        FindClosestRowtoTimeString = i
        Exit Function
    End If
    i = i - 1
Wend

If i = datastart Then
    'MsgBox ("Could not find the correct timestamp")
End If

End Function
Function WhichDateDecider(wb As Workbook)
'should always be Tescor
wb.Sheets("RawData").Activate
lastrow = Range("A1000000").End(xlUp).Row
For i = 1 To 100
    If IsDate(Cells(i, 1).Value) Then
        FirstDate = CDate(Format(Cells(i, 1).Value, "yyyy-mm-dd"))
        Exit For
    End If
Next i

LastDate = CDate(Format(Cells(lastrow, 1).Value, "yyyy-mm-dd"))

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

Sub trimexcessAgdata(wkst As Worksheet, lastrow As Long, firstrow As Integer, channelrow As Integer)

'Remove excess data from agilent raw data sheet, trim last rows first and then top rows

wkst.Range(Cells(lastrow + 1, 1), Cells(wkst.Range("A1000000").End(xlUp).Row + 1, 1)).EntireRow.Delete
If firstrow <> channelrow + 1 Then
    wkst.Range(Cells(channelrow + 1, 1), Cells(firstrow - 1, 1)).EntireRow.Delete
End If
    
End Sub
