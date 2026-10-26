Attribute VB_Name = "AP4"
Public ExitSub
Public DataAq As String
Public rawDataSheet As Worksheet
Public TCsheet As Worksheet
Public timeCol As Integer
Public reportmaker As Workbook
Public rawDataWb As Workbook
Public str As String
Public AddAgilent As Boolean
Public TCsheetCol As Integer 'col with channel number on TC sheet
Public channelCol As Integer 'col for channel on actual raw data
Public datastart As Integer
Public lastrow As Long
Public channelrow As Integer
Public looping As Integer
Public Agstart As Date
Public Aglast As Date
Public S As Range
Public J As Integer 'looping counter, needs to be used in StartChecks and in the regular subs
Public timeinterval As Integer
Public maxTempCol As Integer

Sub CheckStablilityAP4v3()

Set reportmaker = ActiveWorkbook
Set rawDataWb = GetDataV2("Select report file")

userinput = Application.inputbox("Input end time of stabilization, if you want it to pull from the end, leave at 0. Input in military time (hh:mm)", "Time input", 0) 'like for examples 14:10
        If userinput = "" Then Exit Sub

AddAgilent = CheckforSheetName("AgilentData")
If AddAgilent Then
    looping = 2
  Else
    looping = 1
End If
  

For J = 1 To looping
    
    StartChecks
    
    Dim endtime As Date
    'setting up time stamps to use
        If userinput = "0" Then
            endtime = CDate(ConvertTimeFormatv2(rawDataSheet.Cells(lastrow, timeCol).Value))
        Else
            
            endtime = ConvertStringtoDateTimev2(CStr(userinput), rawDataWb)
        End If
        
    Dim sub15min As Date
    Dim sub30min As Date
    
    dateSelect = CDate(Format(endtime, "MM-DD-YYYY"))
    
    sub15min = dateSelect + CDate(TimeValue(endtime) - TimeValue("0:15:00"))
    sub30min = dateSelect + CDate(TimeValue(endtime) - TimeValue("0:30:00"))
    
    rawDataSheet.Activate 'somewhere here to TCsheet.Cells(1, nextCol)
    'finding rows to use, start from bottom for end time
    sub15minRow = FindClosestRowtoTimeStringv2(rawDataSheet, sub15min, datastart)
    sub30minRow = FindClosestRowtoTimeStringv2(rawDataSheet, sub30min, datastart)
    sub0minRow = FindClosestRowtoTimeStringv2(rawDataSheet, endtime, datastart)
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
            
            IDOD = TCsheet.Cells(i, 1).Value 'should just be ID or OD for CTS only
            'need to change locations to look for channel numbers for agilent system
            If J = 1 And CheckforSheetName("Tescor TCs") Then 'for Tescor TC Channel Names
                channelNum = Format(TCsheet.Cells(i, TCsheetCol).Value, "00")  'needs to be in 3 digit format
                columnName = "TC" + channelNum
            ElseIf J = 1 Then 'For CTS TC Channel Names
                channelNum = Format(TCsheet.Cells(i, TCsheetCol).Value, "000")  'needs to be in 3 digit format
                columnName = "Z_UUT_" + IDOD + "_TC_" + channelNum
            Else 'For Agilent TC channel names
                columnName = CStr(TCsheet.Cells(i, TCsheetCol).Value) + " (F)" 'agilent system just takes the # in 1001 format
            End If
          
            'set range of full data
            Set x = rawDataSheet.Range(Cells(channelrow, 1), Cells(channelrow + 1, 700)).Find(columnName, LookIn:=xlValues)
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

Sub StartChecks()
  
timeCol = 1 ' changed on Tescor

DataAq = "" 'reset, will only be Tescor or "" at this point
    'set TC sheet and raw data sheet
    If J = 1 Then
        If CheckforSheetName("Tescor TCs") Then
            Set rawDataSheet = rawDataWb.Sheets("Measured Data")
            Set TCsheet = rawDataWb.Sheets("Tescor TCs")
            str = "Date"
            timeCol = 3
            DataAq = "Tescor"
            TCsheetCol = 1
        Else
            Set rawDataSheet = rawDataWb.Sheets("RawData")
            Set TCsheet = rawDataWb.Sheets("AP4_Chamber")
            str = "Date/Time"
            timeCol = 1
            TCsheetCol = 2
        End If
    Else
        Set rawDataSheet = rawDataWb.Sheets("AgilentData")
        Set TCsheet = rawDataWb.Sheets("AgilentTCs")
        str = "Time"
        timeCol = 1
        TCsheetCol = 2
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
    For i = 1 To 100
        If IsDate(ConvertTimeFormatv2(rawDataSheet.Cells(channelrow + i, timeCol))) Then
            datastart = channelrow + i
                Exit For
        End If
    Next i

    'find time interval to determine a good point to find the corresponding times on agilent system
    timeinterval = CInt(DateDiff("s", ConvertTimeFormatv2(rawDataSheet.Cells(datastart, timeCol)), ConvertTimeFormatv2(rawDataSheet.Cells(datastart + 1, timeCol))))
     
    'find Max temps allowed column
    TCsheet.Activate
    Set x = TCsheet.Range(Cells(1, 1), Cells(1, 100)).Find("Max Temps Allowed (C)", LookAt:=xlPart)
    maxTempCol = x.Column
    
    'find last row
    lastrow = rawDataSheet.Range("A1000000").End(xlUp).Row
    
    'need to match up the agilent data to the CTS/tescor data
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
                If DateDiff("s", rawDataSheet.Cells(k + 1, timeCol), Agstart) <= timeinterval + 3 _
                And DateDiff("s", rawDataSheet.Cells(k, timeCol), Agstart) >= 0 Then
                    datastart = k
                    Exit Do
                End If
                k = k + 1
            Loop
        End If 'should keep datastart as the initial starting row if it doesn't find a better one

        'rawDataSheet.Range(Cells(channelrow + 1, 1), Cells(datastart - 1, 1)).EntireRow.Delete
        k = lastrow
        'skip loop if agilent data was turned off before CTS sw
        If DateDiff("s", rawDataSheet.Cells(k, timeCol), Aglast) <= 0 Then
            Do While k > datastart
                If DateDiff("s", rawDataSheet.Cells(k - 1, timeCol), Aglast) <= timeinterval + 3 _
                And DateDiff("s", rawDataSheet.Cells(k, timeCol), Aglast) >= 0 Then
                    lastrow = k
                    Exit Do
                End If
                k = k - 1
                If k = datastart Then
                    MsgBox ("Did not find last time stamp of consule data, will use end of Agilent Data as end")
                End If
            Loop
        End If 'should keep last row as the ending row if ending time not found
        
       'Remove excess data from agilent raw data sheet
       'Call trimexcessAgdata(rawDataSheet, lastrow, datastart, channelrow)
        'have to reestablish that the first row and last row of the Ag data is now good to use
        'datastart = channelrow + 1
        'lastrow = rawDataSheet.Range("A1000000").End(xlUp).Row
    Else
        'find this out on CTS loop, will need on next loop with agilent system
        Agstart = ConvertTimeFormatv2(rawDataSheet.Cells(datastart, timeCol)) 'time stamp we need to look for on agilent system
        Aglast = ConvertTimeFormatv2(rawDataSheet.Cells(lastrow, timeCol))
    End If
    
  

End Sub
Sub MaxTempsAP4v2()

Set reportmaker = ActiveWorkbook
Set rawDataWb = GetDataV2("Select report file")


AddAgilent = CheckforSheetName("AgilentData")
If AddAgilent Then
    looping = 2
  Else
    looping = 1
End If
  

For J = 1 To looping

    StartChecks
       
    If J = 2 Then
        rawDataSheet.Activate
        'Remove excess data from agilent raw data sheet
        Call trimexcessAgdatav2(rawDataSheet, lastrow, datastart, channelrow)
        'have to reestablish that the first row and last row of the Ag data is now good to use
        datastart = channelrow + 1
        lastrow = rawDataSheet.Range("A1000000").End(xlUp).Row
    End If
    
    

    
    'where finding the max temps really start
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
         
            IDOD = TCsheet.Cells(i, 1).Value 'should just be ID or OD for CTS only
            'need to change locations to look for channel numbers for agilent system
            If J = 1 And CheckforSheetName("Tescor TCs") Then 'for Tescor TC Channel Names
                channelNum = Format(TCsheet.Cells(i, TCsheetCol).Value, "00")  'needs to be in 3 digit format
                columnName = "TC" + channelNum
            ElseIf J = 1 Then 'For CTS TC Channel Names
                channelNum = Format(TCsheet.Cells(i, TCsheetCol).Value, "000")  'needs to be in 3 digit format
                columnName = "Z_UUT_" + IDOD + "_TC_" + channelNum
            Else 'For Agilent TC channel names
                columnName = CStr(TCsheet.Cells(i, TCsheetCol).Value) + " (F)" 'agilent system just takes the # in 1001 format
            End If
           
            Set x = rawDataSheet.Range(Cells(channelrow, 1), Cells(channelrow + 1, 700)).Find(columnName, LookIn:=xlValues)
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



Function CheckforSheetName(sheetname As String)
i = 1
While i <= ActiveWorkbook.Sheets.Count
    If ActiveWorkbook.Sheets(i).Name = sheetname Then
        CheckforSheetName = True
        Exit Function
    End If
    i = i + 1
Wend
End Function


Public Function GetDataV3(message As String)
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

Function ConvertStringtoDateTimev2(userinput As String, wb As Workbook) 'taking in a string that is hh:mm miltary time and translating it to a date time to function later
desiredtimestamp = CDate(userinput)
'10/25/24- adding in extra input- need to know date to use if 2 dates on tecor data
'funciton to test if there are 2 dates
sheetDate = WhichDateDeciderv2(wb)
desiredDate = CDate(Format(sheetDate, "yyyy-mm-dd"))
Dim combinedDateTime As String
combinedDateTime = desiredDate + desiredtimestamp
ConvertStringtoDateTimev2 = combinedDateTime
End Function
Function FindClosestRowtoTimeStringv2(wkst As Worksheet, timetoFind As Date, datastart As Integer)
lastrow = wkst.Range("A1000000").End(xlUp).Row

i = lastrow
While i > datastart

    time1 = CDate(ConvertTimeFormatv2(wkst.Cells(i, timeCol).Value))
    time2 = CDate(ConvertTimeFormatv2(wkst.Cells(i - 1, timeCol).Value))
    d = DateDiff("s", timetoFind, time1) '1795
    e = DateDiff("s", timetoFind, time2)
    z = 5 - 3
    If (DateDiff("s", timetoFind, time1) <= 3 + timeinterval And DateDiff("s", timetoFind, time2) >= 0 - timeinterval) Or (time1 = timetoFind) Then
        FindClosestRowtoTimeStringv2 = i
        Exit Function
    End If
    i = i - 1
Wend

If i = datastart Then
    'MsgBox ("Could not find the correct timestamp")
End If

End Function
Function WhichDateDeciderv2(wb As Workbook)
'should always be Tescor
rawDataSheet.Activate
lastrow = Range("A1000000").End(xlUp).Row
For i = 0 To 100
    If IsDate(ConvertTimeFormatv2(Cells(datastart + i, timeCol).Value)) Then
        FirstDate = CDate(Format(ConvertTimeFormatv2(Cells(datastart + i, timeCol).Value), "yyyy-mm-dd"))
        Exit For
    End If
Next i

LastDate = CDate(Format(ConvertTimeFormatv2(Cells(lastrow, timeCol).Value), "yyyy-mm-dd"))

'if thie first date on the tescor data doesn't match the last then it will prompt you to tell it which date you meant
If FirstDate <> LastDate Then
    answer = MsgBox("This data spans overnight to 2 dates. Which date should be used? Click Yes for: " & FirstDate & " or click No for: " & LastDate & "", vbQuestion + vbYesNo + vbDefaultButton2, "Date in question")
        If answer = vbYes Then
            WhichDateDeciderv2 = FirstDate
        Else
            WhichDateDeciderv2 = LastDate
        End If
Else
    WhichDateDeciderv2 = FirstDate

End If

End Function

Sub trimexcessAgdatav2(wkst As Worksheet, lastrow As Long, firstrow As Integer, channelrow As Integer)

'Remove excess data from agilent raw data sheet, trim last rows first and then top rows

wkst.Range(Cells(lastrow + 1, 1), Cells(wkst.Range("A1000000").End(xlUp).Row + 1, 1)).EntireRow.Delete
If firstrow <> channelrow + 1 Then
    wkst.Range(Cells(channelrow + 1, 1), Cells(firstrow - 1, 1)).EntireRow.Delete
End If
    
End Sub

Function ConvertTimeFormatv2(time As String)
'needs to be configured for all data sets
If time <> "" Then
    If DataAq = "Tescor" Then
    
        Sec = CInt(Right(time, 2))
        Min = CInt(Mid(time, 15, 2))
        hr = CInt(Mid(time, 12, 2))
        Days = CInt(Mid(time, 9, 2))
        Mon = CInt(Mid(time, 6, 2))
        Years = CInt(Left(time, 4))
        
        ConvertTimeFormatv2 = DateSerial(Years, Mon, Days) + TimeSerial(hr, Min, Sec)
        Exit Function
    Else
        ConvertTimeFormatv2 = time
    End If
End If
End Function
