Attribute VB_Name = "Module1"
Option Explicit

Sub CalcAtt()
    ' The sub calculates the attendance information according to raw data given
    
    ' There are several stages in concern...
    '  Obtain Settings
    '  0. Create or Replace the "Result" sheet
    '  I. Read and list raw data
    '  II. Check characteristic of the attendance information
    '  III. Calculate Daily Attendance Details
    '  IV. Evaluate Monthy Attendance Details
    On Error Resume Next
    
    'Stage 0
    Dim messageBox As VbMsgBoxResult
    
    Dim readingSheet As Worksheet
    
    Dim destinationSheet As Worksheet
    Dim settingSheet As Worksheet
    Dim holidaySheet As Worksheet
    Dim settingParameter As Range
    Dim theRange, source As Range
    
    Dim startTime As Double
    Dim endTime As Double
    Dim lunch As Double
    Dim Version As Double
    Dim significance As Integer
    Dim interval As Integer
    Dim expected As Integer
    Dim name As String
    
    Dim startingColumn As Integer
    Dim i, j, k, rowCount, overTimeRowCount, overTimeCount, cursor, cursorBegin, codeASCII, recordCount, readingRowBegins, skippedDays, runCount, holidayCount As Integer
    Dim maximumRecordCount As Integer
    Dim overNight As Boolean
    Dim lunchCell(0 To 1) As String
    Dim expectCell(0 To 1) As String
    Dim workCell(0 To 1) As String
    Dim overtimeCell(0 To 1) As String
    Dim lateCell(0 To 1) As String
    
    Set settingSheet = Sheets("Settings")
    Set holidaySheet = Sheets("Holidays")
    Set destinationSheet = Sheets("Results")
    Set readingSheet = Sheets("Readings")
    
    
    startingColumn = 65 ' denotes A
    cursorBegin = 8 ' cursor starts at row 8
    readingRowBegins = 4 ' the reading sheet contains data from row 4
    skippedDays = 0
    overNight = False
    overTimeRowCount = shtOvertime.Range("I2").Value
    k = 0
    
    maximumRecordCount = 4 ' each line has only 4 records
    ' Set Header Configurations
    lunchCell(0) = "D3"
    lunchCell(1) = "E3"
    expectCell(0) = "D4"
    expectCell(1) = "E4"
    workCell(0) = "F3"
    workCell(1) = "G3"
    overtimeCell(0) = "F4"
    overtimeCell(1) = "G4"
    lateCell(0) = "H3"
    lateCell(1) = "I3"
    ' Obtain settings

    Set settingParameter = settingSheet.Range("A2:A100").Find(what:="Start", After:=settingSheet.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    startTime = settingSheet.Range("B" & settingParameter.Row).Value
    Set settingParameter = settingSheet.Range("A2:A100").Find(what:="End", After:=settingSheet.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    endTime = settingSheet.Range("B" & settingParameter.Row).Value
    Set settingParameter = settingSheet.Range("A2:A100").Find(what:="Lunch", After:=settingSheet.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    lunch = settingSheet.Range("B" & settingParameter.Row).Value
    Set settingParameter = settingSheet.Range("A2:A100").Find(what:="significance", After:=settingSheet.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    significance = settingSheet.Range("B" & settingParameter.Row).Value
    Set settingParameter = settingSheet.Range("A2:A100").Find(what:="interval", After:=settingSheet.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    interval = settingSheet.Range("B" & settingParameter.Row).Value
    Set settingParameter = settingSheet.Range("A2:A100").Find(what:="version", After:=settingSheet.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    Version = settingSheet.Range("B" & settingParameter.Row).Value
    Set settingParameter = settingSheet.Range("A2:A100").Find(what:="Name", After:=settingSheet.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    name = settingSheet.Range("B" & settingParameter.Row).Value
    ' Setting Validation
    If endTime <= startTime Then
        messageBox = MsgBox("Incorrect time detected." & Chr(10) & "Please Check Settings correctly." & Chr(10) & " Operation Halt", vbCritical + vbOKOnly)
        Exit Sub
    End If
    expected = DateDiff("n", startTime, endTime) - lunch

    
    If destinationSheet Is Nothing Then
        ' The sheet does not exist
        Set destinationSheet = Sheets.Add(Null, ActiveSheet)
        ' Name the sheet as "Results"
        destinationSheet.name = "Results"
    End If
    ' Clear the sheet
    destinationSheet.Cells.ClearContents
    
    
    
    'Stage I'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Prepare the Reading Sheet
    
    ' Define a cursor for destination sheet
  
    codeASCII = startingColumn
    cursor = cursorBegin
    ' Note the most left hand column is used

    ' Obtain current rows
    rowCount = readingSheet.Range("D2").Value
    
    ' Set Headers Here
    ' Staff Information
    destinationSheet.Range("B2").Value = "STAFF ID"
    destinationSheet.Range("C2").Value = readingSheet.Range("B2").Value
    destinationSheet.Range("F2").Value = name
    ' Holiday Information
    holidayCount = holidaySheet.Range("B1").Value
    ' Work Inforamtion
    destinationSheet.Range("B3").Value = "Start"
    destinationSheet.Range("C3").Value = startTime
    destinationSheet.Range("B4").Value = "End"
    destinationSheet.Range("C4").Value = endTime
    
    destinationSheet.Range(lunchCell(0)).Value = "Lunch"
    destinationSheet.Range(lunchCell(1)).Value = lunch
    
    destinationSheet.Range(expectCell(0)).Value = "Expect"
    destinationSheet.Range(expectCell(1)).Value = expected
    
    ' These are placed here as labels, their content will be added once all calcuations are completed(see the bottom)
    destinationSheet.Range(workCell(0)).Value = "Worked Day(s)"
    destinationSheet.Range(overtimeCell(0)).Value = "Overtimed Day(s)"
    destinationSheet.Range(lateCell(0)).Value = "VERSION"
    destinationSheet.Range(lateCell(1)).Value = Version
    ' Header for the content
    destinationSheet.Range("A7").Value = "Date"
    destinationSheet.Range("B7").Value = "Records"
    destinationSheet.Range("H7").Value = " Time1"
    destinationSheet.Range("I7").Value = "Lunch"
    destinationSheet.Range("J7").Value = "Time2"
    destinationSheet.Range("L7").Value = "Worked"
    destinationSheet.Range("M7").Value = "Overtime"
    destinationSheet.Range("N7").Value = "Holiday"

    
    ' For Each Row
    recordCount = 1
    i = 0
    Do While i < (rowCount)
    ' For Debug ONLY: NOT FOR PRODUCTION
    'MsgBox "Current Row " & i & ",Row Count: " & rowCount
    'If i = 33 Then
     'MsgBox "just stop!"
    'End If
    ' Out of Debug Area
    runCount = runCount + 1
        If runCount > 10000 Then
            MsgBox "Iteration Protection Activcated. Refer to Developer. Developer: Please check the coding"
            Exit Do
        End If
        'Set source = theRange.SpecialCells(xlCellTypeVisible).Cells(i)
Start:
        ' A date is added to first column of each row
        If recordCount = 1 Then
            ' Add the date at column
            If skippedDays = 0 Then
                destinationSheet.Range(Chr(codeASCII) & cursor).Value = DateValue(readingSheet.Range("A" & i + readingRowBegins).Value)
            End If
            ' switch column to C
            codeASCII = codeASCII + 1
   
            ' Check  Skipping
            ' If there is a skip, add a blank line...
            If cursor > cursorBegin Then ' but not for first row
            'MsgBox destinationSheet.Range(Chr(startingColumn) & cursor).Value & "," & destinationSheet.Range(Chr(startingColumn) & cursor - 1).Value
            'MsgBox DateDiff("d", destinationSheet.Range(Chr(startingColumn) & cursor).Value, destinationSheet.Range(Chr(startingColumn) & cursor - 1).Value) <> -1
                If DateDiff("d", destinationSheet.Range(Chr(startingColumn) & cursor).Value, destinationSheet.Range(Chr(startingColumn) & cursor - 1).Value) <> -1 Then
                skippedDays = skippedDays + 1
                'MsgBox "There is a skip at row " & cursor
                'MsgBox "The difference is " & DateDiff("d", destinationSheet.Range(Chr(startingColumn) & cursor).Value, destinationSheet.Range(Chr(startingColumn) & cursor - 1).Value)
                    destinationSheet.Range(Chr(startingColumn) & cursor).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                    destinationSheet.Range(Chr(startingColumn) & cursor).Value = DateValue(DateAdd("d", skippedDays, readingSheet.Range("A" & i + readingRowBegins - 1).Value))
                    cursor = cursor + 1
                    ' go back to beginning...
                    recordCount = 1
                    ' and reset the column too...
                    codeASCII = startingColumn
                    GoTo Start
                Else
                    ' There is no skip, reset the skip day counter
                    skippedDays = 0
                End If
            End If
        End If
      ' Add Further Validation Rules here... Check your notes!
         ' Once all necessary rules are passed, the next line can be executed.
            If readingSheet.Range("D" & i + readingRowBegins).Value = "Valid" Then
                destinationSheet.Range(Chr(codeASCII) & cursor).Value = readingSheet.Range("A" & i + readingRowBegins).Value
                'switch to next column
                codeASCII = codeASCII + 1
                recordCount = recordCount + 1
                
            End If
            
        ' Check for same day to next row
        ' Once a differnt date is spotted, a new row will be used
        'MsgBox DateDiff("d", readingSheet.Range("A" & i + readingRowBegins).Value, readingSheet.Range("A" & i + readingRowBegins).Value) & " 1: " & readingSheet.Range("A" & i + readingRowBegins).Value & "2: " & readingSheet.Range("A" & i + readingRowBegins).Value
        If DateDiff("d", readingSheet.Range("A" & i + readingRowBegins).Value, readingSheet.Range("A" & i + readingRowBegins + 1).Value) = 0 Then
            ' same day,
           
            If overNight Then
                ' process the obvernight matter
                overNight = False
                GoTo secondStage
            End If
        Else
            ' not the same, check for overnight work capacity (Both Exit and Valid Record)[Version 1.2]
            If readingSheet.Range("C" & i + readingRowBegins + 1).Value = "Exit" And readingSheet.Range("D" & i + readingRowBegins + 1).Value = "Valid" Then
                overNight = True
                GoTo Ending ' skip for making a new row for this moment, but it will be resume once the relevant record entered
            End If
secondStage:
            ' if not overnight work, proeed to stage II'''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Check whether there are sufficient information provided for that day...
            
            ' move to 2 columns away from the records
            codeASCII = startingColumn + maximumRecordCount + 3
            'GoTo Ending ' Skip the stage II main part
            'If recordCount < maximumRecordCount Then
                ' switch the column further in case insufficient records captured
                 'codeASCII = codeASCII + maximumRecordCount - recordCount
            'End If
                      
            ' and stage III'''''''''''''''''''''''''''''''''''''
            'codeASCII = 75 ' move to column I
            ' Calculate duration between record point 1 and 2
            If DateDiff("n", destinationSheet.Range(Chr(startingColumn + 1) & cursor).Value, destinationSheet.Range(Chr(startingColumn + 2) & cursor).Value) > 0 Then
            ' only valid time information can be added (Positive time results)
                destinationSheet.Range(Chr(codeASCII) & cursor).Value = DateDiff("n", destinationSheet.Range(Chr(startingColumn + 1) & cursor).Value, destinationSheet.Range(Chr(startingColumn + 2) & cursor).Value)
            End If
            
            If recordCount >= maximumRecordCount Then
                codeASCII = codeASCII + 1 ' calculate lunch period
                If DateDiff("n", destinationSheet.Range(Chr(startingColumn + 2) & cursor).Value, destinationSheet.Range(Chr(startingColumn + 3) & cursor).Value) > 0 Then
                    destinationSheet.Range(Chr(codeASCII) & cursor).Value = DateDiff("n", destinationSheet.Range(Chr(startingColumn + 2) & cursor).Value, destinationSheet.Range(Chr(startingColumn + 3) & cursor).Value)
                End If
                codeASCII = codeASCII + 1 ' Calculate duration between record point 3 and 4
                If DateDiff("n", destinationSheet.Range(Chr(startingColumn + 3) & cursor).Value, destinationSheet.Range(Chr(startingColumn + 4) & cursor).Value) > 0 Then
                    destinationSheet.Range(Chr(codeASCII) & cursor).Value = DateDiff("n", destinationSheet.Range(Chr(startingColumn + 3) & cursor).Value, destinationSheet.Range(Chr(startingColumn + 4) & cursor).Value)
                End If
            Else
                codeASCII = codeASCII + 2
            End If
            codeASCII = codeASCII + 2
            'codeASCII = 79 ' move to column L Overtime Count
            ' Read overtime working hours here
            overTimeCount = 0
            Do While overTimeRowCount > k And shtOvertime.Range("A" & k + 3).Value = destinationSheet.Range("A" & cursor).Value
                overTimeCount = overTimeCount + shtOvertime.Range("H" & k + 3).Value
                k = k + 1 ' check next row
            Loop
            ' to total work hours
            destinationSheet.Range(Chr(codeASCII) & cursor).Formula = "=" & Chr(codeASCII - 4) & cursor & "+" & Chr(codeASCII - 2) & cursor  ' Total time period
            'move to overtime Count
            codeASCII = codeASCII + 1
            destinationSheet.Range(Chr(codeASCII) & cursor).Value = overTimeCount
            'destinationSheet.Range(Chr(codeASCII) & cursor).Formula = "=IF(AND(" & Chr(codeASCII - 1) & cursor & "- " & expectCell(1) & " > " & significance & "," & Chr(codeASCII + 1) & cursor & "=0)," & Chr(codeASCII - 1) & cursor & "- " & expectCell(1) & ",0)" ' Overtime period
            ' Holiday Period Determine here
            codeASCII = codeASCII + 1
            ' If the date is found on holiday sheet, copy work hours into holiday
            For j = 2 To (holidayCount + 2)
                If DateDiff("d", holidaySheet.Range("A" & j).Value, destinationSheet.Range(Chr(startingColumn) & cursor).Value) = 0 Then
                    destinationSheet.Range(Chr(codeASCII) & cursor).Value = destinationSheet.Range(Chr(codeASCII - 2) & cursor).Value
                    Exit For ' Once a holiday has been found, exit the loop immediately
                End If
            Next
            cursor = cursor + 1
            recordCount = 1
Ending:
        End If

        i = i + 1 ' for the next record
        If i <> rowCount And recordCount = 1 Then ' check whether reaching the end of all records
                ' not reaching the end , then switch row and reset column
                codeASCII = startingColumn
        End If
        
        
    Loop
    'GoTo setStyle ' For Debug Only
    ' Add bottom line here
        ' the worktime sum ups
        destinationSheet.Range(Chr(codeASCII) & cursor).Formula = "=SUM(" & Chr(codeASCII) & cursorBegin & ":" & Chr(codeASCII) & cursor - 1 & ")"
        destinationSheet.Range(Chr(codeASCII - 1) & cursor).Formula = "=SUM(" & Chr(codeASCII - 1) & cursorBegin & ":" & Chr(codeASCII - 1) & cursor - 1 & ")"
        destinationSheet.Range(Chr(codeASCII - 2) & cursor).Formula = "=SUM(" & Chr(codeASCII - 2) & cursorBegin & ":" & Chr(codeASCII - 1) & cursor - 2 & ")"
        ' convert to the hours
        destinationSheet.Range(Chr(codeASCII) & cursor + 1).Formula = "=(" & Chr(codeASCII) & cursor & "/60)"
        destinationSheet.Range(Chr(codeASCII - 1) & cursor + 1).Formula = "=(" & Chr(codeASCII - 1) & cursor & "/60)"
        destinationSheet.Range(Chr(codeASCII - 2) & cursor + 1).Formula = "=(" & Chr(codeASCII - 2) & cursor & "/60)"
        ' Add the header
        destinationSheet.Range(workCell(1)).Formula = "=COUNTA(B" & cursorBegin & ":B" & cursor - 1 & ")"
        destinationSheet.Range(overtimeCell(1)).Formula = "=COUNTIF(" & Chr(codeASCII - 1) & cursorBegin & ":" & Chr(codeASCII - 1) & cursor - 1 & ",""> 0"")"
        
        ' Add styling here
        ' Redirect the user to results
    destinationSheet.Select ' This guarantees the styling can be carried across. VERSION 1.1
        ' The beginning of each line has a date with ISO8601 format
        ' All time for each line are shown as hh:mm
        destinationSheet.Range("A" & cursorBegin & ":A" & cursor).NumberFormat = "yyyy-mm-dd"
        destinationSheet.Range("B" & cursorBegin & ":" & Chr(codeASCII) & cursor).NumberFormat = "h:mm"
        'MsgBox (cursor)
        ' Set minute count for rest of columns
        destinationSheet.Range(Chr(startingColumn + maximumRecordCount + 3) & cursorBegin & ":AZ" & cursor).NumberFormat = "0"
        ' Configure Specific format for the buttom row
        destinationSheet.Range(Chr(startingColumn + maximumRecordCount + 3) & cursor + 1 & ":AZ" & cursor + 1).NumberFormat = "0.00"
setStyle:
        ' clear current style
        destinationSheet.Range("A8:" & Chr(codeASCII) & cursor + 50).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        
' set border style
    destinationSheet.Range("A8:" & Chr(codeASCII) & cursor - 1).Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .weight = xlThin
    End With
        
    
End Sub
Sub prepareReadings()
    ' This sub prepares the readings, prior to stage I
    ' Importantly this will also determine whether the record is valid or not
    ' The main calucation function will use this sheet instead of raw data
    
    Dim rawSheet As Worksheet
    Dim readingSheet As Worksheet
    Dim rawRange, rawRow As Range
    
    Dim cursorBegins, rowCount, codeASCII, startingColumn, currentSize, currentRow, rawCurrentRow, absoluteMaximum As Integer
    Dim validString, invalidString, statusString, validityString As String
    Dim messageBox As VbMsgBoxResult
    
    Set rawSheet = Worksheets("Data")
    Set readingSheet = Worksheets("Readings")
    
    ' Switch Data sheet view, this may be a bug regarding auto-filter
    rawSheet.Select
    
    Set rawRange = rawSheet.Range("A3:A" & Range("A65536").End(xlUp).Row) 'Row shifts down by 1 row. VERSION 1.1
    Set rawRow = rawRange.SpecialCells(xlCellTypeVisible)
    ' Obtain current rows
    rowCount = rawRange.SpecialCells(xlCellTypeVisible).Count
    
    ' Clear all content before working
    
    readingSheet.Cells.ClearContents
    
    ' All essential settings
    rawCurrentRow = 0 '
    cursorBegins = 4 ' this is for the reading sheet
    currentRow = cursorBegins
    startingColumn = 65 ' denotes A
    currentSize = 0 ' the size starts at 0
    absoluteMaximum = 5000 + rawCurrentRow ' limit the application to handle a specific amount of records
    codeASCII = startingColumn

    ' Staff ID at the beginning
    ' Staff Information
    readingSheet.Range("A1").Value = "Record Readings"
    readingSheet.Range("A2").Value = "STAFF ID"
    readingSheet.Range("B2").Value = Range("D" & rawRow.Cells(1).Row).Value
    readingSheet.Range("C2").Value = "Count"
    readingSheet.Range("D2").Formula = "=COUNTA(A" & cursorBegins & ":A" & absoluteMaximum + cursorBegins & ")"
    'add header
    readingSheet.Range("A3").Value = "Time"
    readingSheet.Range("B3").Value = "Machine"
    readingSheet.Range("C3").Value = "Status"
    readingSheet.Range("D3").Value = "Considered"

    
    
    Do While currentSize < rowCount
        rawCurrentRow = rawCurrentRow + 1 ' this is for the raw data
        '==============================================================
        ' validation
        '==============================================================
        ' The cursor will go on forever in case of both id and month are different, this is generally caused by bad filter
        ' therefore a triggering sum will be imposed
        If rawCurrentRow > absoluteMaximum Then
            'Halt the operation
            'MsgBox "Please Refine the ID selection before continuing"
                ' Clear all destination contnt
            'readingSheet.Cells.ClearContents
            'Exit Sub ' ask user to refine the id selection, as no filter has been applied apparently
        End If
        'Validation
        ' Skip or Stop if appropriate
        
        ' check staff id
        If Not Range("D" & rawRow.Cells(rawCurrentRow).Row).Value = Range("D" & rawRow.Cells(1).Row).Value Then
            ' check whether there is a difference in month too
            'If Not DateDiff("m", Range("B" & rawRow.Cells(rawCurrentRow).Row).Value, Range("B" & rawRow.Cells(1).Row).Value) = 0 Then
                '
           ' Else
                GoTo NextIteration
            'End If
        End If
        ' check month , same staff id
        If Not DateDiff("m", Range("A" & rawRow.Cells(rawCurrentRow).Row).Value, Range("A" & rawRow.Cells(1).Row).Value) = 0 Then
            'MsgBox "Please Refine the month selection before continuing"
            'MsgBox "A" & rawRow.Cells(rawCurrentRow).Row & " " & "A" & rawRow.Cells(1).Row
            'MsgBox Range("A" & rawRow.Cells(rawCurrentRow).Row).Value & " " & Range("A" & rawRow.Cells(1).Row).Value
            ' Clear all destination contnt
            'Exit Sub ' ask user to refine the month selection
        End If
        
        'beginning of the line
        
        'copy time and machine id to reading sheet
        'MsgBox Range("B" & rawRow.Cells(rawCurrentRow).Row).Value
        readingSheet.Range(Chr(codeASCII) & currentRow).Value = Range("A" & rawRow.Cells(rawCurrentRow).Row).Value
        codeASCII = codeASCII + 1
        readingSheet.Range(Chr(codeASCII) & currentRow).Value = Range("C" & rawRow.Cells(rawCurrentRow).Row).Value
        'analysis status code
        ' There is no difference identifying mid-day leave
        Select Case Range("B" & rawRow.Cells(rawCurrentRow).Row).Value
            Case "I"
                statusString = "Entry"
               
            Case "O"
                statusString = "Exit"
                
            Case "1"
                'statusString = "Return"
                statusString = "Entry"
               
            Case "0"
                'statusString = "Leave"
                statusString = "Exit"
                
        End Select
            ' Arrangements for 2XX entries, odd numbers are EXIT and even numbers are ENTRY
        If readingSheet.Range("B" & currentRow).Value >= 200 And readingSheet.Range("B" & currentRow).Value < 300 Then
            ' check for odd/even numbers
            If (readingSheet.Range("B" & currentRow).Value Mod 2) = 0 Then
                ' even : regard this field as entry
                statusString = "Entry"
            Else
                ' odd : regard this field as exit
                statusString = "Exit"
            End If
        End If
        codeASCII = codeASCII + 1
        readingSheet.Range(Chr(codeASCII) & currentRow).Value = statusString

        ' return to first column
        codeASCII = startingColumn
        currentRow = currentRow + 1 ' this is for the reading sheet
        currentSize = currentSize + 1
NextIteration:
        
        ' end of the line
    Loop
    
' Validate the results
validateStatus ' disable in development mode
readingSheet.Select
End Sub
Sub validateStatus()
    ' This sub validates each record on readings worksheet, and to determine they are valid or not
    
    Dim readingSheet, settingSheet As Worksheet
    Dim settingParameter As Range
    Dim rowBegins, rowCount, interval, i, dayCount, addedRow, entryCount, exitCount As Integer
    Dim theStatus, validityString, invalidString, validString As String
    Dim entered, outed, validateAgain As Boolean
    
    Set readingSheet = Sheets("Readings")
    Set settingSheet = Sheets("Settings")
    
    rowBegins = 4
    entered = False
    outed = False
    validateAgain = False
    validString = "Valid"
    invalidString = "Invalid"
    
    ' Look for interval settings
    Set settingParameter = settingSheet.Range("A2:A100").Find(what:="Interval", After:=settingSheet.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    interval = settingSheet.Range("B" & settingParameter.Row).Value
Validation:
    ' obtain current row count
    rowCount = readingSheet.Range("D2").Value
    ' set valid row count
    readingSheet.Range("E2").Value = "Valid"
    readingSheet.Range("F2").Formula = "=COUNTIF(D" & rowBegins & ":D" & rowBegins + rowCount & ",""Valid"")"
    ' sorting goes first
    ' Sorting goes here...
    readingSheet.Range("A" & rowBegins & ":E" & rowBegins + rowCount).Sort key1:=readingSheet.Range("A" & rowBegins), order1:=xlAscending, Header:=xlNo
    

    ' Validate the records
    ' read the total record number
    For i = 0 To (rowCount - 1)

        ' read current status
        theStatus = readingSheet.Range("C" & rowBegins + i).Value
        
        ' Status analysis
        Select Case theStatus
            Case "Entry"
                If entered And i <> 0 Then ' The first record of entry is always valid
    '                If (DateDiff("d", Range("B" & rawRow.Cells(rawCurrentRow).Row).Value, Range("B" & rawRow.Cells(rawCurrentRow - 1).Row).Value) = 0 Or rawCurrentRow = rowCount) Then
                        validityString = invalidString
    '                End If
    '               It may still be valid, provided this occurs in the next day...
                    If (DateDiff("d", readingSheet.Range("A" & rowBegins + i).Value, readingSheet.Range("A" & rowBegins + i - 1).Value) <> 0) Then
                        validityString = validString
                    End If
                Else
                    validityString = validString
                    entered = True
                End If
            Case "Exit"
                If Not entered Then
                    validityString = invalidString
                Else
                    validityString = validString
                    entered = False
                End If
            Case "Leave"
                 If outed Or Not entered Then
                   validityString = invalidString
                Else
                   validityString = validString
                    outed = True
                End If
            Case "Return"
                If Not outed Or Not entered Then
                    validityString = invalidString
                Else
                    validityString = validString
                    outed = False
                End If
        End Select
        
        ' Invalidated Entry Mitigation: The record may be valid because the data status entry is incorrect.
        If (validityString = invalidString) Then
            ' DO NOT CHANGE 2XX entries - They are always correct
            If readingSheet.Range("B" & rowBegins + i).Value >= 200 And readingSheet.Range("B" & rowBegins + i).Value < 300 Then GoTo Changed
            ' change from exit to entry
            If readingSheet.Range("C" & rowBegins + i).Value = "Exit" Then
                'MsgBox ("Row:" & (rowBegins + i) & "Min Diff:" & DateDiff("n", readingSheet.Range("A" & rowBegins + i).Value, readingSheet.Range("A" & rowBegins + i - 1).Value))
                ' first day of record
                If (DateDiff("d", readingSheet.Range("A" & rowBegins + i).Value, readingSheet.Range("A" & rowBegins + i - 1).Value) <> 0) Then
                    'validityString = validString
                    readingSheet.Range("C" & rowBegins + i).Value = "Entry"
                    validateAgain = True
                    GoTo Changed
                    'MsgBox ("first day of record:" & i)
                End If
                ' a duplicated Exit entry presented subject to sufficient interval (according to settings)

                If (DateDiff("n", readingSheet.Range("A" & rowBegins + i).Value, readingSheet.Range("A" & rowBegins + i - 1).Value) < (-interval)) Then
                    'validityString = validString
                    readingSheet.Range("C" & rowBegins + i).Value = "Entry"
                    validateAgain = True
                    GoTo Changed
                    'MsgBox ("duplicated Exit entry:" & i)
                End If
            End If
            ' change from entry to exit
            If readingSheet.Range("C" & rowBegins + i).Value = "Entry" Then
                ' a duplicated Entry entry presented subject to sufficient interval (according to settings)
                If (DateDiff("n", readingSheet.Range("A" & rowBegins + i).Value, readingSheet.Range("A" & rowBegins + i - 1).Value) < (-interval)) Then
                    'validityString = validString
                    readingSheet.Range("C" & rowBegins + i).Value = "Exit"
                    validateAgain = True
                    GoTo Changed
                    'MsgBox ("duplicated Exit entry:" & i)
                End If
            End If
        End If
Changed:
        readingSheet.Range("D" & rowBegins + i).Value = validityString
    Next i
        ' Prompt for Re-validation
    If validateAgain Then
        validateAgain = False ' change back to not valdiate again
        GoTo Validation
    End If
    '=============================================
    ' Check for correct row count
    '=============================================
    ' Each day must have an odd number of row entries, if not a new row is inserted
    ' the entire table will look again
    dayCount = 1
    addedRow = 0
    For i = rowBegins To (rowBegins + rowCount - 1)
        If shtReadings.Range("D" & i).Value = "Valid" Then
            If shtReadings.Range("C" & i).Value = "Entry" Then
                entryCount = entryCount + 1
            Else
                exitCount = exitCount + 1
            End If
            If i = (rowBegins + rowCount - 1) Then
                GoTo checkCount
            Else
                If DateDiff("d", shtReadings.Range("A" & i).Value, shtReadings.Range("A" & i + 1).Value) = 0 Or (DateDiff("d", shtReadings.Range("A" & i).Value, shtReadings.Range("A" & i + 1).Value) = 1 And shtReadings.Range("C" & i + 1).Value = "Exit") Then
                    dayCount = dayCount + 1
                Else
checkCount:
                    'MsgBox shtReadings.Range("A" & i).Value & ": " & dayCount
                    If dayCount Mod 2 = 1 Then
                        ' an extra row must be added!
                        'MsgBox shtReadings.Range("A" & i).Value & ": " & dayCount
                        shtReadings.Range("A" & rowBegins + rowCount + addedRow).Value = shtReadings.Range("A" & i).Value
                        shtReadings.Range("B" & rowBegins + rowCount + addedRow).Value = 900
                        If entryCount > exitCount Then
                            shtReadings.Range("C" & rowBegins + rowCount + addedRow).Value = "Exit"
                        Else
                            shtReadings.Range("C" & rowBegins + rowCount + addedRow).Value = "Entry"
                        End If
                        addedRow = addedRow + 1
                        validateAgain = True
                    End If
                    dayCount = 1
                    entryCount = 0
                    exitCount = 0

                End If
            End If
        End If
    Next i
' Prompt for Re-validation
    If validateAgain Then
        validateAgain = False ' change back to not valdiate again
        GoTo Validation
    End If
End Sub

Sub obtainData()
' This sub gathers data in accordance to id and very specific date range ( in month)
    Dim readingSheet, settingSheet, dataSheet As Worksheet
    Dim settingParameter As Range
    Dim id, Year, Month, nextYear, nextMonth As Integer
    
    Set readingSheet = Sheets("Readings")
    Set settingSheet = Sheets("Settings")
    Set dataSheet = Sheets("Data")
' Obtain necessary configurations
' id
    Set settingParameter = settingSheet.Range("A2:A100").Find(what:="id", After:=settingSheet.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    id = settingSheet.Range("B" & settingParameter.Row).Value
' date
    Year = shtReadings.cmbYear.Text ' Sheet3 stands for Readings in the code
    Month = shtReadings.cmbMonth.Text
' Filter content from data sheet
    ActiveWorkbook.RefreshAll ' Refresh all content first
    ' select date range
    ' some little preparations - arrange next month
    nextMonth = Month + 1
    nextYear = Year
    ' In case of December - switch to next year
    If nextMonth = 13 Then
        nextMonth = 1
        nextYear = Year + 1
    End If
    shtData.ListObjects("Table_att_v").Range.AutoFilter Field:=1, Criteria1:=">=" & Month & "/1/" & Year, Operator:=xlAnd, Criteria2:="<" & nextMonth & "/1/" & nextYear
    ' select id
    shtData.ListObjects("Table_att_v").Range.AutoFilter Field:=4, Criteria1:=id
' Perform read data ( from current function)
    prepareReadings
End Sub

Sub overTimeCountCheck()
  ' This sub finds all overtime working hours and place them into overtime worksheet
  ' The worksheet has a designed maximum of 130 rows
  ' It checks overtime against two main criterias: the significance and overtime setting
  ' overtime = -1 denotes automatically approval, but the actual overtime period must exceed the significance to quality
  ' Therefore there are several STATUS after the overtime check...
  ' INSIGN - insignificance overtime period, manual approval possible
  ' MANUAL - manual approval required
  ' AUTO   - overtime approved automatically, manual revoke possible
  ' INVALID - invalid entirs, normally causes by insufficient entires (forget to register)
  ' to assist manual approval, FOUND status is provided to explain the nature of overtime
  ' BEFORE - early shift
  ' AFTER  - late shift, usually working at night
  ' INTER  - noon break, typically for lunch time
  
    '==============================================================
    ' prepare variables
    '==============================================================
     Dim lunch, expected, workTime As Long
     Dim i, j, ROW_START, READ_ROW_START, interval, significance, currentRow, overTimeSetting, weight As Integer
     Dim startTime, endTime, timeDiff As Double
     Dim valid, entered, jumpBack As Boolean
     Dim statusString, currentDate, entryTime, nextDay, overTimeStart, overTimeEnd As String
     Dim settingParameter As Range
     
     
    ROW_START = 3
    READ_ROW_START = 4

    
    Set settingParameter = shtSettings.Range("A2:A100").Find(what:="Start", After:=shtSettings.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    startTime = (shtSettings.Range("B" & settingParameter.Row).Value)
    Set settingParameter = shtSettings.Range("A2:A100").Find(what:="End", After:=shtSettings.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    endTime = (shtSettings.Range("B" & settingParameter.Row).Value)
    Set settingParameter = shtSettings.Range("A2:A100").Find(what:="Lunch", After:=shtSettings.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    lunch = shtSettings.Range("B" & settingParameter.Row).Value
    Set settingParameter = shtSettings.Range("A2:A100").Find(what:="Interval", After:=shtSettings.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    interval = shtSettings.Range("B" & settingParameter.Row).Value
    Set settingParameter = shtSettings.Range("A2:A100").Find(what:="Significance", After:=shtSettings.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    significance = shtSettings.Range("B" & settingParameter.Row).Value
    Set settingParameter = shtSettings.Range("A2:A100").Find(what:="Overtime", After:=shtSettings.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    overTimeSetting = shtSettings.Range("B" & settingParameter.Row).Value
    expected = DateDiff("n", startTime, endTime) - lunch
    
    '===============================================================
    ' clear the current content
    '===============================================================
    shtOvertime.Range("A3:H133").ClearContents
    shtOvertime.Range("A4:H133").Borders.LineStyle = XlLineStyle.xlLineStyleNone

    ' clear the bottom format
    ' obtain data from each reading line (starting from 4)
    
    i = READ_ROW_START
    j = ROW_START
    '===============================================================
    ' Add header
    '===============================================================
    Set settingParameter = shtSettings.Range("A2:A100").Find(what:="Id", After:=shtSettings.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    shtOvertime.Range("D1").Value = shtSettings.Range("B" & settingParameter.Row).Value
    shtOvertime.Range("D1").NumberFormat = "000"
    Set settingParameter = shtSettings.Range("A2:A100").Find(what:="Name", After:=shtSettings.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    shtOvertime.Range("E1").Value = shtSettings.Range("B" & settingParameter.Row).Value
    '================================================================
    ' Looking at each valid entries
    '================================================================
     Do While shtReadings.Range("A" & i).Value <> ""
        ' check for valid only result
        If shtReadings.Range("D" & i).Value = "Valid" Then
        valid = False ' false the status first
            ' check for entry / exit result
            Select Case shtReadings.Range("C" & i).Value
                Case "Entry"
                ' entry : check for start
                ' entry : if earlier than start, list out as BEFORE overtime
                If (startTime) > TimeValue(shtReadings.Range("A" & i).Value) Then
                    valid = True
                    timeDiff = DateDiff("n", TimeValue(shtReadings.Range("A" & i).Value), (startTime))
                    shtOvertime.Range("F" & j).Value = "BEFORE"
                    overTimeStart = Format(shtReadings.Range("A" & i).Value, "HH:nn")
                    overTimeEnd = Format(startTime, "HH:nn")
                End If
                ' register for entry
                If entered = False Then
                    entered = True
                    If (startTime) > TimeValue(shtReadings.Range("A" & i).Value) Then
                        entryTime = Format(shtReadings.Range("A" & i).Value, "yyyy-MM-dd") & " " & Format(startTime, "hh:nn") ' excessive time is chopped off.
                    Else
                        entryTime = shtReadings.Range("A" & i).Value
                    End If
                End If
                Case "Exit"
                ' register for exit and add up the counts
                If entered Then
                    entered = False
                    If Format(entryTime, "yyyy-MM-dd") & " " & Format(endTime, "HH:nn") < Format(shtReadings.Range("A" & i).Value, "yyyy-MM-dd HH:nn") Then
                        workTime = workTime + DateDiff("n", entryTime, Format(entryTime, "yyyy-MM-dd") & " " & Format(endTime, "HH:nn")) ' excessive time is chopped off.
                    Else
                        workTime = workTime + DateDiff("n", entryTime, shtReadings.Range("A" & i).Value)
                    End If
                    ' if the exit is last of the day (represented by either
                    'If shtReadings.Range("A" & i + 1).Value = "" Then
                        'nextDay = "1/" & shtReadings.cmbMonth + 1 & "/" & shtReadings.cmbYear ' the last row of reading is empty, so a next month record is append to ally error
                    'Else
                        'nextDay = shtReadings.Range("A" & i + 1).Value
                    'End If
                    ' enter the intermediate overtime check if next record is next day
                    If DateDiff("d", entryTime, nextDay) >= 1 Then
                        ' check for additional workhour
                        'If (workTime - expected) >= 0 Then
                            ' apply another line for intermediate overtimes
                            timeDiff = workTime - expected
                            overTimeStart = ""
                            overTimeEnd = ""
                            shtOvertime.Range("F" & j).Value = "INTER"
                            jumpBack = True
                            ' reset workhour
                            workTime = 0
                            GoTo overTimeStatus
                        
                        'End If
                        ' reset workhour
                        workTime = 0
                    ' another check for absolute last day of entry, including overnight works
                    
                    End If
                End If
jumpBack:
                ' exit : check for end
                ' exit : if later than end, list out as AFTER overtime
                If Format(entryTime, "yyyy-MM-dd") & " " & Format(endTime, "HH:nn") < Format(shtReadings.Range("A" & i).Value, "yyyy-MM-dd HH:nn") Then
                    valid = True
                    timeDiff = DateDiff("n", Format(entryTime, "yyyy-MM-dd") & " " & Format(endTime, "HH:nn"), Format(shtReadings.Range("A" & i).Value, "yyyy-MM-dd HH:nn"))
                    shtOvertime.Range("F" & j).Value = "AFTER"
                    overTimeStart = Format(endTime, "HH:nn")
                    overTimeEnd = Format(shtReadings.Range("A" & i).Value, "HH:nn")
                End If
            End Select

            If valid Then
overTimeStatus:
                shtOvertime.Range("A" & j).Value = Format(entryTime, "yyyy-MM-dd")
                shtOvertime.Range("B" & j).Value = overTimeStart
                shtOvertime.Range("C" & j).Value = overTimeEnd
                shtOvertime.Range("D" & j).Value = timeDiff
                ' check for validity
                    If (overTimeSetting = -1 Or timeDiff <= overTimeSetting) And timeDiff >= significance Then
                        statusString = "AUTO"
                        weight = 1
                    ElseIf timeDiff < (significance) Then
                        statusString = "INSIGN"
                        weight = 0
                    ElseIf timeDiff > overTimeSetting Then
                        statusString = "MANUAL"
                        weight = 0
                    End If
                    shtOvertime.Range("E" & j).Value = statusString
                    
                    ' add the weigth factor and count formula
                    shtOvertime.Range("G" & j).Value = weight
                    shtOvertime.Range("H" & j).Formula = "=G" & j & "*D" & j
                    ' move to next
                    j = j + 1
                    ' reset
                    valid = False
                    If jumpBack Then
                        jumpBack = False
                        GoTo jumpBack
                    End If
            End If
        End If

        
        ' move to next result row
        i = i + 1
     Loop
     '==============================================
     ' Apply bottom arrangements
     '==============================================
    With shtOvertime.Range("A" & j & ":H" & j).Borders(xlEdgeTop)
        .LineStyle = XlLineStyle.xlContinuous
        .weight = xlThick
    End With
    shtOvertime.Range("D" & ROW_START & ":D" & j - 1).NumberFormat = "0"
    shtOvertime.Range("D" & j).Formula = "=SUM(D" & ROW_START & ":D" & j - 1 & ")"
    shtOvertime.Range("H" & j).Formula = "=SUM(H" & ROW_START & ":H" & j - 1 & ")"
shtOvertime.Select
End Sub


