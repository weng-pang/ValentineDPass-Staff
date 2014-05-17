Attribute VB_Name = "Module2"
Option Explicit
Sub DateSorting()
Attribute DateSorting.VB_Description = "Issue in 1.3 mitigation : to Provide preliminary date sorting"
Attribute DateSorting.VB_ProcData.VB_Invoke_Func = " \n14"
'
' DateSorting Macro
' Issue in 1.3 mitigation : to Provide preliminary date sorting
'

'
    ActiveSheet.ListObjects("Table_att_v").Range.AutoFilter Field:=1
    ActiveWorkbook.Worksheets("Data").ListObjects("Table_att_v").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Data").ListObjects("Table_att_v").Sort.SortFields. _
        Add Key:=Range("A2:A1889"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Data").ListObjects("Table_att_v").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub openFrmSelection()
    frmSelection.Show
End Sub

Sub GetData(startDate As String, endDate As String)
'This sub will retrieve all the data in the "Customers" table in
'Northwind

   'Declare variables
   Dim Db As ADODB.Connection
   Dim Rs As ADODB.Recordset
   Dim Ws As Worksheet
   Dim i, rowStarts As Integer
   Dim Path, SQLQuery, ConnectString, id As String
   Dim ReturnArray
   Dim settingParameter As Range

   'This line will define the Object "Ws" as Sheets("Sheet1")
   'The purpose of this is to save typing Sheets("Sheet1")
   'over and over again
    Set Ws = shtData
    
    ' clear current contents
    rowStarts = 3
    If rowStarts <> Ws.UsedRange.Rows.Count Then
        Ws.Rows(rowStarts & ":" & Ws.UsedRange.Rows.Count).Delete
    End If
   'Set the Path to the database. This line is useful because
   'if your database is in another location, you just need to change
   'it here and the Path Variable will be used throughout the code
   Set settingParameter = shtSettings.Range("A2:A100").Find(what:="File", After:=shtSettings.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    Path = (shtSettings.Range("B" & settingParameter.Row).Value)
    Set settingParameter = shtSettings.Range("A2:A100").Find(what:="Id", After:=shtSettings.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    id = (shtSettings.Range("B" & settingParameter.Row).Value)
   ' Connection Settings
   ConnectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Path & "; Persist Security Info=False;Jet OLEDB:Database Locking Mode=0;"
   
   SQLQuery = "SELECT  checktime, checktype, sensorid ,u.badgenumber " & _
            "FROM userinfo u, checkinout c " & _
            "WHERE c.userid = u.userid " & _
            "AND u.badgenumber ='" & id & "' " & _
            "AND checktime >= #" & startDate & "# AND checktime <= #" & endDate & "#;"
    ' Date format: MM/DD/YYYY
   ' Initiate Connections
   Set Db = New ADODB.Connection
   Set Rs = New ADODB.Recordset
   
   Db.Open ConnectString
   
   Rs.Open SQLQuery, Db
   ' Apply content to Data Sheet
   Ws.Range("A3").CopyFromRecordset Rs
   ' Close Connections
   Rs.Close
   Db.Close
   ' call relevant function for processing
   prepareReadings
   
End Sub

