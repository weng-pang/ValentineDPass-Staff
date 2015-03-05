Attribute VB_Name = "modGetData"
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
   'Dim Db As ADODB.Connection
   'Dim Rs As ADODB.Recordset
   Dim Ws As Worksheet
   Dim rowCount, rowStarts As Integer
   Dim Path, SQLQuery, ConnectString, id As String
   'Dim ReturnArray
   Dim settingParameter As Range


   'This line will define the Object "Ws" as Sheets("Sheet1")
   'The purpose of this is to save typing Sheets("Sheet1")
   'over and over again
    Set Ws = shtReadings
    
    ' clear current contents
    rowStarts = 4
    If rowStarts <> Ws.UsedRange.Rows.count Then
        Ws.Rows(rowStarts & ":" & Ws.UsedRange.Rows.count).Delete
    End If
   'Set the Path to the database. This line is useful because
   'if your database is in another location, you just need to change
   'it here and the Path Variable will be used throughout the code
   Set settingParameter = shtSettings.Range("A2:A100").Find(what:="File", After:=shtSettings.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    Path = (shtSettings.Range("B" & settingParameter.Row).Value)
    Set settingParameter = shtSettings.Range("A2:A100").Find(what:="Id", After:=shtSettings.Range("A2:A100").Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    id = (shtSettings.Range("B" & settingParameter.Row).Value)

   '====================================
' For Retrieving Information from DPASS server
    Dim dpassClient As New RestClient
    Dim dpassRequest As New RestRequest
    Dim dpassResponse As RestResponse

    dpassRequest.Format = formurlencoded
    'MsgBox "{""id"": " & id & ",""startTime"":" & startDate & ",""endTime"":" & endDate & " }"
    dpassRequest.AddParameter "key", "b83bc0d2-b3ca-4dc6-b47f-fec816d2a2f9"
    dpassRequest.AddParameter "content", "{""id"": " & id & ",""startTime"":""" & startDate & """,""endTime"":""" & endDate & """ }"
    dpassClient.BaseUrl = "https://kaugebra.com/dpass-rest/find"
    'dpassRequest.AddBody dpassRequestContent
    dpassRequest.Method = httpPOST
    
    Set dpassResponse = dpassClient.Execute(dpassRequest)
    
    ' Check the reponse
    If dpassResponse.StatusCode = Ok Then
        Dim Route As Dictionary

        ' Apply content to Data Sheet
        '-------------------------------------------------------
        'Set up the class connection
    Dim clsJSON As clsJSParse
    Set clsJSON = New clsJSParse
    Dim intindex As Integer
    Dim strStatus, strErrStatus As String
    
    'The first step is to load the JSON information
    'The loaded information will then be available through the clsJSON.Key and clsJSON.Value data pairs
    clsJSON.LoadString dpassResponse.Content
    'MsgBox Len(dpassResponse.Content)
    strStatus = "Loaded: " & clsJSON.NumElements & " Elements"
    
    'Convert status data to text
    Select Case clsJSON.err
        Case -1
            strErrStatus = "JSON string not loaded."
        Case -2
            strErrStatus = "JSON string cannot be fully parsed."
        Case 1
            strErrStatus = "JSON string has been parsed."
    End Select
    
    'Display status data
    strStatus = strStatus & Chr(10) & strErrStatus & Chr(10) & "See Sheet2 for parsed data"
    Sheet1.Cells(10, 2).Value = strStatus

    'Get all the elements of the parsed JSON text and put the data in cells (in shtReadings)
    rowStarts = 4
    rowCount = 0
    For intindex = 1 To clsJSON.NumElements
        Select Case intindex Mod 10
            Case 3
                shtReadings.Cells(rowStarts + rowCount, 1).Value = clsJSON.Value(intindex)
            Case 4
                shtReadings.Cells(rowStarts + rowCount, 2).Value = clsJSON.Value(intindex)
            Case 5
                shtReadings.Cells(rowStarts + rowCount, 3).Value = clsJSON.Value(intindex)
            Case 0
                rowCount = rowCount + 1
        End Select
        'Sheet1.Cells(intindex, 1).Value = clsJSON.Key(intindex)
        'Sheet1.Cells(intindex, 2).Value = clsJSON.Value(intindex)
    Next
        '-------------------------------------------------------
        
        
        'Ws.Range("L3").Value = dpassResponse.Content
        'Dim recordSet As Variant
        'Set recordSet = dpassResponse.Data
        'MsgBox dpassResponse.Content
        'For Each recordData In recordSet
            'Ws.Range("A" & rowStarts + i).Value = recordData.Data("dateTime")
            'Ws.Range("B" & rowStarts + i).Value = recordData.Data("machineId")
            'Ws.Range("C" & rowStarts + i).Value = recordData.Data("entryId")
            'i = i + 1
        'Next recordData
        'Ws.Range("A4").Value = dpassResponse.Data(1)("dateTime")
        'Ws.Range("A4").Value = singleRecord.Data("dateTime")
        
        ' call relevant function for processing
        ' TODO uncomment this part
        Call prepareReadings(rowCount + rowStarts, id)
    Else
        Ws.Range("A4").Value = dpassResponse.Content
    End If
    '======================================
 
   
   
   ' Close Connections
   'Rs.Close
   'Db.Close
   
   
End Sub

