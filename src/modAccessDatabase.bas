Attribute VB_Name = "modAccessDatabase"
Public Enum ViewOrProc
    View
    Proc
End Enum

Public Sub CreateQuery(ByVal strDBPath As String, _
                       ByVal strSql As String, ByVal strQueryName As String, _
                       ByVal vpType As ViewOrProc)
   Dim objCat As Object
   Dim objCmd As Object

   On Error GoTo exit_point

   Set objCat = CreateObject(Class:="ADOX.Catalog")
   objCat.ActiveConnection = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                             "Data Source=" & strDBPath & ";" & _
                             "Jet OLEDB:Engine Type=5;" & _
                             "Persist Security Info=False;"

    Set objCmd = CreateObject(Class:="ADODB.Command")
    objCmd.CommandText = strSql

    If vpType = View Then
        Call objCat.Views.Append(name:=strQueryName, Command:=objCmd)
    ElseIf vpType = Proc Then
        Call objCat.Procedures.Append(name:=strQueryName, Command:=objCmd)
    End If

exit_point:
    Set objCat = Nothing

    If err.Number Then
        Call err.Raise(Number:=err.Number, Description:=err.Description)
    End If
End Sub

Public Sub ModifyQuery(ByVal strDBPath As String, _
                       ByVal strSql As String, ByVal strQueryName As String)
   Dim objCat As Object
   Dim objCmd As Object
   Dim vpType As ViewOrProc

   On Error GoTo exit_point

   Set objCat = CreateObject(Class:="ADOX.Catalog")
   objCat.ActiveConnection = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                             "Data Source=" & strDBPath & ";" & _
                             "Jet OLEDB:Engine Type=5;" & _
                             "Persist Security Info=False;"

    On Error Resume Next
        Set objCmd = objCat.Views(strQueryName).Command
        If Not objCmd Is Nothing Then
            vpType = View
        Else
            Set objCmd = objCat.Procedures(strQueryName).Command
            If Not objCmd Is Nothing Then
                vpType = Proc
            End If
        End If
    On Error GoTo exit_point

    If objCmd Is Nothing Then GoTo exit_point

    objCmd.CommandText = strSql

    If vpType = View Then
        Set objCat.Views(strQueryName).Command = objCmd
    ElseIf vpType = Proc Then
        Set objCat.Procedures(strQueryName).Command = objCmd
    End If

exit_point:
    Set objCat = Nothing

    If err.Number Then
        Call err.Raise(Number:=err.Number, Description:=err.Description)
    End If
End Sub

Public Sub DeleteQuery(ByVal strDBPath As String, ByVal strQueryName As String)
   Dim objCat As Object
   Dim lngCount As Long

   Set objCat = CreateObject(Class:="ADOX.Catalog")
   objCat.ActiveConnection = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                             "Data Source=" & strDBPath & ";" & _
                             "Jet OLEDB:Engine Type=5;" & _
                             "Persist Security Info=False;"
   
    With objCat
        lngCount = .Procedures.count + .Views.count
        On Error Resume Next
            Call .Procedures.Delete(strQueryName)
            Call .Views.Delete(strQueryName)
        On Error GoTo exit_point
        If .Procedures.count + .Views.count = lngCount Then
            err.Number = 3265
            err.Description = "Item cannot be found in the collection corresponding to the requested name or ordinal."
        End If
    End With

exit_point:
    Set objCat = Nothing
   
    If err.Number Then
        Call err.Raise(Number:=err.Number, Description:=err.Description)
    End If
End Sub

Public Function RunQuery(ByVal strDBPath As String, ByVal strQueryName As String, ParamArray parArgs() As Variant) As Object
    Dim objCmd As Object
    Dim objRec As Object
    Dim varArgs() As Variant

    Set objCmd = CreateObject("ADODB.Command")

    If UBound(parArgs) = -1 Then
        varArgs = VBA.Array("")
    Else
        varArgs = parArgs
    End If

    With objCmd
        .ActiveConnection = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                             "Data Source=" & strDBPath & ";" & _
                             "Jet OLEDB:Engine Type=5;" & _
                             "Persist Security Info=False;"
        .CommandText = strQueryName
        If UBound(parArgs) = -1 Then
            Set objRec = .Execute(Options:=4)
        Else
            varArgs = parArgs
            Set objRec = .Execute(Parameters:=varArgs, Options:=4)
        End If
    End With

    Set RunQuery = objRec

    Set objRec = Nothing
End Function

