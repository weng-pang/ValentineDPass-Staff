VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelection 
   Caption         =   "Date Selection"
   ClientHeight    =   2484
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   3768
   OleObjectBlob   =   "frmSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboDMonthEnd_Change()
    cboDDayEnd.RowSource = selectMonthToDay(cboDMonthEnd.Value, cboDYearEnd.Value)
    cboDDayEnd.ListIndex = 0
End Sub

Private Sub cboDMonthStart_Change()
    cboDDayStart.RowSource = selectMonthToDay(cboDMonthStart.Value, cboDYearStart.Value)
    cboDDayStart.ListIndex = 0
End Sub

Private Function selectMonthToDay(Month As Integer, Year As Integer) As String
    Select Case Month
        Case 1, 3, 5, 7, 8, 10, 12
            selectMonthToDay = "Day31List"
        Case 4, 6, 9, 11
            selectMonthToDay = "Day30List"
        Case 2
            If Year Mod 4 = 0 Then
                If Year Mod 100 = 0 Then
                    If Year Mod 400 = 0 Then
                        selectMonthToDay = "Day28List"
                    Else
                        selectMonthToDay = "Day29List"
                    End If
                Else
                    selectMonthToDay = "Day29List"
                End If
            Else
                selectMonthToDay = "Day28List"
            End If
    End Select
End Function

Private Sub cmdDayGet_Click()
    Dim startDate, endDate As String
    ' get start and end date
    startDate = cboDYearStart.Value & "-" & cboDMonthStart.Value & "-" & cboDDayStart.Value
    endDate = cboDYearEnd.Value & "-" & cboDMonthEnd.Value & "-" & cboDDayEnd.Value
    ' validation - ensure correct date order
    'MsgBox startDate
    'MsgBox Format(startDate, "yyyy/mm/dd")
    'MsgBox DateDiff("d", Format(startDate, "yyyy/mm/dd"), Format(endDate, "yyyy/mm/dd"))
    If DateDiff("d", startDate, endDate) >= 0 Then
        ' call the date method
        GetData Format(startDate, "yyyy-mm-dd") & " 00:00", Format(endDate, "yyyy-mm-dd") & " 23:59"
        End
    Else
        MsgBox "Please enter date order correctly", vbCritical + vbOKOnly
    End If
End Sub

Private Sub cmdMonthGet_Click()
    ' call the date method
    GetData cboMYear.Value & "-" & cboMMonth.Value & "-01 00:00", cboMYear.Value & "-" & cboMMonth.Value & "-" & Mid(selectMonthToDay(cboMMonth.Value, cboMYear.Value), 4, 2) & " 23:59"
    'GetData cboMMonth.Value & "/1/" & cboMYear.Value, cboMMonth.Value & "/" & Mid(selectMonthToDay(cboMMonth.Value, cboMYear.Value), 4, 2) & "/" & cboMYear.Value
    End
End Sub

Private Sub UserForm_Initialize()
    ' load up current year and month into system
    ' for Month Range Page
    cboMYear.Value = Year(Date)
    cboMMonth.Value = Month(Date)
    ' for Day Range Page
    cboDYearStart.Value = Year(Date)
    cboDMonthStart.Value = Month(Date)
    cboDMonthStart_Change
    cboDYearEnd.Value = Year(Date)
    cboDMonthEnd.Value = Month(Date)
    cboDMonthEnd_Change
End Sub
