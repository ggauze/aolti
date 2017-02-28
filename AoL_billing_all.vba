' Change this as needed

Public Const IntepreterRatesFile = "C:/Users/sgringauze.ISTREAMPLANET/Dropbox/Karla/Billing/SCCA Interpreter Rates List.xlsx"

' Some constants - modify as needed
Public Const SCCA_RATE = 48
Public Const SCCA_OVERCHARGE = 7
Public Const INTERPRETER_OVERCHARGE = 5

' Color index

Public Const LIGHT_GREEN = 50
Public Const LIGHT_BLUE = 23

' Modify the following constants if column order is changed'

' Billing table '

Public Const b_INTERPRETER = 1
Public Const b_STATUS = 2
Public Const b_LAST_NAME = 3
Public Const b_FIRST_NAME = 4
Public Const b_LANGUAGE = 5
Public Const b_U_NUMBER = 6
Public Const b_DATE = 7
Public Const b_S_START = 8
Public Const b_S_END = 9
Public Const b_S_MIN = 10
Public Const b_ARRIVAL_TIME = 11
Public Const b_A_START = 12
Public Const b_A_END = 13
Public Const b_A_MIN = 14
Public Const b_DEPARTMENT = 15
Public Const b_TYPE = 16
Public Const b_NOTES = 17
Public Const b_RH_UNITS = 18
Public Const b_AH_UNITS = 19
Public Const b_INTERPRATE = 20
Public Const b_RH_FEE_INTERP = 21
Public Const b_AH_FEE_INTERP = 22
Public Const b_INTERPTOTAL = 23
Public Const b_RH_FEE_SCCA = 24
Public Const b_AH_FEE_SCCA = 25
Public Const b_SCCATOTAL = 26
Public Const b_REASON_FOR_CHANGE = 27
Public Const b_CANC_REASON = 28

' internal use columns '
Public Const b_MIN2 = 29


' Intepreter table'

Public Const i_FIRST_NAME = 1
Public Const i_LAST_NAME = 2
Public Const i_FULL_NAME = 3
Public Const i_LOCATION = 4
Public Const i_RATE = 5
Public Const i_2HRMIN = 6

' Global variables
Dim IRates() As Variant
Dim LastRow As Variant

' Global initializaiton

Private Sub GlobalInit()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlManual

    ' Billing sheet initialization

    ActiveWorkbook.Sheets("InterpBilling").Select

    ' Work in a copy
    ActiveSheet.Copy After:=ActiveSheet

   ' Sort by Intepreter, Date, UNum, Start
    ActiveSheet.UsedRange _
        .Sort key1:=Columns(b_S_START), _
               order1:=xlAscending, Header:=xlYes, MatchCase:=False

    ActiveSheet.UsedRange _
        .Sort key1:=Columns(b_INTERPRETER), key2:=Columns(b_DATE), key3:=Columns(b_U_NUMBER), _
               order1:=xlAscending, Header:=xlYes, MatchCase:=False
       
    LastRow = Range("A" & Rows.Count).End(xlUp).Row

    Columns(b_S_START).NumberFormat = "h:mm AM/PM"
    Columns(b_S_END).NumberFormat = "h:mm AM/PM"
    Columns(b_ARRIVAL_TIME).NumberFormat = "h:mm AM/PM"
    Columns(b_A_START).NumberFormat = "h:mm AM/PM"
    Columns(b_A_END).NumberFormat = "h:mm AM/PM"

    ' !!! ATTENTION: The following line makes assumption about column locations in the table '
    ' !!! It clears all contents of the columns betwen RH Units and SCCATotal, inclusively

    Range(Cells(2, b_RH_UNITS), Cells(LastRow, b_SCCATOTAL)).Clear

    ' !!! ATTENTION: The following line makes assumption about column locations in the table '
    ' !!! It sets format to $ for all the columns betwen Interp RH Fee and SCCATotal, inclusively

    Range(Cells(2, b_RH_FEE_INTERP), Cells(LastRow, b_SCCATOTAL)).NumberFormat = "$#,##0.00"


    'Columns(b_INTERPRATE).NumberFormat = "$#,##0.00"
    'Columns(b_INTERPTOTAL).NumberFormat = "$#,##0.00"
    'Columns(b_SCCARATE).NumberFormat = "$#,##0.00"
    'Columns(b_SCCATOTAL).NumberFormat = "$#,##0.00"

    ' internal use columns
    Cells(1, b_MIN2).Value = "Is Min2"
    Columns(b_MIN2).Hidden = True

End Sub

Public Function IsWeekend(InputDate As Date) As Boolean
    Select Case Weekday(InputDate)
        Case vbSaturday, vbSunday
            IsWeekend = True
        Case Else
            IsWeekend = False
    End Select
End Function

' Validate the column titles

' A - 1 - Interpreter
' B - 2 - Status
' C - 3 - Last Name
' D - 4 - First Name
' E - 5 - Language
' F - 6 - U Number
' G - 7 - Date
' H - 8 - S Start
' I - 9 - S End
' J - 10 - S Min
' K - 11 - Arrival Time
' L - 12 - A Start
' M - 13 - A End
' N - 14 - A Min
' O - 15 - Department
' P - 16 - Type
' Q - 17 - Notes
' R - 18 - RH Units
' S - 19 - AH Units
' T - 20 - InterpRate
' U - 21 - Interp RH Fee
' V - 22 - Interp AH Fee
' W - 23 - InterpTotal
' X - 24 - SCCA RH Fee
' Y - 25 - SCCA AH Fee
' Z - 26 - SCCATotal
' AA - 27 - Reason for Change
' AB - 28 - Canc Reason

Function ValidateCaption(titles As Variant) As Boolean
    ' Validate each value
    Dim i As Long
    For i = LBound(titles) To UBound(titles)
        If LCase(Cells(1, i + 1).Value) <> LCase(titles(i)) Then
            MsgBox ("Cell " & Cells(1, i + 1).Address & " is expected to be " & titles(i) & " but instead it is " & Cells(1, i + 1).Value)
            ValidateCaption = False
            Exit Function
        End If
    Next i
    ValidateCaption = True
End Function

Function ValidateBillingCaption() As Boolean
    Dim titles As Variant
    titles = Array("Interpreter", "Status", "Last Name", "First Name", "Language", "U Number", "Date", "S Start", "S End", "S Min", _
                   "Arrival Time", "A Start", "A End", "A Min", "Department", "Type", "Notes", "RH Units", "AH Units", "InterpRate", _
                   "Interp RH Fee", "Interp AH Fee", "InterpTotal", "SCCA RH Fee", "SCCA AH Fee", "SCCATotal", "Reason for Change", "Canc Reason")
    ValidateBillingCaption = ValidateCaption(titles)
End Function

' Load IntepreterRatesFile and import it to the rates array
Function LoadInterpreterRates() As Boolean
    On Error Resume Next

    If Len(Dir(IntepreterRatesFile)) = 0 Then
          MsgBox "File " & IntepreterRatesFile & " does not exist"
          LoadInterpreterRates = False
          Exit Function
    Else
         On Error Resume Next
         Set WB = Workbooks.Open(IntepreterRatesFile)
         On Error GoTo 0
         If WB Is Nothing Then
            MsgBox IntepreterRatesFile & " is invalid, can't load", vbCritical
            LoadInterpreterRates = False
            Exit Function
         End If
    End If

    ' Sort the entries
    Dim R As Range
    ActiveWorkbook.Sheets(1).Select

    ' Validate caption
    Dim titles As Variant
    titles = Array("First Name", "Last Name", "First Name Last Name", "Location", "SCCARates", "2 Hour MIN")
    If ValidateCaption(titles) = False Then
        LoadInterpreterRates = False
        Exit Function
    End If

    Set R = ActiveCell.CurrentRegion
    R.Sort key1:=Columns(i_FULL_NAME), order1:=xlAscending, Header:=xlYes, MatchCase:=False
    IRates = R
    
    ActiveWorkbook.Close savechanges:=False

    LoadInterpreterRates = True

End Function

Function FindIntepreterIdx(Name As Variant, Department As Variant, LastFoundIdx As Variant) As Long

    If LastFoundIdx = -1 Then
       LastFoundIdx = LBound(IRates)
    End If
 
    For i = LastFoundIdx To UBound(IRates)
        If IRates(i, i_FULL_NAME) = Name And (IRates(i, i_LOCATION) = "" Or IRates(i, i_LOCATION) = Left(Department, Len(IRates(i, i_LOCATION)))) _
        Then
            FindIntepreterIdx = i
            Exit Function
        End If
    Next i
    FindIntepreterIdx = -1
End Function

' Initializing rates column
Private Sub InitRatesColumn()

    i = 2
    interpIdx = -1

    Do While i <= LastRow
        interpName = Cells(i, b_INTERPRETER).Value
        dept = Cells(i, b_DEPARTMENT).Value
        interpIdx = FindIntepreterIdx(interpName, dept, interpIdx)
        Do While i <= LastRow And Cells(i, b_INTERPRETER).Value = interpName
            If (interpIdx <> -1) Then
                Cells(i, b_INTERPRATE).Value = IRates(interpIdx, i_RATE)
                If Not IsEmpty(IRates(interpIdx, i_2HRMIN)) Then
                    Call SetValue(i, b_NOTES, vbRed, "MIN2")
                    Cells(i, b_MIN2).Value = True
                End If
            Else
                Cells(i, b_INTERPRATE).Interior.Color = vbRed
            End If
            
            i = i + 1
        Loop
    Loop

End Sub

Private Sub SetValue(Row As Variant, Column As Variant, Color As Long, Value As Variant)
    Cells(Row, Column).Value = Value
    Cells(Row, Column).Font.Color = Color
    Cells(Row, Column).Font.FontStyle = "Bold"
End Sub

Private Sub ConcatOrSetValue(Row As Variant, Column As Variant, Color As Long, Value As Variant)
    Dim newValue As Variant
    If IsEmpty(Cells(Row, Column).Value) Then
        newValue = Value
    Else
        newValue = CStr(Cells(Row, Column).Value) & ", " & CStr(Value)
    End If

    Call SetValue(Row, Column, Color, newValue)
End Sub

Function FindSeriesEnd(isScheduled As Variant, ByRef startIdx As Variant) As Variant

    If isScheduled Then
        START_TIME = b_S_START
        Duration = b_S_MIN
        END_TIME = b_S_END
        NOTE_SUFFIX = "_s"
    Else
        START_TIME = b_A_START
        Duration = b_A_MIN
        END_TIME = b_A_END
        NOTE_SUFFIX = "_a"
    End If

    i = startIdx

    interpreter = Cells(i, b_INTERPRETER).Value
    UNum = Cells(i, b_U_NUMBER).Value
    StartDate = Cells(i, b_DATE).Value
    startTime = Cells(i, START_TIME).Value
    endTime = Cells(i, END_TIME).Value

    Cells(i, START_TIME).Font.Color = vbBlue ' series start is blue'

    ' Check for after-hours and split if the appointment lasts pas '
    i = i + 1

    Do While i <= LastRow
        nextUNum = Cells(i, b_U_NUMBER).Value
        nextStartDate = Cells(i, b_DATE).Value
        nextInterpreter = Cells(i, b_INTERPRETER).Value
        
        If nextUNum = UNum And StartDate = nextStartDate And interpreter = nextInterpreter Then
        
            nextStartTime = Cells(i, START_TIME).Value
            
            ' First, calculate possible overlaps
            If DateDiff("n", endTime, nextStartTime) < 0 Then
                Call ConcatOrSetValue(i, b_NOTES, vbGreen, "OVERLAP" + NOTE_SUFFIX)
                nextStartTime = endTime
            End If
            
            Gap = DateDiff("n", endTime, nextStartTime)
            ' Adjust gap to minimum 60 min appt duration
            If DateDiff("n", startTime, endTime) < 60 And Gap > 60 Then
                Gap = DateDiff("n", DateAdd("n", 60, startTime), nextStartTime)
            End If

            If Gap <= 60 Then  ' if wait time is <= 60, this appointment is a part of the series'
                newEndTime = Cells(i, END_TIME).Value
                
                If Gap > 15 Then
                    Call ConcatOrSetValue(i, b_NOTES, vbGreen, "WT" + NOTE_SUFFIX)
                End If
                
                If newEndTime > endTime Then
                   endTime = newEndTime
                End If

                Cells(i - 1, b_INTERPRATE).Clear

            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
        i = i + 1
    Loop

    FindSeriesEnd = endTime
    startIdx = i
End Function

'
' Find a total duration of late cancelled appointments in the series
'
Function LCLDuration(startIdx As Variant, endIdx As Variant) As Variant
    LCLDuration = 0
    For i = startIdx To endIdx
        If LCase(Cells(i, b_STATUS)) = "lcl" Then
            LCLDuration = LCLDuration + Cells(i, b_S_MIN)
        End If
    Next i
End Function

'
' Find a max time in A_END columnt between the given indexes
'
Function GetHighlightIndex(startIdx As Variant, endIdx As Variant, endTime As Variant) As Variant
    GetHighlightIndex = startIdx
    maxTime = Cells(startIdx, b_A_END).Value
    For i = startIdx To endIdx
        If Cells(i, b_A_END).Value = endTime Then
            GetHighlightIndex = i
            Exit Function
        Else
            If Cells(i, b_A_END).Value > maxTime Then
                maxTime = Cells(i, b_A_END).Value
                GetHighlightIndex = i
            End If
        End If
    Next i
End Function

Private Sub FindSeries()

    i = 2
    Do While i <= LastRow
        startIdx = i
        startTimeSch = Round(Cells(i, b_S_START).Value, 15)
        startTimeAct = Round(Cells(i, b_A_START).Value, 15)
        endTimeSch = Round(FindSeriesEnd(True, i), 15)
        endTimeSchIdx = i - 1
        i = startIdx
        endTimeAct = Round(FindSeriesEnd(False, i), 15)
        endTimeActIdx = i - 1

        ' Calc if arrival time is later than the scheduled start
        isLateArrival = Cells(startIdx, b_ARRIVAL_TIME).Value > startTimeSch
        lateArrivalPenaltyInMinutes = WorksheetFunction.RoundUp(DateDiff("n", startTimeSch, Cells(startIdx, b_ARRIVAL_TIME).Value) / 15, 0) * 15

        endIdx = WorksheetFunction.Max(endTimeSchIdx, endTimeActIdx)
        startTime = WorksheetFunction.Min(startTimeSch, startTimeAct)
        endTime = WorksheetFunction.Max(endTimeSch, endTimeAct)

        ' Check if arrival time is later than the scheduled start
        If isLateArrival Then
            Duration = DateDiff("n", startTime, endTime) - lateArrivalPenaltyInMinutes
            Call ConcatOrSetValue(startIdx, b_NOTES, vbRed, "LA")
        Else
            Duration = WorksheetFunction.Max(DateDiff("n", startTime, endTime), 60)
        End If
        
        ' Handle MIN2 '
        If Duration < 120 And Cells(startIdx, b_MIN2) Then
            Duration = 120
            endTime = DateAdd("n", Duration, startTime)
        End If

        ' Handle MAX4 '
        schDuration = DateDiff("n", startTimeSch, endTimeSch)
        actDuration = DateDiff("n", startTimeAct, endTimeAct)

        If schDuration > 240 Then
            ' Find if there are LCLs in this series'
            actDuration = actDuration - LCLDuration(startIdx, endIdx)

            ' Handle late arrival
            If isLateArrival Then
                actDuration = actDuration - lateArrivalPenaltyInMinutes
            Else
                ' Appointment started later than scheduled
                If startTimeSch < startTimeAct Then
                    actDuration = DateDiff("n", startTimeSch, endTimeAct)
                End If
            End If

            If actDuration < 240 Then
                Duration = 240  ' MAX4 '
                endTime = DateAdd("n", Duration, startTime)
                Call ConcatOrSetValue(endIdx, b_NOTES, vbRed, "MAX4")
            Else
                Duration = actDuration
                endTime = endTimeAct
            End If

        End If

        ' Highlight series end
        If endTimeSch >= endTimeAct And endTime = endTimeSch Then
            Cells(endIdx, b_S_END).Interior.ColorIndex = LIGHT_BLUE
            Cells(endIdx, b_S_END).Font.Color = vbRed
        Else
            hlIdx = GetHighlightIndex(startIdx, i - 1, endTime)
            Cells(hlIdx, b_A_END).Interior.ColorIndex = LIGHT_BLUE
            Cells(hlIdx, b_A_END).Font.Color = vbRed
        End If

        ' Highlight series start
        If startTimeSch <= startTimeAct Then
            If Cells(startIdx, b_ARRIVAL_TIME).Value > startTimeSch Then
                Cells(startIdx, b_ARRIVAL_TIME).Interior.ColorIndex = LIGHT_GREEN
                Cells(startIdx, b_ARRIVAL_TIME).Font.Color = vbBlue
            Else
                Cells(startIdx, b_S_START).Interior.ColorIndex = LIGHT_GREEN
                Cells(startIdx, b_S_START).Font.Color = vbBlue
            End If
        Else
            Cells(startIdx, b_A_START).Interior.ColorIndex = LIGHT_GREEN
            Cells(startIdx, b_A_START).Font.Color = vbBlue
        End If
        
        ' Calculate units'
        units = WorksheetFunction.RoundUp(Duration / 15, 0) / 4

        ' After hours
        postHoursUnits = 0
        If IsWeekend(Cells(startIdx, b_DATE).Value) Then
            postHoursUnits = units
        Else
            dayStart = TimeValue("8:00 am")
            dayEnd = TimeValue("5:00 pm")
    
            preMins = DateDiff("n", startTime, WorksheetFunction.Min(dayStart, endTime))
            postMins = DateDiff("n", WorksheetFunction.Max(dayEnd, startTime), endTime)

            If (preMins > 0) Then
                postHoursUnits = WorksheetFunction.RoundUp(preMins / 15, 0) / 4
            End If
    
            If (postMins > 0) Then
                postHoursUnits = postHoursUnits + WorksheetFunction.RoundUp(postMins / 15, 0) / 4
            End If
        End If

        units = units - postHoursUnits
        Cells(endIdx, b_RH_UNITS).Value = units
        Cells(endIdx, b_RH_UNITS).Font.Color = vbRed
        If postHoursUnits > 0 Then
            Call ConcatOrSetValue(endIdx, b_AH_UNITS, vbRed, postHoursUnits)
            Cells(endIdx, b_AH_FEE_INTERP).Value = postHoursUnits * (Cells(endIdx, b_INTERPRATE).Value + INTERPRETER_OVERCHARGE)
            If LCase(Cells(i, b_STATUS)) <> "dnc" Then
                Cells(endIdx, b_AH_FEE_SCCA).Value = postHoursUnits * (SCCA_RATE + SCCA_OVERCHARGE)
            End If
        End If

        ' Calculate totals '
        Cells(endIdx, b_RH_FEE_INTERP).Value = units * Cells(endIdx, b_INTERPRATE).Value
        Cells(endIdx, b_INTERPTOTAL).Value = Cells(endIdx, b_RH_FEE_INTERP).Value + Cells(endIdx, b_AH_FEE_INTERP).Value

        If LCase(Cells(i, b_STATUS)) <> "dnc" Then
            Cells(endIdx, b_RH_FEE_SCCA).Value = units * SCCA_RATE
            Cells(endIdx, b_SCCATOTAL).Value = Cells(endIdx, b_RH_FEE_SCCA).Value + Cells(endIdx, b_AH_FEE_SCCA).Value
        Else
            Cells(endIdx, b_SCCATOTAL).Value = 0
        End If

        If postHoursUnits = 0 Then
            Cells(endIdx, b_AH_FEE_INTERP).Value = "-"
            Cells(endIdx, b_AH_FEE_INTERP).HorizontalAlignment = xlCenter
            Cells(endIdx, b_AH_FEE_SCCA).Value = "-"
            Cells(endIdx, b_AH_FEE_SCCA).HorizontalAlignment = xlCenter
        End If

        i = endIdx + 1
    Loop

End Sub

Public Sub RunAOLTIBilling()
    
    Call GlobalInit

    If ValidateBillingCaption And LoadInterpreterRates Then
        Call InitRatesColumn
        Call FindSeries
        MsgBox "Success!"
    End If

End Sub
