'------------------------
' CreateTimeSheets
'------------------------

' Change this as needed [last updated 3/24/17]

Public Const IntepreterRatesFile = "T:\!Interpreter Management\!Operating Procedures\SCCA\Billing\Billing Macro\SCCA Interpreter Rates List.xlsx"

' Public Const IntepreterRatesFile = "C:\Users\Sergei\Downloads\SCCA Interpreter Rates List.xlsx"

' Modify the following constants if column order is changed'

' Master table '

Public Const m_INTERPRETER = 2
Public Const m_STATUS = 3
Public Const m_LAST_NAME = 4
Public Const m_FIRST_NAME = 5
Public Const m_LANGUAGE = 6
Public Const m_U_NUMBER = 7
Public Const m_DATE = 8
Public Const m_S_START = 9
Public Const m_S_END = 10
Public Const m_S_MIN = 11
Public Const m_DEPARTMENT = 12

' Time sheet '

Public Const t_LAST_NAME = 1
Public Const t_FIRST_NAME = 2
Public Const t_U_NUMBER = 3
Public Const t_DATE = 4
Public Const t_DEPARTMENT = 5
Public Const t_STATUS = 6
Public Const t_S_START = 7
Public Const t_S_END = 8
Public Const t_S_MIN = 9
Public Const t_ARRIVAL = 10
Public Const t_A_START = 11
Public Const t_A_END = 12
Public Const t_A_MIN = 13
Public Const t_LCL = 14
Public Const t_NOTES = 15

Public Const t_INTERPRETER = 16
Public Const t_LANGUAGE = 17

' Sheet names
Public Const TIME_TABLE_TEMPLATE = "TimeSheetTemplate"
Public Const TIME_TABLE = "TimeSheet"
Public Const MASTER = "From Master"
Public Const MASTER_COPY = "Validation Report"

' Global variables
Dim IRates() As Variant
Dim LastRow As Variant
Dim TimeSheet As Variant
Dim SavePath As String
Dim FileNamePrefix As String
Dim InterpCount As Integer

'Get number of days in the month
Function dhDaysInMonth(Optional dtmDate As Date = 0) As Integer
    ' Return the number of days in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    dhDaysInMonth = DateSerial(Year(dtmDate), _
     Month(dtmDate) + 1, 1) - _
     DateSerial(Year(dtmDate), Month(dtmDate), 1)
End Function

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim x As Long, y As Long
        
        For x = 2 To Cells(Rows.Count, "A").End(xlUp).Row
            y = Application.CountA(Range(Cells(x, "A"), Cells(x, "J")))
            If y < 10 Then
                MsgBox "You need to fill out all of the required cells in row " & x, vbOKOnly + vbCritical, "Information Required"
                Cells(x, "B").Activate
                Cancel = True
                Exit Sub
            End If
        Next x
        
End Sub


' Global initializaiton

Private Sub MasterInit()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlAutomatic

    ' Create a template sheet for interpreter time table
    
    Dim titles As Variant
    titles = Array("Last Name", "First Name", "U Number", "Date", "Department", "Status", "S Start", "S End", "S Min", _
                   "Arrival", "A Start", "A End", "A Min", "LCL on site only", "Interpreter Notes", "Interpreter", "Language")
    
    ActiveWorkbook.Sheets.Add
    ActiveSheet.Name = TIME_TABLE_TEMPLATE
    
    For i = LBound(titles) To UBound(titles)
        Cells(1, i + 1).Value = titles(i)
    Next i
    
    Columns(t_S_START).NumberFormat = "h:mm AM/PM"
    Columns(t_S_END).NumberFormat = "h:mm AM/PM"
    Columns(t_ARRIVAL).NumberFormat = "h:mm AM/PM"
    Columns(t_A_START).NumberFormat = "h:mm AM/PM"
    Columns(t_A_END).NumberFormat = "h:mm AM/PM"
    
    Rows(1).Font.Bold = True

    ' Billing sheet initialization

    ActiveWorkbook.Sheets(MASTER).Select

    ' Work in a copy
    ActiveSheet.Copy After:=ActiveSheet
    ActiveSheet.Name = MASTER_COPY

   ' Sort by Intepreter, Date, UNum, Start
    ActiveSheet.UsedRange _
        .Sort key1:=Columns(m_S_START), _
               order1:=xlAscending, Header:=xlYes, MatchCase:=False

    ActiveSheet.UsedRange _
        .Sort key1:=Columns(m_INTERPRETER), key2:=Columns(m_DATE), key3:=Columns(m_U_NUMBER), _
               order1:=xlAscending, Header:=xlYes, MatchCase:=False
       
    LastRow = Range("B" & Rows.Count).End(xlUp).Row
    SavePath = ActiveWorkbook.Path & "\"
    
    'Calculate the pay period range
    Dim StartDate As Date
    StartDate = Cells(2, m_DATE).Value
    
    If Day(StartDate) < 16 Then
        StartDay = 1
        EndDay = 15
    Else
        StartDay = 16
        EndDay = dhDaysInMonth(StartDate)
    End If
    
    FileNamePrefix = MonthName(Month(StartDate)) + " " + CStr(StartDay) + " - " + CStr(EndDay) + " "
        
End Sub

Private Sub CopyValuesToTableSheet(masterIdx As Variant, ttIdx As Variant)
        TimeSheet.Cells(ttIdx, t_LAST_NAME).Value = Cells(masterIdx, m_LAST_NAME).Value
        TimeSheet.Cells(ttIdx, t_FIRST_NAME).Value = Cells(masterIdx, m_FIRST_NAME).Value
        TimeSheet.Cells(ttIdx, t_U_NUMBER).Value = Cells(masterIdx, m_U_NUMBER).Value
        TimeSheet.Cells(ttIdx, t_DATE).Value = Cells(masterIdx, m_DATE).Value
        TimeSheet.Cells(ttIdx, t_DEPARTMENT).Value = Cells(masterIdx, m_DEPARTMENT).Value
        TimeSheet.Cells(ttIdx, t_STATUS).Value = Cells(masterIdx, m_STATUS).Value
        TimeSheet.Cells(ttIdx, t_S_START) = Cells(masterIdx, m_S_START)
        TimeSheet.Cells(ttIdx, t_S_START).Font.Color = Cells(masterIdx, m_S_START).Font.Color
        TimeSheet.Cells(ttIdx, t_S_END) = Cells(masterIdx, m_S_END)
        TimeSheet.Cells(ttIdx, t_S_END).Font.Color = Cells(masterIdx, m_S_END).Font.Color
        TimeSheet.Cells(ttIdx, t_S_MIN) = Cells(masterIdx, m_S_MIN)
        
        
        
        TimeSheet.Cells(ttIdx, t_INTERPRETER).Value = Cells(masterIdx, m_INTERPRETER).Value
        TimeSheet.Cells(ttIdx, t_LANGUAGE).Value = Cells(masterIdx, m_LANGUAGE).Value
        
        If Cells(masterIdx, m_STATUS) = "LCL" Then
            TimeSheet.Cells(ttIdx, t_ARRIVAL).Value = "N/A"
            TimeSheet.Cells(ttIdx, t_STATUS).Font.Color = vbRed
            TimeSheet.Cells(ttIdx, t_STATUS).Font.Bold = True
            TimeSheet.Cells(ttIdx, t_A_START).Value = Cells(masterIdx, m_S_START).Value
            TimeSheet.Cells(ttIdx, t_A_END).Value = Cells(masterIdx, m_S_END).Value
            TimeSheet.Cells(ttIdx, t_A_MIN).Value = Cells(masterIdx, m_S_MIN).Value
        Else
            TimeSheet.Cells(ttIdx, t_ARRIVAL).Interior.ColorIndex = 40
            TimeSheet.Cells(ttIdx, t_A_START).Interior.ColorIndex = 44
            TimeSheet.Cells(ttIdx, t_A_END).Interior.ColorIndex = 44
            TimeSheet.Cells(ttIdx, t_A_MIN).Formula = "=(" & TimeSheet.Cells(ttIdx, t_A_END).Address & "-" _
                                                           & TimeSheet.Cells(ttIdx, t_A_START).Address & ")*1440"
            
            TimeSheet.Cells(ttIdx, t_LCL).Interior.ColorIndex = 15
            TimeSheet.Cells(ttIdx, t_NOTES).Interior.ColorIndex = 15
        End If
            
End Sub

' Validate interpreter names
Private Sub ValidateInterpreterNames()

    i = 2
    interpIdx = -1

    Do While i <= LastRow
        interpName = Trim(Cells(i, m_INTERPRETER).Value)
        dept = Trim(Cells(i, m_DEPARTMENT).Value)
        interpIdx = FindIntepreterIdx(interpName, dept, interpIdx)
        Do While i <= LastRow And Trim(Cells(i, m_INTERPRETER).Value) = interpName
            If (interpIdx = -1) Then
                Cells(i, m_INTERPRETER).Font.Color = vbRed
            End If
            
            i = i + 1
        Loop
    Loop

End Sub

Function HandleNextInterpreter(startIdx As Variant) As Variant

    i = startIdx
    j = 2

    interpreter = Trim(Cells(i, m_INTERPRETER).Value)
    nextInterpreter = interpreter

    ' Copy the interpreter values
    Do While i <= LastRow
        If interpreter = nextInterpreter Then
            Status = LCase(Cells(i, m_STATUS).Value)
            If Status <> "cld" And Status <> "ins" Then
                Call CopyValuesToTableSheet(i, j)
                j = j + 1
            End If
        Else
            Exit Do
        End If
        i = i + 1
        nextInterpreter = Trim(Cells(i, m_INTERPRETER).Value)
    Loop
    
    ' Only create timesheet for non-empty data
    If j > 2 Then
        ' Copy TimeSheet to a new workbook
        TimeSheet.Copy
        
        ' Border style
        Range(Cells(2, t_ARRIVAL), Cells(j - 1, t_NOTES)).Borders.LineStyle = xlContinuous
        
        ' Freeze first row
        Rows("2:2").Select
        ActiveWindow.FreezePanes = True
        
        ' Lock some cells
        ActiveSheet.Protect UserInterfaceOnly:=True
        Range(Cells(2, t_ARRIVAL), Cells(j - 1, t_NOTES)).Locked = False
        
        ActiveSheet.EnableCalculation = True
        
        Columns.AutoFit
        Columns(t_INTERPRETER).Hidden = True
        Columns(t_LANGUAGE).Hidden = True
        
        ' Save TimeTable sheet to a separate file
        Fname = FileNamePrefix + interpreter + ".xlsx"
        With ActiveWorkbook
            .SaveAs Filename:=SavePath & Fname
            .Close
        End With
        InterpCount = InterpCount + 1
    End If

    HandleNextInterpreter = i
End Function

Function GetFileName(startIdx As Variant, endIdx As Variant) As String

    StartDate = Cells(startIdx, m_DATE).Value
    EndDate = Cells(endIdx, m_DATE).Value
    
    MonthN = MonthName(Month(StartDate))
    
    GetFileName = MonthN + " " + CStr(Day(StartDate)) + " - " + CStr(Day(EndDate)) + " " + Cells(startIdx, m_INTERPRETER).Value + ".xlsx"

End Function

Private Sub FindSeries()

    i = 2

    Do While i <= LastRow
        
        interpreter = Cells(i, m_INTERPRETER).Value
        UNum = Cells(i, m_U_NUMBER).Value
        StartDate = Cells(i, m_DATE).Value
        startTime = Cells(i, m_S_START).Value
        endTime = Cells(i, m_S_END).Value
    
        Cells(i, m_S_START).Font.Color = vbBlue ' series start is blue'
    
        ' Check for after-hours and split if the appointment lasts pas '
        i = i + 1
    
        Do While i <= LastRow
            nextUNum = Cells(i, m_U_NUMBER).Value
            nextStartDate = Cells(i, m_DATE).Value
            nextInterpreter = Cells(i, m_INTERPRETER).Value
            
            If nextUNum = UNum And StartDate = nextStartDate And interpreter = nextInterpreter Then
            
                nextStartTime = Cells(i, m_S_START).Value
                
                ' First, calculate possible overlaps
                If DateDiff("n", endTime, nextStartTime) < 0 Then
                    nextStartTime = endTime
                End If
                
                Gap = DateDiff("n", endTime, nextStartTime)
                ' Adjust gap to minimum 60 min appt duration
                If DateDiff("n", startTime, endTime) < 60 And Gap > 60 Then
                    Gap = DateDiff("n", DateAdd("n", 60, startTime), nextStartTime)
                End If
    
                If Gap <= 60 Then  ' if wait time is <= 60, this appointment is a part of the series'
                    newEndTime = Cells(i, m_S_END).Value
                    
                    If newEndTime > endTime Then
                       endTime = newEndTime
                    End If
                Else
                    Exit Do
                End If
            Else
                Exit Do
            End If
            i = i + 1
        Loop
        Cells(i - 1, m_S_END).Font.Color = vbRed ' mark series end
    Loop
End Sub

Private Sub ProcessMaster()

    i = 2
    InterpCount = 0
    
    Do While i <= LastRow
        If Cells(i, m_INTERPRETER).Font.Color <> vbRed Then
            ' Create a working copy of the current time sheet
            Sheets(TIME_TABLE_TEMPLATE).Copy After:=Sheets(TIME_TABLE_TEMPLATE)
            ActiveSheet.Name = TIME_TABLE
            Set TimeSheet = ActiveSheet
            
            Sheets(MASTER_COPY).Select
            
            endIdx = HandleNextInterpreter(i)
            
            ' Clear current time sheet
            TimeSheet.Delete
            i = endIdx
        Else
            i = i + 1
        End If
    Loop
        
End Sub

Private Sub Cleanup()

    Sheets(TIME_TABLE_TEMPLATE).Delete
    
    ' Uncomment the following if you want the mastre copy to be deleted
    
    ' Sheets(MASTER_COPY).Delete
    
End Sub
Public Sub CreateTimeSheets()
    
    Call MasterInit

    If LoadInterpreterRates(IntepreterRatesFile) Then
        Call ValidateInterpreterNames
        Call FindSeries
        Call ProcessMaster
        Call Cleanup
        MsgBox CStr(InterpCount) & " timesheets created in " & SavePath
    End If

End Sub

