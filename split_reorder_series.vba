Sub CellSplitter()

    Dim lastrow As Integer
    lastrow = GetLastRow()

    Dim Oarray() As String 'O matches column needing to be split
    Dim Parray() As String 'P matches column needing to be split
    Dim LongestArray As Integer
    Dim TempInt As Integer

    Dim i As Integer
    i = 1
    
    'This script adds new "Time" column and inserts forumla
    Columns("H:H").Insert Shift:=xlToRight, _
      CopyOrigin:=xlFormatFromLeftOrAbove 'or xlFormatFromRightOrBelow
    Worksheets(1).Range("H1:H" & lastrow).Formula = "=TEXT(TRIM(G1),""HH:MM"")"
    
    'This script adds new "Department" column and inserts formula
    Columns("K:K").Insert Shift:=xlToRight, _
      CopyOrigin:=xlFormatFromLeftOrAbove 'or xlFormatFromRightOrBelow
    Worksheets(1).Range("K1:K" & lastrow).Formula = "=IF(ISERR(FIND("","",J1,1)),J1,LEFT(J1,FIND("","",J1,1)-1))"
    
    Do While i <= lastrow

        Oarray = Split(Range("O" & i), Chr(10)) 'Chr(10) = carriage return w/i the cell
        Parray = Split(Range("P" & i), Chr(10))
        LongestArray = GetLongestArray(Oarray, Parray)

        If LongestArray > 0 Then

            ' reset the values of O and P columns
            On Error Resume Next
            Range("O" & i).Value = Oarray(0)
            Range("P" & i).Value = Parray(0)
            Err.Clear
            On Error GoTo 0

            ' duplicate the row multiple times
            For TempInt = 1 To LongestArray

                Rows(i & ":" & i).Select
                Selection.Copy

                Range(i + TempInt & ":" & i + TempInt).Select
                Selection.Insert Shift:=xlDown

                ' as each row is copied, change the values of O and P
                On Error Resume Next
                Range("O" & i + TempInt).Value = Oarray(TempInt)
                If Err.Number > 0 Then Range("O" & i + TempInt).Value = ""
                Err.Clear
                Range("P" & i + TempInt).Value = Parray(TempInt)
                If Err.Number > 0 Then Range("P" & i + TempInt).Value = ""
                Err.Clear
                On Error GoTo 0

                Application.CutCopyMode = False

            Next TempInt

            ' increment the outer FOR loop's counters
            lastrow = lastrow + LongestArray
            i = i + LongestArray

        End If

        i = i + 1
    Loop

End Sub

Function GetLongestArray(ByRef Oarray() As String, ByRef Parray() As String)
    GetLongestArray = UBound(Oarray)
    If UBound(Parray) > GetLongestArray Then GetLongestArray = UBound(Parray)
End Function

Function GetLastRow() As Integer
    Worksheets(1).Select 'Select the first worksheet in the workbook...adjust as needed
    Range("A1").Select
    Selection.End(xlDown).Select
    GetLastRow = Selection.Row
    Range("A1").Select
End Function

Sub ReorderColumns()

    lastrow = Range("A" & Rows.Count).End(xlUp).Row
    
    ' Change formulas to values
    With Range("H1:H" & lastrow)               ' Time
        .Value = .Value
    End With
    Columns("H").NumberFormat = "HH:MM"
    

    With Range("K1:K" & lastrow)               ' Department
        .Value = .Value
    End With
    
    ' Add columns
    Columns("A").Insert
    Cells(1, "A").Value = "Interpreter"
    Columns("A").Insert
    Cells(1, "A").Value = "Notes"
    
    Columns("O").Cut               ' Status
    Columns("C").Insert
    Columns("D").Cut               ' CSN
    Columns("O").Insert
    Columns("H").Cut               ' Date
    Columns("L").Insert
    Columns("H").EntireColumn.Delete    ' original Time
    Columns("K").Cut               ' original Department
    Columns("R").Insert
    Columns("K").Cut               ' new Department
    Columns("H").Insert
    Columns("N").Cut               ' Perm Com
    Columns("Q").Insert
    Columns("M").Cut               ' Type
    Columns("L").Insert
    
    ' Hide columns
    ' Range("Q:Z").EntireColumn.Hidden = True

    ' Capitalize some headers
    Cells(1, "D").Value = StrConv(Cells(1, "D").Value, vbUpperCase)
    Cells(1, "E").Value = StrConv(Cells(1, "E").Value, vbUpperCase)
    Cells(1, "F").Value = StrConv(Cells(1, "F").Value, vbUpperCase)
    Cells(1, "K").Value = StrConv(Cells(1, "K").Value, vbUpperCase)
    
    ' Add END column
    Columns("J").Insert
    
    ' Formula to calculate end time from start time and appt duration
    Range("J2").Formula = "=I2+TIME(0,K2,0)"
    Range("J2").AutoFill Destination:=Range("J2:J" & lastrow)

    ' Convert to values
    With Columns("J")
        .Value = .Value
    End With
    
    Cells(1, "J").Value = "End"
    Columns("J").NumberFormat = "HH:MM"

End Sub


Public Sub CalcSeries()

    Application.DisplayAlerts = False
    ActiveSheet.Columns.Hidden = False

    ' sort by Date and UNum
    ActiveSheet.UsedRange _
        .Sort key1:=Range("K2"), key2:=Range("G2"), order1:=xlAscending, Header:=xlYes, MatchCase:=False

    Dim i As Long
    Dim UNum As String
    Dim nexUNum As String
    Dim startTime As Date
    Dim nextStartTime As Date
    Dim startDate As Date
    Dim nextStartDate As Date
    Dim endTime As Date
    Dim newEndTime As Date
    Dim duration As Long
        
    i = 2
    
    Do While i <= ActiveSheet.UsedRange.Rows.Count
        UNum = Cells(i, "G").Value
        startDate = Cells(i, "L").Value
        startTime = Cells(i, "I").Value
        endTime = DateAdd("n", Cells(i, "K").Value, startTime)
        
        Cells(i, "J").Value = endTime
        Cells(i, "I").Font.Color = vbBlue ' start
        
        i = i + 1
        Do While i <= Rows.Count
            nextUNum = Cells(i, "G").Value
            nextStartDate = Cells(i, "L").Value
            
            If nextUNum = UNum And startDate = nextStartDate Then
            
                nextStartTime = Cells(i, "I").Value
                
                Gap = DateDiff("n", endTime, nextStartTime)
                If Gap <= 60 Then
                    newEndTime = DateAdd("n", Cells(i, "K").Value, nextStartTime)
                    Cells(i, "I").Value = newEndTime
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
        
        duration = DateDiff("n", startTime, endTime)
        Cells(i - 1, "J").Value = endTime
        Cells(i - 1, "J").Font.Color = vbRed
    Loop
End Sub

Sub SplitReorderMakeSeries()

    CellSplitter
    ReorderColumns
    CalcSeries
    
End Sub
