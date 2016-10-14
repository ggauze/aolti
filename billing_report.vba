'-----------------------------------
'  calcualte billing report
'-----------------------------------

Public Sub CalcBillingReport()

    Application.DisplayAlerts = False
    ActiveSheet.Columns.Hidden = False

    ActiveSheet.Copy After:=ActiveSheet
    
    ActiveSheet.Range("A1").Select

    Call AddColumns
    
    ' sort by Intepreter, Date, UNum
    ActiveSheet.UsedRange _
        .Sort key1:=Range("O2"), key2:=Range("E2"), key3:=Range("D2"), _
               order1:=xlAscending, Header:=xlYes, MatchCase:=False

    Dim i As Long
    Dim rRng As Range
    Dim UNum As String
    Dim nexUNum As String
    Dim startTime As Date
    Dim nextStartTime As Date
    Dim startDate As Date
    Dim nextStartDate As Date
    Dim endTime As Date
    Dim newEndTime As Date
    Dim duration As Long
    
    Set rRng = ActiveSheet.UsedRange
        
    rRng.Columns("G").NumberFormat = "h:mm AM/PM"
    rRng.Columns("J").NumberFormat = "h:mm AM/PM"
    
    i = 2
    
    Do While i <= rRng.Rows.Count
        UNum = rRng.Cells(i, "D").Value
        startDate = rRng.Cells(i, "E").Value
        startTime = rRng.Cells(i, "F").Value
        endTime = DateAdd("n", rRng.Cells(i, "H").Value, startTime)
        
        rRng.Cells(i, "G").Value = endTime
        rRng.Cells(i, "J").Value = endTime
        rRng.Cells(i, "F").Font.Color = vbBlue ' S start
        rRng.Cells(i, "I").Font.Color = vbBlue ' A start
        rRng.Cells(i, "R").Value = WorksheetFunction.RoundUp(DateDiff("n", startTime, endTime) / 15, 0) / 4
        rRng.Cells(i, "R").Font.Color = vbBlack
        
        ' rRng.Cells(i, "S").Value = WorksheetFunction.RoundUp(rRng.Cells(i, "K").Value / 15, 0) / 4
        ' rRng.Cells(i, "S").Font.Color = vbBlack
        
        i = i + 1
        Do While i <= rRng.Rows.Count
            nextUNum = rRng.Cells(i, "D").Value
            nextStartDate = rRng.Cells(i, "E").Value
            
            If nextUNum = UNum And startDate = nextStartDate Then
            
                nextStartTime = rRng.Cells(i, "F").Value
                
                ' First, calculate possible overlaps
                If DateDiff("n", endTime, nextStartTime) < 0 Then
                    rRng.Cells(i, "P").Value = "OVERLAP"
                    rRng.Cells(i, "P").Font.Color = vbGreen
                    rRng.Cells(i, "P").Font.FontStyle = "Bold"
                End If
                
                ' Adjust end time to minimum 60 min appt duration
                If DateDiff("n", startTime, endTime) < 60 Then
                    endTime = DateAdd("n", 60, startTime)
                End If
                
                Gap = DateDiff("n", endTime, nextStartTime)
                If Gap <= 60 Then
                    newEndTime = DateAdd("n", rRng.Cells(i, "H").Value, nextStartTime)
                    rRng.Cells(i, "G").Value = newEndTime
                    rRng.Cells(i, "J").Value = newEndTime
                    rRng.Cells(i, "R").Value = WorksheetFunction.RoundUp(DateDiff("n", nextStartTime, newEndTime) / 15, 0) / 4
                    rRng.Cells(i, "R").Font.Color = vbBlack
                    
                    If Gap > 0 Then
                        rRng.Cells(i, "P").Value = "WT"
                        rRng.Cells(i, "P").Font.Color = vbGreen
                        rRng.Cells(i, "P").Font.FontStyle = "Bold"
                    
                    End If
                    
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
        If duration < 60 Then
            endTime = DateAdd("n", 60, startTime)
            duration = 60
        End If
        
        rRng.Cells(i - 1, "R").Value = WorksheetFunction.RoundUp(duration / 15, 0) / 4
        rRng.Cells(i - 1, "R").Font.Color = vbRed
        rRng.Cells(i - 1, "G").Value = endTime
        rRng.Cells(i - 1, "G").Font.Color = vbRed
        rRng.Cells(i - 1, "J").Value = endTime
        rRng.Cells(i - 1, "J").Font.Color = vbRed
                
        ' rRng.Cells(i - 1, 9).Value = DateDiff("n", endTime, startTime)

    Loop
    
    ' Sheets(3).SaveAs "GoogleCal", xlCSV


End Sub

Private Sub AddColumns()

    ' Add columns / rename some / add formulas
    
    Cells(1, "E").Value = "Appt Date"
    Cells(1, "F").Value = "S Start"
    Columns("G").Insert Shift:=xlToRight
    Cells(1, "G").Value = "S End"
    Cells(1, "I").Value = "A Start"
    Columns("J").Insert Shift:=xlToRight
    Cells(1, "J").Value = "A End"
    Cells(1, "K").Value = "A MIN"
    Columns("N").Cut
    Columns("Q").Insert Shift:=xlToRight
    Columns("R").Insert Shift:=xlToRight
    Columns("R").Insert Shift:=xlToRight
    Cells(1, "R").Value = "BK Units"
    Cells(1, "R").Font.Color = vbBlack
    Cells(1, "R").Interior.Color = vbGreen
    
    Cells(1, "S").Value = "W Units"
    Cells(1, "S").Font.Color = vbBlack
    Cells(1, "S").Interior.Color = vbYellow

    lastRow = Range("K" & Rows.Count).End(xlUp).Row
    ' Formula to calculate actuall appt duration
    Range("K2").Formula = "=(J2-I2) * 1440"
    Range("K2").AutoFill Destination:=Range("K2:K" & lastRow)
    
    ' Formula to calculate actual units from actual appt duration
    Range("S2").Formula = "=ROUNDUP(ROUND(K2 / 15, 2), 0) / 4"
    Range("S2").AutoFill Destination:=Range("S2:S" & lastRow)
    
End Sub

