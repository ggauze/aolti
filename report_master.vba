Private Sub ChangeInterpreterFormatting()
    For Each c In Range(Range("J2"), Range("J2").End(xlDown)).Cells
        c.Font.Bold = True
        If c.Value = "Telephonic" Then
            c.Value = "Telephonic interpreter"
            c.Font.Color = RGB(0, 176, 240)
        ElseIf c.Value = "VRI" Then
            c.Value = "Video Remote Interpreter"
            c.Font.Color = RGB(0, 176, 80)
        ElseIf c.Value = "Unfilled" Then
            c.Value = "ULS pending"
            c.Font.Color = RGB(0, 32, 96)
        End If
        ' ElseIf Cells(c.Row,
    Next
    ActiveSheet.Columns.AutoFit
End Sub

Private Sub CopyHospitalToWorksheet(Hospital As String)
    Dim NW As Range
    Columns("H").AutoFilter Field:=1, Criteria1:="=*" & Hospital & "*"
    Set NW = ActiveSheet.UsedRange
    NW.Select
    Selection.Copy
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = Hospital
    ActiveSheet.Paste
    
    If Selection.Rows.Count = 1 Then
        Selection.Delete
        ActiveSheet.Range("A1").Value = "No appointments scheduled for this location"
        ActiveSheet.Columns.AutoFit
        Worksheets("MAIN CAMPUS").Activate
    Else
        '  underline Header
        ActiveSheet.Range("A1:J1").Font.Bold = True
        ActiveSheet.Range("A1:J1").Interior.Color = RGB(0, 176, 240)
        ActiveSheet.Range("A1:J1").BorderAround _
          LineStyle:=xlContinuous, Color:=vbBlack, Weight:=xlThick
    
        ActiveSheet.Columns.AutoFit
        
        ' freeze tope row
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True
        
        ActiveSheet.Range("A1").Select
        Worksheets("MAIN CAMPUS").Activate
        NW.Range(Range("A2"), Range("A2").SpecialCells(xlCellTypeLastCell)).Delete
    End If
        
    Columns("H").AutoFilter
    
End Sub

Private Sub SortColumns()
    Columns("A:J") _
        .Sort key1:=Range("B2"), key2:=Range("C2"), key3:=Range("F2"), _
               order1:=xlAscending, Header:=xlYes, MatchCase:=False
    Columns("A:J") _
        .Sort key1:=Range("A2"), order1:=xlAscending, Header:=xlYes, MatchCase:=False
End Sub

Private Function nextDay() As Date
    
    If Weekday(Now) = 6 Then
        nextDay = DateAdd("d", 3, Now)
    ElseIf Weekday(Now) = 7 Then
        nextDay = DateAdd("d", 2, Now)
    Else
        nextDay = DateAdd("d", 1, Now)
    End If
    
End Function

Private Sub SaveEncrypted()
    Dim reportDate As Date
    reportDate = nextDay()
    ActiveWorkbook.SaveAs _
        Filename:=Format(reportDate, "mmmm dd") & ".xlsx", _
        password:="report" & Format(reportDate, "mmdd")
    ActiveWorkbook.SaveAs _
        Filename:=Format(reportDate, "mmmm dd") & " pager.xlsx"
End Sub

Private Sub AddNew()
Set NewBook = Workbooks.Add
    With NewBook
        .Title = "SCCA daily schedule"
        .Subject = Format(nextDay(), "mm/dd/yyyy")
    End With
End Sub


Public Sub CreateSCCAReport()

    Application.DisplayAlerts = False
    ActiveSheet.Columns.Hidden = False
    
    ActiveSheet.UsedRange.AutoFilter Field:=16, Criteria1:=Array("Scheduled"), _
        Operator:=xlFilterValues
    ActiveSheet.UsedRange.Select
    Selection.Copy
    ' ActiveSheet.UsedRange.AutoFilter Field:=16
    ' Columns("P").AutoFilter
    ActiveSheet.Range("A1").Select
    
    Call AddNew
    
    ActiveSheet.Name = "MAIN CAMPUS"
    ActiveSheet.Paste
    ActiveSheet.Columns.AutoFit
    ActiveWindow.FreezePanes = True
    Cells.ClearComments
    Cells.Font.Bold = False
    Cells.Font.Italic = False
    Cells.Font.Color = vbBlack
    Cells.Interior.ColorIndex = xlNone
    Cells.Borders.LineStyle = xlLineStileNone
    
    ' delete some columns
    Columns("A").Delete
    Columns("I:N").Delete
    Columns("J").Delete
    
    ' cut & paste Appt date
    Columns("E").Cut
    Columns("A").Insert
    
    ' cut & paste U number
    Columns("E").Cut
    Columns("D").Insert
    
    ' delete Notes - the last column
    Columns("K").Delete
    
    Call ChangeInterpreterFormatting
    
    Call SortColumns
    
    ' ActiveSheet.Columns.AutoFit
    
    ' find Evergreen
    Call CopyHospitalToWorksheet("EVERGREEN")
    
    ' find Northwest
    Call CopyHospitalToWorksheet("NORTHWEST")
        
    '  underline Header
    ActiveSheet.Range("A1:J1").Font.Bold = True
    ActiveSheet.Range("A1:J1").Interior.Color = RGB(0, 176, 240)
    ActiveSheet.Range("A1:J1").BorderAround _
      LineStyle:=xlContinuous, Color:=vbBlack, Weight:=xlThick
    
    ' freeze tope row
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    ActiveSheet.Range("A1").Select
    
    ' copy FAQ from report_master
    Set wb = ActiveWorkbook
    ' Set myName = Application.VBE.ActiveCodePane.CodeModule.Parent.Name
    
    Workbooks("report_master.xlsm").Activate
    Worksheets("FAQ").Copy After:=wb.Sheets(wb.Sheets.Count)
    wb.Activate
    Worksheets("MAIN CAMPUS").Activate
    
    ' password protect
    ' Call SaveEncrypted
    
End Sub

Public Sub test()

    Application.DisplayAlerts = False
    ActiveSheet.Columns.Hidden = False
    
    ActiveSheet.UsedRange.AutoFilter Field:=16, Criteria1:=Array("Scheduled"), _
        Operator:=xlFilterValues
    ActiveSheet.UsedRange.Select
    Selection.Copy
    ActiveSheet.Range("A1").Select
    
    Set TmpCsvBook = Workbooks.Add()
    With TmpCsvBook
        .Title = "CSV"
    End With
    
    ActiveSheet.Name = "MAIN COPY"
    ActiveSheet.Paste
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "CSV"
    
    Range("A1").Value = "Subject"
    Range("B1").Value = "Start date"
    Range("C1").Value = "Start time"
    Range("D1").Value = "End date"
    Range("E1").Value = "End time"
    Range("F1").Value = "Description"
    
    Worksheets("MAIN COPY").Activate
    
    ActiveSheet.UsedRange.AutoFilter Field:=4, Criteria1:=Array("SPANISH"), _
        Operator:=xlFilterValues

    ' ActiveSheet.UsedRange.AutoFilter Field:=6, Criteria1:=Array("1/1/2015"), _
    '    Operator:=xlFilterValues

    ActiveSheet.UsedRange _
        .Sort key1:=Range("F2"), key2:=Range("E2"), key3:=Range("G2"), _
               order1:=xlAscending, Header:=xlYes, MatchCase:=False

    Dim i As Long
    Dim csvI As Long
    Dim rRng As Range
    Dim UNum As String
    Dim nexUNum As String
    Dim startTime As Date
    Dim nextStartTime As Date
    Dim startDate As Date
    Dim nextStartDate As Date
    Dim endTime As Date
    Dim duration As Long

    ActiveSheet.UsedRange.Select
    Selection.Copy
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "tmp"
    ActiveSheet.Paste
    
    Set rRng = ActiveSheet.UsedRange
    
    i = 2
    csvI = 2
    
    Do While i <= rRng.Rows.Count
        UNum = rRng.Cells(i, 5).Value
        startDate = rRng.Cells(i, 6).Value
        startTime = rRng.Cells(i, 7).Value
        endTime = DateAdd("n", rRng.Cells(i, 8).Value, startTime)
        
        Sheets(3).Cells(csvI, 1).Value = rRng.Cells(i, 4).Value ' language
        Sheets(3).Cells(csvI, 2).Value = startDate
        Sheets(3).Cells(csvI, 3).Value = startTime
        Sheets(3).Cells(csvI, 4).Value = startDate
        Sheets(3).Cells(csvI, 6).Value = rRng.Cells(i, 17).Value & vbCrLf _
                                       & rRng.Cells(i, 18).Value & vbCrLf _
                                       & rRng.Cells(i, 12).Value
        i = i + 1
        Do While i <= rRng.Rows.Count
            nextUNum = rRng.Cells(i, 5).Value
            nextStartDate = rRng.Cells(i, 6).Value
            If nextUNum = UNum And startDate = nextStartDate Then
                nextStartTime = rRng.Cells(i, 7).Value
                If DateDiff("n", nextStartTime, endTime) < 60 Then
                    endTime = DateAdd("n", rRng.Cells(i, 8).Value, nextStartTime)
                Else
                    Exit Do
                End If
            Else
                Exit Do
            End If
            i = i + 1
        Loop
        
        Sheets(3).Cells(csvI, 5).Value = endTime
        csvI = csvI + 1
    Loop
    
    Sheets(3).SaveAs "GoogleCal", xlCSV
            
End Sub
