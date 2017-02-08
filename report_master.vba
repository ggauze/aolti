'
' In order for this script to work, file report_master.xlxm with a tab 'FAQ' should be open
'
' Also, the original master will remain filtered on the column "Status", this filter should be removed manually


Private Sub ChangeInterpreterFormatting()
    For Each c In Range(Range("A2"), Range("A2").End(xlDown)).Cells
        c.Font.Bold = True
        If c.Value = "Telephonic" Then
            c.Font.Color = vbRed
        ElseIf c.Value = "VRI" Then
            c.Font.Color = RGB(0, 176, 80)
        ElseIf c.Value = "Unfilled" Then
            c.Value = "ULS pending"
            c.Font.Color = RGB(0, 176, 240)
        End If
    Next
    ActiveSheet.Columns.AutoFit
End Sub

Private Sub CopyHospitalToWorksheet(Hospital As String, Pattern As String)
    Dim NW As Range
    Columns("J").AutoFilter Field:=1, Criteria1:="=" & Pattern
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
        Worksheets("Main").Activate
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
        Worksheets("Main").Activate
        NW.Range(Range("A2"), Range("A2").SpecialCells(xlCellTypeLastCell)).Delete
    End If
        
    Columns("J").AutoFilter
    
End Sub

Private Sub SortColumns()
    Columns("A:J") _
        .Sort key1:=Range("F2"), key2:=Range("E2"), key3:=Range("G2"), _
               order1:=xlAscending, Header:=xlYes, MatchCase:=False
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

Private Sub SaveWorksheet(reportDate As Date)
    ActiveWorkbook.SaveAs _
        Filename:=Format(reportDate, "mmmm dd") & ".xlsx"
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
    
    ActiveSheet.UsedRange.AutoFilter Field:=3, Criteria1:=Array("Sch"), _
        Operator:=xlFilterValues
    ActiveSheet.UsedRange.Select
    Selection.Copy
    ActiveSheet.Range("A1").Select

    Call AddNew
    
    ActiveSheet.Name = "Main"
    ActiveSheet.Paste
    ActiveSheet.Columns.AutoFit
    ActiveWindow.FreezePanes = True
    Cells.ClearComments
    ' Cells.Font.Bold = False
    ' Cells.Font.Italic = False
    ' Cells.Font.Color = vbBlack
    ' Cells.Interior.ColorIndex = xlNone
    Cells.Borders.LineStyle = xlLineStileNone
    
    ' delete some columns
    Columns("M:Z").Delete 
    Columns("A").Delete  ' NOTES'
    Columns("B").Delete  ' Status'

    Call ChangeInterpreterFormatting
    
    Call SortColumns
    
    ' ActiveSheet.Columns.AutoFit
    
    ' find Evergreen
    Call CopyHospitalToWorksheet("Evergreen", "S EH*")
    
    ' find Northwest
    Call CopyHospitalToWorksheet("Northwest", "S NWH*")
        
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
    Worksheets("Main").Activate

End Sub
