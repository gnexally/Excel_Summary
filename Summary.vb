Sub CreateSummary()
    ' Get today's date
    Dim today As String
    today = Format(Date, "dd/MM/yyyy")
    today = Format(Date, "dd/MM/yyyy", vbUseSystemDayOfWeek, vbUseSystem)
    
    ' Check if Summary sheet already exist, if not create one
    Dim summarySheet As Worksheet
    On Error Resume Next
    Set summarySheet = ThisWorkbook.Sheets("Summary")
    On Error GoTo 0
    If summarySheet Is Nothing Then
        Set summarySheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        summarySheet.Name = "Summary"
    Else
        ' Clear existing data in Summary sheet
        summarySheet.UsedRange.Clear
    End If

    ' Loop through all sheets in the workbook (excluding "Summary" sheet)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim sheetCounter As Integer
        sheetCounter = 0
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Summary" And sheetCounter < 5 Then
            sheetCounter = sheetCounter + 1
            'Add the sheet name to the summary sheet
            r = summarySheet.Cells(summarySheet.Rows.Count, 1).End(xlUp).Row + 1
            summarySheet.Cells(r, 1).Value = ws.Name
            summarySheet.Cells(r, 1).Font.Bold = True
            summarySheet.Cells(r, 1).Font.Size = 14
            summarySheet.Cells(r, 1).Interior.Color = RGB(255, 153, 0)

            ' Copy headers
            ws.Rows(1).Copy
            summarySheet.Cells(summarySheet.Rows.Count, "A").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
            ' Set headers to bold
            summarySheet.Cells(summarySheet.Rows.Count, "A").End(xlUp).Resize(1, ws.UsedRange.Columns.Count).Font.Bold = True
            ' Get last row of data for today's date
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            ' Loop through each row in the current sheet
            For i = 2 To lastRow
                ' Check if the date in the first column matches today's date
                If ws.Cells(i, 1).Value = Date Then
                    ' Copy the current row and paste it into the "Summary" sheet
                    ws.Rows(i).Copy
                    summarySheet.Cells(summarySheet.Rows.Count, "A").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
                    ' Set the number format for the date column
                    summarySheet.Cells(summarySheet.Rows.Count, 1).End(xlUp).NumberFormat = "dd/mm/yyyy"
                End If
            Next i
        End If
    Next ws
    ' Autofit the columns
    summarySheet.Columns.AutoFit
End Sub
