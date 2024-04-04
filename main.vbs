Sub Split_Sheet()

    Dim currentSheet As Worksheet

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each currentSheet In Worksheets
        SaveSheetAsCSV currentSheet
    Next currentSheet

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Sub SaveSheetAsCSV(sheet As Worksheet)

    Workbooks.Add

    With ActiveWorkbook
        sheet.Copy .Sheets(1)
        .SaveAs Filename:=ThisWorkbook.Path & "/" & sheet.Name, FileFormat:=xlCSV
        .Close
    End With

End Sub
