Sub correlation_coefficient()
    ThisWorkbook.Sheets("Sheet1").Range("G2").Value = WorksheetFunction.Correl(Range("A2:A" & ThisWorkbook.Sheets("Sheet1").Range("A10000").End(xlUp).row), Range("B2:B" & ThisWorkbook.Sheets("Sheet1").Range("A10000").End(xlUp).row))
    ThisWorkbook.Sheets("Sheet1").Range("H2").Value = WorksheetFunction.Correl(Range("D2:D" & ThisWorkbook.Sheets("Sheet1").Range("D10000").End(xlUp).row), Range("E2:E" & ThisWorkbook.Sheets("Sheet1").Range("D10000").End(xlUp).row))
End Sub
