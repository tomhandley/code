Attribute VB_Name = "FormatCSV"
Sub FormatCSV()
'Reformat bathy export csv file and save
Dim seconds As Long
    
    seconds = ActiveSheet.Cells(Rows.count, "A").End(xlUp).row - 1
    Range(Cells(2, 1), Cells(seconds + 1, 1)).NumberFormat = "0.0000"
    Range(Cells(2, 2), Cells(seconds + 1, 2)).NumberFormat = "0.0000"
    Range(Cells(2, 3), Cells(seconds + 1, 3)).NumberFormat = "0.000"
    Range(Cells(2, 4), Cells(seconds + 1, 4)).NumberFormat = "m/d/yyyy hh:mm:ss"
    Range("A1").Select
    'Save as CSV
'    Application.DisplayAlerts = False
'    ActiveWorkbook.SaveAs Filename:=path & record & ".csv", FileFormat:=xlCSV
'    Application.DisplayAlerts = True
'    MsgBox("CSV file reformatted and saved.")

End Sub

