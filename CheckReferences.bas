Attribute VB_Name = "CheckReferences"
Sub CheckReferences()
' Check for possible missing or erroneous links in
' formulas and list possible errors in a summary sheet

  Dim iSh As Integer
  Dim sShName As String
  Dim sht As Worksheet
  Dim c, sChar As String
  Dim rng As Range
  Dim i As Integer, j As Integer
  Dim wks As Worksheet
  Dim sChr As String, addr As String
  Dim sFormula As String, scVal As String
  Dim lNewRow As Long
  Dim vHeaders

  vHeaders = Array("Sheet Name", "Cell", "Cell Value", "Formula")
  'check if 'Summary' worksheet is in workbook
  'and if so, delete it
  With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
    .Calculation = xlCalculationManual
  End With

  For i = 1 To Worksheets.count
    If Worksheets(i).Name = "Summary" Then
      Worksheets(i).Delete
    End If
  Next i

  iSh = Worksheets.count

  'create a new summary sheet
    Sheets.Add After:=Sheets(iSh)
    Sheets(Sheets.count).Name = "Summary"
  With Sheets("Summary")
    Range("A1:D1") = vHeaders
  End With
  lNewRow = 2

  ' this will not work if the sheet is protected,
  ' assume that sheet should not be changed; so ignore it
  On Error Resume Next

  For i = 1 To iSh
    sShName = Worksheets(i).Name
    Application.Goto Sheets(sShName).Cells(1, 1)
    Set rng = Cells.SpecialCells(xlCellTypeFormulas, 23)

    For Each c In rng
      addr = c.Address
      sFormula = c.Formula
      scVal = c.Text

      For j = 1 To Len(c.Formula)
        sChr = Mid(c.Formula, j, 1)

        If sChr = "[" Or sChr = "!" Or _
          IsError(c) Then
          'write values to summary sheet
          With Sheets("Summary")
            .Cells(lNewRow, 1) = sShName
            .Cells(lNewRow, 2) = addr
            .Cells(lNewRow, 3) = scVal
            .Cells(lNewRow, 4) = "'" & sFormula
          End With
          lNewRow = lNewRow + 1
          Exit For
        End If
      Next j
    Next c
  Next i

' housekeeping
  With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .Calculation = xlCalculationAutomatic
  End With

' tidy up
  Sheets("Summary").Select
  Columns("A:D").EntireColumn.AutoFit
  Range("A1:D1").Font.Bold = True
  Range("A2").Select
End Sub

