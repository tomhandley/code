Attribute VB_Name = "Reprocess"
Option Explicit

Sub ReprocessXYZ()
Attribute ReprocessXYZ.VB_ProcData.VB_Invoke_Func = "R\n14"
    'Show userform ProgressBar
    'ProgressBar calls Subdivide_Pings
    ProgressBar.Show
End Sub

Sub Subdivide_Pings()
Attribute Subdivide_Pings.VB_ProcData.VB_Invoke_Func = "R\n14"
'Reprocess sonar data after ProcessData has run to refine position/depth data to a more frequent interval than
'one per second. This is accomplished by interpolating RTK position (as already derived in ProcessData), and
'averaging the raw sonar depth readings by the interval defined as sonavg, using shorter intervals with depth
'is changing rapidly (exceeds maxdeltaZ)
Dim son_id As String  'Sonar record identifier
Dim row, sonrow, exprow, lastson As Long  'Row identifiers
Dim seconds As Long  'Number of records in Export sheet
Dim last As Double  'Time in seconds of final record
Dim pings As Integer  'Interval counter
Dim x1, x2, y1, y2 As Double  'Position interpolation basis
Dim rz1, rz2 As Single  'RTK elevation interpolation basis
Dim avgtime, t1 As Double  'Averaging for time from sonar pings
Dim deltaZ, avgZ As Single  'Averaging for depth from sonar pings
Dim millisec As Single  'Time interval used for interpolation
Dim z1, z2 As Single  'Use to calculate change in depth between readings
Const sonavg = 5  'Number of sonar pings to average in each output record
Const maxdeltaZ = 0.1  'Maximum allowable change in depth within sonar averaging interval

seconds = Worksheets("Export").Range("A1").CurrentRegion.Rows.count - 1  'Number of full-second Export records
last = Round(Worksheets("Export").Cells(seconds + 1, 4) * 24 * 3600, 0) 'Time in seconds of last Export record
lastson = Worksheets("Sonar").Range("A1").CurrentRegion.Rows.count  'Row of last sonar record
'Find last record in Sonar sheet with time less than last record of Export sheet
Do While Worksheets("Sonar").Cells(lastson, 1) * 24 * 3600 > last
    lastson = lastson - 1
Loop

'Initialize variables
t1 = Worksheets("Export").Cells(2, 4)
y1 = Worksheets("Export").Cells(2, 1)
y2 = Worksheets("Export").Cells(3, 1)
x1 = Worksheets("Export").Cells(2, 2)
x2 = Worksheets("Export").Cells(3, 2)
rz1 = Worksheets("Combo").Cells(2, 17)
rz2 = Worksheets("Combo").Cells(3, 17)
z2 = Worksheets("Sonar").Cells(2, 4)
son_id = Worksheets("Export").Cells(2, 5)
row = 2  'Row counter for Combo and Export sheets, interpolation for position is between row and row + 1 on Export sheet
exprow = 2  'Row counter for Export2 sheet
sonrow = 2  'Row counter for Sonar sheet

'Add Export2 sheet if it isn't already present
On Error Resume Next
Application.DisplayAlerts = False
Worksheets("Export2").Delete  'Delete any old data
Application.DisplayAlerts = True
Worksheets.Add(After:=Worksheets("Export")).Name = "Export2"
'If (Worksheets("Export2").Name = "") Then  'Sheet doesn't exist
'End If
On Error GoTo 0

'Format Export2 sheet
Worksheets("Export").Range("A1:F1").Copy Destination:=Worksheets("Export2").Range("A1")  'Copy header row
With Worksheets("Export2")  'Set column widths
    .Range("A1:B1").ColumnWidth = 12
    .Range("C1").ColumnWidth = 7.67
    .Range("D1").ColumnWidth = 21.22
    .Range("E1:F1").ColumnWidth = 8.11
End With

'Loop interpolation of records
Do While sonrow + sonavg <= lastson
    pings = 0
    deltaZ = 0
    avgZ = 0
    avgtime = 0
    'Loop until maximum change in depth or averaging interval are exceeded
    Do Until deltaZ > maxdeltaZ Or pings = sonavg
        z1 = z2
        z2 = Worksheets("Sonar").Cells(sonrow + pings + 1, 4)
        avgZ = avgZ + z1
        avgtime = avgtime + Worksheets("Sonar").Cells(sonrow + pings, 1)
        deltaZ = deltaZ + Abs(z2 - z1)
        pings = pings + 1
    Loop
    
    'Average depth and time data
    avgZ = avgZ / pings
    avgtime = avgtime / pings
    'Determine values to interpolate between -- only changes when seconds digit has incremented
    millisec = Round((avgtime - t1) * 24 * 3600, 5) 'avgtime * 24 * 3600 - Int(avgtime * 24 * 3600)
    If millisec > 1 Then 'Seconds digits do not match
        row = row + 1
        millisec = millisec - 1
        'Time
        t1 = Worksheets("Export").Cells(row, 4)
        'Northing
        y1 = y2
        y2 = Worksheets("Export").Cells(row + 1, 1)
        'Easting
        x1 = x2
        x2 = Worksheets("Export").Cells(row + 1, 2)
        'RTK elevation
        rz1 = rz2
        rz2 = Worksheets("Combo").Cells(row + 1, 17)
    End If    'Calculate average bed elevation
    avgZ = rz1 + (rz2 - rz1) * millisec - Abs(avgZ)  'Bed elevation = Smoothed RTK elev - Abs(Avg. Sonar depth)
    'Write values to Export2
    Worksheets("Export2").Cells(exprow, 1) = y1 + (y2 - y1) * millisec
    Worksheets("Export2").Cells(exprow, 2) = x1 + (x2 - x1) * millisec
    Worksheets("Export2").Cells(exprow, 3) = avgZ
    Worksheets("Export2").Cells(exprow, 4) = avgtime
    Worksheets("Export2").Cells(exprow, 5) = son_id
    Worksheets("Export2").Cells(exprow, 6) = Worksheets("Export").Cells(row, 6)
    'Update processing % on progress bar
    UpdateProgBar Round(sonrow / lastson * 100, 0)
    'Increment row identifiers for Sonar and Export sheets
    sonrow = sonrow + pings
    exprow = exprow + 1
Loop
Unload ProgressBar
Worksheets("Export2").Range("A2:C" & exprow - 1).NumberFormat = "0.000"
Worksheets("Export2").Range("D2:D" & exprow - 1).NumberFormat = "mm/dd/yyyy hh:mm:ss.000"
MsgBox "Processing for " & son_id & " complete!" & vbCrLf & lastson - (sonrow - 1) & " sonar record" & IIf(lastson - (sonrow - 1) = 1, "", "s") & " truncated", , "Processing complete"

End Sub

Private Sub UpdateProgBar(pctcomplete As Single)
    With ProgressBar
        .Text.Caption = Format(pctcomplete, "0") & "% Complete"
        .Bar.Width = Round(pctcomplete, 0) * 2
    End With
    DoEvents
End Sub
