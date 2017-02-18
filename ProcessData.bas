Attribute VB_Name = "ProcessData"
Option Explicit

Dim basepath, path, record As String  'path to R000XX_Final_Template, path to record, and record name
Dim fileprompt As VbMsgBoxResult
Dim mybook As Workbook  '000XX_Final_Template workbook path
Dim row As Long  'Row counter
Dim seconds As Long  'Counts number of full seconds in sonar file
Dim DEfile As String  'file path to Data_explorer index file
Dim RTKfile As String  'file path to RTK data file
Const tlen = 0.114  'Transducer length from center of mounting pole to sonar projector
'Const utc_shift = 8  'Set value to shift sonar times forward/back to match RTK times (UTC)
    'For Pacific time, set to 8 for records during PST (winter Nov-Mar) and 7 during PDT (summer Mar-Nov)
    'Based on a sonar file stated in local time and RTK file stated in UTC time
Private Type linear  'Data type to store linear regression values
    b As Double  'Slope of regression
    a As Double  'Intercept of regression
    n As Integer  'Number of points in regression
End Type

Sub AssembleXYZ()
Attribute AssembleXYZ.VB_Description = "Process data from R000XX.DAT, R000XX.DAT.XYZ.csv and RTK files, interpolate missing data, smooth elevation and export bathymetry to R000XX.csv"
Attribute AssembleXYZ.VB_ProcData.VB_Invoke_Func = "D\n14"
' Process sonar, navigation and RTK data and export to csv
'
' ****Functions****
' Imports data from \R000XX\B002.idx and R000XX.DAT.XYZ.csv to Sonar worksheet;
' imports RTK data and adds lines for missing times;
' averages sonar data over each second and outputs data to Combo worksheet;
' fills missing RTK data based on X- and Y-offsets interpolated from sonar data;
' smooths RTK elevation values within min/max ranges;
' plots XY data in TrackPlot tab;
' plots elevation, depth and calculated channel bottom with error bars in Smoothing tab;
' saves processed record as R000XX_Final.xlsx and Export tab as R000XX.csv
'
' ****Notes****
' Move R000XX_Final_Template.xlsx to the directory containing the R000XX files of
' interest and run from that directory
' Make sure to update the constant 'utc_shift' (above) to account for daylight
' savings and different time zones
Dim IsFile As Boolean

'Read defaults from text file
    Set mybook = ActiveWorkbook
    path = mybook.FullName
    path = Left(path, InStrRev(path, "\"))  'Remove filename from path string
    basepath = path
    IOdefaults "read"

'Import data to Sonar worksheet
    Worksheets("Combo").Activate
    Range("A1").Activate
    If record = "na" Then
        fileprompt = vbNo
    Else
        fileprompt = MsgBox("Process the next record in series (last processed was " & record & ")?", vbYesNoCancel)
    End If
    If fileprompt = vbCancel Then
        Exit Sub
    Else
        If fileprompt = vbYes Then
            record = "R" & Format(Str(Int(Val(Right(record, 5)) + 1)), "00000")
        Else
            ChDrive path
            ChDir path  'Change default directory to last used
            path = Application.GetOpenFilename("Humminbird dat files (*.dat),*.dat", , "Navigate to the sonar .dat file to process")
            If path = "False" Then Exit Sub
            record = Left(Right(path, Len(path) - InStrRev(path, "\")), 6)  'Save record from root path
            path = Left(path, InStrRev(path, "\"))  'Cut filename from root path
        End If
    End If
    'Check whether R000XX_Final.xlsx already exists
    IsFile = False
    On Error Resume Next
    IsFile = GetAttr(path & record & "_Final.xlsx")
    If IsFile Then
        fileprompt = MsgBox(record & "_Final.xlsx already exists! Proceed with processing?", vbOKCancel, "Warning")
        If fileprompt = vbCancel Then Exit Sub
    End If
    
    'Check whether R000XX.DAT.XYX.csv exists (processed SonarTRX file)
    IsFile = False
    On Error Resume Next
    IsFile = GetAttr(path & record & ".DAT.XYZ.csv")
    If Not IsFile Then
        fileprompt = MsgBox(record & ".DAT.XYZ.csv sonar file not found! Processing cancelled.", vbExclamation, "File not found!")
        Exit Sub
    End If
    
    'Remove "Click to Process" button
    ActiveSheet.Shapes.Range(Array("Button 1")).Delete
    On Error GoTo 0  'Reset error handling
    
'Import sonar ping time stamps from IDX file
    DATimport
    If fileprompt = vbCancel Then Exit Sub
    
'Divide sonar data by full seconds and write to Combo sheet
    SonarCom

'Import pertinent RTK data to RTK worksheet and insert missing data lines
    RTKimport Cells(2, 2), Cells(seconds + 1, 2)
    If fileprompt = vbCancel Then Exit Sub

'Extract RTK data to Combo sheet
    RTKcom
    
'Interpolate gaps in RTK data
    InterpolateRTK

'Smooth RTK elevation points
    Critical

'Update Plotting ranges
    Sheets("TrackPlot").Select
    ActiveChart.FullSeriesCollection(1).XValues = "=Combo!$D$2:$D$" & seconds + 1  'Sonar X (Easting)
    ActiveChart.FullSeriesCollection(1).Values = "=Combo!$E$2:$E$" & seconds + 1  'Sonar Y (Northing)
    ActiveChart.FullSeriesCollection(2).XValues = "=Combo!$N$2:$N$" & seconds + 1  'RTK X
    ActiveChart.FullSeriesCollection(2).Values = "=Combo!$M$2:$M$" & seconds + 1  'RTK Y
    ActiveChart.ChartTitle.Text = record & " Navigation Tracks"
    Sheets("Smoothing").Select
    ActiveChart.FullSeriesCollection(1).Values = "=Combo!$Q$2:$Q$" & seconds + 1  'Smoothed elev
    ActiveChart.FullSeriesCollection(2).Values = "=Combo!$R$2:$R$" & seconds + 1  'Min elev
    ActiveChart.FullSeriesCollection(3).Values = "=Combo!$S$2:$S$" & seconds + 1  'Max elev
    ActiveChart.FullSeriesCollection(4).Values = "=Combo!$P$2:$P$" & seconds + 1  'Raw elev
    Sheets("Depth").Select
    ActiveChart.FullSeriesCollection(1).Values = "=Combo!$G$2:$G$" & seconds + 1  'Depth
    ActiveChart.FullSeriesCollection(2).Values = "=Export!$C$2:$C$" & seconds + 1  'Bed_elev

'Extract data to Export sheet
    Worksheets("Export").Activate
    'Write Excel formulas so changes to Combo are dynamically adjusted
    For row = 1 To seconds
        Cells(row + 1, 1).FormulaR1C1 = "=Combo!RC[12]"  'Northing
        Cells(row + 1, 2).FormulaR1C1 = "=Combo!RC[12]"  'Easting
        Cells(row + 1, 3).FormulaR1C1 = "=Combo!RC[14]-Combo!RC[4]"  'Bottom = SmoothedElevation - Depth
        Cells(row + 1, 4).FormulaR1C1 = "=Combo!RC[-2]"  'DateTime
        Cells(row + 1, 5).Value = record  'Sonar_ID
        Cells(row + 1, 6).FormulaR1C1 = "=IF(Combo!RC[5]="""",""N/A"",Combo!RC[5])"  'RTK_ID
    Next row
    Range(Cells(2, 1), Cells(seconds + 1, 1)).NumberFormat = "0.0000"
    Range(Cells(2, 2), Cells(seconds + 1, 2)).NumberFormat = "0.0000"
    Range(Cells(2, 3), Cells(seconds + 1, 3)).NumberFormat = "0.000"
    Range(Cells(2, 4), Cells(seconds + 1, 4)).NumberFormat = "m/d/yyyy hh:mm:ss"
    Range("A1").Select
    
'Save files
Dim SaveOn As VbMsgBoxResult
    Sheets("Combo").Select
    SaveOn = MsgBox("Save record and export files?", vbYesNo, "Processing Complete")
    If SaveOn = vbYes Then
        SaveFiles
        fileprompt = MsgBox("Record " & record & " successfully processed and saved. View processed file?", vbYesNo)
        If fileprompt = vbYes Then
            Workbooks.Open (path & record & "_Final.xlsx")
        End If
    Else
        MsgBox ("Record " & record & " processed but not saved or exported")
    End If
    IOdefaults "write"  'Write defaults to text file
End Sub

Private Sub IOdefaults(Optional IOtype As String = "read")
'Read/write last used record, DEfile, and RTKfile to/from R000XX_defaults.txt
Dim f As Integer  'File index number
    f = FreeFile
    ChDir (path) 'Set default path
    If IOtype = "read" Then
        On Error Resume Next
        Open "R000XX_defaults.txt" For Input As #f
        If Err.Number <> 0 Then
            record = "na"
            DEfile = "na"
            RTKfile = "na"
            Exit Sub
        End If
        On Error GoTo 0  'Reset error handling
        Input #f, record
        Input #f, path
        Input #f, DEfile
        Input #f, RTKfile
        Close #f
    Else  'Overwrite defaults file
        ChDir (basepath)  'Save in initial location
        Open "R000XX_defaults.txt" For Output As #f
        Write #f, record
        Write #f, path
        Write #f, DEfile
        Write #f, RTKfile
        Close #f
    End If
End Sub

Private Sub DATimport()
'Import sonar ping time stamps from IDX file
'   Humminbird sonar files consist of a DAT file, and an IDX and SON file for each channel:
'       DAT file -- holds record info for starting date and time, duration, and beginning coordinates
'       IDX files -- two 4-byte fields per record specifying a time increment and a line index in the SON file
'       SON files -- complex binary file with full navigation data and imagery for each record
Dim bytes() As Byte  'holds binary data from IDX file
Dim f As Integer  'file index number
Dim fLen As Long  'length of IDX file in bytes
Dim DEbook As Workbook  'Data Explorer workbook
Dim DATbook As Workbook  'R000XX.DAT.XYZ.csv path
Dim fulldate As Double  'Initial datetime value (days since 1900 + hr/24 + min/60 + sec/3600 + ms/3600/1000)
Dim utc_shift As Integer  'Value to shift sonar times forward/back to match RTK times (UTC)
    'For Pacific time, use 8 for records during PST (winter Nov-Mar) and 7 during PDT (summer Mar-Nov)
    'Based on a sonar file stated in local time and RTK file stated in UTC time
Dim cog As Double  'Course over ground

'Load binary data from IDX file
    f = FreeFile
    Open path & record & "\B002.IDX" For Binary Access Read As #f
    fLen = LOF(f)
    ReDim bytes(1 To fLen)
    Get f, , bytes
    Close f

'Choose data explorer file (DEfile) path
    fileprompt = MsgBox("Use default Data Explorer file path? (" & DEfile & ")", vbYesNoCancel, "Data Explorer source")
    If fileprompt = vbCancel Then
        Exit Sub
    Else
        If fileprompt = vbYes Then
            If DEfile = "na" Then
                MsgBox "Defaults file not found!" & vbCrLf & "Open file from list.", vbCritical, "File read error!"
            Else
                Application.ScreenUpdating = False
                On Error Resume Next  'Check that file opens without errors
                Set DEbook = Workbooks.Open(Filename:=DEfile, ReadOnly:=True)
                If Err.Number <> 0 Then
                    MsgBox "Default data explorer file not found!", vbCritical, "Invalid file path!"
                    DEfile = "na"
                End If
                On Error GoTo 0  'Reset error handling
            End If
        Else
            DEfile = "na"
        End If
    End If
    If DEfile = "na" Then
    'Select DEfile from explorer
        ChDir (path) 'Set default path
        DEfile = Application.GetOpenFilename("Excel Workbooks (*.xls;*.xlsx),*.xls;*.xlsx", , "Select data explorer file")
        Application.ScreenUpdating = False
        Set DEbook = Workbooks.Open(Filename:=DEfile, ReadOnly:=True)
    End If
'Read data and close file
    row = WorksheetFunction.Match(Val(Right(record, 5)), DEbook.Worksheets(1).Range("D2:D400"), 0) + 1  'Find record number in DEbook
    fulldate = DEbook.Worksheets(1).Cells(row, 1)  'Initial date from DEbook
    If Month(fulldate) > 2 And Month(fulldate) < 11 Then
        utc_shift = 7
    Else
        utc_shift = 8
    End If
    fileprompt = MsgBox("Accept time offset of UTC -" & utc_shift & "?", vbYesNo, "UTC offset")
    If fileprompt = vbNo Then
        MsgBox "Set UTC time offset for date of data collection (" & Month(fulldate) & "/" & Day(fulldate) & "/" & Year(fulldate) & "). Set the offset to 7 for Pacific Standard Time (winter months) or 8 for Pacific Daylight Time (summer months).", vbCritical, "Set UTC offset"
    End If
    fulldate = fulldate + DEbook.Worksheets(1).Cells(row, 6) + utc_shift / 24  'Shift times forward or back for UTC correction
    DEbook.Close
    Set DATbook = Workbooks.Open(Filename:=path & record & ".DAT.XYZ.csv", ReadOnly:=True)
    mybook.Worksheets("Sonar").Activate
    Range("A2").Select
    For row = 0 To fLen \ 8 - 1  'Backslash operator is integer division
        'First 4 bytes of each 8-byte record hold time stamp info
        Cells(row + 2, 1) = fulldate + (bytes(row * 8 + 2) * 65536 + bytes(row * 8 + 3) * CLng(256) + bytes(row * 8 + 4)) / 24 / 3600 / 1000
    Next row
    DATbook.Worksheets(1).Range("A2:C" & row + 1).Copy Destination:=Range("B2")  'Copy XYZ data from DATbook to Sonar sheet
    DATbook.Close
    Application.ScreenUpdating = True
'Calculate Course over ground (COG)
    For row = 1 To fLen \ 8 - 1
        'cog = mod(degrees(atan2(x2 - x1, y2 - y1) + 270), 360)
        On Error Resume Next  'No change in position results in division by zero error (Err.Number = 11)
        cog = WorksheetFunction.Atan2(Cells(row + 2, 2) - Cells(row + 1, 2), Cells(row + 2, 3) - Cells(row + 1, 3))
        If Err.Number <> 0 Then cog = 0  'Set COG to 0 for static position
        On Error GoTo 0  'Reset error handling
        cog = cog * 180 / WorksheetFunction.Pi + 270  'Switch to degrees from North
        'Note: can't use mod op. because VBA mod returns integers only
        If cog >= 360 Then cog = cog - 360  '0 < COG < 360
        Cells(row + 1, 5) = cog
    Next row
End Sub
 
Private Sub SonarCom()
'Divide sonar data by full seconds and write to Combo sheet
Dim srow As Long  'Cumulative number rows <= time being evaluated
Dim prev As Long  'Cumulative rows <= previous time increment
Dim t1, t2 As Double  'Time in milliseconds
Dim i, j As Long  'Counters
Dim avgX, avgY As Double  'Average X and Y position for each full second
Dim dep As Single  'Average depth value for each full second
Dim r As linear  'Store regression values
    
'Remove records at end with no change in position (common record artifact)
    If Round(Cells(row + 1, 2), 6) = Round(Cells(row, 2), 6) Or Round(Cells(row + 1, 3), 6) = Round(Cells(row, 3), 6) Then
    'Milliseconds aren't within one ping of 1000 so remove partial second
        Do While Round(Cells(row + 1, 2), 6) = Round(Cells(row, 2), 6) Or Round(Cells(row + 1, 3), 6) = Round(Cells(row, 3), 6)
        'Seconds digits match
            row = row - 1
        Loop
    End If
'Remove partial seconds at end of record count (row)
    Do While Int(Cells(row + 1, 1) * 24 * 3600) = Int(Cells(row, 1) * 24 * 3600)
    'Seconds digits match
        row = row - 1
    Loop
'Average Sonar records for X, Y, Depth, and calculated COG over full-second intervals
    seconds = Timer(Cells(row, 1), Cells(2, 1))  'Number of full seconds in data record
    Worksheets("Combo").Activate
    prev = 0
    srow = 0
    t1 = Worksheets("Sonar").Cells(srow + 2, 1)
    For i = 1 To seconds
        Cells(i + 1, 1) = i  'Index
        If i = 1 Then
            'Set first value to start time
            Cells(2, 2) = Worksheets("Sonar").Range("A2")
        Else
            'Increment sonar time by 1 second
            Cells(i + 1, 2) = Cells(i, 2) + 1 / 24 / 3600
        End If
        'Count number of rows in Sonar for each second of data
        t2 = Cells(i + 1, 2) + 1 / 24 / 3600  'Set t2 to time plus one second
        Do While (t1 < t2 - 0.000000005) And (Worksheets("Sonar").Cells(srow + 2, 1) <> "")
            '0.000000005 needed to correct for rounding errors
            srow = srow + 1
            t1 = Worksheets("Sonar").Cells(srow + 2, 1)  'Set t1 to ping time
        Loop
        Cells(i + 1, 3) = srow  'Row (starting row of each full-second record)
        avgX = 0
        avgY = 0
        'Correct raw positions from Sonar sheet for the transducer sonar projector offset distance
        'X = x1 + sin(radians(cog)) * tlen
        'Y = y1 - cos(radians(cog)) * tlen
        For j = prev + 1 To srow
            avgX = avgX + Worksheets("Sonar").Cells(j + 1, 2) + Sin(WorksheetFunction.Pi / 180 * Worksheets("Sonar").Cells(j + 1, 5)) * tlen
            avgY = avgY + Worksheets("Sonar").Cells(j + 1, 3) - Cos(WorksheetFunction.Pi / 180 * Worksheets("Sonar").Cells(j + 1, 5)) * tlen
        Next j
        'Write averaged X, Y, COG, and depth
        Cells(i + 1, 4) = avgX / (srow - prev)  'X
        Cells(i + 1, 5) = avgY / (srow - prev)  'Y
        Cells(i + 1, 6) = WorksheetFunction.Sum(Worksheets("Sonar").Range("E" & prev + 2 & ":E" & srow + 1)) / (srow - prev)  'COG
        dep = WorksheetFunction.Sum(Worksheets("Sonar").Range("D" & prev + 2 & ":D" & srow + 1)) / (srow - prev)  'Depth
        'If depth was recorded as Bottom Elevation from SonarTRX, change sign to positive
        If dep < 0 Then Cells(i + 1, 7) = -1 * dep Else Cells(i + 1, 7) = dep
        prev = srow
    Next i
    'In case the final COG is blank (when the record ends with an exact full second), fill in with slope
    If Cells(i, 6) = "" Then
        r = Regress(i - 1, 6, 5, , , row)
        Cells(i, 6) = r.a + r.b * (1 + r.n)
    End If
End Sub

Private Function Regress(yrow As Long, ycol As Integer, n As Integer, Optional direction As Integer = 1, _
  Optional CheckCol As Integer = 0, Optional maxrow As Long = 0) As linear
'Calculate slope of the regression line through n points in column ycol starting with Cells(yrow, ycol)
'  where values in ycol are assumed dependent on their row index
'The direction defaults to base extrapolation on points from yrow and up, providing parameters to fill blanks below
'Pass direction = -1 to base extrapolation on points from yrow and down, providing parameters to fill blanks above
'CheckCol will force consideration only of points with valid values in column CheckCol if it is defined (> 0)
'Maxrow is the last row from which data should be considered. It defaults to global variable seconds
'Returns data of type linear with intercept .a, slope .b, and ending independent variable index .n
'Examples for useage:
'   Dim r as linear
'   r = Regress(90, 4, 60, 1, 5) take 60 values from column 4, working upward from row 90 and
'     only picking values when column 5 is not empty, and stopping at or before row 2
'   r = Regress(5, 4, 10, -1, 5) will take 10 consecutive values from column 4, working downward from
'     row 5 and stopping when 10 values have been picked or maxrow is exceeded
'   To place extrapolated values in the next 3 cells:
'       r = Regress(yrow, ycol, n, 1)  'Direction is forward, use 1
'       For i = 1 To 3
'           Cells(yrow + i, ycol) = r.a + r.b * (r.n + i)
'       Next i
'   To place extrapolated values in the previous 3 cells:
'       r = Regress(yrow, ycol, n, -1)  'Direction is backward, use -1
'       For i = 1 To 3
'           Cells(yrow - i, ycol) = r.a + r.b * (r.n + i)
'       Next i
'
Dim count, x, i As Integer  'Counters
Dim y(), xval() As Double  'y holds values for regression, c holds indices
Dim ybar, xbar As Double  'Average values of y and x
'Dim alpha, beta As Double  'Regression line coefficients
ReDim y(1 To n)
ReDim xval(1 To n)

'Check for valid direction
    If Abs(direction) <> 1 Then
    'Assume invariable
        With Regress
            .b = 0
            .a = Cells(yrow, ycol)
            .n = 1
        End With
        MsgBox "Invalid direction passed to function Regress. Direction must be passed as 1 " & _
            "for slope from preceeding points, or -1 for slope from proceeding points.", vbOKOnly
        Exit Function
    End If
'Initialize
    x = 0
    count = 0
    ybar = 0
    xbar = 0
    If maxrow = 0 Then maxrow = seconds + 1 'set to default
'Gather data
    Do Until count = n Or yrow + x > maxrow Or yrow + x < 2
        If CheckCol > 0 Then
            If Cells(yrow + x, CheckCol) <> "" And Cells(yrow + x, CheckCol).Interior.Color <> 49407 Then
            'Interpolated cells are color 49407
                count = count + 1
                y(count) = Cells(yrow + x, ycol)
                ybar = ybar + y(count)
                xval(count) = Abs(x) + 1
            End If
        Else
            count = count + 1
            y(count) = Cells(yrow + x, ycol)
            ybar = ybar + y(count)
            xval(count) = Abs(x) + 1
        End If
        x = x - direction
    Loop
'Calculate averages
    If CheckCol = 0 Then
        xbar = (count + 1) / 2 'average of 1..n
    Else
        For i = 1 To count
            xbar = xbar + Abs(x) - xval(i) + 1
        Next i
        xbar = xbar / count
    End If
    ybar = ybar / count
'Calculate results
    With Regress
        .n = Abs(x)  'index of independent variable x
        For i = 1 To count
            .b = .b + (.n - xval(i) + 1 - xbar) * (y(i) - ybar) 'sum numerator
            .a = .a + (.n - xval(i) + 1 - xbar) ^ 2  'sum denomenator
        Next i
        'Set final values
        .b = .b / .a  'slope
        .a = ybar - .b * xbar  'intercept
    End With
End Function

Private Function ExSlope(x1, x2, x3 As Variant) As Double
'Returns an incremental value based on the slope to be used for extrapolation
'For extrapolation before known data, pass x1, x2 and x3 in order
'  x1 is the known value after an initial gap to extrapolate
'  x2 and x3 may or may not be present
'Reverse order of points to extrapolate slope after known data
Dim p2, p3 As Byte  'Used to ignore missing data in slope averaging
    If x2 = "" Then
        x2 = 0
        p2 = 0
    Else: p2 = 1
    End If
    If x3 = "" Then
        x3 = 0
        p3 = 0
    Else: p3 = 1
    End If
    'p2 and p3 are used to ignore zero values in slope averaging
    ExSlope = (x1 - x2 + (x2 - x3) * p3) * Round((p2 + p3 + 0.1) / 2, 0) / (1 + p3)
    'Added 0.1 forces roundup behavior for 0.5 (VBA usually rounds 0.5 down)
End Function

Private Sub RTKimport(date1 As Date, date2 As Date)
'Import RTK data to RTK worksheet, delete duplicate entries, and insert lines for missing times
Dim RTKbook As Workbook
Dim RTKstart, RTKend As Long  'Starting and ending row numbers in RTKbook to copy over
Dim rng As Range
'XXX Dim dt As Long  'Number of seconds between consecutive RTK records
'XXX Dim dx, dy, dz As Integer
'XXX Dim i As Long  'Counter
'XXXX new items
'Dim dpos As Single  'distance between two consecutive RTK records
'XXXX

'Clear anything past header row on Worksheets("RTK")
    Set rng = Worksheets("RTK").UsedRange
    Set rng = rng.Offset(1, 0).Resize(rng.Rows.count - 1)
    rng.ClearContents
'Choose RTK file path
    fileprompt = MsgBox("Use default RTK data file path? (" & RTKfile & ")", vbYesNoCancel, "RTK data source")
    If fileprompt = vbCancel Then
        Exit Sub
    Else
        If fileprompt = vbYes Then
            If RTKfile = "na" Then
                MsgBox "Defaults file not found!" & vbCrLf & "Open file from list.", vbCritical, "File read error!"
            Else
                Application.ScreenUpdating = False
                On Error Resume Next  'Check that file opens without errors
                Set RTKbook = Workbooks.Open(Filename:=RTKfile, ReadOnly:=True)
                If Err.Number <> 0 Then
                    MsgBox "Default RTK data file not found!", vbCritical, "Invalid file path!"
                    RTKfile = "na"
                    Application.ScreenUpdating = True
                End If
                On Error GoTo 0  'Reset error handling
            End If
        Else
            RTKfile = "na"
        End If
    End If
    If RTKfile = "na" Then
    'Select RTKfile from explorer
        ChDir (path) 'Set default path
        RTKfile = Application.GetOpenFilename("Excel Workbooks (*.xls;*.xlsx),*.xls;*.xlsx", , "Select RTK data file")
        Application.ScreenUpdating = False
        Set RTKbook = Workbooks.Open(Filename:=RTKfile, ReadOnly:=True)
    End If
'Import RTK data and close file
    RTKstart = FindRow(date1, "initial")
    RTKend = FindRow(date2, "final")
    'Copy data from RTKbook
'    RTKbook.Worksheets(1).Range("A" & RTKstart & ":Q" & RTKend).Copy Destination:=mybook.Worksheets("RTK").Range("A2")
    mybook.Worksheets("RTK").Range("A2:Q" & 2 + RTKend - RTKstart).Value = RTKbook.Worksheets(1).Range("A" & RTKstart & ":Q" & RTKend).Value
    Application.DisplayAlerts = False  'Prevent "Save changes" dialog
    RTKbook.Close
    Application.DisplayAlerts = True
    mybook.Worksheets("RTK").Activate
    Application.ScreenUpdating = True
'Remove duplicates and add rows for missing entries
    row = 2
'XXXX must change to delete duplicated positions, not times since 1/m may have multiple same-second records
'Move entire block to add rows just to the Combo sheet section; don't need blank rows on the RTK sheet
'Old workflow: empty rows added to rtk sheet here for missing records, info on combo sheet later evaluated to see if it was empty using simple for loop
'New workflow: just delete duplicates here, later when the info is read to combo sheet average full-second records and add missing ones
'CHECK TO SEE IF RESULTING ROW VALUE IS USED IN NEXT SUB, IT MAY BE WRONG NOW (COULD USE TIMER FUNCTION TO COUNT END-START SECONDS)
    Do While Cells(row + 1, 1) > 0
        If Cells(row + 1, 4) = Cells(row, 4) Then
            If Cells(row + 1, 5) = Cells(row, 5) Then
                Cells(row, 1).EntireRow.Delete
            End If
        End If
        row = row + 1
    Loop
'XXXXX old dt loop
'    Do While Cells(row + 1, 1) > 0
'        dt = Timer(Cells(row + 1, 3), Cells(row, 3))
'        If dt = 0 Then  'The RTK output often creates duplicate records after gaps, remove duplicates here
'                Cells(row, 1).EntireRow.Delete
'        Else
'            If dt > 1 And dt < 7200 Then  'No more than two hour gap
'                Rows(row + 1 & ":" & row + dt - 1).Insert
'                For i = 1 To dt - 1
'                    Cells(row + i, 3) = Cells(row + i - 1, 3) + 1 / 24 / 3600
'                Next i
'            End If
'        End If
'        row = row + dt
'    Loop
End Sub

Private Sub RTKcom()
'Write RTK data to Combo worksheet
'Assumes RTK data are already adjusted to XYZ position at base of pipe and level of sonar emitter
Dim onepersec As Boolean  'True for RTK set to one reading per second, false for other intervals or distance-based
Dim start As Long
Dim cog As Double  'Course over ground
Dim Xrtk, Yrtk As Double  'X and Y position
Dim i As Long  'Counter
Dim r As linear  'Store regression values
Const maxZ = 0.03  'Maximum allowable RTK precision in Z; 3 cm based on equipment limitations
Const maxXY = 0.15  'Maximum allowable RTK precision in XY; based on limiting overall uncertainty
    
'XXXXX debug setting
onepersec = False
'XXXX
mybook.Worksheets("Combo").Activate
'Set start row in RTK sheet to offset if there is an initial gap in data
start = Round((Range("B2") - Worksheets("RTK").Range("C2")) * 24 * 3600, 0) + 1
For i = 1 To seconds
    cog = 9999  'Initialize to invalid value to test for changes
    If i + start < 2 Then
    'RTK record starts later than sonar, can't interpolate
        Cells(i + 1, 12) = Cells(i + 1, 2) 'No matching RTK time, use time from Combo sheet
        Cells(i + 1, 20) = "> " & Format(Str(maxXY), "0.00")  'StDev_XY
        Cells(i + 1, 21) = maxZ  'StDev_Z set to max for later smoothing
        Cells(i + 1, 22) = "None,Extrapolated"  'RTK Solution Type
    Else
        Cells(i + 1, 12) = Worksheets("RTK").Cells(i + start, 3)  'RTK Time
        If Worksheets("RTK").Cells(i + start, 7) <= maxXY And Worksheets("RTK").Cells(i + start, 1) <> "" Then
        'Horizontal precision OK
            Cells(i + 1, 11) = Worksheets("RTK").Cells(i + start, 2)  'Point_ID
            Cells(i + 1, 20) = Worksheets("RTK").Cells(i + start, 7)  'StDev_XY
            'Calculate COG_RTK, Northing and Easting
            If Worksheets("RTK").Cells(i + start + 1, 7) <= maxXY And Worksheets("RTK").Cells(i + start + 1, 1) <> "" Then
            'Next horizontal precision is OK: calculate COG, X and Y between two valid points
            'cog = mod(degrees(atan2(x2 - x1, y2 - y1)) + 270, 360)
                On Error Resume Next  'No change in position results in division by zero error (Err.Number = 11)
                cog = WorksheetFunction.Atan2( _
                    Worksheets("RTK").Cells(i + start + 1, 5) - Worksheets("RTK").Cells(i + start, 5), _
                    Worksheets("RTK").Cells(i + start + 1, 4) - Worksheets("RTK").Cells(i + start, 4))
                If Err.Number <> 0 Then cog = 0  'Set COG to 0 for static position
                On Error GoTo 0  'Reset error handling
                cog = cog * 180 / WorksheetFunction.Pi + 270  'Switch to zero degrees North
                If cog >= 360 Then cog = cog - 360  'Force 0 < COG < 360
            Else  'No second point: calculate COG from slope of valid points above
                If i > 1 And Cells(i, 15) <> "" Then 'Previous point has valid COG
                    r = Regress(i, 15, 5, , 15)  'Regression through last five valid COGs
                    cog = r.a + r.b * (1 + r.n)  'y = a + bx
                End If
            End If
            If cog <> 9999 Then  'Check if cog was calculated
                Cells(i + 1, 15) = cog  'COG_RTK
                'Use XY position and COG to offset transducer
                Xrtk = Worksheets("RTK").Cells(i + start, 5) + Sin(WorksheetFunction.Pi / 180 * cog) * tlen
                Yrtk = Worksheets("RTK").Cells(i + start, 4) - Cos(WorksheetFunction.Pi / 180 * cog) * tlen
                Cells(i + 1, 8) = Xrtk - Cells(i + 1, 4)  'X-offset = Xrtk - Xsonar
                Cells(i + 1, 9) = Yrtk - Cells(i + 1, 5)  'Y-offset = Yrtk - Ysonar
                Cells(i + 1, 14) = Xrtk  'Easting
                Cells(i + 1, 13) = Yrtk  'Northing
            End If
            If Worksheets("RTK").Cells(i + start, 8) <= maxZ Then
            'Horizontal and vertical precision OK
                Cells(i + 1, 16) = Worksheets("RTK").Cells(i + start, 6)  'Elev
                Cells(i + 1, 21) = Worksheets("RTK").Cells(i + start, 8)  'StDev_Z
                Cells(i + 1, 22) = Worksheets("RTK").Cells(i + start, 17)  'RTK Solution Type
            Else
            'Only horizontal precision OK
                Cells(i + 1, 21) = maxZ  'StDev_Z set to max for later smoothing
                Cells(i + 1, 22) = "Float,Horizontal"  'RTK Solution Type
            End If
        Else
        'Neither horizontal nor vertial precision OK
            Cells(i + 1, 20) = "> " & Format(Str(maxXY), "0.00")  'StDev_XY
            Cells(i + 1, 21) = maxZ  'StDev_Z set to max for later smoothing
            Cells(i + 1, 22) = "None,Interpolated"  'RTK Solution Type
        End If
    End If
    Cells(i + 1, 18).FormulaR1C1 = "=RC[-2] - RC[3]"  'Min = Elev - StDev_Z
    Cells(i + 1, 19).FormulaR1C1 = "=RC[-3] + RC[2]"  'Max = Elev + StDev_Z
Next i

End Sub

Private Sub RTKcomOld()
'Write RTK data to Combo worksheet
'Assumes RTK data are already adjusted to XYZ position at base of pipe and level of sonar emitter
Dim start As Long
Dim cog As Double  'Course over ground
Dim Xrtk, Yrtk As Double  'X and Y position
Dim i As Long  'Counter
Dim r As linear  'Store regression values
Const maxZ = 0.03  'Maximum allowable RTK precision in Z; 3 cm based on equipment limitations
Const maxXY = 0.15  'Maximum allowable RTK precision in XY; based on limiting overall uncertainty
    
mybook.Worksheets("Combo").Activate
'Set start row in RTK sheet to offset if there is an initial gap in data
start = Round((Range("B2") - Worksheets("RTK").Range("C2")) * 24 * 3600, 0) + 1
For i = 1 To seconds
    cog = 9999  'Initialize to invalid value to test for changes
    If i + start < 2 Then
    'RTK record starts later than sonar, can't interpolate
        Cells(i + 1, 12) = Cells(i + 1, 2) 'No matching RTK time, use time from Combo sheet
        Cells(i + 1, 20) = "> " & Format(Str(maxXY), "0.00")  'StDev_XY
        Cells(i + 1, 21) = maxZ  'StDev_Z set to max for later smoothing
        Cells(i + 1, 22) = "None,Extrapolated"  'RTK Solution Type
    Else
        Cells(i + 1, 12) = Worksheets("RTK").Cells(i + start, 3)  'RTK Time
        If Worksheets("RTK").Cells(i + start, 7) <= maxXY And Worksheets("RTK").Cells(i + start, 1) <> "" Then
        'Horizontal precision OK
            Cells(i + 1, 11) = Worksheets("RTK").Cells(i + start, 2)  'Point_ID
            Cells(i + 1, 20) = Worksheets("RTK").Cells(i + start, 7)  'StDev_XY
            'Calculate COG_RTK, Northing and Easting
            If Worksheets("RTK").Cells(i + start + 1, 7) <= maxXY And Worksheets("RTK").Cells(i + start + 1, 1) <> "" Then
            'Next horizontal precision is OK: calculate COG, X and Y between two valid points
            'cog = mod(degrees(atan2(x2 - x1, y2 - y1)) + 270, 360)
                On Error Resume Next  'No change in position results in division by zero error (Err.Number = 11)
                cog = WorksheetFunction.Atan2( _
                    Worksheets("RTK").Cells(i + start + 1, 5) - Worksheets("RTK").Cells(i + start, 5), _
                    Worksheets("RTK").Cells(i + start + 1, 4) - Worksheets("RTK").Cells(i + start, 4))
                If Err.Number <> 0 Then cog = 0  'Set COG to 0 for static position
                On Error GoTo 0  'Reset error handling
                cog = cog * 180 / WorksheetFunction.Pi + 270  'Switch to zero degrees North
                If cog >= 360 Then cog = cog - 360  'Force 0 < COG < 360
            Else  'No second point: calculate COG from slope of valid points above
                If i > 1 And Cells(i, 15) <> "" Then 'Previous point has valid COG
                    r = Regress(i, 15, 5, , 15)  'Regression through last five valid COGs
                    cog = r.a + r.b * (1 + r.n)  'y = a + bx
                End If
            End If
            If cog <> 9999 Then  'Check if cog was calculated
                Cells(i + 1, 15) = cog  'COG_RTK
                'Use XY position and COG to offset transducer
                Xrtk = Worksheets("RTK").Cells(i + start, 5) + Sin(WorksheetFunction.Pi / 180 * cog) * tlen
                Yrtk = Worksheets("RTK").Cells(i + start, 4) - Cos(WorksheetFunction.Pi / 180 * cog) * tlen
                Cells(i + 1, 8) = Xrtk - Cells(i + 1, 4)  'X-offset = Xrtk - Xsonar
                Cells(i + 1, 9) = Yrtk - Cells(i + 1, 5)  'Y-offset = Yrtk - Ysonar
                Cells(i + 1, 14) = Xrtk  'Easting
                Cells(i + 1, 13) = Yrtk  'Northing
            End If
            If Worksheets("RTK").Cells(i + start, 8) <= maxZ Then
            'Horizontal and vertical precision OK
                Cells(i + 1, 16) = Worksheets("RTK").Cells(i + start, 6)  'Elev
                Cells(i + 1, 21) = Worksheets("RTK").Cells(i + start, 8)  'StDev_Z
                Cells(i + 1, 22) = Worksheets("RTK").Cells(i + start, 17)  'RTK Solution Type
            Else
            'Only horizontal precision OK
                Cells(i + 1, 21) = maxZ  'StDev_Z set to max for later smoothing
                Cells(i + 1, 22) = "Float,Horizontal"  'RTK Solution Type
            End If
        Else
        'Neither horizontal nor vertial precision OK
            Cells(i + 1, 20) = "> " & Format(Str(maxXY), "0.00")  'StDev_XY
            Cells(i + 1, 21) = maxZ  'StDev_Z set to max for later smoothing
            Cells(i + 1, 22) = "None,Interpolated"  'RTK Solution Type
        End If
    End If
    Cells(i + 1, 18).FormulaR1C1 = "=RC[-2] - RC[3]"  'Min = Elev - StDev_Z
    Cells(i + 1, 19).FormulaR1C1 = "=RC[-3] + RC[2]"  'Max = Elev + StDev_Z
Next i

End Sub

Private Function FindRow(dateval As Date, matchtype As String) As Long
'Return row number from RTK data file that matches a given date
'RTK data file must have been activated by calling sub
'matchtype should be passed as "initial" or "final"
Dim step, index As Long
Dim lastRow As Long
Dim dt As Integer
    lastRow = ActiveSheet.Cells(Rows.count, "A").End(xlUp).row - 1
    index = Int(lastRow / 2) 'start at midpoint
    step = index 'initial step will become half the size of index in loop below
    Do Until Abs(dateval - Cells(index + 1, 3)) < 1 / 48 / 3600 Or Abs(step) = 1
        step = Abs(step)  'reset step to positive after each iteration
        If dateval < Cells(index + 1, 3) Then
            step = -Int((step + 1) / 2)  'step = -Roundup(step/2)
        Else
            step = Int((step + 1) / 2)  'step = Roundup(step/2)
        End If
        index = index + step
    Loop
    If Abs(step) = 1 Then
        'No exact time match, so move to one row below low value or one above high value
        If Abs(dateval - Cells(index + 1, 3)) * 24 * 3600 > 0.5 Then
            If matchtype = "initial" Then
                If dateval < Cells(index + 1, 3) Then
                    index = index - 1
                End If
            Else
                If dateval > Cells(index + 1, 3) Then
                    index = index + 1
                End If
'                If Abs(dateval - Cells(index + 1, 3)) <= 1 / 24 Then  'Less than one hour gap
'                     index = index + 1
'                End If
            End If
            If Abs(dateval - Cells(index + 1, 3)) <= 1 / 24 Then  'Less than one hour gap
                dt = Round(Abs(dateval - Cells(index + 1, 3)) * 24 * 3600, 0)
                MsgBox ("No matching time for " & matchtype & " bound (" & dateval & ") in RTK file " & Chr(151) & _
                    " the next closest time will be used (" & dt & " second difference)")
            Else
                MsgBox ("No matching time for " & matchtype & " bound (" & dateval & ") in RTK file " & Chr(151) & _
                    " " & matchtype & " values will be extrapolated.")
                'Ignore data more than an hour away from target
                If matchtype = "initial" Then
                    index = index + 1
                Else: index = index - 1
                End If
            End If
        End If
    End If
    FindRow = index + 1  'set value to actual worksheet row number
End Function

Private Sub InterpolateRTK()
'Interpolate RTK X and Y based on linear offset from sonar X and Y
'  with RTK Z interpolated linearly
'X and Y positions and offsets already corrected for transducer position based on COG
'  so interpolated positions do not need to be corrected again
Dim count As Long  'Number of blanks to interpolate or extrapolate
Dim RTKrows As Long  'Number of rows in RTK sheet
Dim incX, incY, incZ As Double  'Increments to offset X, Y and Z for interpolation
Dim i, j As Long  'Counter
Dim span As Long  'Number of seconds between known RTK points
Dim warn() As Boolean  'Store record of interpolated and extrapolated areas for message output
Dim warning As String  'Warning message for output
Dim xr As linear
Dim yr As linear
Dim zr As linear  'Store regression values
Dim normLog As Single  'Used for logarithmic damping of slope extrapolation
Dim xmean, ymean As Double  'Mean of adjacent 120 valid X- and Y-offsets
Dim target As Long  'Start row for interpolation or extrapolation
Dim dir As Integer  'Direction of interpolation or extrapolation
Const damper = 7  'Number of seconds to dampen X-, Y- and Z-offset trends on log scale before they reach zero
Const window = 100  'Number of seconds to average offsets forward and backward to find mean
ReDim warn(1 To 7)

'Interpolate Northing, Easting
i = 1
Do While Cells(i + 1, 1) <> ""  'Index isn't blank
    If Cells(i + 1, 8) = "" Then  'X-offset is blank
        'Count blanks
        count = 0
        Do While Cells(i + count + 1, 8) = "" And i + count <= seconds
            count = count + 1
        Loop
        If count = seconds Then  'No RTK position for whole record
            For j = 1 To seconds
                Cells(j + 1, 14) = Cells(j + 1, 4)  'Copy X from sonar
                Cells(j + 1, 13) = Cells(j + 1, 5)  'Copy Y from sonar
            Next j
            warn(1) = True
        ElseIf i = 1 Or i + count = seconds + 1 Then 'Initial or terminal gap in RTK
            'Find regression slope of X- and Y-offsets from known points and produce a smooth curve towards the average
            If i = 1 Then  'Initial gap
                dir = -1
                target = count + 2
                warn(2) = True
            Else  'Terminal gap
                dir = 1
                target = i
                warn(3) = True
            End If
            xr = Regress(target, 8, 6, dir, 11)  'Regression of X-offset through 6 points with valid Point_ID
            yr = Regress(target, 9, 6, dir, 11)  'Regression of Y-offset through 6 points with valid Point_ID
            xmean = MidMean(window, target, 8)  'Average X-offset within +/- <window> values
            ymean = MidMean(window, target, 9)  'Average Y-offset within +/- <window> values
            For j = 1 To count
            'Calculate X and Y offsets
                If j < damper + 2 Then
                    normLog = WorksheetFunction.Log(damper - j + 2, damper + 1)
                    Cells(target + j * dir, 8) = Cells(target + (j - 1) * dir, 8) + xr.b * normLog
                    Cells(target + j * dir, 9) = Cells(target + (j - 1) * dir, 9) + yr.b * normLog
                Else
                    If j <= damper * 2 Then normLog = WorksheetFunction.Log(j - damper, damper + 1) Else normLog = 1
                    Cells(target + j * dir, 8) = Cells(target + (j - 1) * dir, 8) + (xmean - Cells(target + (j - 1) * dir, 8)) / damper * normLog
                    Cells(target + j * dir, 9) = Cells(target + (j - 1) * dir, 9) + (ymean - Cells(target + (j - 1) * dir, 9)) / damper * normLog
                End If
            Next j
        Else  'Middle gap--interpolate between two valid values
            xr.b = (Cells(i + count + 1, 8) - Cells(i, 8)) / (count + 1)
            yr.b = (Cells(i + count + 1, 9) - Cells(i, 9)) / (count + 1)
            For j = 1 To count
                Cells(i + j, 8) = Cells(i + j - 1, 8) + xr.b  'X-offset
                Cells(i + j, 9) = Cells(i + j - 1, 9) + yr.b  'Y-offset
            Next j
        End If
        'Use formula for Northing and Easting in case X-offset and Y-offset are changed manually
        For j = 1 To count
            Cells(i + j, 13).FormulaR1C1 = "=RC[-8] + RC[-4]"  'Northing
            Cells(i + j, 14).FormulaR1C1 = "=RC[-10] + RC[-6]"  'Easting
        Next j
        'Highlight records with interpolated data in orange
        Range("H" & i + 1, "I" & i + count).Interior.Color = 49407
        Range("M" & i + 1, "N" & i + count).Interior.Color = 49407
        i = i + count
    Else
        i = i + 1
    End If
Loop

'Interpolate Elevation gaps
i = 1
Do While Cells(i + 1, 1) <> ""  'Index isn't blank
    If Cells(i + 1, 16) = "" Then  'Elev is blank
        count = 0
        Do While Cells(i + count + 1, 16) = "" And i + count <= seconds
            count = count + 1
        Loop
        If count = seconds Then  'No RTK elevation for whole record
            'Check starting RTK values
            RTKrows = Worksheets("RTK").Cells(Rows.count, "B").End(xlUp).row  'Last used row in RTK worksheet
            span = Timer(Worksheets("RTK").Cells(RTKrows, 3), Worksheets("RTK").Cells(2, 3))
            If span <= 7200 Then  'Less than two hours between known RTK end points
                warn(4) = True
                zr.b = (Worksheets("RTK").Cells(RTKrows, 6) - Worksheets("RTK").Cells(2, 6)) / span
                zr.n = CInt(Timer(Cells(2, 12), Worksheets("RTK").Cells(2, 3)))
                Cells(2, 16) = Worksheets("RTK").Cells(2, 6) + zr.b * (zr.n + 1)  'First elevation
                For j = 2 To seconds
                    Cells(j + 1, 16) = Cells(j, 16) + zr.b  'Elevation
                Next j
            Else
                warn(5) = True
            End If
        ElseIf i = 1 Or i + count = seconds + 1 Then 'Initial or terminal gap in RTK
            'Use the regression slope of elevation to extrapolate
            If i = 1 Then  'Initial gap
                dir = -1
                target = count + 2
                warn(6) = True
            Else  'Terminal gap
                dir = 1
                target = i
                warn(7) = True
            End If
            zr = Regress(target, 16, 300, dir, 16) 'Regression of Elevation through 300 non-blank points
        Else  'Middle gap--interpolate between two valid values
            dir = 1
            target = i
            zr.b = (Cells(i + count + 1, 16) - Cells(i, 16)) / (count + 1)  'Slope
        End If
        For j = 1 To count
            Cells(target + j * dir, 16) = Cells(target + (j - 1) * dir, 16) + zr.b  'Elev
        Next j
        'Highlight records with interpolated data in orange
        Range("P" & i + 1, "P" & i + count).Interior.Color = 49407
        i = i + count
    Else
        i = i + 1
    End If
Loop
        
'Output warnings
    warning = ""
    If warn(1) Or warn(4) Or warn(5) Then
        warning = "No RTK " & IIf(warn(1), "position " & IIf(warn(4) Or warn(5), "or elevation ", ""), "elevation ") & "data for full record! "
        If warn(1) Then warning = warning & "Position copied from sonar record"
        If warn(4) Then
            warning = warning & IIf(warn(1), "; e", "E") & "levation interpolated between last valid RTK points."
        ElseIf warn(5) Then
            warning = warning & IIf(warn(1), "; n", "N") & "o RTK elevation within one hour of record end points--interpolation not performed."
        Else
            warning = warning & "."
        End If
    Else
        If warn(2) Or warn(6) Then
            warning = "No starting RTK " & IIf(warn(2), "position " & IIf(warn(6), "or elevation ", ""), "elevation ") & "data! "
        End If
        If warn(3) Or warn(7) Then
            warning = warning & "No ending RTK " & IIf(warn(3), "position " & IIf(warn(7), "or elevation ", ""), "elevation ") & "data! "
        End If
        If Len(warning) > 1 Then warning = warning & "Check fit of extrapolated values."
    End If
    If Len(warning) > 1 Then MsgBox warning
    
End Sub

Private Function MidMean(maxcount As Integer, vrow, vcol As Long) As Double
'Calculate average of maxcount values in column vcol both above and below row vrow
Dim k, n1, n2 As Integer 'Counters

MidMean = Cells(vrow, vcol)
k = 0
n1 = 0
Do While k < maxcount And vrow - k > 2
    k = k + 1
    If Cells(vrow - k, vcol) <> "" Then
        n1 = n1 + 1
        MidMean = MidMean + Cells(vrow - k, vcol)
    End If
Loop
k = 0
n2 = 0
Do While k < maxcount And vrow + k <= seconds
    k = k + 1
    If Cells(vrow + k, vcol) <> "" Then
        n2 = n2 + 1
        MidMean = MidMean + Cells(vrow + k, vcol)
    End If
Loop
MidMean = MidMean / (n1 + n2 + 1)

End Function

Private Function Timer(time2, time1 As Date) As Long
'Evaluate time difference in seconds between to dates
    If (Hour(time2) - Hour(time1)) < 0 Then  'Check for UTC time passing midnight
        Timer = (24 - Hour(time1) + Hour(time2)) * CLng(3600) + (Minute(time2) - Minute(time1)) * 60 + Second(time2) - Second(time1)
    Else
        Timer = (Hour(time2) - Hour(time1)) * CLng(3600) + (Minute(time2) - Minute(time1)) * 60 + Second(time2) - Second(time1)
    End If
End Function

Sub Critical()
' Identify critical points in elevation and interpolate linearly
' between critical points. Points are stored in crit() with values
' in yc(), which could be passed as arguments to a more sophisticated
' interpolation technique such as splining.
Dim zmin(), zmax() As Single  'Min/max possible elevation from RTK
Dim k, maxi, maxj As Integer
Dim i, j As Long
Dim z() As Single  'Stores values of critical points
Dim crit() As Integer  'Stores indices of critical points
Dim critChange As Boolean
Dim m, b As Single
ReDim zmin(1 To seconds)
ReDim zmax(1 To seconds)
ReDim z(1 To seconds)
ReDim crit(1 To seconds)

'Determine which starting value stays within min/max bounds the longest
    zmin(1) = Range("R2")
    zmax(1) = Range("S2")
    zmin(2) = Range("R3")
    zmax(2) = Range("S3")
    maxi = 0
    For i = 0 To 9 'Test 10 starting values ranging from zmin to zmax
        z(1) = (zmax(1) - zmin(1)) / 9 * i + zmin(1)
        j = 2
        Do While z(1) >= zmin(j) And z(1) <= zmax(j)
            j = j + 1
            zmin(j) = Cells(j + 1, 18)
            zmax(j) = Cells(j + 1, 19)
        Loop
        If j > maxj Then
            maxj = j
            maxi = i
        Else
            If j = maxj And Abs(i - 4.5) < Abs(maxi - 4.5) Then  'Where two j's are equivalent, choose i closer to the middle
                maxi = i
            End If
        End If
    Next i
    'Set z(1) to best starting value
    z(1) = (zmax(1) - zmin(1)) / 9 * maxi + zmin(1)

'Initialize
    crit(1) = 1
    k = 2
    Cells(2, 17) = z(1)  'Initial smoothed value

'Find critical points and interpolate between
For i = 2 To seconds
    zmin(i) = Cells(i + 1, 18)
    zmax(i) = Cells(i + 1, 19)
    critChange = False
    
    'Check if last critical value exceeds current max or min
    If z(crit(k - 1)) > zmax(i) Then
        'Reset k for two consecutive decreasing critical values
        If crit(k - 1) = i - 1 And (z(i - 1) = zmax(i - 1) And i > maxi) Then k = k - 1
        z(i) = zmax(i)
        crit(k) = i
        critChange = True
    ElseIf z(crit(k - 1)) < zmin(i) Then
        'Reset k for two consecutive increasing critical values
        If crit(k - 1) = i - 1 And (z(i - 1) = zmin(i - 1) And i > maxi) Then k = k - 1
        z(i) = zmin(i)
        crit(k) = i
        critChange = True
    ElseIf i = seconds Then
        'Extrapolate for tail values after last critical point
        '  only if z(seconds) is not already a critical point
        z(i) = m * i + b
        crit(k) = seconds
        critChange = True
    End If
    
    'Check if linear interpolation violates min/max boundaries
    If critChange Then
        m = (z(i) - z(crit(k - 1))) / (crit(k) - crit(k - 1))
        b = z(i) - m * crit(k)
        For j = crit(k - 1) + 1 To crit(k) - 1
            If j = i Then   'Exit loop when i has been reset to lesser value (boundaries were violated)
                Exit For
            End If
            z(j) = m * j + b
            If z(j) > zmax(j) Then
                z(j) = zmax(j)
                crit(k) = j
                maxi = i
                i = j
                m = (z(j) - z(crit(k - 1))) / (crit(k) - crit(k - 1))
                b = z(j) - m * crit(k)
                j = crit(k - 1)
            ElseIf z(j) < zmin(j) Then
                z(j) = zmin(j)
                crit(k) = j
                maxi = i
                i = j
                j = crit(k - 1)
                m = (z(i) - z(j)) / (crit(k) - j)
                b = z(i) - m * crit(k)
            End If
            Cells(j + 1, 17) = z(j)
        Next j
        k = k + 1
    End If
    Cells(i + 1, 17) = z(i)
Next i

End Sub

Private Sub SaveFiles()
'Save worksheet with appropriate record number, save export tab as csv, and reopen Template
    Application.DisplayAlerts = False
    'Save Current workbook with record number
        Worksheets("Combo").Activate
        ActiveWorkbook.SaveAs Filename:=path & record & "_Final.xlsx", FileFormat:=xlOpenXMLWorkbook
    'Activate Export tab and save as CSV
        Worksheets("Export").Activate
        ActiveWorkbook.SaveAs Filename:=path & record & ".csv", FileFormat:=xlCSV
    'Reopen blank R000XX_Final_Template
        ActiveWorkbook.Close False
        Workbooks.Open Filename:=basepath & "R000XX_Final_Template.xlsx"
    Application.DisplayAlerts = True
End Sub
