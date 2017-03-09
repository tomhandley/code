Attribute VB_Name = "process_sonar"
Option Explicit

Dim basepath As String 'path to R000XX_Final_Template
Dim path As String 'path to record
Dim record As String 'record name
Dim fileprompt As VbMsgBoxResult
Dim mybook As Workbook '000XX_Final_Template workbook path
Dim row As Long 'row counter
Dim seconds As Long 'counts number of full seconds in sonar file
Dim DE_file As String 'file path to data_explorer index file
Dim rtk_file As String 'file path to rtk data file
Const tlen = 0.114 'transducer length from center of mounting pole to sonar projector
Const max_z = 0.03 'maximum allowable rtk precision in z; 3 cm based on equipment limitations
Const max_xy = 0.15 'maximum allowable rtk precision in xy; based on limiting overall uncertainty
'Set worksheet columns for input and output
'Const sonar_time = 1, sonar_east = 2, sonar_north = 3, sonar_depth = 4, sonar_cog = 5
'Const rtk_baseid = 1, rtk_pointid = 2, rtk_time = 3, rtk_north = 4, rtk_east = 5, rtk_elev = 6, rtk_horiz_prec = 7, rtk_vert_prec = 8, rtk_soln_type = 17
'Const combo_index = 1, combo_son_time = 2, combo_row = 3, combo_x = 4, combo_y = 5, combo_cog = 6, combo_depth = 7, combo_xoffs = 8, combo_yoffs = 9
'Const combo_pointid = 11, combo_rtk_time = 12, combo_north = 13, combo_east = 14, combo_cog_RTK = 15
'Const combo_elev = 16, combo_smooth = 17, combo_min = 18, combo_max = 19, combo_stdev_xy = 20, combo_stdev_z = 21, combo_soln_type = 22
'Const export_north = 1, export_east = 2, export_elev = 3, export_time = 4, export_sonarid = 5, export_rtkid = 6
Private Type linear  'Data type to store linear regression values
    b As Double 'slope of regression
    a As Double 'intercept of regression
    n As Integer 'number of points in regression
End Type

Sub assembleXYZ()
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
' All functions were defined for the following worksheet columns:
' Sonar: (1)SonarTime, (2)Easting, (3)Northing, (4)Depth, (5)COG
' RTK: (1)Base_ID, (2)Point_ID, (3)Start Time, (4Northing, (5)Easting, (6)Elevation, (7)Horizontal Precision, (8)Vertical Precision, (9)Std Dev n, (10)Std Dev e, (11)Std Dev u, (12)Std Dev Hz, (13)Geoid Separation, (14)dN, (15)dE, (16)dHt, (17)Solution Type
' Combo: (1)Index, (2)Sonar Time, (3)Row, (4)X (Easting), (5)Y (Northing), (6)COG_sonar, (7)Depth, (8)X-offset, (9)Y-offset, (10)BLANK, (11)Point ID, (12)RTK Time, (13)Northing, (14)Easting, (15)COG_RTK, (16)Elev, (17)Smooth, (18)Min, (19)Max, (20)StDev_XY, (21)StDev_Z, (22)Solution Type
' Export: (1)Northing, (2)Easting, (3)Bed_elev, (4)DateTime, (5)Sonar_ID, (6)RTK_ID

Dim IsFile As Boolean
Dim ws As Integer
Dim rng As Range

'Read defaults from text file
    Set mybook = ActiveWorkbook
    path = mybook.FullName
    path = Left(path, InStrRev(path, "\")) 'remove filename from path string
    basepath = path
    IO_defaults "read"

'Import data to sonar worksheet
    combo.Range("A1").Activate
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
            ChDir path 'change default directory to last used
            path = Application.GetOpenFilename("Humminbird dat files (*.dat),*.dat", , "Navigate to the sonar .dat file to process")
            If path = "False" Then Exit Sub
            record = Left(Right(path, Len(path) - InStrRev(path, "\")), 6) 'save record from root path
            path = Left(path, InStrRev(path, "\")) 'cut filename from root path
        End If
    End If
    'check whether R000XX_Final.xlsx already exists
    IsFile = False
    On Error Resume Next
    IsFile = GetAttr(path & record & "_Final.xlsx")
    If IsFile Then
        fileprompt = MsgBox(record & "_Final.xlsx already exists! Proceed with processing?", vbOKCancel, "Warning")
        If fileprompt = vbCancel Then Exit Sub
    End If
    
    'check whether R000XX.DAT.XYX.csv exists (processed SonarTRX file)
    IsFile = False
    On Error Resume Next
    IsFile = GetAttr(path & record & ".DAT.XYZ.csv")
    If Not IsFile Then
        fileprompt = MsgBox(record & ".DAT.XYZ.csv sonar file not found! Processing cancelled.", vbExclamation, "File not found!")
        Exit Sub
    End If
    
'Clear anything past header row on first four worksheets
    For ws = 1 To 4
        Set rng = Sheets(ws).UsedRange
        Set rng = rng.Offset(1, 0).Resize(rng.Rows.count - 1)
        rng.ClearContents
        rng.Interior.ColorIndex = 0
    Next ws

'Import sonar ping time stamps from IDX file
    sonar_import
    If fileprompt = vbCancel Then Exit Sub
    
'Divide sonar data by full seconds and write to combo sheet
    sonar_to_combo

'Import pertinent rtk data to rtk worksheet and insert missing data lines
    rtk_import Cells(2, 2), Cells(seconds + 1, 2)
    If fileprompt = vbCancel Then Exit Sub

'Extract rtk data to combo sheet
    rtk_to_combo
    
'Interpolate gaps in rtk data
    interpolate_rtk

'Smooth rtk elevation points
    critical

'Update Plotting ranges
    Sheets("TrackPlot").Select
    ActiveChart.FullSeriesCollection(1).XValues = "=Combo!$D$2:$D$" & seconds + 1 'sonar x (easting)
    ActiveChart.FullSeriesCollection(1).Values = "=Combo!$E$2:$E$" & seconds + 1 'sonar y (northing)
    ActiveChart.FullSeriesCollection(2).XValues = "=Combo!$N$2:$N$" & seconds + 1 'rtk x
    ActiveChart.FullSeriesCollection(2).Values = "=Combo!$M$2:$M$" & seconds + 1 'rtk y
    ActiveChart.ChartTitle.Text = record & " Navigation Tracks"
    Sheets("Smoothing").Select
    ActiveChart.FullSeriesCollection(1).Values = "=Combo!$Q$2:$Q$" & seconds + 1 'smoothed elev
    ActiveChart.FullSeriesCollection(2).Values = "=Combo!$R$2:$R$" & seconds + 1 'min elev
    ActiveChart.FullSeriesCollection(3).Values = "=Combo!$S$2:$S$" & seconds + 1 'max elev
    ActiveChart.FullSeriesCollection(4).Values = "=Combo!$P$2:$P$" & seconds + 1 'raw elev
    Sheets("Depth").Select
    ActiveChart.FullSeriesCollection(1).Values = "=Combo!$G$2:$G$" & seconds + 1 'Depth
    ActiveChart.FullSeriesCollection(2).Values = "=Export!$C$2:$C$" & seconds + 1 'Bed_elev

'Extract data to export sheet
    expo.Activate
    'write excel formulas so changes to combo are dynamically adjusted
    For row = 1 To seconds
        Cells(row + 1, 1).FormulaR1C1 = "=Combo!RC[12]" 'northing
        Cells(row + 1, 2).FormulaR1C1 = "=Combo!RC[12]" 'easting
        Cells(row + 1, 3).FormulaR1C1 = "=Combo!RC[14]-Combo!RC[4]" 'bottom = smoothed_elevation - depth
        Cells(row + 1, 4).FormulaR1C1 = "=Combo!RC[-2]" 'datetime
        Cells(row + 1, 5).Value = record 'sonar id
        Cells(row + 1, 6).FormulaR1C1 = "=IF(Combo!RC[5]="""",""N/A"",Combo!RC[5])" 'rtk id
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
        save_files
        fileprompt = MsgBox("Record " & record & " successfully processed and saved. View processed file?", vbYesNo)
        If fileprompt = vbYes Then
            Workbooks.Open (path & record & "_Final.xlsx")
        End If
    Else
        MsgBox ("Record " & record & " processed but not saved or exported")
    End If
    IO_defaults "write" 'write defaults to text file

End Sub

Private Sub IO_defaults(Optional IOtype As String = "read")
'read/write last used record, DE_file, and rtk_file to/from R000XX_defaults.txt
Dim f As Integer 'file index number
    f = FreeFile
    ChDir (path) 'set default path
    If IOtype = "read" Then
        On Error Resume Next
        Open "R000XX_defaults.txt" For Input As #f
        If Err.Number <> 0 Then
            record = "na"
            DE_file = "na"
            rtk_file = "na"
            Exit Sub
        End If
        On Error GoTo 0 'reset error handling
        Input #f, record
        Input #f, path
        Input #f, DE_file
        Input #f, rtk_file
        Close #f
    Else 'overwrite defaults file
        ChDir (basepath) 'save in initial location
        Open "R000XX_defaults.txt" For Output As #f
        Write #f, record
        Write #f, path
        Write #f, DE_file
        Write #f, rtk_file
        Close #f
    End If
End Sub

Private Sub sonar_import()
'Import sonar ping time stamps from IDX file
'   Humminbird sonar files consist of a DAT file, and an IDX and SON file for each channel:
'       DAT file -- holds record info for starting date and time, duration, and beginning coordinates
'       IDX files -- two 4-byte fields per record specifying a time increment and a line index in the SON file
'       SON files -- complex binary file with full navigation data and imagery for each record
Dim f As Integer 'file index number
Dim fLen As Long 'length of IDX file in bytes
Dim idx_data() As Byte 'holds binary data from IDX file
Dim DE_book As Workbook 'data explorer workbook
Dim dat_book As Workbook 'R000XX.DAT.XYZ.csv path
Dim fulldate As Double 'initial datetime value (days since 1900 + hr/24 + min/60 + sec/3600 + ms/3600/1000)
Dim utc_shift As Integer 'value to shift sonar times forward/back to match rtk times (UTC)
    'for Pacific time, use 8 for records during PST (winter Nov-Mar) and 7 during PDT (summer Mar-Nov)
    'based on a sonar file stated in local time and rtk file stated in UTC time
Dim cog As Double 'course over ground

'Load binary data from IDX file
    f = FreeFile
    Open path & record & "\B002.IDX" For Binary Access Read As #f
    fLen = LOF(f)
    ReDim idx_data(1 To fLen)
    Get f, , idx_data
    Close f

'Choose data explorer file (DE_file) path
    fileprompt = MsgBox("Use default Data Explorer file path? (" & DE_file & ")", vbYesNoCancel, "Data Explorer source")
    If fileprompt = vbCancel Then
        Exit Sub
    Else
        If fileprompt = vbYes Then
            If DE_file = "na" Then
                MsgBox "Defaults file not found!" & vbCrLf & "Open file from list.", vbCritical, "File read error!"
            Else
                Application.ScreenUpdating = False
                On Error Resume Next 'check that file opens without errors
                Set DE_book = Workbooks.Open(Filename:=DE_file, ReadOnly:=True)
                If Err.Number <> 0 Then
                    MsgBox "Default data explorer file not found!", vbCritical, "Invalid file path!"
                    DE_file = "na"
                End If
                On Error GoTo 0 'reset error handling
            End If
        Else
            DE_file = "na"
        End If
    End If
    If DE_file = "na" Then
    'select DE_file from explorer
        ChDir (path) 'set default path
        DE_file = Application.GetOpenFilename("Excel Workbooks (*.xls;*.xlsx),*.xls;*.xlsx", , "Select data explorer file")
        Application.ScreenUpdating = False
        Set DE_book = Workbooks.Open(Filename:=DE_file, ReadOnly:=True)
    End If
   
'Read data and close file
    row = WorksheetFunction.Match(Val(Right(record, 5)), DE_book.Worksheets(1).Range("D2:D400"), 0) + 1 'find record number in DE_book
    fulldate = DE_book.Worksheets(1).Cells(row, 1) 'initial date from DE_book
    If Month(fulldate) > 2 And Month(fulldate) < 11 Then
        utc_shift = 7
    Else
        utc_shift = 8
    End If
    fileprompt = MsgBox("Accept time offset of UTC -" & utc_shift & "?", vbYesNo, "UTC offset")
    If fileprompt = vbNo Then
        MsgBox "Set UTC time offset for date of data collection (" & Month(fulldate) & "/" & Day(fulldate) & "/" & Year(fulldate) & "). Set the offset to 7 for Pacific Standard Time (winter months) or 8 for Pacific Daylight Time (summer months).", vbCritical, "Set UTC offset"
    End If
    fulldate = fulldate + DE_book.Worksheets(1).Cells(row, 6) + utc_shift / 24  'Shift times forward or back for UTC correction
    DE_book.Close
    Set dat_book = Workbooks.Open(Filename:=path & record & ".DAT.XYZ.csv", ReadOnly:=True)
    mybook.Worksheets("Sonar").Activate
    Range("A2").Select
    For row = 0 To fLen \ 8 - 1 'backslash operator is integer division
        'first 4 bytes of each 8-byte record hold time stamp info
        Cells(row + 2, 1) = fulldate + (idx_data(row * 8 + 2) * 65536 + idx_data(row * 8 + 3) * CLng(256) + idx_data(row * 8 + 4)) / 24 / 3600 / 1000
    Next row
    dat_book.Worksheets(1).Range("A2:C" & row + 1).Copy Destination:=Range("B2") 'copy xyz data from dat_book to sonar sheet
    dat_book.Close
    Application.ScreenUpdating = True

'Calculate course over ground (cog)
    For row = 1 To fLen \ 8 - 1
        Cells(row + 1, 5) = get_cog(Cells(row + 1, 2), Cells(row + 2, 2), Cells(row + 1, 3), Cells(row + 2, 3))
    Next row

End Sub
 
Private Sub sonar_to_combo()
'Divide sonar data by full seconds and write to combo sheet
Dim srow As Long 'cumulative number rows <= time being evaluated
Dim prev As Long 'cumulative rows <= previous time increment
Dim t1 As Double
Dim t2 As Double 'times in milliseconds
Dim i As Long
Dim j As Long
Dim avg_x As Double 'average x position for each full second
Dim avg_y As Double 'average y position for each full second
Dim avg_depth As Single 'average depth value for each full second
Dim r As linear 'regression values
    
'Remove records at end with no change in position (common record artifact)
    If Round(Cells(row + 1, 2), 6) = Round(Cells(row, 2), 6) Or Round(Cells(row + 1, 3), 6) = Round(Cells(row, 3), 6) Then
    'milliseconds aren't within one ping of 1000 so remove partial second
        Do While Round(Cells(row + 1, 2), 6) = Round(Cells(row, 2), 6) Or Round(Cells(row + 1, 3), 6) = Round(Cells(row, 3), 6)
        'seconds digits match
            row = row - 1
        Loop
    End If

'Remove partial seconds at end of record count (row)
    Do While Int(Cells(row + 1, 1) * 24 * 3600) = Int(Cells(row, 1) * 24 * 3600)
    'seconds digits match
        row = row - 1
    Loop

'Average sonar records for x, y, depth, and calculated cog over full-second intervals
    seconds = timer(Cells(row, 1), Cells(2, 1)) 'number of full seconds in data record
    combo.Activate
    prev = 0
    srow = 0
    t1 = sonar.Cells(srow + 2, 1)
    For i = 1 To seconds
        Cells(i + 1, 1) = i 'index
        If i = 1 Then
            'set first value to start time
            Cells(2, 2) = sonar.Range("A2")
        Else
            'increment sonar time by 1 second
            Cells(i + 1, 2) = Cells(i, 2) + 1 / 24 / 3600
        End If
        'count number of rows in sonar for each second of data
        t2 = Cells(i + 1, 2) + 1 / 24 / 3600 'set t2 to time plus one second
        Do While (t1 < t2 - 0.000000005) And (sonar.Cells(srow + 2, 1) <> "")
            '0.000000005 needed to correct for rounding errors
            srow = srow + 1
            t1 = sonar.Cells(srow + 2, 1) 'set t1 to ping time
        Loop
        Cells(i + 1, 3) = srow 'row (starting row of each full-second record)
        avg_x = 0
        avg_y = 0
        'correct raw positions from sonar sheet for the transducer sonar projector offset distance
        'x = x1 + sin(radians(cog)) * tlen
        'y = y1 - cos(radians(cog)) * tlen
        For j = prev + 1 To srow
            avg_x = avg_x + sonar.Cells(j + 1, 2) + Sin(WorksheetFunction.Pi / 180 * sonar.Cells(j + 1, 5)) * tlen
            avg_y = avg_y + sonar.Cells(j + 1, 3) - Cos(WorksheetFunction.Pi / 180 * sonar.Cells(j + 1, 5)) * tlen
        Next j
        'write averaged x, y, cog, and depth
        Cells(i + 1, 4) = avg_x / (srow - prev) 'x
        Cells(i + 1, 5) = avg_y / (srow - prev) 'y
        Cells(i + 1, 6) = WorksheetFunction.Sum(sonar.Range("E" & prev + 2 & ":E" & srow + 1)) / (srow - prev) 'cog
        avg_depth = WorksheetFunction.Sum(sonar.Range("D" & prev + 2 & ":D" & srow + 1)) / (srow - prev) 'depth
        'if depth was recorded as bottom elevation from SonarTRX, change sign to positive
        If avg_depth < 0 Then avg_depth = -1 * avg_depth
        Cells(i + 1, 7) = avg_depth
        prev = srow
    Next i
    'in case the final cog is blank (when the record ends with an exact full second), extrapolate cog with slope
    If Cells(i, 6) = "" Then
        r = regress(i - 1, 6, 5, , , row)
        Cells(i, 6) = r.a + r.b * (1 + r.n)
    End If

End Sub

Private Function regress(yrow As Long, ycol As Integer, n As Integer, Optional direction As Integer = 1, _
  Optional check_col As Integer = 0, Optional maxrow As Long = 0) As linear
'Calculate slope of the regression line through n points in column ycol starting with Cells(yrow, ycol)
'  where values in ycol are assumed dependent on their row index
'the direction defaults to base extrapolation on points from yrow and up, providing parameters to fill blanks below
'pass direction = -1 to base extrapolation on points from yrow and down, providing parameters to fill blanks above
'check_col will force consideration only of points with valid values in column check_col if it is defined (> 0)
'maxrow is the last row from which data should be considered. It defaults to global variable seconds
'returns data of type linear with intercept .a, slope .b, and ending independent variable index .n
'
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
'Define type linear prior to function call:
'    Private Type linear  'Data type to store linear regression values
'        b As Double 'slope of regression
'        a As Double 'intercept of regression
'        n As Integer 'number of points in regression
'    End Type
'
Dim count As Integer
Dim x As Integer 'index value
Dim i As Integer
Dim y() As Double 'values for regression
Dim xval() As Double 'xval holds indices
Dim ybar As Double 'average values for y
Dim xbar As Double 'average values for x
Dim get_data As Boolean
ReDim y(1 To n)
ReDim xval(1 To n)

'Check for valid direction
    If Abs(direction) <> 1 Then
    'assume invariable
        With regress
            .b = 0
            .a = Cells(yrow, ycol)
            .n = 1
        End With
        MsgBox "Invalid direction passed to function regress. Direction must be passed as 1 " & _
            "for slope from preceeding points, or -1 for slope from proceeding points. Slope set to zero.", vbOKOnly
        Exit Function
    End If

'Initialize
    x = 0
    count = 0
    ybar = 0
    xbar = 0
    If maxrow = 0 Then maxrow = seconds + 1 'set to default

'Gather data for regression
    get_data = True
    Do Until count = n Or yrow + x > maxrow Or yrow + x < 2
        If check_col > 0 Then
            If Cells(yrow + x, check_col) <> "" And Cells(yrow + x, check_col).Interior.Color <> 49407 Then
            'interpolated cells are color 49407
                get_data = True
            Else: get_data = False
            End If
        End If
        If get_data Then
            count = count + 1
            y(count) = Cells(yrow + x, ycol)
            ybar = ybar + y(count)
            xval(count) = Abs(x) + 1
        End If
        x = x - direction
    Loop

'Calculate averages
    If check_col = 0 Then
        xbar = (count + 1) / 2 'average of 1..n
    Else
        For i = 1 To count
            xbar = xbar + Abs(x) - xval(i) + 1
        Next i
        xbar = xbar / count
    End If
    ybar = ybar / count

'Calculate results
    With regress
        .n = Abs(x) 'index of independent variable x
        For i = 1 To count
            .b = .b + (.n - xval(i) + 1 - xbar) * (y(i) - ybar) 'sum numerator
            .a = .a + (.n - xval(i) + 1 - xbar) ^ 2 'sum denominator
        Next i
        'set final values
        .b = .b / .a 'slope
        .a = ybar - .b * xbar 'intercept
    End With

End Function

Private Sub rtk_import(date1 As Date, date2 As Date)
'Import rtk data to rtk worksheet, delete duplicate entries, and insert missing rows
Dim rng As Range
Dim rtk_book As Workbook
Dim rtk_start, rtk_end As Long 'starting and ending row numbers in rtk_book to copy over
Dim dt As Long

'Choose rtk file path
    fileprompt = MsgBox("Use default RTK data file path? (" & rtk_file & ")", vbYesNoCancel, "RTK data source")
    If fileprompt = vbCancel Then
        Exit Sub
    Else
        If fileprompt = vbYes Then
            If rtk_file = "na" Then
                MsgBox "Defaults file not found!" & vbCrLf & "Open file from list.", vbCritical, "File read error!"
            Else
                Application.ScreenUpdating = False
                On Error Resume Next 'check that file opens without errors
                Set rtk_book = Workbooks.Open(Filename:=rtk_file, ReadOnly:=True)
                If Err.Number <> 0 Then
                    MsgBox "Default RTK data file not found!", vbCritical, "Invalid file path!"
                    rtk_file = "na"
                    Application.ScreenUpdating = True
                End If
                On Error GoTo 0 'reset error handling
            End If
        Else
            rtk_file = "na"
        End If
    End If
    If rtk_file = "na" Then
    'select rtk_file from explorer
        ChDir (path) 'set default path
        rtk_file = Application.GetOpenFilename("Excel Workbooks (*.xls;*.xlsx),*.xls;*.xlsx", , "Select RTK data file")
        Application.ScreenUpdating = False
        Set rtk_book = Workbooks.Open(Filename:=rtk_file, ReadOnly:=True)
    End If

'Import rtk data and close file
    rtk_book.Worksheets(1).Select
    rtk_start = find_row(date1, "initial")
    rtk_end = find_row(date2, "final")
    'copy data from rtk_book
    mybook.Worksheets("RTK").Range("A2:Q" & 2 + rtk_end - rtk_start).Value = rtk_book.Worksheets(1).Range("A" & rtk_start & ":Q" & rtk_end).Value
    Application.DisplayAlerts = False 'prevent "Save changes" dialog
    rtk_book.Close
    Application.DisplayAlerts = True
    mybook.Worksheets("RTK").Activate
    Application.ScreenUpdating = True

End Sub

Private Sub rtk_to_combo()
'Write rtk data to Combo worksheet
'assumes rtk data are already adjusted to xyz-position at base of pipe and level of sonar emitter
Dim combo_row As Long
Dim rtk_row As Long
Dim dt As Long 'number of seconds elapsed between two records
Dim cog As Double 'course over ground
Dim count As Long
Dim avg As Double
Dim r As linear 'store regression values
Dim rtk_x As Double
Dim rtk_y As Double
Dim increment As Long

combo.Activate
'Import data from rtk sheet, averaging records and leaving gaps as needed
    combo_row = 2
    rtk_row = 2
    Do Until timer(rtk.Cells(rtk_row, 3), Cells(combo_row, 2)) >= 0 'increment rtk_row until rtk time >= sonar time
        rtk_row = rtk_row + 1
    Loop
    Do While combo_row <= seconds + 1
        dt = timer(rtk.Cells(rtk_row, 3), Cells(combo_row, 2))
        If dt = 0 Then 'rtk and sonar times match
            If rtk.Cells(rtk_row + 1, 4) = rtk.Cells(rtk_row, 4) And rtk.Cells(rtk_row + 1, 5) = rtk.Cells(rtk_row, 5) Then 'no change in position
                rtk_row = rtk_row + 1 'skip over duplicate position, occasionally happens after gap in rtk record
            Else 'average one or more rtk records with matching time and write data to Combo sheet
                Cells(combo_row, 12) = rtk.Cells(rtk_row, 3) 'rtk time
                count = 1
                Do While timer(rtk.Cells(rtk_row + count, 3), rtk.Cells(rtk_row, 3)) = 0
                    count = count + 1
                Loop
                cog = 9999 'initialize to invalid value to test for changes
                avg = WorksheetFunction.Average(rtk.Cells(rtk_row, 7).Resize(count)) 'stdev_xy
                If avg <= max_xy Then 'horizontal precision OK
                    Cells(combo_row, 11) = rtk.Cells(rtk_row, 2) 'point id
                    Cells(combo_row, 20) = avg 'stdev_xy
'XXXX this calculates cog between averaged positions, could split up to partial second positions for higher accuracy
'XXXX update to:
'XXXX 1. load cog for each pair in an array
'XXXX 2. correct x and y by cog in an array
'XXXX 3. average the updated x and y values and output
                    'calculate cog_rtk, northing and easting
                    If rtk.Cells(rtk_row + 1, 7) <= max_xy And timer(rtk.Cells(rtk_row + count + 1, 3), rtk.Cells(rtk_row, 3)) <= 5 Then
                    'next horizontal precision is OK and within 5 seconds: calculate cog between two valid points
                        cog = get_cog(rtk.Cells(rtk_row, 5), rtk.Cells(rtk_row + 1, 5), rtk.Cells(rtk_row, 4), rtk.Cells(rtk_row + 1, 4))
                    Else 'no second point: calculate cog from slope of valid points above
                        If combo_row > 2 And Cells(combo_row - 1, 15) <> "" Then 'previous point has valid cog
                            r = regress(combo_row - 1, 15, 5, , 15) 'regression through last five valid cogs
                            cog = r.a + r.b * (1 + r.n) 'y = a + bx
                        End If
                    End If
                    If cog <> 9999 Then 'check if cog was calculated
                        Cells(combo_row, 15) = cog 'cog_rtk
                        'use xy-position and cog to offset transducer
                        rtk_x = WorksheetFunction.Average(rtk.Cells(rtk_row, 5).Resize(count)) + Sin(WorksheetFunction.Pi / 180 * cog) * tlen
                        rtk_y = WorksheetFunction.Average(rtk.Cells(rtk_row, 4).Resize(count)) - Cos(WorksheetFunction.Pi / 180 * cog) * tlen
                        Cells(combo_row, 8) = rtk_x - Cells(combo_row, 4) 'x-offset = rtk_x - sonar_x
                        Cells(combo_row, 9) = rtk_y - Cells(combo_row, 5) 'y-offset = rtk_y - sonar_y
                        Cells(combo_row, 14) = rtk_x 'easting
                        Cells(combo_row, 13) = rtk_y 'northing
                    End If
                    avg = WorksheetFunction.Average(rtk.Cells(rtk_row, 8).Resize(count))
                    If avg <= max_z Then
                    'both horizontal and vertical precision OK
                        Cells(combo_row, 16) = WorksheetFunction.Average(rtk.Cells(rtk_row, 6).Resize(count)) 'elevation
                        Cells(combo_row, 21) = WorksheetFunction.Average(rtk.Cells(rtk_row, 8).Resize(count)) 'stdev_z
                        Cells(combo_row, 22) = rtk.Cells(rtk_row, 17) 'solution type
                    Else
                    'only horizontal precision OK
                        Cells(combo_row, 21) = max_z 'stdev_z set to max for later smoothing
                        Cells(combo_row, 22) = "Float,Horizontal" 'solution type
                    End If
                Else
                'neither horizontal nor vertial precision OK
                    Cells(combo_row, 20) = "> " & Format(Str(max_xy), "0.00") 'stdev_xy
                    Cells(combo_row, 21) = max_z 'stdev_z set to max for later smoothing
                    Cells(combo_row, 22) = "None,Interpolated"  'solution type
                End If
                rtk_row = rtk_row + count
            End If
            count = 1 'use for incrementing combo_row
        Else 'dt > 0 --> gap in rtk record, fill with time and basic "no data" info for later interpolation
            count = WorksheetFunction.Min(dt, seconds + 2 - combo_row) 'limit data to seconds + 1 row
            Cells(combo_row, 2).Resize(count).Copy Cells(combo_row, 12)
'XXX LEAVE stdev_XX and soln type BLANK FOR INTERPOLATION
'XXX            Cells(combo_row, 20).Resize(count) = "> " & Format(Str(max_xy), "0.00") 'stdev_xy
'XXX            Cells(combo_row, 21).Resize(count) = max_z 'stdev_z set to max for later smoothing
'XXX            Cells(combo_row, 22).Resize(count) = "None,Interpolated" 'solution type
        End If
        Cells(combo_row, 18).Resize(count).FormulaR1C1 = "=RC[-2] - RC[3]"  'min = elev - stdev_z
        Cells(combo_row, 19).Resize(count).FormulaR1C1 = "=RC[-3] + RC[2]"  'max = elev + stdev_z
        combo_row = combo_row + count
    Loop
End Sub

Private Function get_cog(x1, x2, y1, y2) As Double
'Calculate course over ground given two points
'cog = mod(degrees(atan2(x2 - x1, y2 - y1)) + 270, 360)
Dim cog As Double
    On Error Resume Next 'no change in position results in division by zero error (Err.Number = 11)
    cog = WorksheetFunction.Atan2(x2 - x1, y2 - y1)
    If Err.Number <> 0 Then cog = 0 'set cog to 0 for static position
    On Error GoTo 0 'reset error handling
    cog = cog * 180 / WorksheetFunction.Pi + 270 'switch to coordinates in zero degrees North
    If cog >= 360 Then cog = cog - 360 'force 0 <= cog < 360
    get_cog = cog
End Function

Private Function find_row(dateval As Date, matchtype As String) As Long
'Return row number from rtk data file that matches the input date and time
'rtk data file must be active
'matchtype should be passed as "initial" or "final"
Dim step, index As Long
Dim last_row As Long
Dim dt As Integer
    last_row = ActiveSheet.Cells(Rows.count, "A").End(xlUp).row - 1
    index = Int(last_row / 2) 'start at midpoint
    step = index 'initial step will become half the size of index in loop below
    Do Until Abs(dateval - Cells(index + 1, 3)) < 1 / 48 / 3600 Or Abs(step) = 1
        step = Abs(step) 'reset step to positive after each iteration
        If dateval < Cells(index + 1, 3) Then
            step = -Int((step + 1) / 2) 'step = -roundup(step/2)
        Else
            step = Int((step + 1) / 2) 'step = roundup(step/2)
        End If
        index = index + step
    Loop
    If Abs(step) = 1 Then
        'no exact time match, so move to one row below low value or one above high value
        If Abs(dateval - Cells(index + 1, 3)) * 24 * 3600 > 0.5 Then
            If matchtype = "initial" Then
                If dateval < Cells(index + 1, 3) Then
                    index = index - 1
                End If
            Else
                If dateval > Cells(index + 1, 3) Then
                    index = index + 1
                End If
            End If
            If Abs(dateval - Cells(index + 1, 3)) <= 1 / 24 Then 'less than one hour gap
                dt = Round(Abs(dateval - Cells(index + 1, 3)) * 24 * 3600, 0)
                MsgBox ("No matching time for " & matchtype & " bound (" & dateval & ") in RTK file " & Chr(151) & _
                    " the next closest time will be used (" & dt & " second difference)")
            Else
                MsgBox ("No matching time for " & matchtype & " bound (" & dateval & ") in RTK file " & Chr(151) & _
                    " " & matchtype & " values will be extrapolated.")
                'ignore data more than an hour away from target
                If matchtype = "initial" Then
                    index = index + 1
                Else: index = index - 1
                End If
            End If
        End If
    Else 'check for repeating values of same second (can occur when rtk collects data @ 1/m)
        If matchtype = "initial" Then
            step = -1
        Else: step = 1
        End If
        Do While Abs(dateval - Cells(index + step + 1, 3)) < 1 / 48 / 3600
            index = index + step
        Loop
    End If
    find_row = index + 1 'set value to actual worksheet row number
End Function

Private Sub interpolate_rtk()
'Interpolate missing rtk xy-position based on linear offset from sonar xy-position,
'Interpolate missing rtk z linearly
'xy-positions and offsets were previously corrected for course over ground,
'  so interpolated positions do not need to be corrected again
Dim x As Integer
Dim warn_id As Integer
Dim warn() As Boolean 'record of interpolated and extrapolated areas for message output
Dim warning As String 'warning message for output
ReDim warn(1 To 7)
Const offset_damping = 7 'number of seconds to dampen x-, y- and z-offset trends on log scale before they reach zero
Const max_damping = 5 'number of seconds to transition from interpolated values of stdev_xy and stdev_z before reaching the maximum

'Dim i As Long
'Dim count As Long
'Dim dir As Integer 'direction of interpolation or extrapolation
'Dim target As Long 'start row for interpolation or extrapolation
'Dim xr As linear 'x-offset regression
'Dim yr As linear 'y-offset regression
''XXXDim xmean As Double 'mean of adjacent x-offsets in window size
''XXXDim ymean As Double 'mean of adjacent y-offsets in window size
'Dim j As Long
'Dim damping As Single 'used for logarithmic damping of slope extrapolation
'Const window = 100 'number of seconds to average offsets forward and backward to find mean

''Interpolate northing and easting
'i = 2
'Do While i <= seconds + 1
'    If Cells(i, 8) = "" Then 'x-offset is blank
'        'count blanks
'        count = 1
'        Do While Cells(i + count, 8) = "" And i + count <= seconds + 1
'            count = count + 1
'        Loop
'        If count = seconds Then 'no rtk position for whole record
'            Cells(2, 4).Resize(count).Copy Cells(2, 14) 'copy x from sonar
'            Cells(2, 5).Resize(count).Copy Cells(2, 13) 'copy y from sonar
'            warn(1) = True
'        ElseIf i = 2 Or i + count - 1 = seconds + 1 Then 'initial or terminal gap in rtk
'            'find regression slope of x- and y-offsets from known points and produce a smooth curve towards the average
'            If i = 2 Then 'initial gap
'                dir = -1
'                target = count + 2
'                warn(2) = True
'            Else 'terminal gap
'                dir = 1
'                target = i - 1
'                warn(3) = True
'            End If
'            xr = regress(target, 8, 6, dir, 11) 'regression of x-offset through 6 points with valid Point_ID
'            yr = regress(target, 9, 6, dir, 11) 'regression of y-offset through 6 points with valid Point_ID
'XXX            xmean = midmean(window, target, 8) 'average x-offset within +/- <window> values
'XXX            ymean = midmean(window, target, 9) 'average y-offset within +/- <window> values
'            For j = 1 To count
'            'calculate x- and y-offsets
'                If j < offset_damping + 2 Then
'                    damping = WorksheetFunction.Log(offset_damping - j + 2, offset_damping + 1)
'                    Cells(target + j * dir, 8) = Cells(target + (j - 1) * dir, 8) + xr.b * damping
'                    Cells(target + j * dir, 9) = Cells(target + (j - 1) * dir, 9) + yr.b * damping
'                Else
'                    damping = 0
'XXX                    If j <= offset_damping * 2 Then damping = WorksheetFunction.Log(j - offset_damping, offset_damping + 1) Else damping = 1
'XXX                    Cells(target + j * dir, 8) = Cells(target + (j - 1) * dir, 8) + (xmean - Cells(target + (j - 1) * dir, 8)) / offset_damping * damping
'XXX                    Cells(target + j * dir, 9) = Cells(target + (j - 1) * dir, 9) + (ymean - Cells(target + (j - 1) * dir, 9)) / offset_damping * damping
'                End If
'            Next j
'        Else 'middle gap--interpolate between two valid values
'            xr.b = (Cells(i + count, 8) - Cells(i - 1, 8)) / (count + 1)
'            yr.b = (Cells(i + count, 9) - Cells(i - 1, 9)) / (count + 1)
'            For j = 0 To count - 1
'                Cells(i + j, 8) = Cells(i + j - 1, 8) + xr.b 'x-offset
'                Cells(i + j, 9) = Cells(i + j - 1, 9) + yr.b 'y-offset
'            Next j
'        End If
'        'use formula for northing and easting in case x-offset and y-offset are changed manually
'        Cells(i, 13).Resize(count).FormulaR1C1 = "=RC[-8] + RC[-4]" 'northing
'        Cells(i, 14).Resize(count).FormulaR1C1 = "=RC[-10] + RC[-6]" 'easting
'        'highlight records with interpolated data in orange
'        Cells(i, 8).Resize(count, 2).Interior.Color = 49407
'        Cells(i, 13).Resize(count, 2).Interior.Color = 49407
'        i = i + count + 1
'    Else
'        i = i + 1
'    End If
'Loop

'Interpolate northing and easting
x = interp(8, offset_damping + 3, , 0, 20) 'interpolate x-offset and easting
x = interp(9, offset_damping + 3, , 0, 20) 'interpolate y-offset and northing

'Interpolate missing rtk data
warn_id = interp(16, max_damping, 6) 'interpolate elevation
If warn_id = 0 Or warn_id > 2 Then
    If warn_id = 5 Then
        warn(6) = True
        warn(7) = True
    ElseIf warn_id > 0 Then
        warn(warn_id + 3) = True
    End If
    x = interp(20, max_damping, 7, max_xy) 'interpolate stdev_xy
    x = interp(21, max_damping, 8, max_z) 'interpolate stdev_z
ElseIf warn_id = 1 Then
    Cells(2, 20).Resize(seconds) = ">" & max_xy
    Cells(2, 21).Resize(seconds) = max_z
    Cells(2, 22).Resize(seconds) = "None,interpolated"
    Cells(2, 21).Resize(seconds, 2).Interior.Color = 49407
Else 'warn_id = 2
    Cells(2, 22).Resize(seconds) = "None"
End If

'Output warnings
    x = interp(20, max_damping, 7, max_xy) 'interpolate stdev_xy
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

Private Function interp(icol As Integer, damping_interval As Integer, Optional rtk_col As Integer = 0, Optional limit_val As Single = -1, Optional window As Integer = 50) As Integer
'Interpolate gaps in input column
'Outputs an identifier depending on which data are missing
'   0: all data present
'   1: no data for selected column, interpolated between outside points from rtk sheet
'   2: no data for whole record, no interpolation performed
'   3: initial gap
'   4: terminal gap
'   5: both initial and terminal gaps
'icol is the column on combo sheet to interpolate gaps in
'damping_interval is the number of seconds to use log-dampened weighting of interpolated value before it becomes constant at limit_val
'   damping_interval * 2 is also used as the maximum time in seconds to interpolate from rtk records beyond the ends instead of using slope trends
'rtk_col is the column with matching data in the rtk sheet
'   if none is specified, extrapolation will be based on regression of slope of change in values only
'limit_val is the value, if any, that the column data will approach over damping_interval number of seconds
'   if none is specified, strict interpolation will be used
'window is the number of values to use in regression of initial and terminal gaps

Dim warn_id As Integer 'function output
Dim i As Long
Dim count As Long
Dim start_val As Double 'starting value for interpolation
Dim rtk_rows As Long
Dim dt As Long 'time difference in seconds
Dim r As linear 'elevation regression
Dim j As Long
Dim dir As Integer 'direction of interpolation or extrapolation
Dim target As Long 'start row for interpolation or extrapolation
Dim extra As Boolean 'true if extrapolating beyond endpoints
Dim d As Long
Dim midpoint As Single
Dim wi As Single 'log damping of standard interpolation weight vs limit_val in average
Dim soln_type As String
Dim shift As Integer

warn_id = 0
i = 2
Do While i <= seconds + 1
    If Cells(i, icol) = "" Then
        'count blanks
        count = 1
        Do While Cells(i + count, icol) = "" And i + count <= seconds + 1
            count = count + 1
        Loop
        start_val = 0
        extra = False
        'set interpolation parameters
        If count = seconds Then 'no data for whole record
            'check starting rtk values
            rtk_rows = rtk.Cells(Rows.count, "B").End(xlUp).row 'last used row in rtk worksheet
            dt = timer(rtk.Cells(rtk_rows, 3), rtk.Cells(2, 3))
            If dt <= 7200 And rtk_col > 0 Then 'less than two hours between known rtk end points
                warn_id = 1
                r.b = (rtk.Cells(rtk_rows, rtk_col) - rtk.Cells(2, rtk_col)) / dt
                r.n = CInt(timer(Cells(2, 12), rtk.Cells(2, 3)))
                Cells(2, icol) = rtk.Cells(2, rtk_col) + r.b * (r.n + 1) 'first elevation
                count = count - 1
                dir = 1
                target = 2
            Else
                interp = 2
                Exit Function
            End If
        ElseIf i = 2 Then 'initial gap in rtk
            dir = -1
            target = count + 2
            dt = timer(Cells(target, 12), rtk.Cells(2, 3))
            If dt <= damping_interval * 2 And rtk_col > 0 Then 'set slope to interpolate from first value in rtk sheet
                r.b = (Cells(target, icol) - rtk.Cells(2, rtk_col)) / dt
                midpoint = (dt + 1) / 2
                shift = dt - count
            Else 'calculate elevation regression and set slope to extrapolate
                extra = True
                r = regress(target, icol, window, dir, 11) 'regression through window size number of non-blank points
                start_val = midmean(window, target, icol, 11)
            End If
            warn_id = 3
        ElseIf i + count - 1 = seconds + 1 Then 'terminal gap
            dir = 1
            target = i - 1
            rtk_rows = rtk.Cells(Rows.count, "B").End(xlUp).row 'last used row in rtk worksheet
            dt = timer(rtk.Cells(rtk_rows, 3), Cells(target, 12))
            If dt <= damping_interval * 2 And rtk_col > 0 Then 'set slope to interpolate to last value in rtk sheet
                r.b = (rtk.Cells(rtk_rows, rtk_col) - Cells(target, icol)) / dt
                midpoint = (dt + 1) / 2
                shift = 0
            Else 'calculate elevation regression and set slope to extrapolate
                extra = True
                r = regress(target, icol, window, dir, 11) 'regression through window size number of non-blank points
                start_val = midmean(window, target, icol, 11)
            End If
            If warn_id = 3 Then warn_id = 5 Else warn_id = 4
        Else 'middle gap--interpolate between two valid values
            dir = 1
            target = i - 1
            r.b = (Cells(i + count, icol) - Cells(target, icol)) / (count + 1)
            midpoint = (count + 1) / 2
            shift = 0
        End If
        If start_val = 0 Then start_val = Cells(target, icol)
        'interpolate empty cells
        wi = 1 'interpolation weighting defaults to 1
        For j = 1 To count
            If limit_val >= 0 Then
                If count = seconds Then
                    wi = 0 'for full record gap, use max values
                Else
                    If extra Then 'extrapolation beyond endpoints
                        d = j
                    Else
                        d = midpoint - Abs(j + shift - midpoint) 'number of records from a known point
                    End If
                    If d < damping_interval Then
                        wi = WorksheetFunction.Log(damping_interval - d + 1, damping_interval)
                    Else: wi = 0
                    End If
                End If
            End If
            Cells(target + j * dir, icol) = (r.b * j * dir + start_val) * wi + limit_val * (1 - wi)
            'during elevation interpolation, fill in rtk solution type
            If icol = 21 Then
                If wi > 0 Then
                    Cells(target + j * dir, 22) = Left(Cells(target, 22), 5) & ",interpolated"
                Else: Cells(target + j * dir, 22) = "None,interpolated"
                End If
            End If
        Next j
        'use formulas for northing and easting in case x-offset and y-offset are changed manually
        If icol = 8 Or icol = 9 Then
            shift = icol - 8
            Cells(i, 14 - shift).Resize(count).FormulaR1C1 = "=RC[" & -(10 - shift * 2) & "] + RC[" & -(6 - shift * 2) & "]"
            Cells(i, 14 - shift).Resize(count).Interior.Color = 49407
        End If
        'highlight records with interpolated data in orange
        Cells(i, icol).Resize(count).Interior.Color = 49407
        i = i + count + 1
    Else
        i = i + 1
    End If
Loop

interp = warn_id

End Function

Private Function midmean(maxcount As Integer, vrow As Long, vcol As Integer, Optional check_col As Integer = 0) As Double
'Calculate average of maxcount values in column vcol both above and below row vrow
'if check_col is passed, only count values when check_col is not empty
Dim k As Integer
Dim n1 As Integer
Dim n2 As Integer

If check_col = 0 Then check_col = vcol
midmean = Cells(vrow, vcol)
k = 0
n1 = 0
Do While k < maxcount And vrow - k > 2
    k = k + 1
    If Cells(vrow - k, check_col) <> "" Then
        n1 = n1 + 1
        midmean = midmean + Cells(vrow - k, vcol)
    End If
Loop
k = 0
n2 = 0
Do While k < maxcount And vrow + k <= seconds
    k = k + 1
    If Cells(vrow + k, check_col) <> "" Then
        n2 = n2 + 1
        midmean = midmean + Cells(vrow + k, vcol)
    End If
Loop
midmean = midmean / (n1 + n2 + 1)

End Function

Private Function timer(time1, time2 As Date) As Long
'Evaluate time difference in seconds between to dates
'evaluates as time1 minus time2, may be negative
    If (Hour(time1) - Hour(time2)) < 0 Then  'Check for UTC time passing midnight
        timer = (24 + Hour(time1) - Hour(time2)) * CLng(3600) + (Minute(time1) - Minute(time2)) * 60 + Second(time1) - Second(time2)
    Else
        timer = (Hour(time1) - Hour(time2)) * CLng(3600) + (Minute(time1) - Minute(time2)) * 60 + Second(time1) - Second(time2)
    End If
End Function

Sub critical()
'Identify critical points in elevation and interpolate linearly
'between critical points. Points are stored in crit() with values
'in yc(), which could be passed as arguments to a more sophisticated
'interpolation technique such as splining.
Dim zmin(), zmax() As Single 'min/max possible elevation from RTK
Dim k, maxi, maxj As Integer
Dim i, j As Long
Dim z() As Single 'stores values of critical points
Dim crit() As Integer 'stores indices of critical points
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
    For i = 0 To 9 'test 10 starting values ranging from zmin to zmax
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
            If j = maxj And Abs(i - 4.5) < Abs(maxi - 4.5) Then 'where two j's are equivalent, choose i closer to the middle
                maxi = i
            End If
        End If
    Next i
    'set z(1) to best starting value
    z(1) = (zmax(1) - zmin(1)) / 9 * maxi + zmin(1)

'Initialize
    crit(1) = 1
    k = 2
    Cells(2, 17) = z(1) 'initial smoothed value

'Find critical points and interpolate between
For i = 2 To seconds
    zmin(i) = Cells(i + 1, 18)
    zmax(i) = Cells(i + 1, 19)
    critChange = False
    
    'check if last critical value exceeds current max or min
    If z(crit(k - 1)) > zmax(i) Then
        'reset k for two consecutive decreasing critical values
        If crit(k - 1) = i - 1 And (z(i - 1) = zmax(i - 1) And i > maxi) Then k = k - 1
        z(i) = zmax(i)
        crit(k) = i
        critChange = True
    ElseIf z(crit(k - 1)) < zmin(i) Then
        'reset k for two consecutive increasing critical values
        If crit(k - 1) = i - 1 And (z(i - 1) = zmin(i - 1) And i > maxi) Then k = k - 1
        z(i) = zmin(i)
        crit(k) = i
        critChange = True
    ElseIf i = seconds Then
        'extrapolate for tail values after last critical point
        '  only if z(seconds) is not already a critical point
        z(i) = m * i + b
        crit(k) = seconds
        critChange = True
    End If
    
    'check if linear interpolation violates min/max boundaries
    If critChange Then
        m = (z(i) - z(crit(k - 1))) / (crit(k) - crit(k - 1))
        b = z(i) - m * crit(k)
        For j = crit(k - 1) + 1 To crit(k) - 1
            If j = i Then 'exit loop when i has been reset to lesser value (boundaries were violated)
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

Private Sub save_files()
'Save worksheet with appropriate record number, save export tab as csv, and reopen Template
    Application.DisplayAlerts = False
    'save Current workbook with record number
        Worksheets("Combo").Activate
        ActiveWorkbook.SaveAs Filename:=path & record & "_Final.xlsx", FileFormat:=xlOpenXMLWorkbook
    'activate Export tab and save as CSV
        Worksheets("Export").Activate
        ActiveWorkbook.SaveAs Filename:=path & record & ".csv", FileFormat:=xlCSV
    'reopen blank R000XX_Final_Template
        ActiveWorkbook.Close False
        Workbooks.Open Filename:=basepath & "R000XX_Final_Template.xlsx"
    Application.DisplayAlerts = True
End Sub




