Attribute VB_Name = "process_sonar"
Option Explicit

Dim basepath As String 'path to R000XX_Final_Template
Dim path As String 'path to record
Dim record As String 'record name
Dim fileprompt As VbMsgBoxResult
Dim mybook As Workbook
Dim row As Long
Dim seconds As Long 'number of full seconds in sonar file
Const tlen = 0.114 'transducer length from center of mounting pole to sonar projector
Const max_z = 0.03 'maximum allowable rtk precision in z; 3 cm based on equipment limitations
Const max_xy = 0.15 'maximum allowable rtk precision in xy; based on limiting overall uncertainty
'Set worksheet columns for input and output
'Const sonar_time = 1, sonar_east = 2, sonar_north = 3, sonar_depth = 4, sonar_cog = 5
'Const rtk_baseid = 1, rtk_pointid = 2, rtk_time = 3, rtk_north = 4, rtk_east = 5, rtk_elev = 6, rtk_horiz_prec = 7, rtk_vert_prec = 8, rtk_soln_type = 17
'Const combo_index = 1, combo_son_time = 2, combo_row = 3, combo_x = 4, combo_y = 5, combo_cog = 6, combo_depth = 7, combo_xoffs = 8, combo_yoffs = 9
'Const combo_flag = 10, combo_pointid = 11, combo_rtk_time = 12, combo_north = 13, combo_east = 14, combo_cog_RTK = 15
'Const combo_elev = 16, combo_smooth = 17, combo_min = 18, combo_max = 19, combo_stdev_xy = 20, combo_stdev_z = 21, combo_soln_type = 22
'Const export_id = 1, export_north = 2, export_east = 3, export_elev = 4, export_time = 5, export_record = 6, export_rtkid = 7, export_flag =8
Private Type linear  'Data type to store linear regression values
    b As Double 'slope of regression
    a As Double 'intercept of regression
    n As Integer 'number of points in regression
End Type

Sub assembleXYZ()
Attribute assembleXYZ.VB_ProcData.VB_Invoke_Func = "A\n14"
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
' All functions were defined for the following worksheet columns:
' Sonar: (1)SonarTime, (2)Easting, (3)Northing, (4)Depth, (5)COG
' RTK: (1)Base_ID, (2)Point_ID, (3)Start Time, (4Northing, (5)Easting, (6)Elevation, (7)Horizontal Precision, (8)Vertical Precision, (9)Std Dev n, (10)Std Dev e, (11)Std Dev u, (12)Std Dev Hz, (13)Geoid Separation, (14)dN, (15)dE, (16)dHt, (17)Solution Type
' Combo: (1)Index, (2)Sonar Time, (3)Row, (4)X (Easting), (5)Y (Northing), (6)COG_sonar, (7)Depth, (8)X-offset, (9)Y-offset, (10)BLANK, (11)Point ID, (12)RTK Time, (13)Northing, (14)Easting, (15)COG_RTK, (16)Elev, (17)Smooth, (18)Min, (19)Max, (20)StDev_XY, (21)StDev_Z, (22)Solution Type
' Export: (1)Northing, (2)Easting, (3)Bed_elev, (4)DateTime, (5)Sonar_ID, (6)RTK_ID

Dim log_file As String 'file path to R000XX.txt processing log file
Dim DE_file As String 'file path to data_explorer index file
Dim rtk_file As String 'file path to rtk data file
Dim IsFile As Boolean
Dim ws As Integer
Dim rng As Range
Dim utc_shift As Integer 'number of hours to shift sonar times forward/back to match rtk times (UTC)
Dim out_points As Long

'Read defaults from text file
    Set mybook = ActiveWorkbook
    path = mybook.FullName
    path = Left(path, InStrRev(path, "\")) 'remove filename from path string
    basepath = path
    IO_defaults log_file, DE_file, rtk_file, "read"
    
'Set file paths
    combo.Activate
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
            If dir(path) <> vbNullString Then
                ChDir (path) 'set default path
            End If
            path = Application.GetOpenFilename("Humminbird dat files (*.dat),*.dat", , "Navigate to the sonar .dat file to process")
            If path = "False" Then Exit Sub
            record = Left(Right(path, Len(path) - InStrRev(path, "\")), 6) 'save record from root path
            path = Left(path, InStrRev(path, "\")) 'cut filename from root path
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
    IsFile = GetAttr(path & record & ".DAT.XYZ.csv")
    If Not IsFile Then
        fileprompt = MsgBox(record & ".DAT.XYZ.csv sonar file not found! Processing cancelled.", vbCritical, "File not found!")
        Exit Sub
    End If
    On Error GoTo 0
    
'Clear anything past header row on first four worksheets
    For ws = 1 To 5
        Set rng = Sheets(ws).UsedRange
        Set rng = rng.Offset(1, 0).Resize(rng.Rows.count - 1)
        rng.ClearContents
        rng.Interior.ColorIndex = 0
    Next ws

'Import sonar ping time stamps from IDX file
    sonar_import DE_file, utc_shift
    If fileprompt = vbCancel Then Exit Sub
    
'Divide sonar data by full seconds and write to combo sheet
    sonar_to_combo

'Flag points listed in processing log for inspection or deletion
    flag_points log_file, utc_shift

'Import pertinent rtk data to rtk worksheet and insert missing data lines
    rtk_import rtk_file, Cells(2, 2), Cells(seconds + 1, 2)
    If fileprompt = vbCancel Then Exit Sub

'Extract rtk data to combo sheet
    rtk_to_combo
    
'Interpolate gaps in rtk data
    interpolate_rtk

'Smooth rtk elevation points
    critical

'Extract data to export sheet
    export_data out_points
    
'Update Plotting ranges
    update_plots out_points

'Save files
Dim SaveOn As VbMsgBoxResult
    IO_defaults log_file, DE_file, rtk_file, "write" 'write defaults file
    SaveOn = MsgBox("Save record and export files?", vbYesNo, "Processing Complete")
    If SaveOn = vbYes Then
        save_files
        fileprompt = MsgBox("Record " & record & " successfully processed and saved. View processed file?", vbYesNo)
        'reopen blank R000XX_Final_Template
        Application.Workbooks.Open basepath & "R000XX_Final_Template.xlsm"
        If fileprompt = vbYes Then Workbooks.Open path & record & "_Final.xlsx"
        mybook.Close False 'close running workbook (now at R000XX.csv)
    Else
        MsgBox ("Record " & record & " processed but not saved or exported")
    End If

End Sub

Private Sub IO_defaults(ByRef log_file As String, ByRef DE_file As String, ByRef rtk_file As String, Optional IOtype As String = "read")
'read/write last used record, DE_file, and rtk_file to/from R000XX_defaults.txt
Dim f As Integer 'file index number
    f = FreeFile
    If IOtype = "read" Then
        ChDir (path) 'set default path
        On Error Resume Next
        Open "R000XX_defaults.txt" For Input As #f
        If Err.Number <> 0 Then
            record = "na"
            log_file = "na"
            DE_file = "na"
            rtk_file = "na"
            Exit Sub
        End If
        On Error GoTo 0 'reset error handling
        Input #f, record
        Input #f, path
        Input #f, log_file
        Input #f, DE_file
        Input #f, rtk_file
        Close #f
    Else 'overwrite defaults file
        ChDir (basepath) 'save in initial location
        Open "R000XX_defaults.txt" For Output As #f
        Write #f, record
        Write #f, path
        Write #f, log_file
        Write #f, DE_file
        Write #f, rtk_file
        Close #f
    End If
End Sub

Private Sub sonar_import(ByRef DE_file As String, ByRef time_shift As Integer)
'Import sonar ping time stamps from IDX file
'   Humminbird sonar files consist of a DAT file, and an IDX and SON file for each channel:
'       DAT file -- holds record info for starting date and time, duration, and beginning coordinates
'       IDX files -- two 4-byte fields per record specifying a time increment and a line index in the SON file
'       SON files -- complex binary file with full navigation data and imagery for each record
'   If there are three channels, B000 is downscan, B001 is sidescan left, and B002 is sidescan right
'   If there are four channels, B000 is downscan01, B001 is downscan-2, B002 is sidescan left, and B003 is sidescan right
'time_shift is the number of hours to shift sonar times forward/back to match rtk times (UTC)
'   for Pacific time, use 8 for records during PST (winter Nov-Mar) and 7 during PDT (summer Mar-Nov)
'   based on a sonar file stated in local time and rtk file stated in UTC time
Dim f As Integer 'file index number
Dim fLen As Long 'length of IDX file in bytes
Dim idx_data() As Byte 'holds binary data from IDX file
Dim DE_book As Workbook 'data explorer workbook
Dim rng As Range
Dim dat_book As Workbook 'R000XX.DAT.XYZ.csv path
Dim fulldate As Double 'initial datetime value (days since 1900 + hr/24 + min/60 + sec/3600 + ms/3600/1000)
Dim cog As Double 'course over ground

'Load binary data from IDX file
    f = FreeFile
    Open path & record & "\B002.IDX" For Binary Access Read As #f 'use sidescan channel for time stamps
    fLen = LOF(f)
    ReDim idx_data(1 To fLen)
    Get f, , idx_data
    Close f

'Open data explorer file
    check_file DE_file, "Data Explorer"
    If fileprompt = vbCancel Then Exit Sub
    Application.ScreenUpdating = False
    DoEvents 'allows macro to continue if started with hotkeys using SHIFT
    Set DE_book = Workbooks.Open(Filename:=DE_file, ReadOnly:=True)

'Read data and close file
    Set rng = DE_book.Worksheets(1).Cells(1, 4).Resize(DE_book.Worksheets(1).Cells(Rows.count, "A").End(xlUp).row, 1) 'last used row
    row = Application.Match(Val(Right(record, 5)), rng, 0) 'find record number in DE_book
    fulldate = DE_book.Worksheets(1).Cells(row, 1) 'initial date from DE_book
    If Month(fulldate) > 2 And Month(fulldate) < 11 Then
        time_shift = 7
    Else
        time_shift = 8
    End If
    fileprompt = MsgBox("Accept time offset of UTC -" & time_shift & "?", vbYesNo, "UTC offset")
    If fileprompt = vbNo Then
        time_shift = InputBox("Set UTC time offset for date of data collection (" & Month(fulldate) & "/" & Day(fulldate) & "/" & Year(fulldate) & "). " & _
            "Set the offset to 7 for Pacific Standard Time (winter months) or 8 for Pacific Daylight Time (summer months).", _
            "Set UTC offset", IIf(time_shift = 7, 8, 7))
    End If
    fulldate = fulldate + DE_book.Worksheets(1).Cells(row, 6) + time_shift / 24  'Shift times forward or back for UTC correction
    DE_book.Close
    Set dat_book = Workbooks.Open(Filename:=path & record & ".DAT.XYZ.csv", ReadOnly:=True)
    For row = 0 To fLen \ 8 - 1 'backslash operator is integer division
        'first 4 bytes of each 8-byte record hold time stamp info
        sonar.Cells(row + 2, 1) = fulldate + (idx_data(row * 8 + 2) * 65536 + idx_data(row * 8 + 3) * CLng(256) + idx_data(row * 8 + 4)) / 24 / 3600 / 1000
    Next row
    dat_book.Worksheets(1).Range("A2:C" & row + 1).Copy Destination:=sonar.Range("B2") 'copy xyz data from dat_book to sonar sheet
    dat_book.Close
    Application.ScreenUpdating = True

'Calculate course over ground (cog)
    For row = 1 To fLen \ 8 - 1
        sonar.Cells(row + 1, 5) = get_cog(sonar.Cells(row + 1, 2), sonar.Cells(row + 2, 2), sonar.Cells(row + 1, 3), sonar.Cells(row + 2, 3))
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
    Do While Round(sonar.Cells(row + 1, 2), 6) = Round(sonar.Cells(row, 2), 6) Or Round(sonar.Cells(row + 1, 3), 6) = Round(sonar.Cells(row, 3), 6)
    'seconds digits match
        row = row - 1
    Loop

'Remove partial seconds at end of record count (row)
    Do While Int(sonar.Cells(row + 1, 1) * 24 * 3600) = Int(sonar.Cells(row, 1) * 24 * 3600)
    'seconds digits match
        row = row - 1
    Loop

'Average sonar records for x, y, depth, and calculated cog over full-second intervals
    seconds = timer(sonar.Cells(row, 1), sonar.Cells(2, 1)) 'number of full seconds in data record
    prev = 0
    srow = 0
    t1 = sonar.Cells(srow + 2, 1)
    For i = 1 To seconds
        Cells(i + 1, 1) = i 'index
        If i = 1 Then
            'set first value to start time
            Cells(2, 2) = sonar.Cells(2, 1)
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

Private Sub flag_points(ByRef log_file As String, time_shift As Integer)
'Flags records for inspection or deletion based on data in log_file
'flag handling:
'  records marked "d":
'    rtk elevation is not imported, elevation is later interpolated
'    position data are unaffected
'    points are not copied to export sheet
'  records marked "i":
'    flag is marked on export sheet
'log_file formatting:
'text file with record ID integer and notes on one line, followed by multiple lines
'  led by a two-space indent, a flag of "i" or "d", the time range affected, and optional comments
'repeat with no gaps between records
'integer ID is just the integer portion of the record, ex. R00092-->92
'Line 1: ID notes
'Line 2:   f hh:mm:ss-hh:mm:ss <comments>

Dim f As Integer 'file index number
Dim response As VbMsgBoxResult
Dim fdata As String
Dim log_data() As String 'data from log file
Dim notes As String 'notes in log file to output in a message box
Dim record_num As String 'numeral value of record
Dim dash As Integer
Dim count As Integer
Dim i As Integer
Dim log_row As New Collection 'pertinent row records from log data
Dim start_time As Date
Dim entry As Variant
Dim op As String
Dim time1 As String, time2 As String
Dim t1 As Date, t2 As Date
Dim dt As Long

'Read log file
    f = FreeFile
    On Error Resume Next
    Open log_file For Binary As #f
    ChDir (path) 'set default path
    Do While Err.Number <> 0
        response = MsgBox("Log file not found! Select a new log file?", vbYesNo, "File not found")
        If response = vbNo Then Exit Sub
        Err.Number = 0
        log_file = Application.GetOpenFilename("Text Files (*.txt),*txt", , "Select log file")
        If log_file = "" Then Exit Sub
        Open log_file For Binary As #f
    Loop
    On Error GoTo 0
    
    fdata = Space$(LOF(f)) 'read
    Get #f, , fdata
    Close #f
    log_data() = Split(fdata, vbCrLf)

'Identify pertinent log data records
    row = 0
    notes = "x"
    record_num = CStr(Val(Right(record, 5)))
    Do While row <= UBound(log_data)
        dash = InStr(1, log_data(row), " ")
        If dash = 0 Then count = Len(log_data(row)) Else count = dash - 1
        If Left(log_data(row), count) = record_num Then
            notes = Right(log_data(row), Len(log_data(row)) - (Len(record_num)))
            If Len(notes) > 0 Then notes = Right(notes, Len(notes) - 1)
            entry = 1
            If row + entry <= UBound(log_data) Then
                Do While Left(log_data(row + entry), 2) = "  "
                    log_row.Add Right(log_data(row + entry), Len(log_data(row + entry)) - 2) 'remove indent
                    entry = entry + 1
                    If row + entry > UBound(log_data) Then Exit Do
                Loop
            End If
            Exit Do
        End If
        row = row + 1
    Loop
    If notes = "x" Then
        MsgBox Prompt:="Record " & Str(record_num) & " not found in log file!", Title:="Record not found"
        Exit Sub
    Else
        If Len(notes) > 0 Then MsgBox "Processing notes: " & notes
    End If

'Flag records in column J on combo sheet
    start_time = combo.Cells(2, 2) - Int(combo.Cells(2, 2))
    On Error Resume Next
    For Each entry In log_row
        Err.Number = 0
        op = Left(entry, 1)
        dash = InStr(3, entry, "-")
        If dash > 0 Then
            time1 = Mid(entry, 3, dash - 3)
        Else
            time1 = Mid(entry, 3, 8)
        End If
        time2 = Mid(entry, dash + 1, InStr(7, entry, " ") - dash - 1)
        If time1 = "start" Then
            t1 = start_time
        Else: t1 = TimeValue(time1) + time_shift / 24
        End If
        If time2 = "end" Then
            t2 = combo.Cells(seconds + 1, 2) - Int(combo.Cells(seconds + 1, 2))
        ElseIf dash > 0 Then
            t2 = TimeValue(time2) + time_shift / 24
        Else 'single point to flag
            t2 = t1
        End If
        If Err.Number = 0 Then
            dt = timer(t2, t1) + 1
            count = timer(t1, start_time)
            If t1 < start_time Or dt + count > seconds Or (t2 < t1 And Hour(t2) <> 0) Then
                MsgBox "Problem reading log entry: '" & entry & "'" & vbCrLf & _
                  "Times outside of record bounds -- record skipped.", vbCritical, "Log file record error"
            Else
                combo.Cells(2, 10).Offset(count).Resize(dt) = op
                'Update plot ranges
                If op = "d" Then
                    plotdata.Cells(2, 1).Offset(count).Resize(dt) = 1
                Else: plotdata.Cells(2, 2).Offset(count).Resize(dt) = 1
                End If
            End If
        Else
            MsgBox "Problem reading log entry: '" & entry & "'" & vbCrLf & _
              "Invalid time format -- record skipped.", vbCritical, "Log file record error"
        End If
    Next entry
    On Error GoTo 0

End Sub

Private Sub check_file(ByRef file_path As String, name As String)
'Check whether to use default file_path, and if so whether the file exists,
'  otherwise request correct file path
Dim IsFile As Boolean

    If file_path = "na" Then
        fileprompt = vbNo
        MsgBox "Defaults file not found! Please navigate to " & name & " file", vbCritical, "File read error"
    Else: fileprompt = MsgBox("Use default " & name & " file path? (" & file_path & ")", vbYesNoCancel, name & " source")
    End If
    If fileprompt = vbCancel Then Exit Sub
    If fileprompt = vbNo Then file_path = Application.GetOpenFilename("Excel Workbooks (*.xls;*.xlsx),*.xls;*.xlsx", , "Select " & name & " file")
    ChDir (path) 'set default path
    On Error Resume Next
    IsFile = False
    IsFile = GetAttr(file_path)
    Do While Not IsFile
        MsgBox "Invalid file path!" & vbCrLf & "Please navigate to " & name & " file.", vbCritical, "File read error"
        file_path = Application.GetOpenFilename("Excel Workbooks (*.xls;*.xlsx),*.xls;*.xlsx", , "Select " & name & " file")
        IsFile = GetAttr(file_path)
    Loop
    On Error GoTo 0

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
        If .a > 0 Then .b = .b / .a Else .b = 0 'slope
        .a = ybar - .b * xbar 'intercept
    End With

End Function

Private Sub rtk_import(ByRef rtk_file As String, date1 As Date, date2 As Date)
'Import rtk data to rtk worksheet, delete duplicate entries, and insert missing rows
Dim rng As Range
Dim rtk_book As Workbook
Dim rtk_start, rtk_end As Long 'starting and ending row numbers in rtk_book to copy over
Dim dt As Long

'Open rtk data file
    check_file rtk_file, "RTK data"
    If fileprompt = vbCancel Then Exit Sub
    Application.ScreenUpdating = False
    Set rtk_book = Workbooks.Open(Filename:=rtk_file, ReadOnly:=True)
    
'Import rtk data and close file
    rtk_book.Worksheets(1).Activate
    rtk_start = find_row(date1, "initial")
    rtk_end = find_row(date2, "final")
    'copy data from rtk_book
    mybook.Worksheets("RTK").Range("A2:Q" & 2 + rtk_end - rtk_start).Value = rtk_book.Worksheets(1).Range("A" & rtk_start & ":Q" & rtk_end).Value
    Application.DisplayAlerts = False 'prevent "Save changes" dialog
    rtk_book.Close
    Application.DisplayAlerts = True
    mybook.Activate
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
                        dt = timer(Cells(combo_row, 2), rtk.Cells(rtk_row - 1, 3)) 'time to previous point
                        If combo_row - dt > 2 And dt <= 5 Then 'point within 5 sec has valid cog
                            r = regress(combo_row - dt, 15, 5, , 15) 'regression through last five valid cogs
                            cog = r.a + r.b * (1 + r.n) 'y = a + bx
                        Else
                            cog = Cells(combo_row, 6)
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
                    If avg <= max_z And combo.Cells(combo_row, 10) <> "d" Then
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
Const offset_damping = 10 'number of seconds to dampen x-, y- and z-offset trends on log scale before they reach zero
Const max_damping = 5 'number of seconds to transition from interpolated values of stdev_xy and stdev_z before reaching the maximum

'Interpolate northing and easting
x = interp(8, offset_damping, , 9999, 20) 'interpolate x-offset and easting
x = interp(9, offset_damping, , 9999, 20) 'interpolate y-offset and northing

'Interpolate missing rtk data
warn_id = interp(16, max_damping + 2, 6, 9999, 20) 'interpolate elevation
If warn_id = 5 Then
    warn(6) = True
    warn(7) = True
ElseIf warn_id > 0 Then
    warn(warn_id + 3) = True
End If
If warn_id = 0 Or warn_id > 2 Then
    x = interp(20, max_damping, 7, max_xy, 50) 'interpolate stdev_xy
    x = interp(21, max_damping, 8, max_z, 50) 'interpolate stdev_z
ElseIf warn_id = 1 Then
    Cells(2, 20).Resize(seconds) = ">" & max_xy
    Cells(2, 21).Resize(seconds) = max_z
    Cells(2, 22).Resize(seconds) = "None,interpolated"
    Cells(2, 21).Resize(seconds, 2).Interior.Color = 49407
Else 'warn_id = 2
    Cells(2, 22).Resize(seconds) = "None"
End If

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
'   gaps smaller than damping_interval are interpolated with no damping
'   damping_interval * 2 is also used as the maximum time in seconds to interpolate from rtk records beyond the ends instead of using slope trends
'rtk_col is the column with matching data in the rtk sheet
'   if none is specified, extrapolation will be based on regression of slope of change in values only
'limit_val is the value, if any, that the column data will approach over damping_interval number of seconds
'   if none is specified, strict interpolation will be used
'   to use average of values within window range, pass 9999 to function
'window is the number of values to use in regression of initial and terminal gaps

Dim warn_id As Integer 'function output
Dim use_avg As Boolean 'set limit_val to average of records within given range
Dim i As Long
Dim count As Long
Dim rtk_row As Long
Dim dt As Long 'time difference in seconds
Dim r As linear 'elevation regression
Dim j As Long
Dim dir As Integer 'direction of interpolation or extrapolation
Dim target As Long 'start row for interpolation or extrapolation
Dim extra As Boolean 'true if extrapolating beyond endpoints
Dim d As Long
Dim midpoint As Single
Dim wi As Single 'log damping of standard interpolation weight vs limit_val in average
Dim shift As Integer
Dim limit As linear 'store slope and intercept of slope between averaged limits for middle gaps

warn_id = 0
If limit_val = 9999 Then use_avg = True Else use_avg = False
i = 2
Do While i <= seconds + 1
    If Cells(i, icol) = "" Then
        'count blanks
        count = 1
        Do While Cells(i + count, icol) = "" And i + count <= seconds + 1
            count = count + 1
        Loop
        extra = False
        'set interpolation parameters
        If count = seconds Then 'no data for whole record
            'check starting rtk values
            rtk_row = rtk.Cells(Rows.count, "B").End(xlUp).row 'last used row in rtk worksheet
            dt = timer(rtk.Cells(rtk_row, 3), rtk.Cells(2, 3))
            If dt <= 7200 And rtk_col > 0 Then 'less than two hours between known rtk end points
                warn_id = 1
                r.b = (rtk.Cells(rtk_row, rtk_col) - rtk.Cells(2, rtk_col)) / dt
                r.n = CInt(timer(Cells(2, 12), rtk.Cells(2, 3)))
                r.a = rtk.Cells(2, rtk_col) + r.b * r.n 'first elevation
                dir = 1
                target = 1
            Else
                interp = 2
                Exit Function
            End If
        Else
            'set parameters based on gap location in record
            If i = 2 Then 'initial gap in rtk
                dir = -1
                target = count + 2
                rtk_row = 2
                warn_id = 3
            Else
                dir = 1
                target = i - 1
                If i + count - 1 = seconds + 1 Then 'terminal gap
                    rtk_row = rtk.Cells(Rows.count, "B").End(xlUp).row 'last used row in rtk worksheet
                    If warn_id = 3 Then warn_id = 5 Else warn_id = 4
                Else 'middle gap--interpolate between two valid values
                    r.b = (Cells(i + count, icol) - Cells(target, icol)) / (count + 1)
                    midpoint = (count + 1) / 2
                End If
            End If
            
            'set values for interpolation or extrapolation of endpoints
            If i = 2 Or warn_id > 3 Then 'initial or terminal gap
                If Cells(target + count * dir, 10) = "d" Then
                    extra = True
                Else
                    dt = timer(rtk.Cells(rtk_row, 3), Cells(target, 12), dir)
                    If dt - count > damping_interval * 2 Or rtk_col = 0 Then extra = True
                End If
                If extra Then
                    'extrapolate beyond known points
                    r = regress(target, icol, window, dir, 11) 'regression through window size number of non-blank points
                Else
                    'set slope to interpolate from known point in rtk sheet
                    r.b = (dir * (Cells(target, icol) - rtk.Cells(rtk_row, rtk_col))) / dt
                    midpoint = (dt + 1) / 2
                End If
            End If
        End If
        
        'set interpolation starting value unless all records empty
        If Not warn_id = 1 Then r.a = Cells(target, icol)
        
        'adjust limit values
        limit.b = 0
        If use_avg And count >= damping_interval Then
            limit.a = midmean(window, target, icol, -dir, 11)
            If i > 2 And warn_id < 4 Then
                limit.b = (midmean(window, target + count + 1, icol, 1, 11) - limit.a) / (count + 1)
            End If
        Else: limit.a = limit_val
        End If

        'interpolate empty cells
        wi = 1 'interpolation weighting defaults to 1
        For j = 1 To count
            If limit_val >= 0 And count >= damping_interval Then
                If count = seconds Then
                    wi = 0 'for full record gap, use max values
                Else
                    If extra Then 'extrapolation beyond endpoints
                        d = j
                    Else
                        d = midpoint - Abs(j - midpoint)  'number of records from a known point
                    End If
                    If d < damping_interval Then
                        wi = WorksheetFunction.Log(damping_interval - d + 1, damping_interval)
                    Else: wi = 0
                    End If
                End If
            End If
            Cells(target + j * dir, icol) = (r.b * j * dir + r.a) * wi + (limit.a + limit.b * j) * (1 - wi)
            'during stdev_z interpolation, fill in rtk solution type
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

Private Function midmean(maxcount As Integer, vrow As Long, vcol As Integer, Optional dir As Integer = 0, Optional check_col As Integer = 0) As Double
'Calculate average of maxcount values in column vcol both above and below row vrow (inclusive)
'dir is the direction to search, -1 for up, 1 for down, or 0 for both
'if check_col is passed, only count values when check_col is not empty
Dim k As Integer
Dim n1 As Integer
Dim n2 As Integer

If check_col = 0 Then check_col = vcol
midmean = Cells(vrow, vcol)
n1 = 0
n2 = 0

If dir <= 0 Then
    k = 0
    Do While k < maxcount And vrow - k > 2
        k = k + 1
        If Cells(vrow - k, check_col) <> "" And Cells(vrow - k, vcol) <> "" Then
            n1 = n1 + 1
            midmean = midmean + Cells(vrow - k, vcol)
        End If
    Loop
End If
If dir >= 0 Then
    k = 0
    Do While k < maxcount And vrow + k <= seconds
        k = k + 1
        If Cells(vrow + k, check_col) <> "" Then
            n2 = n2 + 1
            midmean = midmean + Cells(vrow + k, vcol)
        End If
    Loop
End If
midmean = midmean / (n1 + n2 + 1)

End Function

Private Function timer(time1 As Date, time2 As Date, Optional ord As Integer = 1) As Long
'Evaluate time difference in seconds between to dates
'by default evaluates to time1 minus time2, may be negative
'if ord = -1, evaluates to time2 minus time1
    If Abs(ord) <> 1 Then
        MsgBox ("Order of evaluation 'ord' passed to Function timer must be 1 or -1!" & vbCrLf & "Defaulting to 1")
        ord = 1
    End If
    timer = ord * (Hour(time1) - Hour(time2)) * CLng(3600) + ord * (Minute(time1) - Minute(time2)) * 60 + ord * (Second(time1) - Second(time2))
    If ord * (Hour(time1) - Hour(time2)) < 0 Then 'Check for UTC time passing midnight
        timer = timer + 24 * CLng(3600)
    End If
End Function

Private Sub critical()
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

Private Sub export_data(ByRef records_out As Long)
Dim count As Long

'write excel formulas to export sheet so changes to combo are dynamically adjusted
    expo.Cells(2, 1).Resize(seconds).Formula = "=" & CStr(Val(Right(record, 5))) & " & ""."" & ROW() - 1" 'point_ID
    expo.Cells(2, 1).Resize(seconds).Copy
    expo.Cells(2, 1).PasteSpecial xlPasteValues 'paste values
    expo.Cells(2, 2).Resize(seconds, 2).FormulaR1C1 = "=Combo!RC[11]" 'northing, easting
    expo.Cells(2, 4).Resize(seconds).FormulaR1C1 = "=Combo!RC[13]-Combo!RC[3]" 'bottom = smoothed_elevation - depth
    expo.Cells(2, 5).Resize(seconds).FormulaR1C1 = "=Combo!RC[-3]" 'datetime
    expo.Cells(2, 6).Resize(seconds) = record 'sonar id
    expo.Cells(2, 7).Resize(seconds).FormulaR1C1 = "=IF(Combo!RC[4]="""",""N/A"",Combo!RC[4])" 'rtk id
    combo.Cells(2, 10).Resize(seconds).Copy Destination:=expo.Cells(2, 8) 'flag
    
'remove rows flagged "d"
    On Error Resume Next
    records_out = seconds
    row = Application.Match("d", expo.Cells(2, 8).Resize(records_out), 0) + 1
    Do While Err.Number = 0
        count = 0
        Do While expo.Cells(row + count, 8) = "d"
            count = count + 1
        Loop
        expo.Cells(row, 1).Resize(count).EntireRow.Delete
        records_out = records_out - count
        row = Application.Match("d", expo.Cells(2, 8).Resize(records_out), 0) + 1
    Loop
    On Error GoTo 0
    
'update formatting
    expo.Cells(2, 2).Resize(records_out, 2).NumberFormat = "0.0000"
    expo.Cells(2, 4).Resize(records_out).NumberFormat = "0.000"
    expo.Cells(2, 5).Resize(records_out).NumberFormat = "m/d/yyyy hh:mm:ss"
    
'update plot range
    count = Len(CStr(Val(Right(record, 5))))
    plotdata.Cells(2, 3).Resize(records_out).FormulaR1C1 = "=VALUE(RIGHT(Export!RC[-2],LEN(Export!RC[-2])-" & count + 1 & "))"

End Sub

Private Sub update_plots(records_out As Long)
    With Chart5 'TrackPlot
        .FullSeriesCollection(1).XValues = "=Combo!$D$2:$D$" & seconds + 1 'sonar x (easting)
        .FullSeriesCollection(1).Values = "=Combo!$E$2:$E$" & seconds + 1 'sonar y (northing)
        .FullSeriesCollection(2).XValues = "=Combo!$N$2:$N$" & seconds + 1 'rtk x
        .FullSeriesCollection(2).Values = "=Combo!$M$2:$M$" & seconds + 1 'rtk y
        .ChartTitle.Text = record & " Navigation Tracks"
    End With
    With Chart6 'Smoothing
        .FullSeriesCollection(1).Values = "=Combo!$Q$2:$Q$" & seconds + 1 'smoothed elev
        .FullSeriesCollection(2).Values = "=Combo!$R$2:$R$" & seconds + 1 'min elev
        .FullSeriesCollection(3).Values = "=Combo!$S$2:$S$" & seconds + 1 'max elev
        .FullSeriesCollection(4).Values = "=Combo!$P$2:$P$" & seconds + 1 'raw elev
        .FullSeriesCollection(5).Values = "=PlotData!$A$2:$A$" & seconds + 1 'deleted
        .FullSeriesCollection(6).Values = "=PlotData!$B$2:$B$" & seconds + 1 'inspect
        .Axes(xlCategory).MaximumScale = seconds
    End With
    With Chart7 'Depth
        .FullSeriesCollection(1).XValues = "=Combo!$A$2:$A$" & seconds + 1 'depth
        .FullSeriesCollection(1).Values = "=Combo!$G$2:$G$" & seconds + 1
        .FullSeriesCollection(2).XValues = "=PlotData!$C$2:$C$" & records_out + 1 'bed elev
        .FullSeriesCollection(2).Values = "=Export!$D$2:$D$" & records_out + 1
        .FullSeriesCollection(3).Values = "=PlotData!$A$2:$A$" & seconds + 1 'deleted
        .FullSeriesCollection(4).Values = "=PlotData!$B$2:$B$" & seconds + 1 'inspect
        .Axes(xlCategory).MinimumScale = 1
        .Axes(xlCategory).MaximumScale = seconds
    End With
End Sub

Private Sub save_files()
'Save worksheet with appropriate record number, save export tab as csv, and reopen Template
    Application.DisplayAlerts = False
    'save Current workbook with record number
        combo.SaveAs Filename:=path & record & "_Final.xlsx", FileFormat:=xlOpenXMLWorkbook
    'activate Export tab and save as CSV
        expo.SaveAs Filename:=path & record & ".csv", FileFormat:=xlCSV
        Set mybook = ActiveWorkbook
    Application.DisplayAlerts = True
End Sub