Attribute VB_Name = "Import_idx"
Option Explicit

Sub Import_idx()
'
'Humminbird sonar files consist of a DAT file, and an IDX and SON file for each channel:
'   DAT file -- holds record info for date and time of the beginning, duration, and beginning coordinates
'   IDX files -- indices with two 4-byte fields per record specifying a time increment and a line location in the SON file
'   SON files -- contain full navigation data for each record
'This function imports hex data from sonar idx files and calculates the time offset for each record

'Import data to Sonar worksheet
Dim path, record As String  'root path containing R000XX files
Dim bytes() As Byte
Dim f As Integer
Dim i, row, fLen As Long
Dim DATAfile As String
Dim mybook, DATAbook As Workbook
Dim utm_shift As Byte
Dim di As Long
Dim ti, fulldate As Double

utm_shift = 8  'Set value to shift sonar times forward/back for UTM adjustment (to match RTK times)
    
    '--------------REDUNDANT
    path = "C:\Users\thandley.AD3\Desktop\Thesis\Bathymetry\2015-01-21_lindsey\Sonar\RECORD\"
    record = "R00027"
    Set mybook = ActiveWorkbook
    '--------------
    
    'Import nav data and timestamps
    f = FreeFile
    Open path & record & "\B002.IDX" For Binary Access Read As #f
    fLen = LOF(f)
    ReDim bytes(1 To fLen)
    Get f, , bytes
    Close f
    'Calculate time difference without printing columns of hex data

'XXXX----Initialize DATAfile to blank--will represent default inputbox entry
    DATAfile = ""
'XXXX----change default to read last used file path from R000XX, also change for RTK default
    If DATAfile = "" Then DATAfile = "C:\Users\thandley.AD3\Desktop\Thesis\Bathymetry\Data_explorer.xlsx"
    Set DATAbook = Workbooks.Open(DATAfile)
    row = WorksheetFunction.Match(Val(Right(record, 5)), DATAbook.Worksheets(1).Range("D2:D400"), 0) + 1
    di = DATAbook.Worksheets(1).Cells(row, 1)
    ti = DATAbook.Worksheets(1).Cells(row, 6) + utm_shift / 24  'Shift times forward or back for UTM correction
    
    DATAbook.Close
    mybook.Worksheets(1).Activate
    fulldate = di + ti
    For row = 0 To fLen \ 8 - 1 'Backslash operator is integer division
        Cells(row + 1, 11) = fulldate + (bytes(row * 8 + 2) * 65536 + bytes(row * 8 + 3) * CLng(256) + bytes(row * 8 + 4)) / 24 / 3600 / 1000
    Next row
End Sub
