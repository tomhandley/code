Attribute VB_Name = "Import_dat"
Option Explicit

Sub Import_dat()
'
'Humminbird sonar files consist of a DAT file, and an IDX and SON file for each channel:
'   DAT file -- holds record info for date and time of the beginning, duration, and beginning coordinates
'   IDX files -- indices with two 4-byte fields per record specifying a time increment and a line location in the SON file
'   SON files -- contain full navigation data for each record
'This function imports hex data from sonar idx files and calculates the time offset for each record

'Import data to Sonar worksheet
Dim path, record As String  'root path containing R000XX files
Dim rtype As String  'Specify IDX or DAT file
Dim bytes() As Byte
Dim f As Integer
Dim i, row, fLen As Long

    path = "C:\Users\thandley.AD3\Desktop\Thesis\Bathymetry\2015-01-21_lindsey\Sonar\RECORD\"
    record = "R00027"
    rtype = "IDX"
    'Import nav data and timestamps
    f = FreeFile
    Open path & record & IIf(rtype = "IDX", "\B002.IDX", ".DAT") For Binary Access Read As #f
    fLen = LOF(f)
    ReDim bytes(1 To fLen)
    Get f, , bytes
    Close f
    If rtype = "DAT" Then
        For i = 1 To fLen
            Cells(1, i) = bytes(i)
            Cells(2, i) = IIf(bytes(i) < 16, "0", "") & Hex$(bytes(i))
        Next i
    Else
        For row = 1 To fLen \ 8  'Backslash operator is integer division
            For i = 1 To 8
                Cells(row, i) = IIf(bytes((row - 1) * 8 + i) < 16, "0", "") & Hex$(bytes((row - 1) * 8 + i))
            Next i
        Next row
        'Calculate time difference without printing columns of hex data
        For row = 0 To fLen \ 8 - 1 'Backslash operator is integer division
                Cells(row + 1, 10) = bytes(row * 8 + 2) * 65536 + bytes(row * 8 + 3) * CLng(256) + bytes(row * 8 + 4)
        Next row
    End If
End Sub
