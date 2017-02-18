Attribute VB_Name = "Interpolate"
Sub Interpolate()
Attribute Interpolate.VB_ProcData.VB_Invoke_Func = "i\n14"
' Interpolate missing values linearly
' Run sub from any cell in the leftmost column to interpolate
' There must be valid starting and ending values to interpolate between
Dim i, j As Integer  'row and column counters
Dim index As Range  'starting cell value
Dim rspan, cspan As Integer  'number of rows and columns for interpolation
Dim startr, startc As Long  'starting and ending rows and columns for interpolation
Dim val1, val2 As Double  'starting and ending values to interpolate between
Dim inc As Double  'incremental difference between each interpolated value
Const maxcol = 9  'last column to interpolate missing values

If ActiveCell > 0 Then
    i = MsgBox("Invalid cell selected! Highlight an empty cell and restart.", vbExclamation, "Invalid cell")
Else
    Set index = Range(ActiveCell, ActiveCell)
    'count blank rows
    i = 1
    Do While Cells(index.row - i, index.Column) = 0
        i = i + 1
    Loop
    rspan = i
    startr = index.row - i + 1
    i = 1
    Do While Cells(index.row + i, index.Column) = 0
        i = i + 1
    Loop
    rspan = rspan + i - 1
    
    'count blank columns
    startc = index.Column
    cspan = maxcol - index.Column + 1
    
    'interpolate missing values
    For j = 0 To cspan - 1
        val1 = Cells(startr - 1, startc + j)
        val2 = Cells(startr + rspan, startc + j)
        inc = (val2 - val1) / (rspan + 1)
        For i = 1 To rspan
            Cells(startr + i - 1, startc + j) = val1 + inc * i
        Next i
    Next j
End If

End Sub


