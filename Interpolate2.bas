Attribute VB_Name = "Interpolate2"
Sub Interpolate2()
' Interpolate missing values linearly
' Only for gaps in column icol shorter than max
Dim i As Integer  'row counter
Dim irow As Long  'row index
Dim rspan As Integer  'number of rows for interpolation
Dim startr As Long  'starting row for interpolation
Dim val1, val2 As Double  'starting and ending values to interpolate between
Dim inc As Double  'incremental difference between each interpolated value
Const icol = 5  'column to interpolate
Const imax = 15  'maximum missing values to fill
Const maxval = 400  'maximum allowable value for data in icol
Const minval = 0  'minimum allowable value for data in icol

'Delete bad values
For irow = 4 To 35423
    If Cells(irow, icol) <= minval Then Cells(irow, icol).ClearContents
    If Cells(irow, icol) > maxval Then
        val1 = 0
        For i = irow - 10 To irow - 1
            val1 = val1 + Cells(irow, icol)  'running sum of values for averaging
        Next i
        val1 = val1 / 10 'average of previous 10 values
        Cells(irow, icol).ClearContents
    End If
Next irow

'Interpolate missing values
For irow = 4 To 35423
    If Cells(irow, icol) = "" Then
        rspan = 1
        Do While Cells(irow + rspan, icol) = ""
            rspan = rspan + 1
        Loop
        If rspan <= imax Then
            val1 = Cells(irow - 1, icol)
            val2 = Cells(irow + rspan, icol)
            inc = (val2 - val1) / (rspan + 1)
            For i = 1 To rspan
                Cells(irow + i - 1, icol) = val1 + inc * i
            Next i
        End If
        irow = irow + rspan - 1
    End If
Next irow

End Sub



