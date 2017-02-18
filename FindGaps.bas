Attribute VB_Name = "FindGaps"
Sub FindGaps()
' Identify gaps or repeated values in time-series data
'
Dim i, last As Long
Dim index As Integer
Dim diff As Single  'Minimum difference between two times to flag as a gap
Dim gaps() As Long  'Array to store row indices of flagged time data
Const maxgap = 20  'Max number of gaps to store
Const outcol = 11  'Column to output gap data to
ReDim gaps(1 To maxgap)

'Search for gaps in time record
last = Cells(Rows.count, 1).End(xlUp).row
index = 0
diff = Round(Cells(5, 1) - Cells(4, 1), 5) + 0.00001 'consider gaps larger time differences than this
For i = 4 To last - 1
    If Cells(i, 1) > 0 Then
        If Cells(i + 1, 1) - Cells(i, 1) > diff Then
            If index < maxgap Then
                index = index + 1
                gaps(index) = i + 1
            End If
        Else
            If Cells(i + 1, 1) - Cells(i, 1) < diff - 0.00005 Then
                If index < maxgap Then
                    index = index + 1
                    gaps(index) = i + 1
                End If
            End If
        End If
    End If
Next i

'Output row numbers of gaps to column outcol
If index > 0 Then
    With Cells(3, outcol)
        .Value = "Gaps"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    For i = 1 To index
        Cells(i + 3, outcol) = gaps(i)
    Next i
End If

'Show errors and processing messages
If index = maxgap Then
    i = MsgBox("More than " & maxgap & " gaps detected!", vbExclamation, "Maximum exceeded")
Else
    i = MsgBox("Processing " & last & " rows completed without errors; " & IIf(index = 0, "no", index) & " gap" & IIf(index <> 1, "s", "") & " detected.", vbokayonly, "Processing complete")
End If

End Sub

