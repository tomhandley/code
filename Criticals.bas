Attribute VB_Name = "Criticals"
Sub Criticals_alone() 'seconds As Integer
'
' Identify critical points in elevation and interpolate linearly
' between critical points. Points are stored in crit() with values
' in yc(), which can be passed as arguments to a more sophisticated
' interpolation technique such as splining.
'
Dim n As Integer

n = 0 'Starting with 0 excludes header row
Do While Cells(n + 2, 1).Value > 0
    n = n + 1
Loop

Dim i, j, k, maxi, maxj As Integer
Dim zmin(), zmax() As Single
ReDim zmin(1 To n)
ReDim zmax(1 To n)
Dim z() As Single
Dim crit() As Integer
ReDim z(1 To n)    'stores values of critical points
ReDim crit(1 To n)  'stores indices of critical points
Dim critChange As Boolean
Dim m, b As Single

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
        zmin(j) = Cells(j + 1, 18).Value
        zmax(j) = Cells(j + 1, 19).Value
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
Cells(2, 17).Value = z(1)  'Initial smoothed value

For i = 2 To n
    zmin(i) = Cells(i + 1, 18).Value
    zmax(i) = Cells(i + 1, 19).Value
    critChange = False
    
    'Check if last critical value exceed current max or min
    If z(crit(k - 1)) > zmax(i) Then
        'Reset k for two consecutive decreasing critical values
        If crit(k - 1) = i - 1 And (z(i - 1) = zmax(i - 1) And i > maxi) Then
            k = k - 1
        End If
        z(i) = zmax(i)
        crit(k) = i
        critChange = True
    Else
        If z(crit(k - 1)) < zmin(i) Then
            'Reset k for two consecutive increasing critical values
            If crit(k - 1) = i - 1 And (z(i - 1) = zmin(i - 1) And i > maxi) Then
                k = k - 1
            End If
            z(i) = zmin(i)
            crit(k) = i
            critChange = True
        Else
            'Extrapolate for tail values after last critical point
            'only if z(n) is not already a critical point
            If i = n Then
                z(i) = m * i + b
                crit(k) = n
                critChange = True
            End If
        End If
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
            Else
                If z(j) < zmin(j) Then
                    z(j) = zmin(j)
                    crit(k) = j
                    maxi = i
                    i = j
                    j = crit(k - 1)
                    m = (z(i) - z(j)) / (crit(k) - j)
                    b = z(i) - m * crit(k)
                    'need a marker to say whether this condition was true
                End If
            End If
            Cells(j + 1, 17).Value = z(j)
        Next j
        k = k + 1
    End If
    Cells(i + 1, 17).Value = z(i)
Next i

'Write critical values to worksheet
'ReDim Preserve crit(1 To k - 1)
'Dim zc() As Single
'ReDim zc(1 To k - 1)
'For i = 1 To k - 1
'    zc(i) = z(crit(i))
'    Cells(i + 2, 11).Value = crit(i)
'    Cells(i + 2, 12).Value = zc(i)
'Next i

'For i = 1 To n
'    Cells(i + 1, 10).Value = z(i)
'Next i

End Sub

