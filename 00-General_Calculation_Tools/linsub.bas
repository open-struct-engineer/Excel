Attribute VB_Name = "linsub"
Option Explicit

'LU decomposition including partial pivoting
Function lu_decomp(a() As Variant) As Variant()
    Dim n As Long
    Dim m As Long
    Dim pvt As Double
    Dim new_pvt_row As Long
    Dim org_pvt_row As Long
    Dim j As Long
    Dim i As Long
    Dim k As Long
    Dim Aov(0 To 1) As Variant
    Dim nmo As Long
    Dim temp_pvt_row As Long
    Dim temp_pvt_row2 As Long
    
    n = UBound(a, 1)
    m = UBound(a, 2)
    
    nmo = n - 1
    
    Dim ov() As Variant
    ReDim ov(n)
    
    'initialize order vector
    For i = 0 To n
        ov(i) = i
    Next i
    
    If n <> m Then
        Err.Raise 5, "LU Decomposition", "Non-sqaure matrix; can't decompose"
    End If
    
    For j = 0 To nmo
        pvt = Abs(a(ov(j), j))
        new_pvt_row = -1
        org_pvt_row = j
        
        For i = j + 1 To n
            If Abs(a(ov(i), j)) > pvt Then
                pvt = Abs(a(ov(i), j))
                new_pvt_row = i
            End If
        Next
        
        If new_pvt_row <> -1 And org_pvt_row <> new_pvt_row Then
            temp_pvt_row = ov(org_pvt_row)
            temp_pvt_row2 = ov(new_pvt_row)
            ov(org_pvt_row) = temp_pvt_row2
            ov(new_pvt_row) = temp_pvt_row
        End If
        
        For i = j + 1 To n
            a(ov(i), j) = a(ov(i), j) / a(ov(j), j)
        Next
        
        For i = j + 1 To n
            For k = j + 1 To n
                a(ov(i), k) = a(ov(i), k) - a(ov(i), j) * a(ov(j), k)
            Next
        Next
    Next
    
    'Packing up things to return
    'VBA can't return more than one item
    Aov(0) = a
    Aov(1) = ov
    
    lu_decomp = Aov

End Function

Private Function forward_elim(lu() As Variant, ov() As Variant, b() As Variant) As Variant()
    Dim n As Integer
    Dim m As Integer
    Dim j As Integer
    Dim i As Integer
    
    n = UBound(lu, 1)
    m = UBound(lu, 2)
    
    For j = 0 To n
        For i = j + 1 To n
            b(ov(i)) = b(ov(i)) - lu(ov(i), j) * b(ov(j))
        Next
    Next
    
    forward_elim = b

End Function

Private Function back_sub(lu() As Variant, ov() As Variant, b() As Variant) As Variant()
    Dim n As Integer
    Dim m As Integer
    Dim j As Integer
    Dim k As Integer
    Dim i As Integer
    
    n = UBound(lu, 1)
    m = UBound(lu, 2)
    
    Dim x() As Variant
    ReDim x(n)
    
    For i = 0 To n
        x(i) = 0
    Next i
    
    x(ov(n)) = b(ov(n)) / lu(ov(n), n)
    For j = n - 1 To 0 Step -1
        x(ov(j)) = b(ov(j))
        For k = n To j + 1 Step -1
            x(ov(j)) = x(ov(j)) - x(ov(k)) * lu(ov(j), k)
        Next
        x(ov(j)) = x(ov(j)) / lu(ov(j), j)
    Next
    
    back_sub = x

End Function
'Reorders solutions according to order vector
Private Function reorder(x() As Variant, ov() As Variant) As Variant()

    Dim n As Integer
    Dim i As Integer
    
    n = UBound(x, 1)
    
    Dim y() As Variant
    ReDim y(n)
    
    For i = 0 To n
        y(i) = x(ov(i))
    Next
    
    reorder = y
    
End Function

'Performs LU decomp, forward and back substitution returns solution vector
'to a linear set of equations
Function solve(lu() As Variant, ov() As Variant, b() As Variant) As Variant()
    Dim n As Integer
    Dim m As Integer
    Dim xo() As Variant
    
    n = UBound(lu, 1)
    m = UBound(lu, 2)
    
    Dim c() As Variant
    ReDim c(n)
    Dim x() As Variant
    ReDim x(n)
    
    c = forward_elim(lu, ov, b)
    x = back_sub(lu, ov, c)
    xo = reorder(x, ov)
    
    solve = xo

End Function

'Matrix-matrix multiplication
Function matmul(a() As Variant, b() As Variant) As Variant()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim m As Integer
    Dim x As Integer
    Dim y As Integer
    Dim z As Integer
    Dim c() As Variant

    i = UBound(a, 1)
    j = UBound(a, 2)
    k = UBound(b, 1)
    m = UBound(b, 2)
    
    If j <> k Then
        Err.Raise 5, "Matrix Multiplication", "Dims don't match"
    End If
    
    ReDim c(i, m)

    For x = 0 To i
        For y = 0 To m
            For z = 0 To j
                c(x, y) = c(x, y) + a(x, z) * b(z, y)
            Next z
        Next y
    Next x

    matmul = c
        
End Function

'Matrix-vector multiplication
Function matvecmul(a() As Variant, b() As Variant) As Variant()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim x As Integer
    Dim z As Integer
    Dim c() As Variant

    i = UBound(a, 1)
    j = UBound(a, 2)
    k = UBound(b, 1)
    
    If j <> k Then
        Err.Raise 5, "Matrix Vector Multiplication", "Dims don't match"
    End If
    
    ReDim c(i)

    For x = 0 To i
        For z = 0 To j
            c(x) = c(x) + a(x, z) * b(z)
        Next z
    Next x

    matvecmul = c
    
End Function

'Transpose matrix
Function mattran(a() As Variant) As Variant()
    Dim i As Integer
    Dim j As Integer
    Dim x As Integer
    Dim y As Integer
    Dim b As Variant
    
    i = UBound(a, 1)
    j = UBound(a, 2)
    
    ReDim b(j, i)
    
    For x = 0 To i
        For y = 0 To j
            b(y, x) = a(x, y)
        Next y
    Next x
    
    mattran = b
    
End Function

'Matrix-matrix addition
Function matadd(a() As Variant, b() As Variant) As Variant()
    Dim i As Integer
    Dim j As Integer
    Dim am As Long
    Dim an As Long
    Dim bm As Long
    Dim bn As Long
    Dim c() As Variant
    
    am = UBound(a, 1)
    an = UBound(a, 2)
    bm = UBound(b, 1)
    bn = UBound(b, 2)
    
    If am <> bm Or an <> bn Then
        Err.Raise "Matrix Addition", "Matrices not same dimensions"
    End If
    
    ReDim c(am, an)
    
    For i = 0 To am
        For j = 0 To an
            c(i, j) = a(i, j) + b(i, j)
        Next j
    Next i
    
    matadd = c
End Function

'Vector-vector addition
Function vecadd(a() As Variant, b() As Variant) As Variant()
    Dim i As Integer
    Dim am As Long
    Dim bm As Long
    Dim c() As Variant
    
    am = UBound(a, 1)
    bm = UBound(b, 1)
    
    If am <> bm Then
        Err.Raise "Vector Addition", "Vectors not same dimensions"
    End If
    
    ReDim c(am)
    
    For i = 0 To am
            c(i) = a(i) + b(i)
    Next i
    
    vecadd = c
End Function

