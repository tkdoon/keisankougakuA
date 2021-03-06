Attribute VB_Name = "Module1"
Sub circle1_Click()
    Dim Main(), B(), Bt(), D(), x(), y(), setsu(), f(), f_output(), u_() As Variant
    Nodes = 6
    ReDim B(3, Nodes)
    ' h0=20
    ' half_h0=h0/2
    ' L=200
    element = 4
    E = Cells(7, 2)
    nu = Cells(8, 2)
    t = Cells(9, 2)
    ReDim f(Nodes * 2)
    For i = 1 To Nodes * 2
        f(i) = Cells(i + 1, 11)
    Next i
    ReDim x(Nodes)
    ReDim y(Nodes)
    x(1) = Cells(2, 2)
    x(3) = Cells(3, 2)
    x(4) = Cells(4, 2)
    x(6) = Cells(5, 2)
    x(2) = (x(1) + x(3)) / 2
    x(5) = (x(4) + x(6)) / 2
    y(1) = Cells(2, 3)
    y(3) = Cells(3, 3)
    y(4) = Cells(4, 3)
    y(6) = Cells(5, 3)
    y(2) = (y(1) + y(3)) / 2
    y(5) = (y(4) + y(6)) / 2

    ReDim Main(Nodes * 2, Nodes * 2 + 1)
    For i = 1 To Nodes * 2
        For j = 1 To Nodes * 2 + 1
            Main(i, j) = 0
        Next j
    Next i
    ReDim D(3, 3)
    D(1, 1) = 1
    D(1, 2) = nu
    D(1, 3) = 0
    D(2, 1) = nu
    D(2, 2) = 1
    D(2, 3) = 0
    D(3, 1) = 0
    D(3, 2) = 0
    D(3, 3) = (1 - nu) / 2
    setsu = Array(0, Array(0, 1, 2, 5), Array(0, 2, 3, 6), Array(0, 1, 5, 4), Array(0, 2, 6, 5))

    For i = 1 To Nodes * 2
        Main(i, Nodes * 2 + 1) = f(i)
    Next i
    For l = 1 To element
        x1 = x(setsu(l)(1))

        x2 = x(setsu(l)(2))
            
        x3 = x(setsu(l)(3))

        y1 = y(setsu(l)(1))
        y2 = y(setsu(l)(2))
        y3 = y(setsu(l)(3))
                 
        A = (x2 * y3 + x1 * y2 + x3 * y1 - x1 * y3 - y2 * x3 - y1 * x2) / 2
        beta1 = y2 - y3
        beta2 = y3 - y1
        beta3 = y1 - y2
        gamma1 = x3 - x2
        gamma2 = x1 - x3
        gamma3 = x2 - x1
        B(1, 1) = beta1
        B(1, 2) = beta2
        B(1, 3) = beta3
        B(1, 4) = 0
        B(1, 5) = 0
        B(1, 6) = 0
        B(2, 1) = 0
        B(2, 2) = 0
        B(2, 3) = 0
        B(2, 4) = gamma1
        B(2, 5) = gamma2
        B(2, 6) = gamma3
        B(3, 1) = gamma1
        B(3, 2) = gamma2
        B(3, 3) = gamma3
        B(3, 4) = beta1
        B(3, 5) = beta2
        B(3, 6) = beta3

        Bt = matrix_t(B)
        For i = 1 To 3
            For j = 1 To Nodes
                B(i, j) = B(i, j) * E * t / 4 / A / (1 - nu * nu)
            Next j
        Next i

        k = matrix_cross(matrix_cross(Bt, D), B)

        For i = 1 To 3
            For j = 1 To 3
                Main(setsu(l)(i), setsu(l)(j)) = Main(setsu(l)(i), setsu(l)(j)) + k(i, j)
                Main(setsu(l)(i) + 6, setsu(l)(j)) = Main(setsu(l)(i) + 6, setsu(l)(j)) + k(i + 3, j)
                Main(setsu(l)(i) + 6, setsu(l)(j) + 6) = Main(setsu(l)(i) + 6, setsu(l)(j) + 6) + k(i + 3, j + 3)
                Main(setsu(l)(i), setsu(l)(j) + 6) = Main(setsu(l)(i), setsu(l)(j) + 6) + k(i, j + 3)
            Next j
        Next i
    Next l
Call print_array(Main)
    deep_main = ShrinkArray(Main)
    small_matrix = ShrinkArray(ShrinkArray(ShrinkArray(Main, 7, 7), 4, 4), 1, 1)

    hakidasi_matrix = hakidasi(small_matrix, Nodes * 2 - 3, Nodes * 2 - 2)
    u = answer_of_hakidasi(hakidasi_matrix, Nodes * 2 - 3, Nodes * 2 - 2)
    ReDim u_(Nodes * 2)
    u_(1) = 0
    u_(2) = u(1)
    u_(3) = u(2)
    u_(4) = 0
    u_(5) = u(3)
    u_(6) = u(4)
    u_(7) = 0
    u_(8) = u(5)
    u_(9) = u(6)
    u_(10) = u(7)
    u_(11) = u(8)
    u_(12) = u(9)

Call print_array(deep_main)

    ReDim f_output(Nodes * 2)
    For i = 1 To Nodes * 2
    f_output(i) = 0
        For j = 1 To Nodes * 2
            f_output(i) = f_output(i) + deep_main(i, j) * u_(j)
        Next j
    Next i

    For i = 1 To Nodes * 2
        Cells(1 + i, 12) = u_(i)
        Cells(1 + i, 13) = f_output(i)
    Next i


End Sub
Function hakidasi(AF, row_size, col_size)
'colsize=rowsize+1

    For k = 1 To row_size
        '(i,i)???
        Key = AF(k, k)
        For i = k To col_size
            AF(k, i) = AF(k, i) / Key
        Next i
        '1???0?
        For i = k + 1 To row_size
            num = AF(i, k)
            For j = k To col_size
                AF(i, j) = AF(i, j) - AF(k, j) * num
            Next j
        Next i
    Next k
    
    '???0?
    For i = 1 To row_size - 1
        For k = i + 1 To row_size
        num = AF(i, k)
            For j = i + 1 To col_size
                AF(i, j) = AF(i, j) - AF(k, j) * num
            Next j
        Next k
    Next i


    hakidasi = AF
End Function

Function answer_of_hakidasi(hakidasi_AF, row_size, col_size)
    Dim u() As Variant
    ReDim u(row_size)
    For i = 1 To row_size
        u(i) = hakidasi_AF(i, col_size)
    Next i
    answer_of_hakidasi = u
End Function

Function ShrinkArray(arr, Optional row As Integer = -1, Optional col As Integer = -1)
   'row???
   If row >= 0 Then
        For i = 0 To UBound(arr, 2)
            For j = row To UBound(arr) - 1
                arr(j, i) = arr(j + 1, i)
            Next j
        Next i
        For i = 0 To UBound(arr, 2)
            arr(UBound(arr), i) = ""
        Next i
    End If
    ' col ???
    If col >= 0 Then
        For i = 0 To UBound(arr)
            For j = col To UBound(arr, 2) - 1
                arr(i, j) = arr(i, j + 1)
            Next j
        Next i
        For i = 0 To UBound(arr)
            arr(i, UBound(arr, 2)) = ""
        Next i
    End If
    ShrinkArray = arr
End Function

Sub print_array(arr, Optional msg As String)
    Debug.Print ("--")
    For i = 0 To UBound(arr)
        Dim tmp
        tmp = ""
        For j = 0 To UBound(arr, 2)
            tmp = tmp & arr(i, j) & "  "
        Next j
        Debug.Print (tmp)
    Next i
    Debug.Print ("--- " & Now & " " & msg & " ---")
End Sub

Function matrix_t(m)
  ' ?s?????]?u
  row = UBound(m)
  col = UBound(m, 2)
  Dim ans()
  ReDim ans(col, row)
  For i = 1 To row
    For j = 1 To col
      ans(j, i) = m(i, j)
    Next j
  Next i
  matrix_t = ans
End Function

Function matrix_cross(m1, m2)
  ' ?s?????m???|???Z
  row1 = UBound(m1)
  col1 = UBound(m1, 2)
  row2 = UBound(m2)
  col2 = UBound(m2, 2)
  Dim ans()
  ReDim ans(row1, col2)
  For i = 1 To row1
    For j = 1 To col2
      sum_ = 0
      For k = 1 To col1
        sum_ = sum_ + m1(i, k) * m2(k, j)
      Next k
      ans(i, j) = sum_
    Next j
  Next i
  matrix_cross = ans

End Function

