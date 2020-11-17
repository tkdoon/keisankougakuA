Attribute VB_Name = "Module1"
Sub Click()
    Dim Main(), B(), Bt(), D(), x(), y(), setsu(), f(), f_output(), u_() As Variant
    Nodes = 22
    ReDim B(3, 6)
    ' h0=20
    ' half_h0=h0/2
    ' L=200
    element = 20
    E = Cells(7, 2)
    nu = Cells(8, 2)
    t = Cells(9, 2)
    ReDim f(Nodes * 2)
    For i = 1 To Nodes * 2
        f(i) = Cells(i + 1, 11)
    Next i
    ReDim x(Nodes)
    ReDim y(Nodes)
    x1 = Cells(2, 2)
    x2 = Cells(3, 2)
    x3 = Cells(4, 2)
    x4 = Cells(5, 2)
    y1 = Cells(2, 3)
    y2 = Cells(3, 3)
    y3 = Cells(4, 3)
    y4 = Cells(5, 3)
    For i = 1 To Nodes / 2
        x(i) = (x2 + x1) / (Nodes / 2 - 1) * (i - 1)
        x(i + 11) = (x2 + x1) / (Nodes / 2 - 1) * (i - 1)
        y(i + 11) = 10 - (y2 - y1) / (Nodes / 2 - 1) * (i - 1)
        y(i) = -10 + (y3 - y4) / (Nodes / 2 - 1) * (i - 1)
        Debug.Print (x(i) & y(i))
        Debug.Print (x(i + 11) & y(i + 11))
    Next i
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
    setsu = Array(0, Array(0, 1, 2, 13), Array(0, 2, 3, 14), Array(0, 3, 4, 15), Array(0, 4, 5, 16), Array(0, 5, 6, 17), Array(0, 6, 7, 18), Array(0, 7, 8, 19), Array(0, 8, 9, 20), Array(0, 9, 10, 21), Array(0, 10, 11, 22), Array(0, 1, 13, 12), Array(0, 2, 14, 13), Array(0, 3, 15, 14), Array(0, 4, 16, 15), Array(0, 5, 17, 16), Array(0, 6, 18, 17), Array(0, 7, 19, 18), Array(0, 8, 20, 19), Array(0, 9, 21, 20), Array(0, 10, 22, 21))

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
            For j = 1 To 6
                B(i, j) = B(i, j) * E * t / 4 / A / (1 - nu * nu)
            Next j
        Next i

        k = matrix_cross(matrix_cross(Bt, D), B)

        For i = 1 To 3
            For j = 1 To 3
                Main(setsu(l)(i), setsu(l)(j)) = Main(setsu(l)(i), setsu(l)(j)) + k(i, j)
                Main(setsu(l)(i) + 22, setsu(l)(j)) = Main(setsu(l)(i) + 22, setsu(l)(j)) + k(i + 3, j)
                Main(setsu(l)(i) + 22, setsu(l)(j) + 22) = Main(setsu(l)(i) + 22, setsu(l)(j) + 22) + k(i + 3, j + 3)
                Main(setsu(l)(i), setsu(l)(j) + 22) = Main(setsu(l)(i), setsu(l)(j) + 22) + k(i, j + 3)
            Next j
        Next i
    Next l
Call print_array(Main)
    deep_main = ShrinkArray(Main)
    small_matrix = ShrinkArray(ShrinkArray(ShrinkArray(Main, 23, 23), 12, 12), 1, 1)

    hakidasi_matrix = hakidasi(small_matrix, Nodes * 2 - 3, Nodes * 2 - 2)
    u = answer_of_hakidasi(hakidasi_matrix, Nodes * 2 - 3, Nodes * 2 - 2)
    ReDim u_(Nodes * 2)
    u_(1) = 0
    u_(2) = u(1)
    u_(3) = u(2)
    u_(4) = u(3)
    u_(5) = u(4)
    u_(6) = u(5)
    u_(7) = u(6)
    u_(8) = u(7)
    u_(9) = u(8)
    u_(10) = u(9)
    u_(11) = u(10)
    u_(12) = 0
    u_(13) = u(11)
    u_(14) = u(12)
    u_(15) = u(13)
    u_(16) = u(14)
    u_(17) = u(15)
    u_(18) = u(16)
    u_(19) = u(17)
    u_(20) = u(18)
    u_(21) = u(19)
    u_(22) = u(20)
    u_(23) = 0
    u_(24) = u(21)
    u_(25) = u(22)
    u_(26) = u(23)
    u_(27) = u(24)
    u_(28) = u(25)
    u_(29) = u(26)
    u_(30) = u(27)
    u_(31) = u(28)
    u_(32) = u(29)
    u_(33) = u(30)
    u_(34) = u(31)
    u_(35) = u(32)
    u_(36) = u(33)
    u_(37) = u(34)
    u_(38) = u(35)
    u_(39) = u(36)
    u_(40) = u(37)
    u_(41) = u(38)
    u_(42) = u(39)
    u_(43) = u(40)
    u_(44) = u(41)

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

