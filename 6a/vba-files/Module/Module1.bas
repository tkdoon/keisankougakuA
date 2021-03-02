Attribute VB_Name = "Module1"
Sub circle_Click()
    Dim Main(), B(), Bt(), D(), x(), y(), setsu(), f(), f_output(), u_(), fix_u(4) As Variant
    Nodes = 5
    For i = 1 To 4
        fix_u(i) = Cells(i + 1, 9)
    Next i
    ReDim B(3, Nodes)
    ' h0=20
    ' half_h0=h0/2
    ' L=200
    element = 6
    E = Cells(7, 2)
    nu = Cells(8, 2)
    r = Cells(9, 2)

    ReDim f(Nodes * 2)
    For i = 1 To Nodes * 2
        f(i) = Cells(i + 1, 11)
    Next i
    ReDim x(Nodes)
    ReDim y(Nodes)
    x(1) = Cells(2, 2)
    x(2) = Cells(3, 2)
    x(3) = Cells(4, 2)
    x(4) = Cells(5, 2)
    x(5)=400
    y(1) = Cells(2, 3)
    y(2) = Cells(3, 3)
    y(3) = Cells(4, 3)
    y(4) = Cells(5, 3)
    y(5)=0

    ReDim Main(Nodes * 2, Nodes * 2 + 1)
    For i = 1 To Nodes * 2
        For j = 1 To Nodes * 2 + 1
            Main(i, j) = 0
        Next j
    Next i

    For i = 1 To Nodes * 2
        Main(i, Nodes * 2 + 1) = f(i)
    Next i

    For l = 1 To element
        setsu = Array(0, Array(0, 1, 2), Array(0, 3, 4), Array(0, 2, 4), Array(0, 3, 2),Array(0,2,5),Array(0,4,5))
        x1 = x(setsu(l)(1))
        x2 = x(setsu(l)(2))
        y1 = y(setsu(l)(1))
        y2 = y(setsu(l)(2))
        k_matrix = k_matrix_factory(E, r, x1, x2, y1, y2)
        For i = 1 To 2
            For j = 1 To 2
                Main(2 * setsu(l)(i) - 1, 2 * setsu(l)(j) - 1) = Main(2 * setsu(l)(i) - 1, 2 * setsu(l)(j) - 1) + k_matrix(2 * i - 1, 2 * j - 1)
                Main(2 * setsu(l)(i) - 1, 2 * setsu(l)(j)) = Main(2 * setsu(l)(i) - 1, 2 * setsu(l)(j)) + k_matrix(2 * i - 1, 2 * j)
                Main(2 * setsu(l)(i), 2 * setsu(l)(j)) = Main(2 * setsu(l)(i), 2 * setsu(l)(j)) + k_matrix(2 * i, 2 * j)
                Main(2 * setsu(l)(i), 2 * setsu(l)(j) - 1) = Main(2 * setsu(l)(i), 2 * setsu(l)(j) - 1) + k_matrix(2 * i, 2 * j - 1)
            Next j
        Next i
    Next l
    
Call print_array(Main)
    deep_main = ShrinkArray(Main)
    small_matrix = ShrinkArray(ShrinkArray(ShrinkArray(ShrinkArray(Main, fix_u(4), fix_u(4)), fix_u(3), fix_u(3)), fix_u(2), fix_u(2)), fix_u(1), fix_u(1))

    hakidasi_matrix = hakidasi(small_matrix, Nodes * 2 - 4, Nodes * 2 - 3)
    u = answer_of_hakidasi(hakidasi_matrix, Nodes * 2 - 4, Nodes * 2 - 3)
    ReDim u_(Nodes * 2)
    For i = 1 To 4
        u_(fix_u(i)) = 0
    Next i
    For j = 1 To fix_u(1) - 1
        u_(j) = u(j)
    Next j
    For i = 1 To 3
        For j = fix_u(i) + 1 To fix_u(i + 1) - 1
            u_(j) = u(j - i)
        Next j
    Next i
    For j = fix_u(4) + 1 To Nodes * 2
        u_(j) = u(j - 4)
    Next j
For i = 1 To 8
 Debug.Print (u_(i))
Next i
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

Function ShrinkArray(arr, Optional row = -1, Optional col = -1)
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

Function k_matrix_factory(E, r, x1, x2, y1, y2)
    l = sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    s = (y2 - y1) / l
    c = (x2 - x1) / l
    pi = 4 * atn(1)
    k = E * r ^ 2 * pi / l
    Dim ans(4, 4)
    ans(1, 1) = c ^ 2
    ans(1, 2) = c * s
    ans(1, 3) = -c ^ 2
    ans(1, 4) = -c * s
    ans(2, 1) = c * s
    ans(2, 2) = s ^ 2
    ans(2, 3) = -c * s
    ans(2, 4) = -s ^ 2
    ans(3, 1) = -c ^ 2
    ans(3, 2) = -c * s
    ans(3, 3) = c ^ 2
    ans(3, 4) = c * s
    ans(4, 1) = -c * s
    ans(4, 2) = -s ^ 2
    ans(4, 3) = c * s
    ans(4, 4) = s ^ 2
    For i = 1 To 4
        For j = 1 To 4
            ans(i, j) = ans(i, j) * k
        Next j
    Next i
    k_matrix_factory = ans
End Function

