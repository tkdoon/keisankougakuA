Attribute VB_Name = "Module2"
Sub junikakkei2_Click()
    Dim b, D, Bt, BtD, BtDB As Variant
    
    b = read_matrix_from_sheet(2, 2, 3, 6)
    D = read_matrix_from_sheet(2, 8, 3, 3)
    Bt = matrix_t(b)
    Dim x As Integer
    For x = 6 To 11
        Sheets("Sheet1").Range("B" & x, "D" & x) = Bt(x - 6)
    Next x
        BtD = matrix_cross(Bt, D)
        BtDB = matrix_cross(BtD, b)
    Dim y As Integer
    For y = 6 To 11
        Sheets("Sheet1").Range("E" & y, "J" & y) = BtDB(y - 6)
    Next y
End Sub
Function read_matrix_from_sheet(row_origin, col_origin, row_size, col_size)
    ' シートから原点と縦横のサイズを指定し、2次元状にデータを取得する
    ' Range関数の代わりに使える。

    Dim b As Variant
    b = create_matrix(row_size, col_size)
    For i = 0 To UBound(b)
        For j = 0 To UBound(b(0))
            b(i)(j) = Cells(row_origin + i, col_origin + j)
            
        Next j
    Next i
    read_matrix_from_sheet = b

End Function
Function create_matrix(row_size, col_size)
    ' 任意サイズの行列を作成する
    Dim ans, row As Variant
    
    ans = Array()
    ReDim ans(row_size - 1)
    For i = 0 To row_size - 1
        row = Array() ' 新しいオブジェクトのインスタンスが代入される
        ReDim row(col_size - 1)
        ans(i) = row
    Next i
    
    create_matrix = ans
End Function
Function matrix_t(m)
  ' 行列の転置
  ans = create_matrix(UBound(m(0)) + 1, UBound(m) + 1)
  For i = 0 To UBound(ans)
    For j = 0 To UBound(ans(0))
      ans(i)(j) = m(j)(i)
    Next j
  Next i
  matrix_t = ans
End Function
Function matrix_cross(m1, m2)
  ' 行列同士の掛け算
  ans = create_matrix(UBound(m1) + 1, UBound(m2(0)) + 1)
  
  For i = 0 To UBound(ans)
    For j = 0 To UBound(ans(0))
      sum_ = 0
      For k = 0 To UBound(m1(0))
        sum_ = sum_ + m1(i)(k) * m2(k)(j)
      Next k
      ans(i)(j) = sum_
    Next j
  Next i
  matrix_cross = ans
End Function
