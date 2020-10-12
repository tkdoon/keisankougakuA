Attribute VB_Name = "Module1"

Sub 正方形長方形1_Click()
    Dim A, B, AB As Variant
    
    p = Worksheets("Sheet1").Cells(1, 2)
    q = Worksheets("Sheet1").Cells(1, 3)
    r = Worksheets("Sheet2").Cells(1, 3)
    
    A = read_matrix_from_sheet(2, 1, p, q, 1)
    B = read_matrix_from_sheet(2, 1, q, r, 2)
    AB = matrix_cross(A, B)
    Dim x As Integer
     Dim y As Integer
    For x = 1 To p
        For y = 1 To r
        Worksheets("Sheet3").Cells(1 + x, y) = AB(x - 1)(y - 1)
        Next y
    Next x

End Sub
Function read_matrix_from_sheet(row_origin, col_origin, row_size, col_size, num)
    ' シートから原点と縦横のサイズを指定し、2次元状にデータを取得する
    ' Range関数の代わりに使える。

    Dim A(), B As Variant
    B = create_matrix(row_size, col_size)
    For i = 0 To UBound(B)
        For j = 0 To UBound(B(0))
            B(i)(j) = Worksheets("Sheet" & num).Cells(row_origin + i, col_origin + j)
        Next j
    Next i
    read_matrix_from_sheet = B

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

