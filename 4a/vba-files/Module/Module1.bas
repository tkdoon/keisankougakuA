Attribute VB_Name = "Module1"
Sub circle1_Click()
Dim A_Nodes(), A66(), f(), Iz(), k2() As Variant
    Nodes = 6
    E = Cells(1, 2)
    H0 = Cells(2, 2)
    H1 = Cells(3, 2)
    b = Cells(4, 2)
    L = Cells(5, 2)
    ReDim f(Nodes * 3)
    For i = 1 To Nodes * 3
        f(i) = Cells(i + 1, 10)
    Next i
    l_ = L / (Nodes - 1)
    ReDim A_Nodes(Nodes * 3, Nodes * 3 + 1)
    For i = 1 To Nodes * 3
        For j = 1 To Nodes * 3
            A_Nodes(i, j) = 0
        Next j
        A_Nodes(i, Nodes * 3 + 1) = f(i)
    Next i
    ReDim Iz(Nodes - 1)
    ReDim k2(Nodes - 1)
    For i = 1 To Nodes - 1
        Iz(i) = b * (H0 + (1 - 2 * i) * (H0 - H1) / 10) ^ 3 / 12
        k2(i) = E * b / l_ / 2 * (2 * H0 - 2 * i * l_ / L * (H0 - H1) + l_ / L * (H0 - H1))
    Next i
    ReDim A66(Nodes - 1, 6, 6)
    For i = 1 To Nodes - 1
        k = E * Iz(i) / l_ ^ 3
        A66(i, 1, 1) = k2(i)
        A66(i, 1, 2) = 0
        A66(i, 1, 3) = 0
        A66(i, 1, 4) = -k2(i)
        A66(i, 1, 5) = 0
        A66(i, 1, 6) = 0
        A66(i, 2, 1) = 0
        A66(i, 2, 2) = 12 * k
        A66(i, 2, 3) = 6 * k * l_
        A66(i, 2, 4) = 0
        A66(i, 2, 5) = -12 * k
        A66(i, 2, 6) = 6 * k * l_
        A66(i, 3, 1) = 0
        A66(i, 3, 2) = 6 * k * l_
        A66(i, 3, 3) = 4 * k * l_ ^ 2
        A66(i, 3, 4) = 0
        A66(i, 3, 5) = -6 * k * l_
        A66(i, 3, 6) = 2 * k * l_ ^ 2
        A66(i, 4, 1) = -k2(i)
        A66(i, 4, 2) = 0
        A66(i, 4, 3) = 0
        A66(i, 4, 4) = k2(i)
        A66(i, 4, 5) = 0
        A66(i, 4, 6) = 0
        A66(i, 5, 1) = 0
        A66(i, 5, 2) = -12 * k
        A66(i, 5, 3) = -6 * k * l_
        A66(i, 5, 4) = 0
        A66(i, 5, 5) = 12 * k
        A66(i, 5, 6) = -6 * k * l_
        A66(i, 6, 1) = 0
        A66(i, 6, 2) = 6 * k * l_
        A66(i, 6, 3) = 2 * k * l_ ^ 2
        A66(i, 6, 4) = 0
        A66(i, 6, 5) = -6 * k * l_
        A66(i, 6, 6) = 4 * k * l_ ^ 2
    Next i

    For i = 1 To Nodes * 3 - 5 Step 3
        For j = 0 To 5
            For k = 0 To 5
                A_Nodes(i + j, i + k) = A_Nodes(i + j, i + k) + A66((i + 2) / 3, j + 1, k + 1)
            Next k
        Next j
    Next i
Call print_array(A_Nodes, "aaaaa")
    key2 = A_Nodes(1, 4)
    key5 = A_Nodes(2, 5)
    key6 = A_Nodes(2, 6)
    key9 = A_Nodes(3, 5)
    key10 = A_Nodes(3, 6)
    A_small_size = ShrinkArray(ShrinkArray(ShrinkArray(A_Nodes, 1, 1), 1, 1), 1, 1)
    hakidasi_matrix = hakidasi(A_small_size, 3 * (Nodes - 1), 3 * Nodes - 2)
    u = answer_of_hakidasi(hakidasi_matrix, 3 * (Nodes - 1), 3 * Nodes - 2)
    For i = 1 To 3 * (Nodes - 1)
        Cells(i + 4, 11) = u(i)
    Next i
    Cells(2, 11) = 0
    Cells(3, 11) = 0
    Cells(4, 11) = 0
    Cells(2, 12) = key2 * u(1)
    Cells(3, 12) = key5 * u(2) + key6 * u(3)
    Cells(4, 12) = key9 * u(2) + key10 * u(3)
    For i = 1 To 3 * Nodes - 3
        Cells(i + 4, 12) = f(i + 3)
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


