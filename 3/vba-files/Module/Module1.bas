Attribute VB_Name = "Module1"
Sub circle_Click()
    Dim A_Nodes(), A22(), f(), k() As Variant
    Nodes = Cells(1, 2)
    E = Cells(2, 2)
    H0 = Cells(3, 2)
    H1 = Cells(4, 2)
    b = Cells(5, 2)
    L = Cells(6, 2)
    P = Cells(7, 2)
    ReDim f(Nodes)
    For i = 1 To Nodes
        f(i) = Cells(i + 1, 6)
    Next i
    l_ = L / (Nodes - 1)

    ReDim A_Nodes(Nodes, Nodes + 1)
    For i = 1 To Nodes
        For j = 1 To Nodes
            A_Nodes(i, j) = 0
        Next j
        A_Nodes(i, Nodes + 1) = f(i)
    Next i
    ReDim k(Nodes - 1)
    For i = 1 To Nodes - 1
        k(i) = E * b / l_ / 2 * (2 * H0 - 2 * i * l_ / L * (H0 - H1) + l_ / L * (H0 - H1))
        Debug.Print (k(i))
        
    Next i
    ReDim A22(Nodes - 1, 2, 2)
    For i = 1 To Nodes - 1
        A22(i, 1, 1) = k(i)
        A22(i, 2, 1) = -k(i)
        A22(i, 1, 2) = -k(i)
        A22(i, 2, 2) = k(i)
    Next i
    For i = 1 To Nodes - 1
        A_Nodes(i, i) = A_Nodes(i, i) + A22(i, 1, 1)
        A_Nodes(i + 1, i) = A_Nodes(i + 1, i) + A22(i, 2, 1)
        A_Nodes(i, i + 1) = A_Nodes(i, i + 1) + A22(i, 1, 2)
        A_Nodes(i + 1, i + 1) = A_Nodes(i + 1, i + 1) + A22(i, 2, 2)
    Next i
    Call print_array(A_Nodes, "")
    A_small_size = ShrinkArray(A_Nodes, 1, 1)
    Call print_array(A_small_size, "")
    hakidasi_matrix = hakidasi(A_small_size, Nodes - 1, Nodes)
    u = answer_of_hakidasi(hakidasi_matrix, Nodes - 1, Nodes)
    For i = 1 To Nodes - 1
        Cells(i + 2, 7) = u(i)
    Next i
End Sub
Function hakidasi(AF, row_size, col_size)
'colsize=rowsize+1

    For k = 1 To row_size
        '(i,i)???
        key = AF(k, k)
        For i = k To col_size
            AF(k, i) = AF(k, i) / key
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
    Call print_array(AF, "aaa")

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
