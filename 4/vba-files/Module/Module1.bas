Attribute VB_Name = "Module1"
Sub circle1_Click()
    Dim A_Nodes(), A44(), f(), Iz() As Variant
    Nodes = 6
    E = Cells(1, 2)
    H0 = Cells(2, 2)
    H1 = Cells(3, 2)
    b = Cells(4, 2)
    L = Cells(5, 2)
    ReDim f(Nodes * 2)
    For i = 1 To Nodes * 2
        f(i) = Cells(i + 1, 10)
    Next i
    l_ = L / (Nodes - 1)
    ReDim A_Nodes(Nodes * 2, Nodes * 2 + 1)
    For i = 1 To Nodes * 2
        For j = 1 To Nodes * 2
            A_Nodes(i, j) = 0
        Next j
        A_Nodes(i, Nodes * 2 + 1) = f(i)
    Next i
    ReDim Iz(Nodes - 1)
    For i = 1 To Nodes - 1
        Iz(i) = b * (H0 + (1 - 2 * i) * (H0 - H1) / 10) ^ 3 / 12
    Next i
    ReDim A44(Nodes - 1, 4, 4)
    For i = 1 To Nodes - 1
        k = E * Iz(i) / l_ ^ 3
        A44(i, 1, 1) = 12 * k
        A44(i, 1, 2) = 6 * k * l_
        A44(i, 1, 3) = -12 * k
        A44(i, 1, 4) = 6 * k * l_
        A44(i, 2, 1) = 6 * k * l_
        A44(i, 2, 2) = 4 * k * l_ ^ 2
        A44(i, 2, 3) = -6 * k * l_
        A44(i, 2, 4) = 2 * k * l_ ^ 2
        A44(i, 3, 1) = -12 * k
        A44(i, 3, 2) = -6 * k * l_
        A44(i, 3, 3) = 12 * k
        A44(i, 3, 4) = -6 * k * l_
        A44(i, 4, 1) = 6 * k * l_
        A44(i, 4, 2) = 2 * k * l_ ^ 2
        A44(i, 4, 3) = -6 * k * l_
        A44(i, 4, 4) = 4 * k * l_ ^ 2
    Next i

    For i = 1 To Nodes * 2 - 3 Step 2
        For j = 0 To 3
            For k = 0 To 3
                A_Nodes(i + j, i + k) = A_Nodes(i + j, i + k) + A44((i + 1) / 2, j + 1, k + 1)
            Next k
        Next j
    Next i

    key1 = A_Nodes(1, 3)
    key2 = A_Nodes(1, 4)
    key3 = A_Nodes(2, 3)
    key4 = A_Nodes(2, 4)
    A_small_size = ShrinkArray(ShrinkArray(A_Nodes, 1, 1), 1, 1)
    hakidasi_matrix = hakidasi(A_small_size, 2 * (Nodes - 1), 2 * Nodes - 1)
    u = answer_of_hakidasi(hakidasi_matrix, 2 * (Nodes - 1), 2 * Nodes - 1)
    For i = 1 To 2 * (Nodes - 1)
        Cells(i + 3, 11) = u(i)
    Next i
    Cells(2, 11) = 0
    Cells(3, 11) = 0
    Cells(2, 12) = key1 * u(1) + key2 * u(2)
    Cells(3, 12) = key3 * u(1) + key4 * u(2)
    For i = 1 To 2 * Nodes - 2
        Cells(i + 3, 12) = f(i + 2)
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
