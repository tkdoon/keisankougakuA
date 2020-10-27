Attribute VB_Name = "Module1"
Sub circle_Click()
    Dim AF(4, 5), A
    AF(1, 1) = Sheets("Sheet1").Cells(3, 2)
    AF(1, 2) = Sheets("Sheet1").Cells(3, 3)
    AF(1, 3) = Sheets("Sheet1").Cells(3, 5)
    AF(1, 4) = Sheets("Sheet1").Cells(3, 6)
    AF(1, 5) = Sheets("Sheet1").Cells(3, 7)
    AF(2, 1) = Sheets("Sheet1").Cells(4, 2)
    AF(2, 2) = Sheets("Sheet1").Cells(4, 3)
    AF(2, 3) = Sheets("Sheet1").Cells(4, 5)
    AF(2, 4) = Sheets("Sheet1").Cells(4, 6)
    AF(2, 5) = Sheets("Sheet1").Cells(4, 7)
    AF(3, 1) = Sheets("Sheet1").Cells(6, 2)
    AF(3, 2) = Sheets("Sheet1").Cells(6, 3)
    AF(3, 3) = Sheets("Sheet1").Cells(6, 5)
    AF(3, 4) = Sheets("Sheet1").Cells(6, 6)
    AF(3, 5) = Sheets("Sheet1").Cells(6, 7)
    AF(4, 1) = Sheets("Sheet1").Cells(7, 2)
    AF(4, 2) = Sheets("Sheet1").Cells(7, 3)
    AF(4, 3) = Sheets("Sheet1").Cells(7, 5)
    AF(4, 4) = Sheets("Sheet1").Cells(7, 6)
    AF(4, 5) = Sheets("Sheet1").Cells(7, 7)
hakidasi0 = hakidasi(AF, 4, 5)
u = answer_of_hakidasi(hakidasi0, 4, 5)
A = read(2, 1, 6, 6)
Dim b As Double

b = A(2, 2) * u(1) + A(2, 3) * u(2) + A(2, 5) * u(3)
c = A(2, 6) * u(4)
'MsgBox (Round(b, 3) & "  " & Round(c, 4))
Call print_array(A)
debug.print(A(5, 2) * u(1))
Sheets("Sheet1").Cells(2, 7) = A(1, 2) * u(1) + A(1, 3) * u(2) + A(1, 5) * u(3) + A(1, 6) * u(4)
Sheets("Sheet1").Cells(5, 7) = A(4, 2) * u(1) + A(4, 3) * u(2) + A(4, 5) * u(3) + A(4, 6) * u(4)

Sheets("Sheet1").Cells(3, 8) = u(1)
Sheets("Sheet1").Cells(4, 8) = u(2)
Sheets("Sheet1").Cells(6, 8) = u(3)
Sheets("Sheet1").Cells(7, 8) = u(4)

End Sub
Function read(row_origin, col_origin, row_size, col_size)
    Dim AF() As Variant
    ReDim AF(row_size, col_size)
    For i = 1 To row_size
        For j = 1 To col_size
            AF(i, j) = Sheets("Sheet1").Cells(i + row_origin - 1, j + col_origin - 1)
        Next j
    Next i
read = AF
End Function
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
