Attribute VB_Name = "Module1"
Sub click_circle()
    Dim u
    u= read_to_calculate_answer(2, 2, 6, 7)
    For i=1 to 6
        Sheets("Sheet1").Cells(i+1,9)=u(i)
    next i
End Sub
'-------------------------------------------------------------------
Function read(row_origin, col_origin, row_size, col_size)
    Dim AF() As variant
    ReDim AF(row_size, col_size)
    For i = 1 To row_size
        For j = 1 To col_size
            AF(i, j) = Sheets("Sheet1").Cells(i + row_origin - 1, j + col_origin - 1)
        Next j
    Next i
read = AF
End Function
'---------------------------------------------------------------------
function hakidasi(AF,row_size,col_size)
'colsize=rowsize+1

    For k = 1 To row_size
        '(i,i)を１に
        key=AF(k,k)
        For i = k To col_size
            AF(k, i) = AF(k, i) / key
        Next i
        '1の左を0に
        For i = k + 1 To row_size
            num=AF(i,k)
            For j = k To col_size
                AF(i, j) = AF(i, j) - AF(k, j) * num
            Next j
        Next i
    Next k

    '右側も0に
    for i=1 to row_size-1
        for k=i+1 to row_size
        num=AF(i,k)
            for j=i+1 to col_size
                AF(i,j)=AF(i,j)-AF(k,j)*num
            next j
        next k
    next i
    call print_array(AF, "aaa")

    hakidasi=AF
end function

function answer_of_hakidasi(hakidasi_AF,row_size,col_size)
    Dim u() as variant
    Redim u(row_size)
    for i=1 to row_size
        u(i)=hakidasi_AF(i,col_size)
    next i
    answer_of_hakidasi=u
end function
'----------------------------------------------------------------
function read_to_calculate_answer(row_origin, col_origin, row_size, col_size)
    AF=read(row_origin, col_origin, row_size, col_size)
    hakidasi_matrix=hakidasi(AF,row_size, col_size)
    read_to_calculate_answer=answer_of_hakidasi(hakidasi_matrix,row_size,col_size)
end function

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