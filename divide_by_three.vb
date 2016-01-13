Option Explicit
 
Public Function divide_by_three() As Boolean


    Dim arr_static_array(6)         As Variant
    Dim arr_temp_array()            As Variant
    Dim arr_current                 As Variant
    Dim arr_elements                As Variant
    Dim arr_current_copy            As Variant

    Dim l_target                    As Long
    Dim l_current_sum               As Long
    Dim l_counter                   As Long
    Dim l_max                       As Long
    
    'Required output:
    'T
    'F
    'T
    'T
    'F
    'F

    arr_temp_array = Array(1, 3, 4, 5, 3, 2)
    arr_static_array(0) = arr_temp_array

    arr_temp_array = Array(4, 2, 5, 8, 3)
    arr_static_array(1) = arr_temp_array

    arr_temp_array = Array(5, 1, 7, 4, 3, 6, 1)
    arr_static_array(2) = arr_temp_array
    
    'This breaks the greedy algorithm
    arr_temp_array = Array(4, 5, 2, 5, 3, 4, 2, 5)
    arr_static_array(3) = arr_temp_array

    arr_temp_array = Array(7, 9, 3, 8, 3)
    arr_static_array(4) = arr_temp_array

    arr_temp_array = Array(5, 2, 1, 3, 2, 5)
    arr_static_array(5) = arr_temp_array
    
    arr_temp_array = Array(1, 2, 3, 4, 5, 6)
    arr_static_array(6) = arr_temp_array

    For Each arr_current In arr_static_array
        arr_current_copy = arr_current
        l_target = sum_array(arr_current) / 3

        If (sum_array(arr_current) / 3) / l_target <> 1 Then
            Call print_result(0, arr_current_copy, sum_array(arr_current) / 3)
            GoTo place_to_go
        End If

        'arr_current = bubble_sort(arr_current) 'Sorting

        For l_counter = 0 To 2
        
            l_current_sum = l_target
            
            While (l_current_sum > 0)
                arr_elements = return_array_with_smaller_numbers(arr_current, l_current_sum)
                                    
                If IsArrayEmpty(arr_elements) Then
                    Call print_result(0, arr_current_copy, l_target)
                    GoTo place_to_go
                End If
                
'                If UBound(arr_elements) = 0 Then
'                    Call print_result(0, arr_current, l_target)
'                    GoTo place_to_go
'                End If
                
                l_max = WorksheetFunction.Max(arr_elements)
                Call Decrement(l_current_sum, l_max)

                arr_current = remove_from_array(arr_current, l_max)
                
            Wend
        Next l_counter
        Call print_result(1, arr_current_copy, l_target)
        
place_to_go:
    Next arr_current

End Function

Public Function remove_from_array(ByVal my_array As Variant, ByVal l_to_remove As Long) As Variant

    Dim l_counter           As Long
    Dim arr_result()        As Long
    Dim b_found             As Boolean

    If UBound(my_array) = 0 Then
        Exit Function
    End If
    ReDim arr_result(UBound(my_array) - 1)
    
    For l_counter = LBound(my_array) To UBound(my_array)
    
        
        If (my_array(l_counter) = l_to_remove And Not b_found) Then
            b_found = True
        Else
            If b_found Then
                arr_result(l_counter - 1) = my_array(l_counter)
            Else
                arr_result(l_counter) = my_array(l_counter)
            End If
        End If
        
    Next l_counter
    
    remove_from_array = arr_result
    'Call print_array(arr_result)
End Function

Public Sub Decrement(ByRef value_to_decrement, Optional l_minus As Long = 1)
    
    value_to_decrement = value_to_decrement - l_minus

End Sub


Public Function return_array_with_smaller_numbers(ByRef my_array As Variant, ByRef l_current_sum As Long) As Variant
    
    Dim my_array_result()   As Variant
    Dim l_counter           As Long
    Dim l_counter_2         As Long
    
    For l_counter = LBound(my_array) To UBound(my_array)
        If my_array(l_counter) <= l_current_sum Then
            ReDim Preserve my_array_result(l_counter_2)
            my_array_result(l_counter_2) = my_array(l_counter)
            Call Increment(l_counter_2)
        End If
    Next l_counter
    
    'Call print_array(my_array_result)
    return_array_with_smaller_numbers = my_array_result
    
End Function

Public Sub Increment(ByRef value_to_increment, Optional l_plus As Long = 1)

       value_to_increment = value_to_increment + l_plus

End Sub

Public Sub print_result(ByVal b_result As Boolean, ByVal my_array As Variant, ByVal d_target As Double)
    
    Debug.Print (CBool(b_result)); " -> "; print_array_one_line(my_array); " avg-> "; d_target
    
End Sub

Public Function sum_array(my_array As Variant, Optional last_values_not_to_calculate As Long = 0) As Double
'For unknown reasons, WorksheetFunction.sum(my_array) does not work always,
'when we sum currency, long and double...

    Dim l_counter As Long

    For l_counter = LBound(my_array) To UBound(my_array) - last_values_not_to_calculate
        sum_array = sum_array + my_array(l_counter)
    Next l_counter

End Function

Public Function print_array_one_line(my_array As Variant) As String

    Dim counter     As Integer
    Dim s_array     As String
    
    For counter = LBound(my_array) To UBound(my_array)
        
        s_array = s_array & my_array(counter)
    
    Next counter
    
End Function

Function bubble_sort(ByRef TempArray As Variant) As Variant

    Dim Temp            As Variant
    Dim i               As Long
    Dim NoExchanges     As Long

    ' Loop until no more "exchanges" are made.
    Do
        NoExchanges = True

        ' Loop through each element in the array.
        For i = LBound(TempArray) To UBound(TempArray) - 1

            ' If the element is greater than the element
            ' following it, exchange the two elements.

            If CLng(TempArray(i)) > CLng(TempArray(i + 1)) Then
                NoExchanges = False
                Temp = TempArray(i)
                TempArray(i) = TempArray(i + 1)
                TempArray(i + 1) = Temp
            End If
            
        Next i
    
    Loop While Not (NoExchanges)
    bubble_sort = TempArray
    
End Function

Public Sub print_array(my_array As Variant)
    Dim counter As Integer
    
    For counter = LBound(my_array) To UBound(my_array)
        Debug.Print counter & " --> " & my_array(counter)
    Next counter
    
End Sub


'http://www.cpearson.com/excel/vbaarrays.htm
Public Function IsArrayAllocated(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayAllocated
' Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
' sized with Redim) or FALSE if the array is not allocated (a dynamic that has not yet
' been sized with Redim, or a dynamic array that has been Erased). Static arrays are always
' allocated.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
' This function is just the reverse of IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim N As Long
On Error Resume Next

' if Arr is not an array, return FALSE and get out.
If IsArray(Arr) = False Then
    IsArrayAllocated = False
    Exit Function
End If

' Attempt to get the UBound of the array. If the array has not been allocated,
' an error will occur. Test Err.Number to see if an error occurred.
N = UBound(Arr, 1)
If (Err.Number = 0) Then
    ''''''''''''''''''''''''''''''''''''''
    ' Under some circumstances, if an array
    ' is not allocated, Err.Number will be
    ' 0. To acccomodate this case, we test
    ' whether LBound <= Ubound. If this
    ' is True, the array is allocated. Otherwise,
    ' the array is not allocated.
    '''''''''''''''''''''''''''''''''''''''
    If LBound(Arr) <= UBound(Arr) Then
        ' no error. array has been allocated.
        IsArrayAllocated = True
    Else
        IsArrayAllocated = False
    End If
Else
    ' error. unallocated array
    IsArrayAllocated = False
End If

End Function


Public Function IsArrayEmpty(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayEmpty
' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
' This function is really the reverse of IsArrayAllocated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim LB As Long
Dim UB As Long

Err.Clear
On Error Resume Next
If IsArray(Arr) = False Then
    ' we weren't passed an array, return True
    IsArrayEmpty = True
End If

' Attempt to get the UBound of the array. If the array is
' unallocated, an error will occur.
UB = UBound(Arr, 1)
If (Err.Number <> 0) Then
    IsArrayEmpty = True
Else
    ''''''''''''''''''''''''''''''''''''''''''
    ' On rare occassion, under circumstances I
    ' cannot reliably replictate, Err.Number
    ' will be 0 for an unallocated, empty array.
    ' On these occassions, LBound is 0 and
    ' UBoung is -1.
    ' To accomodate the weird behavior, test to
    ' see if LB > UB. If so, the array is not
    ' allocated.
    ''''''''''''''''''''''''''''''''''''''''''
    Err.Clear
    LB = LBound(Arr)
    If LB > UB Then
        IsArrayEmpty = True
    Else
        IsArrayEmpty = False
    End If
End If

End Function

