Attribute VB_Name = "modArraySupport"

Option Explicit
Option Private Module
Option Compare Text

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'modArraySupport
'By Chip Pearson, chip@cpearson.com, www.cpearson.com
'
'This module contains procedures that provide information about and manipulate
'VB/VBA arrays. NOTE: These functions call one another. It is strongly
'suggested that you import this entire module to a VBProject rather then
'copy/pasting individual procedures.
'
'For details on these functions, see www.cpearson.com/excel/VBAArrays.htm
'
'This module contains the following functions:
'     AreDataTypesCompatible           --> changed order of arguments
'     ChangeBoundsOfVector             --> renamed from 'ChangeBoundsOfArray'
'     CombineTwoDArrays
'     CompareVectors                   --> renamed from 'CompareArrays'
'     ConcatenateArrays
'     CopyArray                        --> changed order of arguments
'     CopyNonNothingObjectsToVector    --> renamed from 'CopyNonNothingObjectsToArray'
'     CopyVectorSubSetToVector         --> renamed from 'CopyArraySubSetToArray'
'     DataTypeOfArray
'     DeleteVectorElement              --> renamed from 'DeleteArrayElement'
'     ExpandArray
'     FirstNonEmptyStringIndexInVector --> renamed from 'FirstNonEmptyStringIndexInArray'
'     GetColumn
'     GetRow
'     InsertElementIntoVector          --> renamed from 'InsertElementIntoArray'
'     IsArrayAllDefault
'     IsArrayAllNumeric
'     IsArrayAllocated
'     IsArrayDynamic
'     (IsArrayEmpty)                   --> = Not IsArrayAllocated
'     IsArrayObjects
'     IsNumericDataType
'     IsVariantArrayConsistent
'     (IsVariantArrayNumeric)          --> merged into `IsArrayAllNumeric'
'     IsVectorSorted                   --> renamed from 'IsArraySorted'
'     MoveEmptyStringsToEndOfVector    --> renamed from 'MoveEmptyStringsToEndOfArray'
'     NumberOfArrayDimensions
'     NumElements
'     ResetVariantArrayToDefaults
'     ReverseVectorInPlace             --> renamed from 'ReverseArrayInPlace'
'     ReverseVectorOfObjectsInPlace    --> renamed from 'ReverseArrayOfObjectsInPlace'
'     SetObjectArrayToNothing
'     SetVariableToDefault
'     SwapArrayColumns
'     SwapArrayRows
'     TransposeArray
'     VectorsToArray
'
'Function documentation is above each function.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Error Number Constants
Private Const C_ERR_NO_ERROR As Long = 0
Private Const C_ERR_SUBSCRIPT_OUT_OF_RANGE As Long = 9
Private Const C_ERR_ARRAY_IS_FIXED_OR_LOCKED As Long = 10


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'AreDataTypesCompatible
'This function determines if 'SourceVar' is compatible with 'DestVar'. If the
'two data types are the same, they are compatible. If the value of 'SourceVar'
'can be stored in 'DestVar' with no loss of precision or an overflow, they are
'compatible.
'For example, if 'DestVar' is a 'Long' and 'SourceVar' is an 'Integer', they
'are compatible because an 'Integer' can be stored in a 'Long' with no loss of
'information. If 'DestVar' is a 'Long' and 'SourceVar' is a 'Double', they are
'not compatible because information will be lost converting from a 'Double' to
'a 'Long' (the decimal portion will be lost).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AreDataTypesCompatible( _
    ByVal SourceVar As Variant, _
    ByVal DestVar As Variant _
        ) As Boolean
    
    Dim SVType As VbVarType
    Dim DVType As VbVarType
    
    Dim LongLongType As Byte
    LongLongType = DeclareLongLong
    
    
    'Set the default return value
    AreDataTypesCompatible = False
    
    'If one variable is an array and the other is not an array, they are incompatible
    If (IsArray(SourceVar) And Not IsArray(DestVar)) Or _
            (Not IsArray(SourceVar) And IsArray(DestVar)) Then
        Exit Function
    End If
    
    'If 'SourceVar' is an array, get the type of array. If it is an array its
    ''VarType' is 'vbArray + VarType(element)' so we subtract 'vbArray' to get
    'the data type of the array. E.g., the 'VarType' of an array of 'Long's is
    '8195 = vbArray + vbLong,
    '8195 - vbArray = vbLong (= 3).
    If IsArray(SourceVar) Then
        SVType = VarType(SourceVar) - vbArray
    Else
        SVType = VarType(SourceVar)
    End If
    'If 'DestVar' is an array, get the type of array
    If IsArray(DestVar) Then
        DVType = VarType(DestVar) - vbArray
    Else
        DVType = VarType(DestVar)
    End If
    
    'Test the data type of 'DestVar' and return a result if 'SourceVar' is
    'compatible with that type.
    If SVType = DVType Then
        'The variable types are the same --> they are compatible
        AreDataTypesCompatible = True
    'If the data types are not the same, determine whether they are compatible
    Else
        Select Case DVType
            Case vbInteger
                'there is no compatible match for that
                '(that isn't already caught above)
            Case vbLong, LongLongType
                Select Case SVType
                    Case vbInteger, vbLong, LongLongType
                        AreDataTypesCompatible = True
                End Select
            Case vbSingle
                Select Case SVType
                    Case vbInteger, vbLong, LongLongType, vbSingle
                        AreDataTypesCompatible = True
                End Select
            Case vbDouble
                Select Case SVType
                    Case vbInteger, vbLong, LongLongType, vbSingle, vbDouble
                        AreDataTypesCompatible = True
                End Select
'            'this is already covered above
'            Case vbString
'                Select Case SVType
'                    Case vbString
'                        AreDataTypesCompatible = True
'                End Select
'            'this is already covered above
'            Case vbObject
'                Select Case SVType
'                    Case vbObject
'                        AreDataTypesCompatible = True
'                End Select
            Case vbBoolean
                Select Case SVType
                    Case vbBoolean, vbInteger
                        AreDataTypesCompatible = True
                End Select
'            'this is already covered above
'            Case vbByte
'                Select Case SVType
'                    Case vbByte
'                        AreDataTypesCompatible = True
'                End Select
            Case vbCurrency
                Select Case SVType
                    Case vbInteger, vbLong, LongLongType, vbSingle, vbDouble
                        AreDataTypesCompatible = True
                End Select
            Case vbDecimal
                Select Case SVType
                    Case vbInteger, vbLong, LongLongType, vbSingle, vbDouble
                        AreDataTypesCompatible = True
                End Select
            Case vbDate
                Select Case SVType
                    Case vbLong, LongLongType, vbSingle, vbDouble
                        AreDataTypesCompatible = True
                End Select
            Case vbEmpty
                Select Case SVType
                    Case vbVariant
                        AreDataTypesCompatible = True
                End Select
            Case vbError
            Case vbNull
'            'this is already covered above
'            Case vbObject
'                Select Case SVType
'                    Case vbObject
'                        AreDataTypesCompatible = True
'                End Select
            Case vbVariant
                'everything is compatible to a 'Variant'
                AreDataTypesCompatible = True
        End Select
    End If
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ChangeBoundsOfVector
'This function changes the upper and lower bounds of the specified vector.
''InputVector' MUST be a single-dimensional dynamic array.
'If the new size of the vector (NewUpperBound - NewLowerBound + 1) is greater
'than the original vector, the unused elements on the right side of the vector
'are the default values for the data type of the vector. If the new size is
'less than the original size, only the first (left-most) 'N' elements are
'included in the new vector.
'The elements of the vector may be simple variables ('String's, 'Long's, etc.),
'objects, or arrays. User-Defined Types are not supported.
'The function returns 'True' if successful, 'False' otherwise.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ChangeBoundsOfVector( _
    ByRef InputVector As Variant, _
    ByVal NewLowerBound As Long, _
    Optional ByVal NewUpperBound As Variant _
        ) As Boolean
    
    Dim TempVector() As Variant
    Dim InNdx As Long
    Dim OutNdx As Long
    Dim TempNdx As Long
    Dim FirstIsObject As Boolean
    
    
    'Set the default return value
    ChangeBoundsOfVector = False
    
    If IsMissing(NewUpperBound) Or IsEmpty(NewUpperBound) Then
        NewUpperBound = NewLowerBound + UBound(InputVector) - LBound(InputVector)
    ElseIf Not IsNumeric(NewUpperBound) Then
        Exit Function
    ElseIf NewUpperBound <> CLng(NewUpperBound) Then
        Exit Function
    End If
    
    If NewLowerBound > NewUpperBound Then Exit Function
    If Not IsArrayDynamic(InputVector) Then Exit Function
    If NumberOfArrayDimensions(InputVector) <> 1 Then Exit Function
    
    'We need to save the 'IsObject' status of the first element of 'InputVector'
    'to properly handle 'Empty' variables if we are making the vector larger
    'than it was before.
    FirstIsObject = IsObject(InputVector(LBound(InputVector)))
    
    
    'Resize 'TempVector' and save the values in 'InputVector' in 'TempVector'.
    ''TempVector' will have an 'LBound' of 1 and a 'UBound' of the size of
    '(NewUpperBound - NewLowerBound +1)
    ReDim TempVector(1 To (NewUpperBound - NewLowerBound + 1))
    'Load up 'TempVector'
    TempNdx = 0
    For InNdx = LBound(InputVector) To UBound(InputVector)
        TempNdx = TempNdx + 1
        If TempNdx > UBound(TempVector) Then
            Exit For
        End If
        
        If (IsObject(InputVector(InNdx)) = True) Then
            If InputVector(InNdx) Is Nothing Then
                Set TempVector(TempNdx) = Nothing
            Else
                Set TempVector(TempNdx) = InputVector(InNdx)
            End If
        Else
            TempVector(TempNdx) = InputVector(InNdx)
        End If
    Next
    
    'Now, erase 'InputVector', resize it to the new bounds, and load up the
    'values from 'TempVector' to the new 'InputVector'
    Erase InputVector
    ReDim InputVector(NewLowerBound To NewUpperBound)
    OutNdx = LBound(InputVector)
    For TempNdx = LBound(TempVector) To UBound(TempVector)
        If OutNdx <= UBound(InputVector) Then
            If IsObject(TempVector(TempNdx)) Then
                Set InputVector(OutNdx) = TempVector(TempNdx)
            Else
                If FirstIsObject = True Then
                    If IsEmpty(TempVector(TempNdx)) Then
                        Set InputVector(OutNdx) = Nothing
                    Else
                        Set InputVector(OutNdx) = TempVector(TempNdx)
                    End If
                Else
                    InputVector(OutNdx) = TempVector(TempNdx)
                End If
            End If
        Else
            Exit For
        End If
        OutNdx = OutNdx + 1
    Next
    
    ChangeBoundsOfVector = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CombineTwoDArrays
'This takes two 2-dimensional arrays, 'Arr1' and 'Arr2', and returns an array
'combining the two. The number of rows in the result is 'NumRows(Arr1)' +
''NumRows(Arr2)'. 'Arr1' and 'Arr2' must have the same number of columns, and
'the result array will have that many columns as well. All the 'LBounds' must
'be the same. E.g.,
'The following arrays are legal:
'    Dim Arr1(0 To 4, 0 To 10)
'    Dim Arr2(0 To 3, 0 To 10)
'The following arrays are illegal
'    Dim Arr1(0 To 4, 1 To 10)
'    Dim Arr2(0 To 3, 0 To 10)
'
'The returned result array is 'Arr1' with additional rows appended from 'Arr2'.
'For example, the arrays
'    a    b        and     e    f
'    c    d                g    h
'become
'    a    b
'    c    d
'    e    f
'    g    h
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CombineTwoDArrays( _
    ByVal Arr1 As Variant, _
    ByVal Arr2 As Variant _
        ) As Variant
    
    'Upper and lower bounds of 'Arr1'
    Dim LBoundRow1 As Long
    Dim UBoundRow1 As Long
    Dim LBoundCol1 As Long
    Dim UBoundCol1 As Long
    
    'Upper and lower bounds of 'Arr2'
    Dim LBoundRow2 As Long
    Dim UBoundRow2 As Long
    Dim LBoundCol2 As Long
    Dim UBoundCol2 As Long
    
    'Upper and lower bounds of Result
    Dim UBoundRowResult As Long
    Dim LBoundColResult As Long
    Dim UBoundColResult As Long
    
    'Index Variables
    Dim RowNdx1 As Long
    Dim ColNdx1 As Long
    Dim RowNdx2 As Long
    Dim ColNdx2 As Long
    Dim RowNdxResult As Long
    
    'Array Sizes
    Dim NumRows1 As Long
    Dim NumCols1 As Long
    
    Dim NumRows2 As Long
    Dim NumCols2 As Long
    
    Dim Done As Boolean
    Dim Result() As Variant
    
    Dim V As Variant
    
    
    'Set the default return value
    CombineTwoDArrays = Null
    
    If Not IsArray(Arr1) Then Exit Function
    If Not IsArray(Arr2) Then Exit Function
    If NumberOfArrayDimensions(Arr1) <> 2 Then Exit Function
    If NumberOfArrayDimensions(Arr2) <> 2 Then Exit Function
    
    'Get the existing bounds
    LBoundRow1 = LBound(Arr1, 1)
    UBoundRow1 = UBound(Arr1, 1)
    
    LBoundCol1 = LBound(Arr1, 2)
    UBoundCol1 = UBound(Arr1, 2)
    
    LBoundRow2 = LBound(Arr2, 1)
    UBoundRow2 = UBound(Arr2, 1)
    
    LBoundCol2 = LBound(Arr2, 2)
    UBoundCol2 = UBound(Arr2, 2)
    
    'Get the total number of rows for the result array
    NumRows1 = UBoundRow1 - LBoundRow1 + 1
    NumCols1 = UBoundCol1 - LBoundCol1 + 1
    NumRows2 = UBoundRow2 - LBoundRow2 + 1
    NumCols2 = UBoundCol2 - LBoundCol2 + 1
    
    'Ensure the number of columns are equal
    If NumCols1 <> NumCols2 Then Exit Function
    
    'Ensure that ALL the 'LBound's are equal
    If (LBoundRow1 <> LBoundRow2) Or _
            (LBoundRow1 <> LBoundCol1) Or _
            (LBoundRow1 <> LBoundCol2) Then _
                    Exit Function
    
    'Set the bounds of the columns of the result array
    LBoundColResult = LBoundRow1
    UBoundColResult = UBoundCol1
    UBoundRowResult = LBoundRow1 + NumRows1 + NumRows2 - 1
    
    'ReDim the result array to have number of rows equal to
    ''number-of-rows(Arr1) + number-of-rows(Arr2)'
    'and number-of-columns equal to number-of-columns(Arr1)
    ReDim Result(LBoundRow1 To UBoundRowResult, LBoundColResult To UBoundColResult)
    
    RowNdxResult = LBound(Result, 1) - 1
    
    Done = False
    Do
        'Copy elements of 'Arr1' to 'Result'
        For RowNdx1 = LBoundRow1 To UBoundRow1
            RowNdxResult = RowNdxResult + 1
            For ColNdx1 = LBoundCol1 To UBoundCol1
                V = Arr1(RowNdx1, ColNdx1)
                Result(RowNdxResult, ColNdx1) = V
            Next
        Next
        
        'Copy elements of 'Arr2' to 'Result'
        For RowNdx2 = LBoundRow2 To UBoundRow2
            RowNdxResult = RowNdxResult + 1
            For ColNdx2 = LBoundCol2 To UBoundCol2
                V = Arr2(RowNdx2, ColNdx2)
                Result(RowNdxResult, ColNdx2) = V
            Next
        Next
        
        Done = RowNdxResult >= UBoundRowResult
    Loop Until Done
    
    CombineTwoDArrays = Result
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CompareVectors
'This function compares two vectors, 'Vector1' and 'Vector2', element by
'element, and puts the results of the comparisons in 'ResultVector' with the
'same 'LBound' as 'Vector1'. Each element of 'ResultVector' will be -1, 0, or
'+1. A -1 indicates that the element in 'Vector1' was less than the
'corresponding element in 'Vector2'. A 0 indicates that the elements are equal,
'and +1 indicates that the element in 'Vector1' is greater than 'Vector2'.
'
'Both 'Vector1' and 'Vector2' must be allocated single-dimensional arrays, and
''ResultVector' must be dynamic array of a numeric data type (typically 'Long').
''Vector1' and 'Vector2' must contain the same number of elements, and have the
'same lower bound. Also 'Vector1' and 'Vector2' are not allowed to contain an
'Object or User Defined Type. The function will return 'False' if not all of
'the previous conditions are met.
'
'When comparing elements, the procedure does the following:
'- If both elements are numeric data types, they are compared arithmetically.
'- If one element is a numeric data type and the other is a string and that
'  string is numeric, then both elements are converted to 'Doubles' and
'  compared arithmetically. If the string is not numeric, both elements are
'  converted to strings and compared using 'StrComp', with the compare mode set
'  by 'CompareMode'.
'- If both elements are numeric strings, they are converted to 'Doubles' and
'  compared arithmetically.
'- If either element is not a numeric string, the elements are converted and
'  compared with 'StrComp'.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CompareVectors( _
    ByVal Vector1 As Variant, _
    ByVal Vector2 As Variant, _
    ByRef ResultVector As Variant, _
    Optional ByVal CompareMode As VbCompareMethod = vbTextCompare _
        ) As Boolean
    
    Dim i As Long
    Dim S1 As String
    Dim S2 As String
    Dim D1 As Double
    Dim D2 As Double
    Dim Compare As VbCompareMethod
    
    
    'Set the default return value
    CompareVectors = False
    
    'Ensure we have a compare mode value
    If CompareMode = vbBinaryCompare Then
        Compare = vbBinaryCompare
    Else
        Compare = vbTextCompare
    End If
    
    If Not IsArray(Vector1) Then Exit Function
    If Not IsArray(Vector2) Then Exit Function
    If Not IsArrayDynamic(ResultVector) Then Exit Function
    If NumberOfArrayDimensions(Vector1) <> 1 Then Exit Function
    If NumberOfArrayDimensions(Vector2) <> 1 Then Exit Function
    
    'Ensure the LBounds are the same and size of the vectors is the same
    If LBound(Vector1) <> LBound(Vector2) Then Exit Function
    If UBound(Vector1) <> UBound(Vector2) Then Exit Function
    
    'ReDim ResultVector to the number of elements in 'Vector1'
    ReDim ResultVector(LBound(Vector1) To UBound(Vector1))
    
    'Scan each vector to see if it contains objects or User-Defined Types
    'If found, exit with 'False'
    For i = LBound(Vector1) To UBound(Vector1)
        If IsObject(Vector1(i)) Then Exit Function
        If VarType(Vector1(i)) >= vbArray Then Exit Function
        If VarType(Vector1(i)) = vbUserDefinedType Then Exit Function
    Next
    For i = LBound(Vector2) To UBound(Vector2)
        If IsObject(Vector2(i)) Then Exit Function
        If VarType(Vector2(i)) >= vbArray Then Exit Function
        If VarType(Vector2(i)) = vbUserDefinedType Then Exit Function
    Next
    
    
    'test each entry
    For i = LBound(Vector1) To UBound(Vector1)
        If IsNumeric(Vector1(i)) And IsNumeric(Vector2(i)) Then
            D1 = CDbl(Vector1(i))
            D2 = CDbl(Vector2(i))
            If D1 = D2 Then
                ResultVector(i) = 0
            ElseIf D1 < D2 Then
                ResultVector(i) = -1
            Else
                ResultVector(i) = 1
            End If
        Else
            S1 = CStr(Vector1(i))
            S2 = CStr(Vector2(i))
            ResultVector(i) = StrComp(S1, S2, Compare)
        End If
    Next
    
    CompareVectors = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ConcatenateArrays
'This function appends 'ArrayToAppend' to the end of 'ResultArray', increasing
'the size of 'ResultArray' as needed. 'ResultArray' must be a dynamic array,
'but it need not be allocated. 'ArrayToAppend' may be either static or dynamic,
'and if dynamic it may be unallocated. If 'ArrayToAppend' is unallocated,
''ResultArray' is left unchanged.
'
'The data types of 'ResultArray' and 'ArrayToAppend' must be either the same
'data type or 'compatible numeric types. A compatible numeric type is a type
'that will not cause a loss of precision or cause an overflow. For example,
''ReturnArray' may be 'Long', and 'ArrayToAppend' may by 'Long' or 'Integer',
'but not 'Single' or 'Double' because information might be lost when converting
'from 'Double' to 'Long' (the decimal portion would be lost).
'
'To skip the compatibility check and allow any variable type in 'ResultArray'
'and 'ArrayToAppend', set the 'CompatibilityCheck' parameter to 'False'. If
'you do this, be aware that you may loose precision and you may will get an
'overflow error which will cause a result of 0 in that element of 'ResultArray'.
'
'Both 'ResultArray' and 'ArrayToAppend' must be one-dimensional arrays.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ConcatenateArrays( _
    ByRef ResultArray As Variant, _
    ByVal ArrayToAppend As Variant, _
    Optional ByVal CompatibilityCheck As Boolean = True _
        ) As Boolean
    
    Dim i As Long
    Dim NumElementsToAdd As Long
    Dim AppendNdx As Long
    Dim ResultLB As Long
    Dim ResultUB As Long
    Dim ResultWasAllocated As Boolean
    
    
    'Set the default result
    ConcatenateArrays = False
    
    If Not IsArray(ArrayToAppend) Then Exit Function
    If Not IsArrayDynamic(ResultArray) Then Exit Function
    
    'Ensure both arrays are single dimensional
    '0 indicates an unallocated array, which is allowed.
    If NumberOfArrayDimensions(ResultArray) > 1 Then Exit Function
    If NumberOfArrayDimensions(ArrayToAppend) > 1 Then Exit Function
    
    'Ensure 'ArrayToAppend' is allocated. If 'ArrayToAppend' is not allocated,
    'we have nothing to append, so exit with a 'True' result.
    If Not IsArrayAllocated(ArrayToAppend) Then
        ConcatenateArrays = True
        Exit Function
    End If
    
    
    If CompatibilityCheck Then
        'Ensure the array are compatible data types
        If Not AreDataTypesCompatible(ArrayToAppend, ResultArray) Then Exit Function
        
        'If one array is an array of objects, ensure the other contains all
        'objects (or 'Nothing')
        If VarType(ResultArray) - vbArray = vbObject Then
            If IsArrayAllocated(ArrayToAppend) Then
                For i = LBound(ArrayToAppend) To UBound(ArrayToAppend)
                    If Not IsObject(ArrayToAppend(i)) Then Exit Function
                Next
            End If
        End If
    End If
    
    
    'Get the number of elements in 'ArrayToAppend'
    NumElementsToAdd = UBound(ArrayToAppend) - LBound(ArrayToAppend) + 1
    
    'Get the bounds for resizing the 'ResultArray'. If ResultArray is allocated
    'use the 'LBound' and 'UBound+1'. If 'ResultArray' is not allocated, use
    'the 'LBound' of 'ArrayToAppend' for both the 'LBound' and 'UBound' of
    ''ResultArray'.
    If IsArrayAllocated(ResultArray) Then
        ResultLB = LBound(ResultArray)
        ResultUB = UBound(ResultArray)
        ResultWasAllocated = True
        ReDim Preserve ResultArray(ResultLB To ResultUB + NumElementsToAdd)
    Else
        ResultUB = UBound(ArrayToAppend)
        ResultWasAllocated = False
        ReDim ResultArray(LBound(ArrayToAppend) To UBound(ArrayToAppend))
    End If
    
    '''Copy the data from 'ArrayToAppend' to 'ResultArray'.
    'If 'ResultArray' was allocated, we have to put the data from 'ArrayToAppend'
    'at the end of the 'ResultArray'.
    If ResultWasAllocated = True Then
        AppendNdx = LBound(ArrayToAppend)
        For i = ResultUB + 1 To UBound(ResultArray)
            If IsObject(ArrayToAppend(AppendNdx)) Then
                Set ResultArray(i) = ArrayToAppend(AppendNdx)
            Else
                ResultArray(i) = ArrayToAppend(AppendNdx)
            End If
            AppendNdx = AppendNdx + 1
            If AppendNdx > UBound(ArrayToAppend) Then
                Exit For
            End If
        Next
    'If 'ResultArray' was not allocated, we simply copy element by element from
    ''ArrayToAppend' to 'ResultArray'.
    Else
        For i = LBound(ResultArray) To UBound(ResultArray)
            If IsObject(ArrayToAppend(i)) Then
                Set ResultArray(i) = ArrayToAppend(i)
            Else
                ResultArray(i) = ArrayToAppend(i)
            End If
        Next
    End If
    
    ConcatenateArrays = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CopyArray
'This function copies the contents of 'SourceArray' to the 'ResultArray'.
'Both 'SourceArray' and 'ResultArray' may be either static or dynamic and
'either or both may be unallocated.
'
'If 'ResultArray' is dynamic, it is resized to match 'SourceArray'. The
''LBound' and 'UBound' of 'ResultArray' will be the same as 'SourceArray',
'and all elements of 'SourceArray' will be copied to 'ResultArray'.
'
'If 'ResultArray' is static and has more elements than 'SourceArray', all
'of 'SourceArray' is copied to 'ResultArray' and the right-most elements
'of 'ResultArray' are left intact.
'
'If 'ResultArray' is static and has fewer elements that 'SourceArray',
'only the left-most elements of 'SourceArray' are copied to fill out
''ResultArray'.
'
'If 'SourceArray' is an unallocated array, 'ResultArray' remains unchanged
'and the procedure terminates.
'
'If both 'SourceArray' and 'ResultArray' are unallocated, no changes are
'made to either array and the procedure terminates.
'
''SourceArray' may contain any type of data, including 'Object's and 'Object's
'that are 'Nothing' (the procedure does not support arrays of 'User Defined
'Types' since these cannot be coerced to 'Variant's -- use classes instead of
'types).
'
'The function tests to ensure that the data types of the arrays are the same or
'are compatible. See the function 'AreDataTypesCompatible' for information
'about compatible data types. To skip this compatibility checking, set the
''CompatibilityCheck' parameter to 'False'. Note that you may lose information
'during data conversion (e.g., losing decimal places when converting a 'Double'
'to a 'Long') or you may get an overflow (storing a 'Long' in an 'Integer')
'which will result in that element in 'ResultArray' having a value of 0.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CopyArray( _
    ByVal SourceArray As Variant, _
    ByRef ResultArray As Variant, _
    Optional ByVal CompatibilityCheck As Boolean = True _
        ) As Boolean
    
    Dim SrcNdx As Long
    Dim ResNdx As Long
    
    
    'Set the default return value
    CopyArray = False
    
    If Not IsArray(ResultArray) Then Exit Function
    If Not IsArray(SourceArray) Then Exit Function
    
    'Ensure both arrays are single dimensional
    '0 indicates an unallocated array, which is allowed.
    If NumberOfArrayDimensions(SourceArray) > 1 Then Exit Function
    If NumberOfArrayDimensions(ResultArray) > 1 Then Exit Function
    
    'If 'SourceArray' is not allocated, leave 'ResultArray' intact and return a
    'result of 'True'.
    If Not IsArrayAllocated(SourceArray) Then
        CopyArray = True
        Exit Function
    End If
    
    If CompatibilityCheck Then
        'Ensure both arrays are the same type or compatible data types. See the
        'function 'AreDataTypesCompatible' for information about compatible types.
        If Not AreDataTypesCompatible(SourceArray, ResultArray) Then
            Exit Function
        End If
        'If one array is an array of objects, ensure the other contains all
        'objects (or 'Nothing')
        If VarType(ResultArray) - vbArray = vbObject Then
            If IsArrayAllocated(SourceArray) Then
                For SrcNdx = LBound(SourceArray) To UBound(SourceArray)
                    If Not IsObject(SourceArray(SrcNdx)) Then Exit Function
                Next
            End If
        End If
    End If
    
    'If both arrays are allocated, copy from 'SourceArray' to 'ResultArray'.
    'If 'SourceArray' is smaller that 'ResultArray', the right-most elements
    'of 'ResultArray' are left unchanged. If 'SourceArray' is larger than
    ''ResultArray', the right most elements of 'SourceArray' are not copied.
    If IsArrayAllocated(ResultArray) Then
        ResNdx = LBound(ResultArray)
        On Error Resume Next
        For SrcNdx = LBound(SourceArray) To UBound(SourceArray)
            If IsObject(SourceArray(SrcNdx)) Then
                Set ResultArray(ResNdx) = SourceArray(SrcNdx)
            Else
                ResultArray(ResNdx) = SourceArray(SrcNdx)
            End If
            ResNdx = ResNdx + 1
            If ResNdx > UBound(ResultArray) Then
                Exit For
            End If
        Next
        On Error GoTo 0
    'If (only) 'ResultArray' is not allocated, 'ReDim ResultArray' to
    'the same size as 'SourceArray' and copy the elements from 'SourceArray' to
    ''ResultArray'.
    Else
        On Error Resume Next
        ReDim ResultArray(LBound(SourceArray) To UBound(SourceArray))
        For SrcNdx = LBound(SourceArray) To UBound(SourceArray)
            If IsObject(SourceArray(SrcNdx)) Then
                Set ResultArray(SrcNdx) = SourceArray(SrcNdx)
            Else
                ResultArray(SrcNdx) = SourceArray(SrcNdx)
            End If
        Next
        On Error GoTo 0
    End If
    
    CopyArray = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CopyNonNothingObjectsToVector
'This function copies all objects that are not Nothing from 'SourceVector'
'to 'ResultVector'. 'ResultVector' MUST be a dynamic array of type 'Object' or
''Variant', e.g.,
'    Dim ResultVector() As Object
'or
'    Dim ResultVector() as Variant
'
''ResultVector' will be erased and then resized to hold the non-Nothing
'elements from 'SourceVector'. The 'LBound' of 'ResultVector' will be the same
'as the 'LBound' of 'SourceVector', regardless of what its 'LBound' was prior
'to calling this procedure.
'
'This function returns 'True' if the operation was successful or 'False' if an
'error occurs.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CopyNonNothingObjectsToVector( _
    ByVal SourceVector As Variant, _
    ByRef ResultVector As Variant _
        ) As Boolean
    
    Dim SrcNdx As Long
    Dim ResNdx As Long
    
    
    'Set the default return value
    CopyNonNothingObjectsToVector = False
    
    If Not IsArrayDynamic(ResultVector) Then Exit Function
    'Ensure 'ResultVector' is unallocated or single-dimensional
    If NumberOfArrayDimensions(ResultVector) > 1 Then Exit Function
    
    'Ensure that all the elements of 'SourceVector' are in fact objects
    If Not IsArrayObjects(SourceVector) Then Exit Function
    
    'Erase the 'ResultVector'. Since 'ResultVector' is dynamic, this will
    'release the memory used by 'ResultVector' and return the array to an
    'unallocated state.
    Erase ResultVector
    'Now, size 'ResultVector' to the size of 'SourceVector'. After moving all
    'the non-Nothing elements, we'll do another resize to get 'ResultVector' to
    'the used size. This method allows us to avoid 'ReDim Preserve' for every
    'element.
    ReDim ResultVector(LBound(SourceVector) To UBound(SourceVector))
    
    ResNdx = LBound(SourceVector)
    For SrcNdx = LBound(SourceVector) To UBound(SourceVector)
        If Not SourceVector(SrcNdx) Is Nothing Then
            Set ResultVector(ResNdx) = SourceVector(SrcNdx)
            ResNdx = ResNdx + 1
        End If
    Next
    
    'Now that we've copied all the non-Nothing elements we call 'ReDim Preserve'
    'to resize the 'ResultVector' to the size actually used. Test 'ResNdx' to
    'see if we actually copied any elements.
    '
    'If 'ResNdx > LBound(SourceVector)' then we copied at least one element out
    'of 'SourceVector' ...
    If ResNdx > LBound(SourceVector) Then
        ReDim Preserve ResultVector(LBound(ResultVector) To ResNdx - 1)
    '... otherwise we didn't copy any elements from 'SourceVector'
    '(all elements in 'SourceVector' were 'Nothing'). In this case,
    ''Erase ResultVector'.
    Else
        Erase ResultVector
    End If
    
    CopyNonNothingObjectsToVector = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CopyVectorSubSetToVector
'This function copies elements of 'SourceVector' to 'ResultVector'. It takes
'the elements from 'FirstElementToCopy' to 'LastElementToCopy' (inclusive) from
''SourceVector' and copies them to 'ResultVector', starting at 'DestinationElement'.
'Existing data in 'ResultVector' will be overwritten. If 'ResultVector' is a
'dynamic array, it will be resized if needed. If 'ResultVector' is a static
'array and it is not large enough to copy all the elements, no elements are
'copied and the function returns 'False'.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'2do: add type compatibility checking (as optional argument)?
Public Function CopyVectorSubSetToVector( _
    ByVal SourceVector As Variant, _
    ByRef ResultVector As Variant, _
    ByVal FirstElementToCopy As Long, _
    ByVal LastElementToCopy As Long, _
    ByVal DestinationElement As Long _
        ) As Boolean
    
    Dim SrcNdx As Long
    Dim ResNdx As Long
    Dim LBoundOrgResultVector As Long
    Dim UBoundOrgResultVector As Long
    Dim NumElementsToCopy As Long
    Dim FinalIndexToCopyInResultVector As Long
    Dim TempArray() As Variant
    
    
    'Set the default return value
    CopyVectorSubSetToVector = False
    
    If Not IsArray(SourceVector) Then Exit Function
    If Not IsArray(ResultVector) Then Exit Function
    If NumberOfArrayDimensions(SourceVector) <> 1 Then Exit Function
    'Ensure 'ResultVector' is unallocated or single dimensional
    If NumberOfArrayDimensions(ResultVector) > 1 Then Exit Function
    
    'Ensure the bounds and indices are valid
    If FirstElementToCopy < LBound(SourceVector) Then Exit Function
    If LastElementToCopy > UBound(SourceVector) Then Exit Function
    If FirstElementToCopy > LastElementToCopy Then Exit Function
    
    
    'Store bounds of (original) 'ResultVector'
        'in case 'ResultVector' is unallocated and thus has no bounds
        On Error Resume Next
    LBoundOrgResultVector = LBound(ResultVector)
    UBoundOrgResultVector = UBound(ResultVector)
        On Error GoTo 0
    
    'Calculate the number of elements we'll copy from 'SourceVector' to 'ResultVector'
    NumElementsToCopy = LastElementToCopy - FirstElementToCopy + 1
    
    'Calculate the final element/index to copy in 'ResultVector'
    FinalIndexToCopyInResultVector = DestinationElement + NumElementsToCopy - 1
    
    If Not IsArrayDynamic(ResultVector) Then
        If (FirstElementToCopy < LBoundOrgResultVector) Or _
                (FinalIndexToCopyInResultVector <= UBoundOrgResultVector) Then
            ''ResultVector' is static and can't be resized.
            'There is not enough room in the vector to copy all the data.
            Exit Function
        End If
    ''ResultVector' is dynamic and can be resized
    Else
        'Test whether we need to resize the ResultVector, and resize it if required
        If Not IsArrayAllocated(ResultVector) Then
            ''ResultVector' is unallocated. Resize it to 'FinalIndexToCopyInResultVector'.
            'This provides empty elements to the left of the 'DestinationElement'
            'and room to copy 'NumElementsToCopy',
            'if 'DestinationElement' is larger than 'Option Base' ...
            If DestinationElement > 1 Then
                ReDim ResultVector(1 To FinalIndexToCopyInResultVector)
            '... and maybe empty elements to the right, if the largest element is
            'smaller than 'Option Base'
            ElseIf FinalIndexToCopyInResultVector < 1 Then
                ReDim ResultVector(DestinationElement To 1)
            Else
                ReDim ResultVector(DestinationElement To FinalIndexToCopyInResultVector)
            End If
        ''ResultVector' is allocated.
        Else
            If (DestinationElement >= LBoundOrgResultVector) And _
                    (FinalIndexToCopyInResultVector <= UBoundOrgResultVector) Then
                'nothing to do in this case
            ElseIf (DestinationElement <= LBoundOrgResultVector) And _
                    (FinalIndexToCopyInResultVector >= UBoundOrgResultVector) Then
                'in this case all elements of 'ResultVector' will be overwritten
                'just 'ReDim ResultVector'
                ReDim ResultVector(DestinationElement To FinalIndexToCopyInResultVector)
            ElseIf DestinationElement < LBoundOrgResultVector Then
                'when we ReDim the 'LBound' the data are shifted to the new indices
                'as well, e.g. a former 'ResultVector(0) = 10' would become
                ''ResultVector(-2) = 10' if 'DestinationElement = -2' etc.
                'Thus, we have to restore the elements that are not overwritten.
                
                'before 'ReDim'ing 'ResultVector' make a dummy copy of it
                If Not CopyArray(ResultVector, TempArray) Then Exit Function
                ReDim Preserve ResultVector(DestinationElement To UBoundOrgResultVector)
                
                'only copy the elements back that will not be overwritten
                For ResNdx = FinalIndexToCopyInResultVector + 1 To UBoundOrgResultVector
                    ResultVector(ResNdx) = TempArray(ResNdx)
                Next
            ElseIf FinalIndexToCopyInResultVector > UBoundOrgResultVector Then
                ReDim Preserve ResultVector(LBoundOrgResultVector To FinalIndexToCopyInResultVector)
            End If
        End If
    End If
    
    'Copy the elements from 'SourceVector' to 'ResultVector'.
    'Note that there is no type compatibility checking when copying the elements.
    ResNdx = DestinationElement
    For SrcNdx = FirstElementToCopy To LastElementToCopy
        If IsObject(SourceVector(SrcNdx)) Then
            Set ResultVector(ResNdx) = SourceVector(SrcNdx)
        Else
            On Error Resume Next
            ResultVector(ResNdx) = SourceVector(SrcNdx)
            On Error GoTo 0
        End If
        ResNdx = ResNdx + 1
    Next
    
    CopyVectorSubSetToVector = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DataTypeOfArray
'Returns a 'VbVarType' value indicating data type of the elements of 'Arr'.
'The 'VarType' of an array is the value 'vbArray' plus the 'VbVarType' value of
'the data type of the array. For example the 'VarType' of an array of 'Long's
'is 8195, which equal to 'vbArray + vbLong'. This code subtracts the value of
''vbArray' to return the native data type.
'If 'Arr' is a simple array, either one- or two-dimensional, the function
'returns the data type of the array. 'Arr' may be an unallocated array. We can
'still get the data type of an unallocated array.
'If 'Arr' is an array of arrays, the function returns 'vbArray'. To retrieve
'the data type of a subarray, pass into the function one of the sub-arrays.
'E.g.,
'    Dim R As VbVarType
'    R = DataTypeOfArray(A(LBound(A)))
'This function supports one- and multi-dimensional arrays. It does not support
'user-defined types. If 'Arr' is an array of empty variants ('vbEmpty') it
'returns 'vbVariant'.
'Returns -1 if 'Arr' is not an array.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DataTypeOfArray( _
    ByVal Arr As Variant _
        ) As VbVarType
    
    Dim Element As Variant
    Dim StoredElement As Variant
    
    
    If Not IsArray(Arr) Then
        DataTypeOfArray = -1
        Exit Function
    End If
    
    'If the array is unallocated, we can still get its data type.
    'The result of 'VarType' of an array is 'vbArray' + the 'VarType' of
    'elements of the array (e.g., the 'VarType' of an array of 'Long's is 8195,
    'which is 'vbArray + vbLong'). Thus, to get the basic data type of the
    'array, we subtract the value 'vbArray'.
    If Not IsArrayAllocated(Arr) Then
        DataTypeOfArray = VarType(Arr) - vbArray
    Else
        '(We use this for loop to get the first element of an array of arbitrary
        'dimensionality)
        For Each Element In Arr
            If IsObject(Element) Then
                DataTypeOfArray = vbObject
                Exit Function
            End If
            StoredElement = Element
            Exit For
        Next
        
        'If we were passed an array of arrays, 'IsArray(StoredElement)' will be
        'true. Therefore, return 'vbArray'. If 'IsArray(StoredElement)' is false,
        'we weren't passed an array of arrays, so simply return the data type of
        ''StoredElement'.
        If IsArray(StoredElement) Then
            DataTypeOfArray = vbArray
        Else
            If VarType(StoredElement) = vbEmpty Then
                DataTypeOfArray = vbVariant
            Else
                DataTypeOfArray = VarType(StoredElement)
            End If
        End If
    End If
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DeleteVectorElement
'This function deletes an element from 'InputVector', and shifts elements that
'are to the right of the deleted element to the left. If 'InputVector' is a
'dynamic array, and the 'ResizeDynamic' parameter is 'True', the array will be
'resized one element smaller. Otherwise, the right-most entry in the array is
'set to the default value appropriate to the data type of the array
'(0, vbNullString, Empty, or Nothing). If the array is an array of 'Variant'
'types, the default data type is the data type of the last element in the
'array. The function returns 'True' if the element was successfully deleted and
''False' otherwise. This procedure works only on single-dimensional arrays.
'(In case the only element is deleted, 'InputVector' is dynamic and
''ResizeDynamic' is 'True' 'InputVector' will be erased.)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DeleteVectorElement( _
    ByRef InputVector As Variant, _
    ByVal ElementNumber As Long, _
    Optional ByVal ResizeDynamic As Boolean = False _
        ) As Boolean
    
    Dim i As Long
    Dim VType As VbVarType
    
    Dim LongLongType As Byte
    LongLongType = DeclareLongLong
    
    
    'Set the default return value
    DeleteVectorElement = False
    
    If Not IsArray(InputVector) Then Exit Function
    If NumberOfArrayDimensions(InputVector) <> 1 Then Exit Function
    
    'Ensure we have a valid 'ElementNumber'
    If ElementNumber < LBound(InputVector) Then Exit Function
    If ElementNumber > UBound(InputVector) Then Exit Function
    
    'Get the variable data type of the element we are deleting
    VType = VarType(InputVector(UBound(InputVector)))
    If IsObject(InputVector(UBound(InputVector))) Then
        VType = vbObject
    ElseIf VType >= vbArray Then
        VType = VType - vbArray
    End If
    
    'Shift everything to the left
    For i = ElementNumber To UBound(InputVector) - 1
        If IsObject(InputVector(i)) Then
            Set InputVector(i) = InputVector(i + 1)
        Else
            InputVector(i) = InputVector(i + 1)
        End If
    Next
    
    If IsArrayDynamic(InputVector) And ResizeDynamic = True Then
        If UBound(InputVector) > LBound(InputVector) Then
            ReDim Preserve InputVector(LBound(InputVector) To UBound(InputVector) - 1)
        Else
            Erase InputVector
        End If
    Else
        'Set the last element of the 'InputVector' to the proper default value
        Select Case VType
            Case vbByte, vbInteger, vbLong, LongLongType, vbSingle, vbDouble, vbDate, vbCurrency, vbDecimal
                InputVector(UBound(InputVector)) = 0
            Case vbString
                InputVector(UBound(InputVector)) = vbNullString
            Case vbArray, vbVariant, vbEmpty, vbError, vbNull, vbUserDefinedType
                InputVector(UBound(InputVector)) = Empty
            Case vbBoolean
                InputVector(UBound(InputVector)) = False
            Case vbObject
                Set InputVector(UBound(InputVector)) = Nothing
            Case Else
                InputVector(UBound(InputVector)) = 0
        End Select
    End If
    
    DeleteVectorElement = True
    
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''IsVariantArrayNumeric
''This function returns 'True' if all the elements of an array of variants are
''numeric data types. They need not all be the same data type. You can have a
''mix of 'Integer's, 'Long's, 'Double's, and 'Single's.
''As long as they are all numeric data types, the function will return 'True'.
''If a non-numeric data type is encountered, the function will return 'False'.
''Also, it will return 'False' if 'InputArray' is not an array, or if
'''InputArray' has not been allocated. 'InputArray' may be a multi-dimensional
''array. This procedure uses the 'IsNumericDataType' function to determine
''whether a variable is a numeric data type. If there is an uninitialized
''variant ('VarType = vbEmpty') in the array, it is skipped and not used in the
''comparison (i.e., 'Empty' is considered a valid numeric data type since you
''can assign a number to it).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Function IsVariantArrayNumeric( _
'    InputArray As Variant _
'        ) As Boolean
'
'    Dim Element As Variant
'
'
'    'Set the default return value
'    IsVariantArrayNumeric = False
'
'    If Not IsArray(InputArray) Then Exit Function
'    If Not IsArrayAllocated(InputArray) Then Exit Function
'
'    For Each Element In InputArray
'        If IsObject(Element) Then Exit Function
'
'        Select Case VarType(Element)
'            Case vbEmpty
'                'allowed
'            Case Else
'                If Not IsNumericDataType(Element) Then Exit Function
'        End Select
'    Next
'
'    'If we made it up to here, then the array is entirely numeric
'    IsVariantArrayNumeric = True
'
'End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ExpandArray
'This expands a two-dimensional array in either dimension. It returns the
'result array if successful, or 'Null' if an error occurred. The original array
'is never changed.
'Parameters:
'- Arr                  is the array to be expanded
'- WhichDim             is either 1 for additional rows or
'                       2 for additional columns
'- AdditionalElements   is the number of additional rows or columns to create.
'- FillValue            is the value to which the new array elements should be
'                       initialized
'You can nest calls to expand array to expand both the number of rows and
'columns, e.g.
'
'C = ExpandArray( _
'        ExpandArray( _
'            Arr:=A, _
'            WhichDim:=1, _
'            AdditionalElements:=3, _
'            FillValue:="R") _
'        , _
'        WhichDim:=2, _
'        AdditionalElements:=4, _
'        FillValue:="C")
'
'This first adds three rows at the bottom of the array, and then adds four
'columns on the right of the array.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'2do:
'- create a type for 'WhichDim' and also replace 'ROWs_' then?
'- should this work for objects as well?
Public Function ExpandArray( _
    ByVal Arr As Variant, _
    ByVal WhichDim As Long, _
    ByVal AdditionalElements As Long, _
    ByVal FillValue As Variant _
        ) As Variant
    
    Dim Result As Variant
    Dim RowNdx As Long
    Dim ColNdx As Long
    
    '==========================================================================
    Const ROWS_ As Long = 1
    '==========================================================================
    
    
    'Set the default return value
    ExpandArray = Null
    
    If Not IsArray(Arr) Then Exit Function
    If NumberOfArrayDimensions(Arr) <> 2 Then Exit Function
    
    'Ensure the dimension is 1 or 2
    Select Case WhichDim
        Case 1, 2
        Case Else
            Exit Function
    End Select
    
    If AdditionalElements < 0 Then Exit Function
    If AdditionalElements = 0 Then
        ExpandArray = Arr
        Exit Function
    End If
    
    If WhichDim = ROWS_ Then
        'ReDim 'Result'
        ReDim Result(LBound(Arr, 1) To UBound(Arr, 1) + AdditionalElements, _
                LBound(Arr, 2) To UBound(Arr, 2))
        
        'Transfer 'Arr' array to 'Result'
        For RowNdx = LBound(Arr, 1) To UBound(Arr, 1)
            For ColNdx = LBound(Arr, 2) To UBound(Arr, 2)
                Result(RowNdx, ColNdx) = Arr(RowNdx, ColNdx)
            Next
        Next
        
        'Fill the rest of the result array with 'FillValue'
        For RowNdx = UBound(Arr, 1) + 1 To UBound(Result, 1)
            For ColNdx = LBound(Arr, 2) To UBound(Arr, 2)
                Result(RowNdx, ColNdx) = FillValue
            Next
        Next
    Else
        'ReDim 'Result'
        ReDim Result(LBound(Arr, 1) To UBound(Arr, 1), _
                LBound(Arr, 2) To UBound(Arr, 2) + AdditionalElements)
        
        'Transfer 'Arr' array to 'Result'
        For RowNdx = LBound(Arr, 1) To UBound(Arr, 1)
            For ColNdx = LBound(Arr, 2) To UBound(Arr, 2)
                Result(RowNdx, ColNdx) = Arr(RowNdx, ColNdx)
            Next
        Next
        
        'Fill the rest of the result array with 'FillValue'
        For RowNdx = LBound(Arr, 1) To UBound(Arr, 1)
            For ColNdx = UBound(Arr, 2) + 1 To UBound(Result, 2)
                Result(RowNdx, ColNdx) = FillValue
            Next
        Next
    End If
    
    ExpandArray = Result
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FirstNonEmptyStringIndexInVector
'This returns the index in 'InputVector' of the first non-empty string.
'This is generally used when 'InputVector' is the result of a sort operation,
'which puts empty strings at the beginning of the array.
'Returns -1 if an error occurred or if the entire array has no empty string.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FirstNonEmptyStringIndexInVector( _
    ByVal InputVector As Variant _
        ) As Long
    
    Dim i As Long
    
    
    'Set the default return value
    FirstNonEmptyStringIndexInVector = -1
    
    If Not IsArray(InputVector) Then Exit Function
    If NumberOfArrayDimensions(InputVector) <> 1 Then Exit Function
    
    For i = LBound(InputVector) To UBound(InputVector)
        If InputVector(i) <> vbNullString Then
            FirstNonEmptyStringIndexInVector = i
            Exit Function
        End If
    Next
    
    FirstNonEmptyStringIndexInVector = -1
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GetColumn
'This populates 'ResultArr' with a one-dimensional array that is the specified
'column of 'Arr'. The existing contents of 'ResultArr' are erased.
''ResultArr' must be a dynamic array. Returns 'True' or 'False' indicating
'success.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetColumn( _
    ByVal Arr As Variant, _
    ByRef ResultArr As Variant, _
    ByVal ColumnNumber As Long _
        ) As Boolean
    
    Dim RowNdx As Long
    
    
    'Set the default return value
    GetColumn = False
    
    If Not IsArray(Arr) Then Exit Function
    If NumberOfArrayDimensions(Arr) <> 2 Then Exit Function
    If Not IsArrayDynamic(ResultArr) Then Exit Function
    
    'Ensure 'ColumnNumber' is less than or equal to the number of columns
    If UBound(Arr, 2) < ColumnNumber Then Exit Function
    If LBound(Arr, 2) > ColumnNumber Then Exit Function
    
    Erase ResultArr
    ReDim ResultArr(LBound(Arr, 1) To UBound(Arr, 1))
    For RowNdx = LBound(ResultArr) To UBound(ResultArr)
        If IsObject(Arr(RowNdx, ColumnNumber)) Then
            Set ResultArr(RowNdx) = Arr(RowNdx, ColumnNumber)
        Else
            ResultArr(RowNdx) = Arr(RowNdx, ColumnNumber)
        End If
    Next
    
    GetColumn = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GetRow
'This populates 'ResultArr' with a one-dimensional array that is the specified
'row of 'Arr'. The existing contents of 'ResultArr' are erased. 'ResultArr'
'must be a dynamic array. Returns 'True' or 'False' indicating success.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetRow( _
    ByVal Arr As Variant, _
    ByRef ResultArr As Variant, _
    ByVal RowNumber As Long _
        ) As Boolean
    
    Dim ColNdx As Long
    
    
    'Set the default return value
    GetRow = False
    
    If Not IsArray(Arr) Then Exit Function
    If NumberOfArrayDimensions(Arr) <> 2 Then Exit Function
    If Not IsArrayDynamic(ResultArr) Then Exit Function
    
    'Ensure 'RowNumber' is less than or equal to the number of rows
    If UBound(Arr, 1) < RowNumber Then Exit Function
    If LBound(Arr, 1) > RowNumber Then Exit Function
    
    Erase ResultArr
    ReDim ResultArr(LBound(Arr, 2) To UBound(Arr, 2))
    For ColNdx = LBound(ResultArr) To UBound(ResultArr)
        If IsObject(Arr(RowNumber, ColNdx)) Then
            Set ResultArr(ColNdx) = Arr(RowNumber, ColNdx)
        Else
            ResultArr(ColNdx) = Arr(RowNumber, ColNdx)
        End If
    Next
    
    GetRow = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'InsertElementIntoVector
'This function inserts an element with a value of 'Value' into 'InputVector' at
'location 'Index'.
''InputVector' must be a dynamic array. The 'Value' is stored in location
''Index', and everything to the right of 'Index' is shifted to the right.
'The array is resized to make room for the new element. The value of 'Index'
'must be greater than or equal to the 'LBound' of 'InputVector' and less than
'or equal to 'UBound + 1'.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function InsertElementIntoVector( _
    ByRef InputVector As Variant, _
    ByVal Index As Long, _
    ByVal Value As Variant _
        ) As Boolean
    
    Dim i As Long
    
    
    'Set the default return value
    InsertElementIntoVector = False
    
    If Not IsArrayDynamic(InputVector) Then Exit Function
    If NumberOfArrayDimensions(InputVector) <> 1 Then Exit Function
    
    'Ensure 'Index' is a valid element index. We allow 'Index' to be equal to
    ''UBound + 1' to facilitate inserting a value at the end of the array, e.g.
    '    InsertElementIntoVector(Arr,UBound(Arr) + 1, 123)
    'will insert "123" at the end of the array.
    If Index < LBound(InputVector) Then Exit Function
    If Index > UBound(InputVector) + 1 Then Exit Function
    
    'Resize the array
    ReDim Preserve InputVector(LBound(InputVector) To UBound(InputVector) + 1)
    
'---
'2do:
'can't this be handled with the function 'AreDataTypesCompatible' of this module?
'---
    'First, we set the newly created last element of 'InputVector' to 'Value'.
    'This is done to trap an "error 13, type mismatch". This last entry will be
    'overwritten when we shift elements to the right, and the 'Value' will be
    'inserted at 'Index'.
    On Error Resume Next
    Err.Clear
    If IsObject(Value) Then
        Set InputVector(UBound(InputVector)) = Value
    Else
        InputVector(UBound(InputVector)) = Value
    End If
    If Err.Number <> 0 Then
        'An error occurred, most likely an error 13, type mismatch.
        'ReDim the array back to its original size and exit the function.
        ReDim Preserve InputVector(LBound(InputVector) To UBound(InputVector) - 1)
        Exit Function
    End If
'---
    
    'Shift everything to the right
    For i = UBound(InputVector) To Index + 1 Step -1
        If IsObject(InputVector(i - 1)) Then
            Set InputVector(i) = InputVector(i - 1)
        Else
            InputVector(i) = InputVector(i - 1)
        End If
    Next
    
    'Insert 'Value' at 'Index'
    If IsObject(Value) Then
        Set InputVector(Index) = Value
    Else
        InputVector(Index) = Value
    End If
    
    InsertElementIntoVector = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IsArrayAllDefault
'Returns 'True' if the array contains all default values for its data type:
'  Variable Type           Value
'  -------------           -------------------
'  Variant                 Empty
'  String                  vbNullString
'  Numeric                 0
'  Object                  Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsArrayAllDefault( _
    ByVal InputArray As Variant _
        ) As Boolean
    
    Dim Element As Variant
    Dim DefaultValue As Variant
    
    
    'Set the default return value
    IsArrayAllDefault = False
    
    If Not IsArray(InputArray) Then Exit Function
    
    'Ensure array is allocated. An unallocated array is considered to be all the
    'same type. Return 'True'.
    If Not IsArrayAllocated(InputArray) Then
        IsArrayAllDefault = True
        Exit Function
    End If
    
    'Test the type of variable
    Select Case VarType(InputArray)
        Case vbArray + vbVariant
            DefaultValue = Empty
        Case vbArray + vbString
            DefaultValue = vbNullString
        'for all (remaining/)numeric variable types
        Case Is > vbArray
            DefaultValue = 0
    End Select
    
    For Each Element In InputArray
        If IsObject(Element) Then
            If Not Element Is Nothing Then Exit Function
        Else
            If VarType(Element) <> vbEmpty Then
                If Element <> DefaultValue Then Exit Function
            End If
        End If
    Next
    
    'If we make it up to here, the array is all defaults.
    IsArrayAllDefault = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IsArrayAllNumeric
'This function returns 'True' if 'Arr' is entirely numeric and 'False'
'otherwise. The 'AllowNumericStrings' parameter indicates whether strings
'containing numeric data are considered numeric. If this parameter is 'True', a
'numeric string is considered a numeric variable. If this parameter is omitted
'or 'False', a numeric string is not considered a numeric variable. Variants
'that are numeric or empty are allowed. Variants that are objects or
'non-numeric data are not allowed. With the 'AllowArrayElements' parameter it
'can be stated, if (sub-)arrays should also be tested for numeric data.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsArrayAllNumeric( _
    ByVal Arr As Variant, _
    Optional ByVal AllowNumericStrings As Boolean = False, _
    Optional ByVal AllowArrayElements As Boolean = False _
        ) As Boolean
    
    Dim Element As Variant
    
    'Set the default return value
    IsArrayAllNumeric = False
    
    If Not IsArray(Arr) Then Exit Function
    If Not IsArrayAllocated(Arr) Then Exit Function
    
    'Loop through the array
    For Each Element In Arr
        If IsObject(Element) Then Exit Function
        
        Select Case VarType(Element)
            Case vbEmpty
                'is (also) allowed
            Case vbString
                'For strings, check the 'AllowNumericStrings' parameter.
                'If True and the element is a numeric string, allow it.
                'If it is a non-numeric string, exit with 'False'.
                'If 'AllowNumericStrings' is 'False', all strings, even
                'numeric strings, will cause a result of 'False'.
                If AllowNumericStrings = True Then
                    If Not IsNumeric(Element) Then Exit Function
                Else
                    Exit Function
                End If
            Case Is >= vbVariant
                'For Variants, disallow Objects.
                If IsObject(Element) Then Exit Function
                'If the element is an array ...
                If IsArray(Element) Then
                    '... only test the elements, if (numeric) array elements are
                    'allowed
                    If AllowArrayElements Then
                        'Test the elements (recursively) with the same rules as the
                        'main array
                        If Not IsArrayAllNumeric( _
                                Element, AllowNumericStrings, AllowArrayElements) Then _
                                        Exit Function
                    Else
                        Exit Function
                    End If
                'If the element is not an array, test, if it is of numeric type.
                Else
                    If Not IsNumeric(Element) Then Exit Function
                End If
            Case Else
                If Not IsNumeric(Element) Then Exit Function
        End Select
    Next
    
    IsArrayAllNumeric = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IsArrayAllocated
'Returns 'True' if the array is allocated (either a static or a dynamic array
'that has been sized with 'ReDim') or 'False' if the array is not allocated
'(a dynamic that has not yet been sized with 'ReDim', or a dynamic array that
'has been erased). Static arrays are always allocated.
'
'The VBA 'IsArray' function indicates whether a variable is an array, but it
'does not distinguish between allocated and unallocated arrays. It will return
''True' for both allocated and unallocated arrays. This function tests whether
'the array has actually been allocated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsArrayAllocated( _
    ByVal Arr As Variant _
        ) As Boolean
    
    Dim DummyVariable As Long
    
    
    'Set the default return value
    IsArrayAllocated = False
    
    On Error Resume Next
    
    If Not IsArray(Arr) Then Exit Function
    
    'Attempt to get the UBound of the array. If the array has not been allocated,
    'an error will occur. Test Err.Number to see if an error occurred.
    DummyVariable = UBound(Arr, 1)
    If Err.Number = 0 Then
        'Under some circumstances, if an array is not allocated, Err.Number
        'will be 0. To accommodate this case, we test whether LBound <= UBound.
        'If this is True, the array is allocated. Otherwise, the array is not
        'allocated.
        IsArrayAllocated = (LBound(Arr) <= UBound(Arr))
    Else
        'error. unallocated array
    End If
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IsArrayDynamic
'This function returns 'True' or 'False' indicating whether 'Arr' is a dynamic
'array.
'Note: If you attempt to 'ReDim' a static array in the same procedure in which
'it is declared, you'll get a compiler error and your code  won't run at all.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsArrayDynamic( _
    ByRef Arr As Variant _
        ) As Boolean
    
    Dim ArrUBound As Long
    
    
    'Set the default return value
    IsArrayDynamic = False
    
    If Not IsArray(Arr) Then Exit Function
    
    'If the array is unallocated, we know it must be a dynamic array
    If Not IsArrayAllocated(Arr) Then
        IsArrayDynamic = True
        Exit Function
    End If
    
    'Save the UBound of Arr.
    'This value will be used to restore the original UBound if Arr is a
    'single-dimensional dynamic array. Unused if Arr is multi-dimensional,
    'or if 'Arr' is a static array.
    ArrUBound = UBound(Arr)
    
    On Error Resume Next
    Err.Clear
    
    'Attempt to increase the 'UBound' of 'Arr' and test the value of
    ''Err.Number'. If 'Arr' is a static array, either single- or
    'multi-dimensional, we'll get a 'C_ERR_ARRAY_IS_FIXED_OR_LOCKED' error. In
    'this case, return 'False'.
    'If 'Arr' is a single-dimensional dynamic array, we'll get 'C_ERR_NO_ERROR'
    'error.
    'If 'Arr' is a multi-dimensional dynamic array, we'll get a
    ''C_ERR_SUBSCRIPT_OUT_OF_RANGE' error.
    ReDim Preserve Arr(LBound(Arr) To ArrUBound + 1)
    Select Case Err.Number
        Case C_ERR_NO_ERROR
            'We successfully increased the 'UBound' of 'Arr'.
            'Do a 'ReDim Preserve' to restore the original 'UBound'.
            ReDim Preserve Arr(LBound(Arr) To ArrUBound)
            IsArrayDynamic = True
        Case C_ERR_SUBSCRIPT_OUT_OF_RANGE
            ''Arr' is a multi-dimensional dynamic array.
            IsArrayDynamic = True
        Case C_ERR_ARRAY_IS_FIXED_OR_LOCKED
            ''Arr' is a static single- or multi-dimensional array.
            IsArrayDynamic = False
        Case Else
            'We should never get here.
            'Some unexpected error occurred. Be safe and return 'False'.
            IsArrayDynamic = False
    End Select
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IsArrayObjects
'Returns 'True' if 'InputArray' is entirely objects ('Nothing' objects are
'optionally allowed -- default it 'True', allow 'Nothing' objects).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsArrayObjects( _
    ByVal InputArray As Variant, _
    Optional ByVal AllowNothing As Boolean = True _
        ) As Boolean
    
    Dim Element As Variant
    
    
    'Set the default return value
    IsArrayObjects = False
    
    If Not IsArray(InputArray) Then Exit Function
    
    For Each Element In InputArray
        If Not IsObject(Element) Then Exit Function
        If Element Is Nothing Then
            If Not AllowNothing Then Exit Function
        End If
    Next
    
    IsArrayObjects = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IsNumericDataType
'This function returns 'True' or 'False' indicating whether the data type of a
'variable is a numeric data type. It will return 'True' for the data types
'  - vbCurrency
'  - vbDecimal
'  - vbDouble
'  - vbInteger
'  - vbLong, LongLongType
'  - vbSingle
'and 'False' for any other data type, including empty 'Variant's and 'Object's.
'If 'TestVar' is an unallocated array, it will test the data type of the array
'and return 'True' or 'False' for that data type. If 'TestVar' is an allocated
'array, it tests all elements, if they are numeric data type using the
''IsArrayAllNumeric' function.
'Use this procedure instead of VBA's 'IsNumeric' function because 'IsNumeric'
'will return 'True' if the variable is a string containing numeric data. This
'will cause problems with code like
'    Dim V1 As Variant
'    Dim V2 As Variant
'    V1 = "1"
'    V2 = "2"
'    If IsNumeric(V1) Then
'        If IsNumeric(V2) Then
'            Debug.Print V1 + V2
'        End If
'    End If
'The output of the 'Debug.Print' statement will be "12", not 3, because 'V1'
'and 'V2' are strings and the '+' operator acts like the '&' operator when used
'with strings. This can lead to unexpected results.
''IsNumeric' should only be used to test strings for numeric content when
'converting a string value to a numeric variable.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsNumericDataType( _
    ByVal TestVar As Variant _
        ) As Boolean
    
    Dim LongLongType As Byte
    LongLongType = DeclareLongLong
    
    
    'Set the default return value
    IsNumericDataType = False
    
    If Not IsArray(TestVar) Then
        Select Case VarType(TestVar)
            Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, LongLongType, vbSingle
                IsNumericDataType = True
        End Select
    Else
        If Not IsArrayAllocated(TestVar) Then
            Select Case VarType(TestVar) - vbArray
                Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, LongLongType, vbSingle
                    IsNumericDataType = True
            End Select
        Else
            IsNumericDataType = IsArrayAllNumeric(TestVar, False, True)
        End If
    End If
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IsVariantArrayConsistent
'This returns 'True' or 'False' indicating whether an array of variants
'contains all the same data types. Returns 'False' under the following
'circumstances:
'    'Arr' is not an array,
'    'Arr' is an array but is unallocated,
'    'Arr' is a multi-dimensional array,
'    'Arr' is allocated but does not contain consistent data types.
'If 'Arr' is an array of objects, objects that are 'Nothing' are ignored. As
'long as all non-'Nothing' objects are the same object type, the function
'returns 'True'.
'It returns 'True' if all the elements of the array have the same data type.
'If 'Arr' is an array of a specific data types, not 'Variant's, e.g.
'    Dim V(1 To 3) As Long
'the function will return 'True'. If an array of variants contains an
'uninitialized element ('VarType = vbEmpty') that element is skipped and not
'used in the comparison. The reasoning behind this is that an empty variable
'will return the data type of the variable to which it is assigned (e.g. it
'will return 'vbNullString' to a 'String' and '0' to a 'Double').
'The function does not support arrays of User Defined Types.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsVariantArrayConsistent( _
    ByVal Arr As Variant _
        ) As Boolean
    
    Dim FirstDataType As VbVarType
    Dim Element As Variant
    
    
    'Set the default return value
    IsVariantArrayConsistent = False
    
    If Not IsArray(Arr) Then Exit Function
    If Not IsArrayAllocated(Arr) Then Exit Function
    
    'Test if we have an array of a specific type rather than 'Variant's. If so,
    'return 'True' and get out.
    If VarType(Arr) - vbArray <> vbVariant Then
        IsVariantArrayConsistent = True
        Exit Function
    End If
    
    'Get the data type of the first element
    For Each Element In Arr
        FirstDataType = VarType(Element)
        Exit For
    Next
    
    'Loop through the array and exit if a differing data type if found.
    For Each Element In Arr
        If VarType(Element) <> vbEmpty Then
            If IsObject(Element) Then
                If Not Element Is Nothing Then
                    If VarType(Element) <> FirstDataType Then Exit Function
                End If
            Else
                If VarType(Element) <> FirstDataType Then Exit Function
            End If
        End If
    Next
    
    'If we make it up to here, then the array is consistent
    IsVariantArrayConsistent = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IsVectorSorted
'This function determines whether a single-dimensional array is sorted. Because
'sorting is an expensive operation, especially so on a large array of 'Variant's,
'you may want to determine if an array is already in sorted order prior to
'doing an actual sort.
'This function returns 'True' if an array is in sorted order (either ascending
'or descending, depending on the value of the 'Descending' parameter -- default
'is 'False' = Ascending). The decision to do a string comparison (with 'StrComp')
'or a numeric comparison (with < or >) is based on the data type of the first
'element of the array.
'If 'InputVector' is not an array, is an unallocated array, or has more than
'one dimension, or the VarType of 'InputVector' is not compatible, the function
'returns 'Null'. Thus, one knows that there is nothing to sort.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsVectorSorted( _
    ByVal InputVector As Variant, _
    Optional ByVal Descending As Boolean = False _
        ) As Variant
    
    Dim StrCompResultFail As Long
    Dim NumericResultFail As Boolean
    Dim i As Long
    Dim NumCompareResult As Boolean
    Dim StrCompResult As Long
    
    Dim IsString As Boolean
    Dim VType As VbVarType
    
    
    'Set the default return value
    IsVectorSorted = Null
    
    If Not IsArray(InputVector) Then Exit Function
    If NumberOfArrayDimensions(InputVector) <> 1 Then Exit Function
    
    'Determine whether we are going to do a string comparison or a numeric
    'comparison
    VType = VarType(InputVector(LBound(InputVector)))
    Select Case VType
        Case vbArray, vbDataObject, vbEmpty, vbError, vbNull, vbObject, vbUserDefinedType
            'Unsupported types.
            Exit Function
        Case vbString, vbVariant
            'Compare as string
            IsString = True
        Case Else
            'Compare as numeric
            IsString = False
    End Select
    
    'The following code sets the values of comparison that will indicate that
    'the array is unsorted. Is the result of 'StrComp' (for strings) or ">="
    '(for numerics) equal the value specified below, we know that the array is
    'unsorted.
    If Descending = True Then
        StrCompResultFail = -1
        NumericResultFail = False
    Else
        StrCompResultFail = 1
        NumericResultFail = True
    End If
    
    For i = LBound(InputVector) To UBound(InputVector) - 1
        If IsString Then
            StrCompResult = StrComp(InputVector(i), InputVector(i + 1))
            If StrCompResult = StrCompResultFail Then
                IsVectorSorted = False
                Exit Function
            End If
        Else
            NumCompareResult = (InputVector(i) >= InputVector(i + 1))
            If NumCompareResult = NumericResultFail Then
                IsVectorSorted = False
                Exit Function
            End If
        End If
    Next
    
    'If we made it up to here, then the array is in sorted order.
    IsVectorSorted = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MoveEmptyStringsToEndOfVector
'This procedure takes the SORTED array 'InputVector', which, if sorted in
'ascending order, will have all empty strings at the front of the array. This
'procedure moves those strings to the end of the array, shifting the non-empty
'strings forward in the array.
'Note that 'InputVector' MUST be sorted in ascending order.
'Returns 'True' if the array was correctly shifted (if necessary) and 'False'
'if an error occurred.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MoveEmptyStringsToEndOfVector( _
    ByRef InputVector As Variant _
        ) As Boolean
    
    Dim Ndx As Long
    Dim NonEmptyNdx As Long
    Dim LBoundArr As Long
    Dim UBoundArr As Long
    Dim FirstNonEmptyNdx As Long
    Dim LastNewNonEmptyNdx As Long
    
    
    'Set the default return value
    MoveEmptyStringsToEndOfVector = False
    
    If Not IsArray(InputVector) Then Exit Function
    If NumberOfArrayDimensions(InputVector) <> 1 Then Exit Function
    
    LBoundArr = LBound(InputVector)
    UBoundArr = UBound(InputVector)
    
    FirstNonEmptyNdx = FirstNonEmptyStringIndexInVector(InputVector)
    If FirstNonEmptyNdx <= LBoundArr Then
        'No empty strings at the beginning of the array. Get out now.
        MoveEmptyStringsToEndOfVector = True
        Exit Function
    End If
    
    LastNewNonEmptyNdx = UBoundArr + LBoundArr - FirstNonEmptyNdx
    
    'Loop through the array and move non-empty strings to the front
    NonEmptyNdx = FirstNonEmptyNdx
    For Ndx = LBoundArr To LastNewNonEmptyNdx
        InputVector(Ndx) = InputVector(NonEmptyNdx)
        NonEmptyNdx = NonEmptyNdx + 1
    Next
    
    'Set last entries entries 'vbNullString's
    For Ndx = LastNewNonEmptyNdx + 1 To UBoundArr
        InputVector(Ndx) = vbNullString
    Next
    
    MoveEmptyStringsToEndOfVector = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'NumberOfArrayDimensions
'This function returns the number of dimensions of an array. An unallocated
'dynamic array has 0 dimensions.
'(This condition can also be tested with 'Not IsArrayAllocated'.)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NumberOfArrayDimensions( _
    ByVal Arr As Variant _
        ) As Long
    
    Dim i As Long
    Dim Res As Long
    
    
    'it seems that an unallocated 'Object' array returns 1, so it is needed a
    'special handler for this case
    If DataTypeOfArray(Arr) = vbObject Then
        If Not IsArrayAllocated(Arr) Then
            NumberOfArrayDimensions = 0
            Exit Function
        End If
    End If
    
    On Error Resume Next
    'Loop, increasing the dimension index 'i', until an error occurs.
    'An error will occur when 'i' exceeds the number of dimension in the array.
    'Return 'i' - 1.
    Do
        i = i + 1
        Res = UBound(Arr, i)
    Loop Until Err.Number <> 0
    
    NumberOfArrayDimensions = i - 1
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'NumElements
'Returns the number of elements in the specified dimension ('Dimension') of the
'array in 'Arr'. If you omit 'Dimension', the first dimension is used. The
'function will return 0 under the following circumstances:
'- 'Arr' is not an array, or
'- 'Arr' is an unallocated array, or
'- 'Dimension' is less than 1, or
'- 'Dimension' is greater than the number of dimension of 'Arr'.
'This function does not support arrays of user-defined Type variables.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NumElements( _
    ByVal Arr As Variant, _
    Optional ByVal Dimension As Long = 1 _
        ) As Long
    
    Dim NumDimensions As Long
    
    
    'Set the default return value
    NumElements = 0
    
    If Not IsArray(Arr) Then Exit Function
    If Not IsArrayAllocated(Arr) Then Exit Function
    If Dimension < 1 Then Exit Function
    
    'check if 'Dimension' is not larger than 'NumDimensions'
    NumDimensions = NumberOfArrayDimensions(Arr)
    If NumDimensions < Dimension Then Exit Function
    
    'returns the number of elements in the array
    NumElements = UBound(Arr, Dimension) - LBound(Arr, Dimension) + 1
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ResetVariantArrayToDefaults
'This resets all the elements of an array of 'Variant's back to their
'appropriate default values. The elements of the array may be of mixed types
'(e.g., some 'Long's, some 'Object's, some 'String's, etc.). Each data type
'will be set to the appropriate default value ('0', 'vbNullString', 'Empty', or
''Nothing'). It returns 'True' if the array was set to defaults, or 'False' if
'an error occurred. 'InputArray' must be an allocated single-dimensional array.
'This function differs from the 'Erase' function in that it preserves the
'original data types, while 'Erase' sets every element to 'Empty'.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ResetVariantArrayToDefaults( _
    ByRef InputArray As Variant _
        ) As Boolean
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    
    'Set the default return value
    ResetVariantArrayToDefaults = False
    
    If Not IsArray(InputArray) Then Exit Function
    
    Select Case NumberOfArrayDimensions(InputArray)
        Case 1
            For i = LBound(InputArray) To UBound(InputArray)
                SetVariableToDefault InputArray(i)
            Next
        Case 2
            For i = LBound(InputArray, 1) To UBound(InputArray, 1)
                For j = LBound(InputArray, 2) To UBound(InputArray, 2)
                    SetVariableToDefault InputArray(i, j)
                Next
            Next
        Case 3
            For i = LBound(InputArray, 1) To UBound(InputArray, 1)
                For j = LBound(InputArray, 2) To UBound(InputArray, 2)
                    For k = LBound(InputArray, 3) To UBound(InputArray, 3)
                        SetVariableToDefault InputArray(i, j, k)
                    Next
                Next
            Next
        Case Else
            Exit Function
    End Select
    
    ResetVariantArrayToDefaults = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ReverseVectorInPlace
'This procedure reverses the order of an array in place -- this is, the array
'variable in the calling procedure is reversed. This works only on
'single-dimensional arrays of simple data types ('String', 'Single', 'Double',
''Integer', 'Long'). It will not work on arrays of objects. Use
''ReverseVectorOfObjectsInPlace' to reverse an array of objects.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'2do:
'- combine with 'ReverseVectorOfObjectsInPlace'?
Public Function ReverseVectorInPlace( _
    ByRef InputVector As Variant _
        ) As Boolean
    
    Dim Temp As Variant
    Dim Ndx As Long
    Dim Ndx2 As Long
    Dim LBoundArr As Long
    Dim UBoundArr As Long
    Dim NoOfElements As Long
    Dim MidPoint As Long
    
    
    'Set the default return value
    ReverseVectorInPlace = False
    
    If Not IsArray(InputVector) Then Exit Function
    If NumberOfArrayDimensions(InputVector) <> 1 Then Exit Function
    
    LBoundArr = LBound(InputVector)
    UBoundArr = UBound(InputVector)
    NoOfElements = UBoundArr - LBoundArr + 1
    
    'calculate midpoint index of 'InputVector'
    MidPoint = LBoundArr + (NoOfElements \ 2) - 1
    
    'initialize 'Ndx2'
    Ndx2 = UBoundArr
    
    For Ndx = LBoundArr To MidPoint
        'swap the elements
        Temp = InputVector(Ndx)
        InputVector(Ndx) = InputVector(Ndx2)
        InputVector(Ndx2) = Temp
        'decrement the upper index
        Ndx2 = Ndx2 - 1
    Next
    
    ReverseVectorInPlace = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ReverseVectorOfObjectsInPlace
'This procedure reverses the order of an array in place -- this is, the array
'variable in the calling procedure is reversed. This works only with arrays of
'objects. It does not work on simple variables. Use 'ReverseVectorInPlace' for
'simple variables. An error will occur if an element of the array is not an
'object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ReverseVectorOfObjectsInPlace( _
    ByRef InputVector As Variant _
        ) As Boolean
    
    Dim Temp As Variant
    Dim Ndx As Long
    Dim Ndx2 As Long
    Dim LBoundArr As Long
    Dim UBoundArr As Long
    Dim NoOfElements As Long
    Dim MidPoint As Long
    
    
    'Set the default return value
    ReverseVectorOfObjectsInPlace = False
    
    If Not IsArray(InputVector) Then Exit Function
    If NumberOfArrayDimensions(InputVector) <> 1 Then Exit Function
    If Not IsArrayObjects(InputVector, True) Then Exit Function
    
    LBoundArr = LBound(InputVector)
    UBoundArr = UBound(InputVector)
    NoOfElements = UBoundArr - LBoundArr + 1
    
    'calculate midpoint index of 'InputVector'
    MidPoint = LBoundArr + (NoOfElements \ 2) - 1
    
    Ndx2 = UBoundArr
    
    For Ndx = LBoundArr To MidPoint
        'swap the elements
        Set Temp = InputVector(Ndx)
        Set InputVector(Ndx) = InputVector(Ndx2)
        Set InputVector(Ndx2) = Temp
        'decrement the upper index
        Ndx2 = Ndx2 - 1
    Next
    
    ReverseVectorOfObjectsInPlace = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SetObjectArrayToNothing
'This sets all the elements of 'InputArray' to 'Nothing'. Use this function
'rather than 'Erase' because if 'InputArray' is an array of 'Variants', 'Erase'
'will set each element to 'Empty', not 'Nothing', and the element will cease
'to be an object.
'The function returns 'True' if successful, 'False' otherwise.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SetObjectArrayToNothing( _
    ByRef InputArray As Variant _
        ) As Boolean
    
    Dim NoOfArrayDimensions As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    
    'Set the default return value
    SetObjectArrayToNothing = False
    
    If Not IsArray(InputArray) Then Exit Function
    
    NoOfArrayDimensions = NumberOfArrayDimensions(InputArray)
    
    If NoOfArrayDimensions < 1 Then Exit Function
    If NoOfArrayDimensions > 3 Then Exit Function
    If Not IsArrayObjects(InputArray, True) Then Exit Function
    
    'Set each element of 'InputArray' to 'Nothing'
    Select Case NoOfArrayDimensions
        Case 1
            For i = LBound(InputArray) To UBound(InputArray)
                Set InputArray(i) = Nothing
            Next
        Case 2
            For i = LBound(InputArray, 1) To UBound(InputArray, 1)
                For j = LBound(InputArray, 2) To UBound(InputArray, 2)
                    Set InputArray(i, j) = Nothing
                Next
            Next
        Case 3
            For i = LBound(InputArray, 1) To UBound(InputArray, 1)
                For j = LBound(InputArray, 2) To UBound(InputArray, 2)
                    For k = LBound(InputArray, 3) To UBound(InputArray, 3)
                        Set InputArray(i, j, k) = Nothing
                    Next
                Next
            Next
        Case Else
            Exit Function
    End Select
    
    SetObjectArrayToNothing = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SetVariableToDefault
'This procedure sets 'Variable' to the appropriate default value for its data
'type. Note that it cannot change User-Defined Types.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetVariableToDefault( _
    ByRef Variable As Variant _
)
    
    Dim LongLongType As Byte
    LongLongType = DeclareLongLong
    
    
    'We test with 'IsObject' here so that the object itself, not the default
    'property of the object, is evaluated.
    If IsObject(Variable) Then
        Set Variable = Nothing
    Else
        Select Case VarType(Variable)
            Case Is >= vbArray
                'The 'VarType' of an array is equal to
                '  vbArray + VarType(ArrayElement).
                'Here we check for anything '>=vbArray'
                Erase Variable
            Case vbBoolean
                Variable = False
            Case vbByte
                Variable = CByte(0)
            Case vbCurrency
                Variable = CCur(0)
            Case vbDataObject
'---
'2do: how can this be set/tested?
                Set Variable = Nothing
'---
            Case vbDate
                Variable = CDate(0)
            Case vbDecimal
                Variable = CDec(0)
            Case vbDouble
                Variable = CDbl(0)
            Case vbEmpty
                Variable = Empty
            Case vbError
                Variable = Empty
            Case vbInteger
                Variable = CInt(0)
            Case vbLong, LongLongType
                Variable = CLngPtr(0)
            Case vbNull
                Variable = Empty
            Case vbObject
'---
'2do: this was already checked above
                Set Variable = Nothing
'---
            Case vbSingle
                Variable = CSng(0)
            Case vbString
                Variable = vbNullString
            Case vbUserDefinedType
                'User-Defined-Types cannot be set to a general default value.
                'Each element must be explicitly set to its default value. No
                'assignment takes place in this procedure.
            Case vbVariant
                'This case is included for consistency, but we will never get
                'here. If the 'Variant' contains data, 'VarType' returns the type
                'of that data. An empty 'Variant' is type 'vbEmpty'.
                Variable = Empty
        End Select
    End If
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SwapArrayColumns
'This function returns an array based on 'Arr' with 'Col1' and 'Col2' swapped.
'It returns the result array or 'Null' if an error occurred.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SwapArrayColumns( _
    ByRef Arr As Variant, _
    ByVal Col1 As Long, _
    ByVal Col2 As Long _
        ) As Variant
    
    Dim Temp As Variant
    Dim Result As Variant
    Dim RowNdx As Long
    
    
    'Set the default return value
    SwapArrayColumns = Null
    
    If Not IsArray(Arr) Then Exit Function
    If NumberOfArrayDimensions(Arr) <> 2 Then Exit Function
    
    'Ensure 'Col1' and 'Col2' are valid column numbers
    If Col1 < LBound(Arr, 2) Then Exit Function
    If Col1 > UBound(Arr, 2) Then Exit Function
    If Col2 < LBound(Arr, 2) Then Exit Function
    If Col2 > UBound(Arr, 2) Then Exit Function
    
    'If 'Col1 = Col2', just return the array and exit. Nothing to do.
    If Col1 = Col2 Then
        SwapArrayColumns = Arr
        Exit Function
    End If
    
    'Set 'Result' to 'Arr'
    Result = Arr
    
    'ReDim 'Temp' to the number of columns
    ReDim Temp(LBound(Arr, 1) To UBound(Arr, 1))
    For RowNdx = LBound(Arr, 1) To UBound(Arr, 1)
        Temp(RowNdx) = Arr(RowNdx, Col1)
        Result(RowNdx, Col1) = Arr(RowNdx, Col2)
        Result(RowNdx, Col2) = Temp(RowNdx)
    Next
    
    SwapArrayColumns = Result
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SwapArrayRows
'This function returns an array based on 'Arr' with 'Row1' and 'Row2' swapped.
'It returns the result array or 'Null' if an error occurred.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SwapArrayRows( _
    ByRef Arr As Variant, _
    ByVal Row1 As Long, _
    ByVal Row2 As Long _
        ) As Variant
    
    Dim Temp As Variant
    Dim Result As Variant
    Dim ColNdx As Long
    
    
    'Set the default return value
    SwapArrayRows = Null
    
    If Not IsArray(Arr) Then Exit Function
    If NumberOfArrayDimensions(Arr) <> 2 Then Exit Function
    
    'Ensure 'Row1' and 'Row2' are valid numbers
    If Row1 < LBound(Arr, 1) Then Exit Function
    If Row1 > UBound(Arr, 1) Then Exit Function
    If Row2 < LBound(Arr, 1) Then Exit Function
    If Row2 > UBound(Arr, 1) Then Exit Function
    
    'If 'Row1 = Row2', just return the array and exit. Nothing to do.
    If Row1 = Row2 Then
        SwapArrayRows = Arr
        Exit Function
    End If
    
    'Set 'Result' to 'Arr'
    Result = Arr
    
    'ReDim 'Temp' to the number of columns
    ReDim Temp(LBound(Arr, 2) To UBound(Arr, 2))
    For ColNdx = LBound(Arr, 2) To UBound(Arr, 2)
        Temp(ColNdx) = Arr(Row1, ColNdx)
        Result(Row1, ColNdx) = Arr(Row2, ColNdx)
        Result(Row2, ColNdx) = Temp(ColNdx)
    Next
    
    SwapArrayRows = Result
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'TransposeArray
'This transposes a two-dimensional array. It returns 'True' if successful or
''False' if an error occurs. 'SourceArr' must be two-dimensional. 'ResultArr'
'must be a dynamic array. It will be erased and resized, so any existing
'content will be destroyed.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TransposeArray( _
    ByVal SourceArr As Variant, _
    ByRef ResultArr As Variant _
        ) As Boolean
    
    Dim RowNdx As Long
    Dim ColNdx As Long
    Dim LB1 As Long
    Dim LB2 As Long
    Dim UB1 As Long
    Dim UB2 As Long
    
    
    'Set the default return value
    TransposeArray = False
    
    If Not IsArray(SourceArr) Then Exit Function
    If NumberOfArrayDimensions(SourceArr) <> 2 Then Exit Function
    If Not IsArrayDynamic(ResultArr) Then Exit Function
    
    'Get the Lower and Upper bounds of 'SourceArr'
    LB1 = LBound(SourceArr, 1)
    LB2 = LBound(SourceArr, 2)
    UB1 = UBound(SourceArr, 1)
    UB2 = UBound(SourceArr, 2)
    
    'Erase and 'ReDim ResultArr'
    'Note the that the 'LBound' and 'UBound' values are preserved.
    Erase ResultArr
    ReDim ResultArr(LB2 To UB2, LB1 To UB1)
    'Loop through the elements of 'SourceArr' and put each value in the proper
    'element of the transposed array
    For RowNdx = LB2 To UB2
        For ColNdx = LB1 To UB1
            ResultArr(RowNdx, ColNdx) = SourceArr(ColNdx, RowNdx)
        Next
    Next
    
    TransposeArray = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VectorsToArray
'This function takes one or more single-dimensional arrays (vectors) and
'converts them into a single two-dimensional array. Each array in 'Vectors'
'comprises one row of the new array. The number of columns in the new array is
'the maximum of the number of elements in each vector.
''Arr' MUST be a dynamic array of a data type compatible with ALL the
'elements in each vector. The code does NOT trap for an error
'13 - Type Mismatch.
'If the 'Vectors' are of differing sizes, 'Arr' is sized to hold the maximum
'number of elements in a vector. The procedure erases the 'Arr' array, so when
'it is reallocated with 'ReDim', all elements will be the reset to their
'default value ('0', 'vbNullString' or 'Empty').
'Unused elements in the new array will remain the default value for that data
'type.
'Each vector in 'Vectors' must be a single-dimensional array, but the vectors
'may be of different sizes and 'LBound's.
'Each element in each vector must be a simple data type. The elements may NOT
'be 'Object's, 'Array's, or 'User-Defined Types'.
'The rows and columns of the result array are 0-based, regardless of
'the 'LBound' of each vector and regardless of the 'Option Base' statement.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function VectorsToArray( _
    ByRef Arr As Variant, _
    ParamArray Vectors() _
        ) As Boolean
    
    Dim Vector As Variant
    Dim NumRows As Long
    Dim NumCols As Long
    Dim NoOfElements As Long
    Dim LBoundVector As Long
    Dim RowNdx As Long
    Dim ColNdx As Long
    Dim VType As VbVarType
    
    Dim LongLongType As Byte
    LongLongType = DeclareLongLong
    
    
    'Set the default return value
    VectorsToArray = False
    
    If Not IsArrayDynamic(Arr) Then Exit Function
    
    'Ensure that at least one vector was passed in 'Vectors'
    If IsMissing(Vectors) Then Exit Function
    
    NumRows = 0
    NumCols = 0
    
    'Loop through 'Vectors' to determine the size of the result array.
    '(We do this loop first to prevent having to do a 'ReDim Preserve'. This
    ' requires looping through 'Vectors' a second time, but this is still faster
    ' than doing 'ReDim Preserve's.)
    For Each Vector In Vectors
        If Not IsArray(Vector) Then Exit Function
        If NumberOfArrayDimensions(Vector) <> 1 Then Exit Function
        
        'Increment the number of rows. Each 'Vector' is one row or the result array.
        NumCols = NumCols + 1
        
        LBoundVector = LBound(Vector)
        
        'Store number of elements in 'Vector' and use the larger value for
        ''NumRows'.
        NoOfElements = UBound(Vector) - LBoundVector + 1
        NumRows = Application.WorksheetFunction.Max(NumRows, NoOfElements)
    Next
    
    'ReDim 'Arr' to the appropriate size. 'Arr' is 0-based in both directions,
    'regardless of the 'LBound' of the original 'Arr' and regardless of the
    ''LBounds' of the 'Vectors'.
    ReDim Arr(0 To NumRows - 1, 0 To NumCols - 1)
    
    For ColNdx = 0 To NumCols - 1
        For RowNdx = 0 To NumRows - 1
            'Set 'Vector' (a Variant) to the 'Vectors(ColNdx)' array. We declare
            ''Vector' as a variant so it can take an array of any simple data type.
            Vector = Vectors(ColNdx)
            
            LBoundVector = LBound(Vector)
            
            VType = VarType(Vector(LBoundVector + RowNdx))
            'define allowed data types and exit function for all others
            Select Case VType
                Case vbByte, vbInteger, vbLong, LongLongType, vbSingle, vbDouble, vbDate, vbCurrency, vbDecimal
                Case vbString
'                Case vbArray, vbVariant, vbEmpty, vbError, vbNull, vbUserDefinedType
'                    Exit Function
                Case vbBoolean
'                Case vbObject
'                    Exit Function
                Case Else
                    Exit Function
            End Select
            'transfer value to 'Arr'
            Arr(RowNdx, ColNdx) = Vector(LBoundVector + RowNdx)
        Next
    Next
    
    VectorsToArray = True
    
End Function


'------------------------------------------------------------------------------

'2do:
'- add to upper list
'- add some parameter checking
'- add unit tests
Public Function VectorTo1DArray( _
    ByVal InputVector As Variant, _
    Optional ByVal LowerBoundOfSecondDimension As Long = 0 _
        ) As Variant
    
    Dim ResultArray() As Variant
    Dim i As Long
    
    
    ReDim ResultArray(LBound(InputVector) To UBound(InputVector), LowerBoundOfSecondDimension To LowerBoundOfSecondDimension)
    For i = LBound(InputVector) To UBound(InputVector)
        ResultArray(i, LowerBoundOfSecondDimension) = InputVector(i)
    Next
    
    VectorTo1DArray = ResultArray
    
End Function
