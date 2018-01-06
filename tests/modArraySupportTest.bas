Attribute VB_Name = "modArraySupportTest"

Option Explicit
Option Compare Text
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.PermissiveAssertClass
Private Fakes As Rubberduck.FakesProvider


'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.PermissiveAssertClass
    Set Fakes = New Rubberduck.FakesProvider
    
    PrepareWorkbook4Tests
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


'==============================================================================
'unit tests for 'CompareArrays'
'==============================================================================

'@TestMethod
Public Sub CompareArrays_UnallocatedArrays_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Arr1() As String
    Dim Arr2() As String
    Dim ResArr() As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport.CompareArrays(Arr1, Arr2, ResArr)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CompareArrays_LegalAndTextCompare_ReturnsTrueAndResArr()
    On Error GoTo TestFail
    
    Dim Arr1(1 To 5) As String
    Dim Arr2(1 To 5) As String
    Dim ResArr() As Long
    Dim i As Long
    
    '==========================================================================
    Dim aExpected(1 To 5) As Long
        aExpected(1) = -1
        aExpected(2) = 1
        aExpected(3) = -1
        aExpected(4) = 0
        aExpected(5) = 0
    '==========================================================================
    
    
    'Arrange:
    Arr1(1) = "2"
    Arr1(2) = "c"
    Arr1(3) = vbNullString
    Arr1(4) = "."
    Arr1(5) = "B"
    
    Arr2(1) = "4"
    Arr2(2) = "a"
    Arr2(3) = "x"
    Arr2(4) = "."
    Arr2(5) = "b"
    
    'Act:
    If Not modArraySupport.CompareArrays(Arr1, Arr2, ResArr, vbTextCompare) _
            Then GoTo TestFail
    
    'Assert:
    For i = LBound(ResArr) To UBound(ResArr)
        Assert.AreEqual CLng(aExpected(i)), ResArr(i)
    Next
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CompareArrays_LegalAndBinaryCompare_ReturnsTrueAndResArr()
    On Error GoTo TestFail
    
    Dim Arr1(1 To 5) As String
    Dim Arr2(1 To 5) As String
    Dim ResArr() As Long
    Dim i As Long
    
    '==========================================================================
    Dim aExpected(1 To 5) As Long
        aExpected(1) = -1
        aExpected(2) = 1
        aExpected(3) = -1
        aExpected(4) = 0
        aExpected(5) = -1
    '==========================================================================
    
    
    'Arrange:
    Arr1(1) = "2"
    Arr1(2) = "c"
    Arr1(3) = vbNullString
    Arr1(4) = "."
    Arr1(5) = "B"
    
    Arr2(1) = "4"
    Arr2(2) = "a"
    Arr2(3) = "x"
    Arr2(4) = "."
    Arr2(5) = "b"
    
    'Act:
    If Not modArraySupport.CompareArrays(Arr1, Arr2, ResArr, vbBinaryCompare) _
            Then GoTo TestFail
    
    'Assert:
    For i = LBound(ResArr) To UBound(ResArr)
        Assert.AreEqual CLng(aExpected(i)), ResArr(i)
    Next
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'ConcatenateArrays'
'==============================================================================

'@TestMethod
Public Sub ConcatenateArrays_StaticResultArray_ResultsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ResultArray(1) As Long          'MUST be dynamic
    Dim ArrayToAppend(1) As Long
    
    
    ResultArray(1) = 8
    ArrayToAppend(1) = 111
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport.ConcatenateArrays(ResultArray, ArrayToAppend)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ConcatenateArrays_BothArraysUnallocated_ResultsTrueAndUnallocatedArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ResultArray() As Long           'MUST be dynamic
    Dim ArrayToAppend() As Long
    
    
    'Act:
    If Not modArraySupport.ConcatenateArrays(ResultArray, ArrayToAppend) Then _
            GoTo TestFail
    
    'Assert:
    Assert.IsFalse IsArrayAllocated(ResultArray)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ConcatenateArrays_ArrayToAppendUnallocated_ResultsTrueAndUnchangedResultArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ResultArray() As Long           'MUST be dynamic
    Dim ArrayToAppend() As Long
    Dim i As Long
    
    '==========================================================================
    Dim aExpected(1 To 2) As Long
        aExpected(1) = 8
        aExpected(2) = 9
    '==========================================================================
    
    
    ReDim ResultArray(1 To 2)
    ResultArray(1) = 8
    ResultArray(2) = 9
    
    'Act:
    If Not modArraySupport.ConcatenateArrays(ResultArray, ArrayToAppend) Then _
            GoTo TestFail
    
    'Assert:
    For i = LBound(ResultArray) To UBound(ResultArray)
        Assert.AreEqual CLng(aExpected(i)), CLng(ResultArray(i))
    Next
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ConcatenateArrays_LegalLong_ResultsTrueAndResultArray()
    On Error GoTo TestFail
    
    Dim ResultArray() As Long           'MUST be dynamic
    Dim ArrayToAppend(1 To 3) As Integer
    Dim i As Long
    
    '==========================================================================
    Dim aExpected(1 To 6) As Long
        aExpected(1) = 8
        aExpected(2) = 9
        aExpected(3) = 10
        aExpected(4) = 111
        aExpected(5) = 112
        aExpected(6) = 113
    '==========================================================================
    
    
    'Arrange:
    ReDim ResultArray(1 To 3)
    ResultArray(1) = 8
    ResultArray(2) = 9
    ResultArray(3) = 10
    
    ArrayToAppend(1) = 111
    ArrayToAppend(2) = 112
    ArrayToAppend(3) = 113
    
    'Act:
    If Not modArraySupport.ConcatenateArrays(ResultArray, ArrayToAppend) Then _
            GoTo TestFail
    
    'Assert:
    For i = LBound(ResultArray) To UBound(ResultArray)
        Assert.AreEqual CLng(aExpected(i)), CLng(ResultArray(i))
    Next
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


''2do: add a test that involves objects
''     (have a look at <https://stackoverflow.com/a/11254505>
''@TestMethod
'Public Sub ConcatenateArrays_LegalVariant_ResultsTrueAndResultArray()
'    On Error GoTo TestFail
'
'    Dim ResultArray() As Range          'MUST be dynamic
'    Dim ArrayToAppend(0 To 0) As Range
'    Dim i As Long
'
'    '==========================================================================
'    Dim wks As Worksheet
'    Set wks = tblFunctions
'    Dim aExpected(1 To 2) As Range
'    With wks
'        Set aExpected(1) = .Cells(1, 1)
'        Set aExpected(2) = .Cells(1, 2)
'    End With
'    '==========================================================================
'
'
'    'Arrange:
'    With wks
'        ReDim ResultArray(1 To 1)
'        Set ResultArray(1) = .Cells(1, 1)
'        Set ArrayToAppend(0) = .Cells(1, 2)
'    End With
'
'    'Act:
'    If Not modArraySupport.ConcatenateArrays(ResultArray, ArrayToAppend) Then _
'            GoTo TestFail
'
'    'Assert:
'    For i = LBound(ResultArray) To UBound(ResultArray)
'Debug.Print aExpected(i) Is ResultArray(i)
'        Assert.AreSame aExpected(i), ResultArray(i)
'    Next
'
''    If B = True Then
''        If modArraySupport.IsArrayAllocated(ResultArray) = True Then
''            For i = LBound(ResultArray) To UBound(ResultArray)
''                If IsObject(ResultArray(i)) = True Then
''Debug.Print CStr(i), "is object", TypeName(ResultArray(i))
''                Else
''Debug.Print CStr(i), ResultArray(i)
''                End If
''            Next
''        Else
''Debug.Print "Result Array Is Not Allocated."
''        End If
''    Else
''Debug.Print "ConcatenateArrays returned False"
''    End If
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub


'==============================================================================
'unit tests for 'CopyArray'
'==============================================================================

'@TestMethod
Public Sub CopyArray_UnallocatedSrc_ResultsTrueAndUnchangedDest()
    On Error GoTo TestFail
    
    Dim Src() As Long
    Dim Dest(0) As Integer
    
    '==========================================================================
    Dim aExpected(0) As Integer
        aExpected(0) = 50
    '==========================================================================
    
    
    'Arrange:
    Dest(0) = 50
    
    'Act:
    If Not modArraySupport.CopyArray(Src, Dest) Then _
            GoTo TestFail
    
    'Assert:
    Assert.AreEqual aExpected(0), Dest(0)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyArray_IncompatibleDest_ResultsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Src(1 To 2) As Long
    Dim Dest(1 To 2) As Integer
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport.CopyArray(Src, Dest)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyArray_AllocatedDestLessElementsThenSrc_ResultsTrueAndDestArray()
    On Error GoTo TestFail
    
    Dim Src(1 To 3) As Long
    Dim Dest(10 To 11) As Long
    Dim i As Long
    
    '==========================================================================
    Dim aExpected(10 To 11) As Long
        aExpected(10) = 1
        aExpected(11) = 2
    '==========================================================================
    
    
    'Arrange:
    Src(1) = 1
    Src(2) = 2
    Src(3) = 3
    
    'Act:
    If Not modArraySupport.CopyArray(Src, Dest) Then _
            GoTo TestFail
    
    'Assert:
    For i = LBound(Dest) To UBound(Dest)
        Assert.AreEqual CLng(aExpected(i)), CLng(Dest(i))
    Next
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyArray_AllocatedDestMoreElementsThenSrc_ResultsTrueAndDestArray()
    On Error GoTo TestFail
    
    Dim Src(1 To 3) As Long
    Dim Dest(10 To 13) As Long
    Dim i As Long
    
    '==========================================================================
    Dim aExpected(10 To 13) As Long
        aExpected(10) = 1
        aExpected(11) = 2
        aExpected(12) = 3
        aExpected(13) = 0
    '==========================================================================
    
    
    'Arrange:
    Src(1) = 1
    Src(2) = 2
    Src(3) = 3
    
    'Act:
    If Not modArraySupport.CopyArray(Src, Dest) Then _
            GoTo TestFail
    
    'Assert:
    For i = LBound(Dest) To UBound(Dest)
        Assert.AreEqual CLng(aExpected(i)), CLng(Dest(i))
    Next
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyArray_NoCompatibilityCheck_ResultsTrueAndDestArrayWithOverflow()
    On Error GoTo TestFail
    
    Dim Src(1 To 2) As Long
    Dim Dest(1 To 2) As Integer
    Dim i As Long
    
    '==========================================================================
    Dim aExpected(1 To 2) As Integer
        aExpected(1) = 1234
        aExpected(2) = 0
    '==========================================================================
    
    
    'Arrange:
    Src(1) = 1234
    Src(2) = 655360
    
    'Act:
    If Not modArraySupport.CopyArray(Src, Dest, True) Then _
            GoTo TestFail
    
    'Assert:
    For i = LBound(Dest) To UBound(Dest)
        Assert.AreEqual aExpected(i), Dest(i)
    Next
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'2do: Add tests with Objects


Public Sub DemoCopyArraySubSetToArray()

    Dim InputArray(1 To 10) As Long
    Dim ResultArray() As Long

    Dim StartNdx As Long
    Dim EndNdx As Long
    Dim DestNdx As Long
    Dim B As Boolean
    Dim N As Long
    
    
    For N = LBound(InputArray) To UBound(InputArray)
        InputArray(N) = N * 10
    Next N

    ReDim ResultArray(1 To 10)
    For N = LBound(ResultArray) To UBound(ResultArray)
        ResultArray(N) = -N
    Next N

    StartNdx = 1
    EndNdx = 5
    DestNdx = 3
    
    B = modArraySupport.CopyArraySubSetToArray( _
            InputArray, ResultArray, StartNdx, EndNdx, DestNdx)
    
    If B = True Then
        If modArraySupport.IsArrayAllocated(ResultArray) = True Then
            For N = LBound(ResultArray) To UBound(ResultArray)
                If IsObject(ResultArray(N)) = True Then
Debug.Print CStr(N), "is object"
                Else
Debug.Print CStr(N), ResultArray(N)
                End If
            Next N
        Else
Debug.Print "ResultArray is not allocated"
        End If
    Else
Debug.Print "CopyArraySubSetToArray returned False"
    End If

End Sub


Public Sub DemoCopyNonNothingObjectsToArray()

    Dim SourceArray(1 To 5) As Object
    Dim ResultArray() As Object
    Dim B As Boolean
    Dim N As Long
    
    
    Set SourceArray(1) = Range("a1")
    Set SourceArray(2) = Range("A2")
    Set SourceArray(3) = Nothing
    Set SourceArray(4) = Nothing
    Set SourceArray(5) = Range("A5")
    
    B = modArraySupport.CopyNonNothingObjectsToArray(SourceArray, ResultArray, False)
    
    If B = True Then
        For N = LBound(ResultArray) To UBound(ResultArray)
Debug.Print CStr(N), ResultArray(N).Address
        Next N
    Else
Debug.Print "CopyNonNothingObjectsToArray returned False"
    End If

End Sub


Public Sub DemoDataTypeOfArray()

    Dim A(1 To 4) As String
    Dim T As VbVarType


    T = modArraySupport.DataTypeOfArray(A)
Debug.Print T

End Sub


Public Sub DemoDeleteArrayElement()

    Dim Stat(1 To 3) As Long
    Dim Dyn() As Variant
    Dim N As Long
    Dim B As Boolean


    ReDim Dyn(1 To 3)
    Stat(1) = 1
    Stat(2) = 2
    Stat(3) = 3
    Dyn(1) = "abc"
    Dyn(2) = 1234
    Dyn(3) = "ABC"

    B = modArraySupport.DeleteArrayElement(Stat, 1, False)
    
    If B = True Then
        For N = LBound(Stat) To UBound(Stat)
Debug.Print CStr(N), Stat(N)
        Next N
    Else
Debug.Print "DeleteArrayElement returned false"
    End If


    B = modArraySupport.DeleteArrayElement(Dyn, 2, False)
    
    If B = True Then
        For N = LBound(Dyn) To UBound(Dyn)
Debug.Print CStr(N), Dyn(N)
        Next N
    Else
Debug.Print "DeleteArrayElement returned false"
    End If

End Sub


Public Sub DemoFirstNonEmptyStringIndexInArray()

    Dim A(1 To 4) As String
    Dim R As Long
    
    
    A(1) = vbNullString
    A(2) = vbNullString
    A(3) = "A"
    A(4) = "B"
    
    R = modArraySupport.FirstNonEmptyStringIndexInArray(A)
Debug.Print "FirstNonEmptyStringIndexInArray", CStr(R)

End Sub


Public Sub DemoInsertElementIntoArray()

    Dim Arr() As Long
    Dim N As Long
    Dim B As Boolean
    
    
    ReDim Arr(1 To 10)
    For N = LBound(Arr) To UBound(Arr)
        Arr(N) = N * 10
    Next N

    B = modArraySupport.InsertElementIntoArray(Arr, 5, 12345)
    
    If B = True Then
        For N = LBound(Arr) To UBound(Arr)
Debug.Print CStr(N), Arr(N)
        Next N
    Else
Debug.Print "InsertElementIntoArray returned false."
    End If

End Sub


Public Sub DemoIsArrayAllDefault()

    Dim L(1 To 4) As Long
    Dim Obj(1 To 4) As Object
    Dim B As Boolean


    B = modArraySupport.IsArrayAllDefault(L)
Debug.Print "IsArrayAllDefault L", B

    B = modArraySupport.IsArrayAllDefault(Obj)
Debug.Print "IsArrayAllDefault Obj", B

    Set Obj(1) = Range("A1")
    B = modArraySupport.IsArrayAllDefault(Obj)
Debug.Print "IsArrayAllDefault Obj", B

End Sub


Public Sub DemoIsArrayAllNumeric()
    
    Dim V(1 To 3) As Variant
    Dim B As Boolean
    
    
    V(1) = "abc"
    V(2) = 2
    V(3) = Empty
    
    B = modArraySupport.IsArrayAllNumeric(V, True)
Debug.Print "IsArrayAllNumeric:", B

End Sub


Public Sub DemoIsArrayAllocated()
    
    Dim B As Boolean
    Dim AllocArray(1 To 3) As Variant
    Dim UnAllocArray() As Variant
    
    
    B = modArraySupport.IsArrayAllocated(AllocArray)
Debug.Print "IsArrayAllocated AllocArray:", B

    B = modArraySupport.IsArrayAllocated(UnAllocArray)
Debug.Print "IsArrayAllocated UnAllocArray:", B

End Sub


Public Sub DemoIsArrayDynamic()
    
    Dim B As Boolean
    Dim StaticArray(1 To 3) As Long
    Dim DynArray() As Long
    
    
    ReDim DynArray(1 To 3)
    
    B = modArraySupport.IsArrayDynamic(StaticArray)
Debug.Print "IsArrayDynamic StaticArray:", B

    B = modArraySupport.IsArrayDynamic(DynArray)
Debug.Print "IsArrayDynamic DynArray:", B

End Sub


Public Sub DemoIsArrayEmpty()

    Dim EmptyArray() As Long
    Dim NonEmptyArray() As Long
    Dim B As Boolean

    
    ReDim NonEmptyArray(1 To 3)
    
    B = modArraySupport.IsArrayEmpty(EmptyArray)
Debug.Print "IsArrayEmpty: EmptyArray:", B

    B = modArraySupport.IsArrayEmpty(NonEmptyArray)
Debug.Print "IsArrayEmpty: NonEmptyArray:", B

End Sub


Public Sub DemoIsArrayObjects()
    
    Dim V(1 To 3) As Variant
    Dim B As Boolean
    
    
    Set V(1) = Nothing
    Set V(2) = Range("A1")
    V(3) = Range("A1")
    
    B = modArraySupport.IsArrayObjects(V, True)
Debug.Print "IsArrayObjects With AllowNothing = True:", B

    B = modArraySupport.IsArrayObjects(V, False)
Debug.Print "IsArrayObjects With AllowNothing = False:", B

End Sub


Public Sub DemoIsNumericDataType()
    
    Dim V As Variant
    Dim VEmpty As Variant
    Dim S As String
    Dim B As Boolean
    
    
    V = 123
    S = "123"

    B = modArraySupport.IsNumericDataType(V)
Debug.Print "IsNumericDataType:", B

    B = modArraySupport.IsNumericDataType(S)
Debug.Print "IsNumericDataType:", B

    B = modArraySupport.IsNumericDataType(VEmpty)
Debug.Print "IsNumericDataType:", B

    V = Array(1, 2, 3)
    B = modArraySupport.IsNumericDataType(V)
Debug.Print "IsNumericDataType:", B

    V = Array("a", "b", "c")
    B = modArraySupport.IsNumericDataType(V)
Debug.Print "IsNumericDataType:", B
    
End Sub


Public Sub DemoIsVariantArrayConsistent()

    Dim B As Boolean
    Dim V(1 To 3) As Variant
    
    
    Set V(1) = Range("A1")
    Set V(2) = Nothing
    Set V(3) = Range("A3")

    B = modArraySupport.IsVariantArrayConsistent(V)
Debug.Print "IsVariantArrayConsistent:", B

End Sub


Public Sub DemoIsVariantArrayNumeric()

    Dim B As Boolean
    Dim V(1 To 3) As Variant
    
    
    V(1) = 123
    Set V(2) = Range("A1")
    V(3) = 789
    
    B = modArraySupport.IsVariantArrayNumeric(V)
Debug.Print "IsVariantArrayNumeric", B

End Sub


Public Sub DemoMoveEmptyStringsToEndOfArray()
    
    Dim B As Boolean
    Dim N As Long
    Dim S(1 To 5) As String
    
    
    S(1) = vbNullString
    S(2) = vbNullString
    S(3) = "C"
    S(4) = "D"
    S(5) = "E"
    
    B = modArraySupport.MoveEmptyStringsToEndOfArray(S)
    
    If B = True Then
        For N = LBound(S) To UBound(S)
            If S(N) = vbNullString Then
Debug.Print CStr(N), "is vbNullString"
            Else
Debug.Print CStr(N), S(N)
            End If
        Next N
    Else
Debug.Print "MoveEmptyStringsToEndOfArray returned False"
    End If

End Sub


Public Sub DemoNumberOfArrayDimensions()

    Dim EmptyArray() As Long
    Dim OneArray(1 To 3) As Long
    Dim ThreeArray(1 To 3, 1 To 2, 1 To 1)
    Dim N As Long


    N = modArraySupport.NumberOfArrayDimensions(EmptyArray)
Debug.Print "NumberOfArrayDimensions EmptyArray", N

    N = modArraySupport.NumberOfArrayDimensions(OneArray)
Debug.Print "NumberOfArrayDimensions OneArray", N

    N = modArraySupport.NumberOfArrayDimensions(ThreeArray)
Debug.Print "NumberOfArrayDimensions ThreeArray", N

End Sub


Public Sub DemoNumElements()

    Dim N As Long
    Dim EmptyArray() As Long
    Dim OneArray(1 To 3) As Long
    Dim ThreeArray(1 To 3, 1 To 2, 1 To 1)


    N = modArraySupport.NumElements(EmptyArray, 1)
Debug.Print "NumElements EmptyArray", N

    N = modArraySupport.NumElements(OneArray, 1)
Debug.Print "NumElements OneArray", N

    N = modArraySupport.NumElements(ThreeArray, 3)
Debug.Print "NumElements ThreeArray", N

End Sub


Public Sub DemoResetVariantArrayToDefaults()

    Dim V(1 To 5) As Variant
    Dim B As Boolean
    Dim N As Long


    V(1) = CInt(123)
    V(2) = "abcd"
    Set V(3) = Range("A1")
    V(4) = CDec(123)
    V(5) = Null
    
    B = modArraySupport.ResetVariantArrayToDefaults(V)
    
    If B = True Then
        For N = LBound(V) To UBound(V)
            If IsObject(V(N)) = True Then
                If V(N) Is Nothing Then
Debug.Print CStr(N), "Is Nothing"
                Else
Debug.Print CStr(N), "Is Object"
                End If
            Else
Debug.Print CStr(N), TypeName(V(N)), V(N)
            End If
        Next N
    Else
Debug.Print "ResetVariantArrayToDefaults  returned false"
    End If

End Sub


Public Sub DemoReverseArrayInPlace()

    Dim V(1 To 5) As Long
    Dim N As Long
    Dim B As Boolean
    
    
    V(1) = 1
    V(2) = 2
    V(3) = 3
    V(4) = 4
    V(5) = 5
    
    B = modArraySupport.ReverseArrayInPlace(V)
    
    If B = True Then
Debug.Print "REVERSED ARRAY --------------------------------------"
        For N = LBound(V) To UBound(V)
Debug.Print V(N)
        Next N
    End If

End Sub


Public Sub DemoReverseArrayOfObjectsInPlace()

    Dim B As Boolean
    Dim N As Long
    Dim V(1 To 5) As Object
    
    
    Set V(1) = Range("A1")
    Set V(2) = Nothing
    Set V(3) = Range("A3")
    Set V(4) = Range("A4")
    Set V(5) = Range("A5")
    
    B = modArraySupport.ReverseArrayOfObjectsInPlace(V)
    
    If B = True Then
Debug.Print "REVERSED ARRAY --------------------------------------"
        For N = LBound(V) To UBound(V)
            If V(N) Is Nothing Then
Debug.Print CStr(N), "Is Nothing"
            Else
Debug.Print CStr(N), V(N).Address
            End If
        Next N
    End If
End Sub


Public Sub DemoSetObjectArrayToNothing()
    
    Dim StaticArray(1 To 2) As Range
    Dim DynamicArray(1 To 2) As Range
    Dim B As Boolean
    Dim N As Long
    
    
    Set StaticArray(1) = Range("A1")
    Set StaticArray(2) = Nothing
    Set DynamicArray(1) = Range("A1")
    Set DynamicArray(2) = Range("A2")
    
    B = modArraySupport.SetObjectArrayToNothing(StaticArray)
    
    If B = True Then
        For N = LBound(StaticArray) To UBound(StaticArray)
            If StaticArray(N) Is Nothing Then
Debug.Print CStr(N), "is nothing "
            End If
        Next N
    End If
    
    
    B = modArraySupport.SetObjectArrayToNothing(DynamicArray)
    
    If B = True Then
        For N = LBound(DynamicArray) To UBound(DynamicArray)
            If DynamicArray(N) Is Nothing Then
Debug.Print CStr(N), "is nothing "
            End If
        Next N
    End If

End Sub


Public Sub DemoVectorsToArray()

    Dim A() As Variant
    Dim B As Boolean
    Dim R As Long
    Dim C As Long
    Dim S As String
    
    Dim AA()
    Dim BB()
    Dim CC() As String
    
    
    ReDim AA(0 To 2)
    ReDim BB(1 To 5)
    ReDim CC(2 To 5)
    
    
    AA(0) = 16
    AA(1) = 2
    AA(2) = 3
    'AA(3) = 3
    BB(1) = 11
    BB(2) = 22
    BB(3) = 33
    BB(4) = 44
    BB(5) = 55
    CC(2) = "A"
    CC(3) = "B"
    CC(4) = "C"
    CC(5) = "D"
    
    B = modArraySupport.VectorsToArray(A, AA, BB, CC)
    
    If B = True Then
        For R = LBound(A, 1) To UBound(A, 1)
            S = vbNullString
            For C = LBound(A, 2) To UBound(A, 2)
                S = S & A(R, C) & " "
            Next C
Debug.Print S
        Next R
    Else
Debug.Print "VectorsToArray Failed"
    End If

End Sub


Public Sub DemoTransposeArray()

    Dim A() As Long
    Dim B() As Long
    Dim Res As Boolean

    Dim RowNdx As Long
    Dim ColNdx As Long
    Dim S As String


    ReDim A(1 To 3, 2 To 5)
    A(1, 2) = 1
    A(1, 3) = 2
    A(1, 4) = 3
    A(1, 5) = 33
    A(2, 2) = 4
    A(2, 3) = 5
    A(2, 4) = 6
    A(2, 5) = 66
    A(3, 2) = 7
    A(3, 3) = 8
    A(3, 4) = 9
    A(3, 5) = 100

Debug.Print "LBound1: " & CStr(LBound(A, 1)) & " Ubound1: " & CStr(UBound(A, 1)), _
            "LBound2: " & CStr(LBound(A, 2)) & " UBound2: " & CStr(UBound(A, 2))

    For RowNdx = LBound(A, 1) To UBound(A, 1)
        S = vbNullString
        For ColNdx = LBound(A, 2) To UBound(A, 2)
            S = S & A(RowNdx, ColNdx) & " "
        Next ColNdx
Debug.Print S
    Next RowNdx
Debug.Print "Transposed Array:"
    
    Res = modArraySupport.TransposeArray(A, B)
Debug.Print "LBound1: " & CStr(LBound(B, 1)) & " Ubound1: " & CStr(UBound(B, 1)), _
            "LBound2: " & CStr(LBound(B, 2)) & " UBound2: " & CStr(UBound(B, 2))
    If Res = True Then
        S = vbNullString
        For RowNdx = LBound(B, 1) To UBound(B, 1)
            S = vbNullString
            For ColNdx = LBound(B, 2) To UBound(B, 2)
                S = S & B(RowNdx, ColNdx) & " "
            Next ColNdx
Debug.Print S
        Next RowNdx
    Else
Debug.Print "Error In Transpose Array"
    End If

End Sub


Public Sub DemoChangeBoundsOfArray()

    Dim NewLB As Long
    Dim NewUB As Long
    Dim B As Boolean
    Dim N As Long
    Dim M As Long
    'Dim Arr() As Range
    'Dim Arr() As Long
    'Dim Arr() As Variant
    Dim Arr() As clsDummy_4_modArraySupportTest


    ReDim Arr(5 To 7)
    'Set Arr(5) = Range("A1")
    'Set Arr(6) = Range("A2")
    'Set Arr(7) = Range("A3")
    'Arr(5) = 11
    'Arr(6) = 22
    'Arr(7) = 33
    'Arr(5) = Array(1, 2, 3)
    'Arr(6) = Array(4, 5, 6)
    'Arr(7) = Array(7, 8, 9)

    Set Arr(5) = New clsDummy_4_modArraySupportTest
    Set Arr(6) = New clsDummy_4_modArraySupportTest
    Set Arr(7) = New clsDummy_4_modArraySupportTest
    Arr(5).Name = "Name 1"
    Arr(5).Value = 1
    Arr(6).Name = "Name 2"
    Arr(6).Value = 3
    Arr(7).Name = "Name 3"
    Arr(7).Value = 3

    NewLB = 20
    NewUB = 25
    
    B = modArraySupport.ChangeBoundsOfArray(Arr, NewLB, NewUB)

Debug.Print "New LBound: " & CStr(LBound(Arr)), "New UBound: " & CStr(UBound(Arr))
    For N = LBound(Arr) To UBound(Arr)
        If IsObject(Arr(N)) = True Then
'Debug.Print "Object: " & TypeName(Arr(N))
            If Arr(N) Is Nothing Then
Debug.Print "Object Is Nothing"
            Else
'Debug.Print "Object: " & Arr(N).Name, Arr(N).Value
Debug.Print "Object: " & TypeName(Arr(N))
            End If
        Else
'            If IsArray(Arr(N)) = True Then
'                For M = LBound(Arr(N)) To UBound(Arr(N))
'                    Debug.Print Arr(N)(M)
'                Next M
'            Else
            If IsEmpty(Arr(N)) = True Then
Debug.Print "Empty"
            ElseIf Arr(N) = vbNullString Then
Debug.Print "vbNullString"
            Else
Debug.Print Arr(N)
            End If
'            End If
        End If
    Next N

End Sub


Public Sub DemoIsArraySorted()

    Dim S(1 To 3) As String
    Dim L(1 To 3) As Long
    Dim R As Variant
    Dim Desc As Boolean
    
    
    Desc = True
    S(1) = "B"
    S(2) = "B"
    S(3) = "A"

    L(1) = 1
    L(2) = 2
    L(3) = 3
    
    R = modArraySupport.IsArraySorted(S, Desc)
    
    If IsNull(R) = True Then
Debug.Print "Error From IsArraySorted"
    Else
        If R = True Then
Debug.Print "Array Is Sorted"
        Else
Debug.Print "Array is Unsorted"
        End If
    End If

End Sub


Public Sub DemoCombineTwoDArrays()

    Dim X As Long
    Dim Y As Long
    Dim N As Long
    Dim S As String
    Dim V As Variant
    Dim E As Variant
    
    Dim A() As String
    Dim B() As String
    Dim C() As String
    Dim D() As String
    
    
    'Ensure it works on 1-Based arrays
    ReDim A(1 To 2, 1 To 2)
    ReDim B(1 To 2, 1 To 2)
    A(1, 1) = "a"
    A(1, 2) = "b"
    A(2, 1) = "c"
    A(2, 2) = "d"
    B(1, 1) = "e"
    B(1, 2) = "f"
    B(2, 1) = "g"
    B(2, 2) = "h"

Debug.Print "--- 1 BASED ARRAY -----------------------"
    V = modArraySupport.CombineTwoDArrays(A, B)
    Call DebugPrint2DArray(V)

    'Ensure it works on 0-Based arrays
    ReDim A(0 To 1, 0 To 1)
    ReDim B(0 To 1, 0 To 1)
    A(0, 0) = "a"
    A(0, 1) = "b"
    A(1, 0) = "c"
    A(1, 1) = "d"
    
    B(0, 0) = "e"
    B(0, 1) = "f"
    B(1, 0) = "g"
    B(1, 1) = "h"
    
Debug.Print "--- 0 BASED ARRAY -----------------------"
    V = modArraySupport.CombineTwoDArrays(A, B)
    Call DebugPrint2DArray(V)
    
    'Ensure it works on Positive-Based arrays
    ReDim A(5 To 6, 5 To 6)
    ReDim B(5 To 6, 5 To 6)
    A(5, 5) = "a"
    A(5, 6) = "b"
    A(6, 5) = "c"
    A(6, 6) = "d"

    B(5, 5) = "e"
    B(5, 6) = "f"
    B(6, 5) = "g"
    B(6, 6) = "h"
    
Debug.Print "--- POSITIVE BASED ARRAY -----------------------"
    V = modArraySupport.CombineTwoDArrays(A, B)
    Call DebugPrint2DArray(V)
    
    'Ensure it works on Negative-Based arrays
    ReDim A(-6 To -5, -6 To -5)
    ReDim B(-6 To -5, -6 To -5)
    A(-6, -6) = "a"
    A(-6, -5) = "b"
    A(-5, -6) = "c"
    A(-5, -5) = "d"

    B(-6, -6) = "e"
    B(-6, -5) = "f"
    B(-5, -6) = "g"
    B(-5, -5) = "h"
    
Debug.Print "--- NEGATIVE BASED ARRAY -----------------------"
    V = modArraySupport.CombineTwoDArrays(A, B)
    Call DebugPrint2DArray(V)
    
    'Ensure Nesting Works
    ReDim A(1 To 2, 1 To 2)
    ReDim B(1 To 2, 1 To 2)
    ReDim C(1 To 2, 1 To 2)
    ReDim D(1 To 2, 1 To 2)
    
    A(1, 1) = "a"
    A(1, 2) = "b"
    A(2, 1) = "c"
    A(2, 2) = "d"
    
    B(1, 1) = "e"
    B(1, 2) = "f"
    B(2, 1) = "g"
    B(2, 2) = "h"

    C(1, 1) = "i"
    C(1, 2) = "j"
    C(2, 1) = "k"
    C(2, 2) = "l"
    
    D(1, 1) = "m"
    D(1, 2) = "n"
    D(2, 1) = "o"
    D(2, 2) = "p"
    
Debug.Print "--- NESTED CALLS -----------------------"
    V = modArraySupport.CombineTwoDArrays( _
            modArraySupport.CombineTwoDArrays( _
                modArraySupport.CombineTwoDArrays(A, B), _
                    C), _
    D)
    Call DebugPrint2DArray(V)

End Sub


Public Sub DebugPrint2DArray(Arr As Variant)

    Dim Y As Long
    Dim X As Long
    Dim S As String


    For Y = LBound(Arr, 1) To UBound(Arr, 1)
        S = vbNullString
        For X = LBound(Arr, 2) To UBound(Arr, 2)
            S = S & Arr(Y, X) & " "
        Next X
Debug.Print S
    Next Y

End Sub


Public Sub DemoExpandArray()

    Dim A As Variant
    Dim B As Variant
    Dim RowNdx As Long
    Dim ColNdx As Long
    Dim S As String


    'ReDim A(-5 To -3, 0 To 3)
    'A(-5, 0) = "a"
    'A(-5, 1) = "b"
    'A(-5, 2) = "c"
    'A(-5, 3) = "d"
    'A(-4, 0) = "e"
    'A(-4, 1) = "f"
    'A(-4, 2) = "g"
    'A(-4, 3) = "h"
    'A(-3, 0) = "i"
    'A(-3, 1) = "j"
    'A(-3, 2) = "k"
    'A(-3, 3) = "l"
    '

    ReDim A(1 To 2, 1 To 4)
    A(1, 1) = "a"
    A(1, 2) = "b"
    A(1, 3) = "c"
    A(1, 4) = "d"
    A(2, 1) = "e"
    A(2, 2) = "f"
    A(2, 3) = "g"
    A(2, 4) = "h"

    Dim C As Variant

Debug.Print "BEFORE:================================="
    For RowNdx = LBound(A, 1) To UBound(A, 1)
        S = vbNullString
        For ColNdx = LBound(A, 2) To UBound(A, 2)
            S = S & A(RowNdx, ColNdx) & " "
        Next ColNdx
Debug.Print S
    Next RowNdx

    S = vbNullString
    B = modArraySupport.ExpandArray(A, 1, 3, "x")

    C = modArraySupport.ExpandArray( _
            ExpandArray(A, 1, 3, "F"), _
                    2, 4, "S")

Debug.Print "AFTER:================================="
    For RowNdx = LBound(B, 1) To UBound(B, 1)
        S = vbNullString
        For ColNdx = LBound(B, 2) To UBound(B, 2)
            S = S & B(RowNdx, ColNdx) & " "
        Next ColNdx
Debug.Print S
    Next RowNdx

'Debug.Print "AFTER:================================="
'    For RowNdx = LBound(C, 1) To UBound(C, 1)
'         S = vbNullString
'         For ColNdx = LBound(C, 2) To UBound(C, 2)
'              S = S & C(RowNdx, ColNdx) & " "
'         Next ColNdx
'         Debug.Print S
'    Next RowNdx

End Sub


Public Sub DemoSwapArrayRows()
    
    Dim R As Long
    Dim C As Long
    Dim S As String
    Dim A(1 To 3, 1 To 2)
    Dim B()
    
    
    A(1, 1) = "a"
    A(1, 2) = "b"
    A(2, 1) = "c"
    A(2, 2) = "d"
    A(3, 1) = "e"
    A(3, 2) = "f"

Debug.Print "BEFORE============================"
    For R = LBound(A, 1) To UBound(A, 1)
        S = vbNullString
        For C = LBound(A, 2) To UBound(A, 2)
            S = S & A(R, C) & " "
        Next C
Debug.Print S
    Next R

    B = modArraySupport.SwapArrayRows(A, 2, 3)

Debug.Print "AFTER============================"
    For R = LBound(B, 1) To UBound(B, 1)
        S = vbNullString
        For C = LBound(B, 2) To UBound(B, 2)
            S = S & B(R, C) & " "
        Next C
Debug.Print S
    Next R

End Sub


Public Sub DemoSwapArrayColumns()
    
    Dim R As Long
    Dim C As Long
    Dim S As String
    Dim A(1 To 3, 1 To 2)
    Dim B()
    
    
    A(1, 1) = "a"
    A(1, 2) = "b"
    A(2, 1) = "c"
    A(2, 2) = "d"
    A(3, 1) = "e"
    A(3, 2) = "f"

Debug.Print "BEFORE============================"
    For R = LBound(A, 1) To UBound(A, 1)
        S = vbNullString
        For C = LBound(A, 2) To UBound(A, 2)
            S = S & A(R, C) & " "
        Next C
Debug.Print S
    Next R

    B = modArraySupport.SwapArrayColumns(A, 1, 2)

Debug.Print "AFTER============================"
    For R = LBound(B, 1) To UBound(B, 1)
        S = vbNullString
        For C = LBound(B, 2) To UBound(B, 2)
            S = S & B(R, C) & " "
        Next C
Debug.Print S
    Next R

End Sub


Public Sub DemoGetColumn()

    Dim InputArr(1 To 2, 1 To 3)
    Dim Result() As Long
    Dim B As Boolean
    Dim N As Long
    
    
    InputArr(1, 1) = 1
    InputArr(1, 2) = 2
    InputArr(1, 3) = 3
    InputArr(2, 1) = 4
    InputArr(2, 2) = 5
    InputArr(2, 3) = 6
    
    B = modArraySupport.GetColumn(InputArr, Result, 3)
    
    If B = True Then
        For N = LBound(Result) To UBound(Result)
Debug.Print Result(N)
        Next N
    Else
Debug.Print "Error from GetColumn"
    End If

End Sub


Public Sub DemoGetRow()

    Dim InputArr(1 To 2, 1 To 3)
    Dim Result() As Long
    Dim B As Boolean
    Dim N As Long
    
    
    InputArr(1, 1) = 1
    InputArr(1, 2) = 2
    InputArr(1, 3) = 3
    InputArr(2, 1) = 4
    InputArr(2, 2) = 5
    InputArr(2, 3) = 6
    
    B = modArraySupport.GetRow(InputArr, Result, 2)
    
    If B = True Then
        For N = LBound(Result) To UBound(Result)
Debug.Print Result(N)
        Next N
    Else
Debug.Print "Error from GetRow"
    End If

End Sub
