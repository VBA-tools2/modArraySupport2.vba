Attribute VB_Name = "TestPreparation"

'@Folder("Tests")

Option Explicit
Option Private Module

Public Sub PrepareWorkbook4Tests()
    
    Dim wkb As Workbook
    Set wkb = ThisWorkbook
    
    Dim wks As Worksheet
    Set wks = GetTestWorksheet(wkb)
    
    InitializeWorksheet4Tests wks

End Sub

Private Function GetTestWorksheet( _
    ByVal wkb As Workbook _
        ) As Worksheet
    
    '==========================================================================
    Const TestWorksheetName As String = "__modArraySupportTest__"
    '==========================================================================
    
    Dim ws As Worksheet
    For Each ws In wkb.Worksheets
        If ws.Name = TestWorksheetName Then
            Set GetTestWorksheet = ws
            Exit Function
        End If
    Next
    
    Dim wks As Worksheet
    Set wks = wkb.Worksheets.Add(wkb.Worksheets(1))
    wks.Name = TestWorksheetName
    
    Set GetTestWorksheet = wks

End Function

Public Sub InitializeWorksheet4Tests( _
    ByVal wks As Worksheet _
)
    
    With Application
        Dim ScreenUpdatingEnabled As Boolean
        ScreenUpdatingEnabled = .ScreenUpdating
        .ScreenUpdating = False
    End With
    
    With wks
        .UsedRange.Clear
        
        .Cells(2, 1).Value = "SEE VBA CODE MODULE 'modArraySupport' for code"
        With .Cells(4, 1)
            .Value = "Functions in the VBA Code"
            With .Interior
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.05
            End With
        End With
        
        .Cells(5, 1).Value = "AreDataTypesCompatible"
        .Cells(6, 1).Value = "ChangeBoundsOfArray"
        .Cells(7, 1).Value = "CombineTwoDArrays"
        .Cells(8, 1).Value = "CompareArrays"
        .Cells(9, 1).Value = "ConcatenateArrays"
        .Cells(10, 1).Value = "CopyArray"
        .Cells(11, 1).Value = "CopyArraySubSetToArray"
        .Cells(12, 1).Value = "CopyNonNothingObjectsToArray"
        .Cells(13, 1).Value = "DataTypeOfArray"
        .Cells(14, 1).Value = "DeleteArrayElement"
        .Cells(15, 1).Value = "ExpandArray"
        .Cells(16, 1).Value = "FirstNonEmptyStringIndexInArray"
        .Cells(17, 1).Value = "GetColumn"
        .Cells(18, 1).Value = "GetRow"
        .Cells(19, 1).Value = "InsertElementIntoArray"
        .Cells(20, 1).Value = "IsArrayAllDefault"
        .Cells(21, 1).Value = "IsArrayAllNumeric"
        .Cells(22, 1).Value = "IsArrayAllocated"
        .Cells(23, 1).Value = "IsArrayDynamic"
        With .Cells(24, 1)
            .Value = "IsArrayEmpty"
            .Font.Strikethrough = True
        End With
        .Cells(25, 1).Value = "IsArrayObjects"
        .Cells(26, 1).Value = "IsArraySorted"
        .Cells(27, 1).Value = "IsNumericDataType"
        .Cells(28, 1).Value = "IsVariantArrayConsistent"
        With .Cells(29, 1)
            .Value = "IsVariantArrayNumeric"
            .Font.Strikethrough = True
        End With
        .Cells(30, 1).Value = "MoveEmptyStringsToEndOfArray"
        .Cells(31, 1).Value = "NumberOfArrayDimensions"
        .Cells(32, 1).Value = "NumElements"
        .Cells(33, 1).Value = "ResetVariantArrayToDefaults"
        .Cells(34, 1).Value = "ReverseArrayInPlace"
        .Cells(35, 1).Value = "ReverseArrayOfObjectsInPlace"
        .Cells(36, 1).Value = "SetObjectArrayToNothing"
        .Cells(37, 1).Value = "SetVariableToDefault"
        .Cells(38, 1).Value = "SwapArrayColumns"
        .Cells(39, 1).Value = "SwapArrayRows"
        .Cells(40, 1).Value = "TransposeArray"
        .Cells(41, 1).Value = "VectorsToArray"
        
        .Columns(1).EntireColumn.AutoFit
    End With
    
    Application.ScreenUpdating = ScreenUpdatingEnabled

End Sub
