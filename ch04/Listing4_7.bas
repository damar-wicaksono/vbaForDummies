Attribute VB_Name = "Listing4_7"
'
' Listing 4-7 Demonstrating the Differences in Data Type Ranges
'
Public Sub DataRange()
    
    ' Declare the numeric variables
    Dim MyInt As Integer
    Dim MySgl As Single
    Dim MyDbl As Double
    Dim MyCur As Currency
    Dim MyDec As Variant    ' Cannot define Decimal type directly
    
    ' Define values for each variable
    MyInt = 30 + 0.00010001000111   ' Forced assignment, cut-off
    MySgl = 30 + 0.00010001000111
    MyDbl = 30 + 0.00010001000111
    MyCur = 30 + 0.00010001000111
    MyDec = CDec(30 + 0.0001000111)
    
    ' Display the actual content
    MsgBox "Integer: " & TwoTab & CStr(MyInt) & _
        vbCrLf & "Single: " & TwoTab & CStr(MySgl) & _
        vbCrLf & "Double: " & TwoTab & CStr(MyDbl) & _
        vbCrLf & "Currency: " & TwoTab & CStr(MyCur) & _
        vbCrLf & "Decimal: " & TwoTab & CStr(MyDec), _
        vbOKOnly, "VBA Data Types"
End Sub
