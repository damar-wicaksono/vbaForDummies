Attribute VB_Name = "Listing5_4"
'
' Listing 5-4 Using Iif to Make Inline Decisions
' Uses an Excel worksheet
'
Public Sub IifDemo()

    ' Create variables to hold the two numbers
    Dim Input1 As Double
    Dim Input2 As Double
    
    ' Create an output string
    Dim Output As String
    
    ' Fill the variables with input from the worksheet
    Input1 = Sheet1.Cells(3, 2)
    Input2 = Sheet1.Cells(4, 2)
    
    ' Use nested Iif functions to check all three conditions
    Output = IIf(Input1 = Input2, "The values are equal.", _
                 IIf(Input1 > Input2, "First Number is greater than Second.", _
                     "Second Number is greater."))
                 
                 
End Sub
