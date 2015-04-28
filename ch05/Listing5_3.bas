Attribute VB_Name = "Listing5_3"
'
' Listing 5-3 Using the If...Then...ElseIf Statement for Comparisons
' Uses an Excel worksheet
'
Public Sub CompareNumbers2()

    ' Create variables to hold the two numbers
    Dim Input1 As Double
    Dim Input2 As Double
    
    ' Create an output as string
    Dim Output As String
    
    ' Fill the variables with input from the worksheet
    Input1 = Sheet1.Cells(3, 2)
    Input2 = Sheet1.Cells(4, 2)
    
    ' Determine if the first number is greater than the second number
    If Input1 > Input2 Then
    
        ' Tell the user the first number is greater
        Output = "First Number is greater than Second Number."
        
    ' Determine if they are equal
    ElseIf Input1 = Input2 Then
        
        ' Tell the user they are equal
        Output = "The values are equal."
        
    Else
        
        ' The first number is less than the second.
        Output = "First Number is less than Second Number."
    
    End If
    
    ' Place the output on the worksheet
    Sheet1.Cells(6, 2) = Output
End Sub
