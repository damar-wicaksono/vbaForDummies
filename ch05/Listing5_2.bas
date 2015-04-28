Attribute VB_Name = "Listing5_2"
'
' Listing 5-2 Using the If...Then...Else Statement for Comparisons
' Uses an Excel worksheet
'
Public Sub CompareNumbers()
    
    ' Create variables to hold the two numbers
    Dim Input1 As Double
    Dim Input2 As Double
    
    ' Create an output string
    Dim Output As String
    
    ' Fill the variables with input from the worksheet
    Input1 = Sheet1.Cells(3, 2)
    Input2 = Sheet1.Cells(4, 2)
    
    ' Determine if the first number is greater than or
    ' equal to the second number
    If Input1 >= Input2 Then
    
        ' Determine if they are equal
        If Input1 = Input2 Then
        
            ' Tell the user they are equal
            Output = "The values are equal."
            
        Else
            
            ' The first number is greater than the second number
            Output = "First Number is greater than Second Number."
        
        End If
            
    Else
    
        ' The second number is greater than the first number
        Output = "First Number is smaller than the Second Number."
    
    End If
        
    ' Place the output on the worksheet
    Sheet1.Cells(6, 2) = Output
End Sub
