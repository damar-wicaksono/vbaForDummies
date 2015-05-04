Attribute VB_Name = "Listing6_3"
'
' Listing 6-3 Using the Debug Object
'
Public Sub UseDebug()

    ' The variable that receives the input
    Dim InNumber As Byte
    
    ' Ask the user for some input
    InNumber = InputBox("Type a number between 1 and 10.", _
                        "Numeric Input", _
                        "1")
    
    ' Print the value of InNumber to the Immediate window
    Debug.Print "InNumber = " & CStr(InNumber)
    
    ' Stop program execution if InNumber is not in the correct range
    Debug.Assert (InNumber >= 1) And (InNumber <= 10)
    
    ' Display the result
    MsgBox "The Number you typed: " & CStr(InNumber), _
           vbOKOnly Or vbInformation, _
           "Successful Input"
                        
End Sub
