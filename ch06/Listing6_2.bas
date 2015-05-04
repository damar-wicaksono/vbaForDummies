Attribute VB_Name = "Listing6_2"
'
' Listing 6-2 Defining a Custom Error Handler
'
Public Sub ErrorHandle()

    ' The variable that receives the input
    Dim InNumber As Byte
    
    ' Tell VBA about the error handle
    On Error GoTo MyHandler
    
    ' Ask the user for some input
    InNumber = InputBox("Type a number between 1 and " & _
                        "10.", "Numeric Input", "1")
    
    ' Determine whether the input is correct
    If (InNumber < 1) Or (InNumber > 10) Then
    
        ' If Input is incorrect, then raise an error
        Err.Raise vbObjectError + 1, _
                  "ErrorCheck.ErrorCondition.ErrorHandle", _
                  "Incorrect Numeric Input. The Number " & _
                  "must be between 1 and 10."
    Else
    
        ' Otherwise display the result
        MsgBox "The Number you Typed: " & CStr(InNumber), _
               vbOKOnly Or vbInformation, "Successful Input"
    End If
    
    ' Exit the Sub
    Exit Sub

' The Start of the Error handler
MyHandler:
    
    ' Display an error message box
    MsgBox "The program experienced an error." & vbCrLf & _
           "Error Number: " & CStr(Err.Number) & vbCrLf & _
           "Description: " & Err.Description & vbCrLf & _
           "Source: " & Err.Source, _
           vbOKOnly Or vbExclamation, _
           "Program Error"
    
    ' Always clear the error after you process it
    Err.Clear
           
End Sub
