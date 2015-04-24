Attribute VB_Name = "Listing4_5"
'
' Listing 4-5 Remove Spaces from a String
'
Public Sub RemoveSpace()
    ' Declare a string with spaces
    Dim IStr As String
    
    ' Declare an output string
    Dim Output As String
    
    ' Add a string to IStr
    IStr = "    Hello   "
    
    ' Show the original string length
    Output = "Original String Length: " & CStr(Len(IStr))
    
    ' Get rid of the spaces on the left
    Output = Output & vbCrLf & _
            "LTrim Length: " & CStr(Len(LTrim(IStr))) & _
            " Value: " & Chr(&H22) & LTrim(IStr) & Chr(&H22)
    
    ' Get rid of the spaces on the right
    Output = Output & vbCrLf & _
            "RTrim Length: " & CStr(Len(RTrim(IStr))) & _
            " Value: " & Chr(&H22) & RTrim(IStr) & Chr(&H22)
    
    ' Get rid of all the spaces
    Output = Output & vbCrLf & _
            "Trim Length: " & CStr(Len(Trim(IStr))) & _
            " Value: " & Chr(&H22) & Trim(IStr) & Chr(&H22)
            
    ' Display the result
    MsgBox Output, vbOKOnly, "Trimming Extra Spaces"
End Sub
