Attribute VB_Name = "Listing4_6"
'
' Listing 4-6 Finding Information in Strings by Using Parsing
'
Public Sub ParseString()

    ' Create a string with elements the program can parse
    Dim MyStr As String
    
    ' Create an output string
    Dim Output As String
    
    ' Fill the input string with data
    MyStr = "A string to parse"
    
    ' Display the whole string
    Output = "The whole string is: " & MyStr
    
    ' Obtain the first word
    Output = Output & vbCrLf & "The First Word: " & _
             Left(MyStr, InStr(1, MyStr, " "))
    
    ' Obtain the second word
    Output = Output & vbCrLf & "The Second Word: " & _
             Right(MyStr, Len(MyStr) - InStrRev(MyStr, " "))
            
    ' Obtain the word "string"
    Output = Output & vbCrLf & "The Word String: " & _
             Trim(Mid(MyStr, InStr(1, MyStr, "string"), Len(MyStr) - InStr(1, MyStr, "to")))
            
    ' Output the result
    MsgBox Output, vbOKOnly, "parsing a String"
End Sub
