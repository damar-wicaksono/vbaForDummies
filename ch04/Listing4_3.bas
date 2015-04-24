Attribute VB_Name = "Listing4_3"
'
' Listing 4-3 Creating Special Characters
'
Public Sub ShowCharacter()
    ' Declare the string.
    Dim MyChar As String
    
    ' Tell what type of character the code displays
    MyChar = "Latin Capital Letter A with Circumflex: "
    
    ' Add the character
    MyChar = MyChar + Chr(&HC2)
    
    ' Display the result
    MsgBox MyChar, vbOKOnly, "Special Character"
End Sub

'
' Listing 4-4 Getting the numeric value of a character
'
Public Sub GetCharacter()
    ' Declare the output variables
    Dim MyChar As String
    Dim CharNum As Integer
    
    ' Add the special character to MyChar
    MyChar = Chr(&HC2)
    
    ' Determine the Unicode number for the character
    CharNum = Asc(MyChar)
    
    ' Display the result as a decimal value
    MsgBox "Character " & MyChar & _
            " = Decimal Value " & CStr(CharNum), _
            vbOKOnly, _
            "Special Character Decimal Value"
End Sub
