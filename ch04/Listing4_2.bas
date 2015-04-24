Attribute VB_Name = "Listing4_2"
'
' Listing 4-2 Adding versus Concatenating Strings
'
Public Sub AddVersusConcatenate()
    ' Create three strings for testing.
    Dim Address1 As String
    Dim Address2 As Variant
    Dim OtherInfo As String
    
    ' Place a value into the strings
    Address1 = "123 First Street"
    OtherInfo = "Somewhere, NV 12345"
    
    ' Place a NULL value into the second address line
    Address2 = Null
    
    ' Concatenate the string to a null
    Dim ConString As String
    ConString = Address1 & vbCrLf & _
                Address2 & vbCrLf & _
                OtherInfo
    
    ' Display the result. You see a blank line for the null
    MsgBox ConString
    
    ' Add the string to a null
    Dim AddString As String
    AddString = Address1 & vbCrLf & _
                (Address2 + vbCrLf) & _
                OtherInfo
    
    ' Display the result. I do not see any blank line
    MsgBox AddString
    
    ' Show that the results are correct when Address2 contains a value
    Address2 = "Apt 3G"
    AddString = Address1 & vbCrLf & _
                Address2 + vbCrLf & _
                OtherInfo
                
    ' Display the result which is a proper one
    MsgBox AddString
End Sub
