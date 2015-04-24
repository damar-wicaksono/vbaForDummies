Attribute VB_Name = "Listing4_9"
'
' Listing 4-9 Converting between Numbers and Strings
'
Public Sub NumberConvert()
    ' Create some variables for use in conversion
    Dim MyInt As Integer
    Dim MySgl As Single
    Dim MyStr As String
    
    ' Conversion between Integer and Single is direct with no data loss
    MyInt = 30
    MySgl = MyInt
    MsgBox "MyInt = " & CStr(MyInt) & vbCrLf & "MySgl = " & CStr(MySgl), _
            vbOKOnly, "Current Data Values"
            
    ' Conversion between Single and Integer is also direct but incurs data loss
    MySgl = 35.01
    MyInt = MySgl
    MsgBox "MyInt = " & CStr(MyInt) & vbCrLf & "MySgl = " & CStr(MySgl), _
            vbOKOnly, "Current Data Values"
            
    ' Conversion between a String and a Single or an Integer can rely on use of
    ' special function. The Conversion can also incur data loss
    MyStr = "40.05"
    MyInt = CInt(MyStr)
    MySgl = CSng(MyStr)
    MsgBox "MyInt = " & CStr(MyInt) & vbCrLf & "MySgl = " & CStr(MySgl), _
            vbOKOnly, "Current Data Values"
            
    ' Conversion between a Single or Integer and a String can rely on use of
    ' a special function when making a direct conversion.
    ' The conversion does not incur any data loss.
    MyInt = 45
    MySgl = 45.05
    MyStr = MyInt
    MsgBox MyStr, vbOKOnly, "Current Data Values"
    
    ' You must use a special function in mixed data situations
    MyStr = "MyInt = " & CStr(MyInt) & vbCrLf & "MySgl = " & CStr(MySgl)
    MsgBox MyStr, vbOKOnly, "Current Data Values"
End Sub
