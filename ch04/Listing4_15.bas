Attribute VB_Name = "Listing4_15"
'
' Listing 4-15 Changing the Format of a Date
'
Public Sub FormatDemo()
    ' Create the Date variable
    Dim MyDate As Date
    
    ' Fill MyDate with the current date and time
    MyDate = Now
    
    ' Display the date using standard "named" formats
    MsgBox "Standard Format = " & vbTab & CStr(MyDate) & vbCrLf & _
           "Long Date = " & vbTab & Format(MyDate, "Long Date") & vbCrLf & _
           "Short Time = " & vbTab & Format(MyDate, "Short Time"), _
           vbOKOnly, "VBA Named Formats"
End Sub
