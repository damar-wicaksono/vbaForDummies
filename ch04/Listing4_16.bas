Attribute VB_Name = "Listing4_16"
'
' Listing 4-16 Defining a Custom Date Format
'
Public Sub CustomFormatDemo()
    ' Create the date variable
    Dim MyDate As Date
    
    ' Fill MyDate with the current date and time
    MyDate = Now
    
    ' Display the date using standard formats
    MsgBox "Custom Date/Time = " & Format(MyDate, "dd mmmm yyyy Hh:Mm:Ss"), _
           vbOKOnly, "VBA Custom Formats"
End Sub
