Attribute VB_Name = "Listing4_8"
'
' Listing 4-8 Converting between Numeric Bases
'
Public Sub ShowBase()
    ' Define the three number bases
    Dim OctNum As Integer
    Dim DecNum As Integer
    Dim HexNum As Integer
    
    ' Define an output string
    Dim Output As String
    
    ' Assign an octal number
    OctNum = &O110
    
    ' Assign a decimal number
    DecNum = 110
    
    ' Assign a hexadecimal number
    HexNum = &H110
    
    ' Create a heading
    Output = vbTab & vbTab & vbTab & "Oct" & _
             vbTab & "Dec" & _
             vbTab & "Hex" & vbCrLf
    
    ' Create an output string
    Output = Output & "Octal Number: " & _
             vbTab & vbTab & Oct$(OctNum) & _
             vbTab & CStr(OctNum) & _
             vbTab & Hex$(OctNum) & _
             vbCrLf & "Decimal Number: " & _
             vbTab & vbTab & Oct$(DecNum) & _
             vbTab & CStr(DecNum) & _
             vbTab & Hex$(DecNum) & _
             vbCrLf & "Hexadecimal Number: " & _
             vbTab & Oct$(HexNum) & _
             vbTab & CStr(HexNum) & _
             vbTab & Hex$(HexNum)
             
    ' Display the actual numbers
    MsgBox Output, vbInformation Or vbOKOnly, "Data Type Output"
End Sub
