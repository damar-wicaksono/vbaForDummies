Attribute VB_Name = "Listing5_9"
'
' Listing 5-9 Using the GoTo Statement
' Uses an Excel worksheet
'
Public Sub MakeChoice()
    
    Dim CursorPosition As Integer   ' Current row selection
    Dim BinValue As Integer         ' Bin for selected row
    Dim Output As Integer           ' Storage room number
    
    ' The restart check point
RestartCheck:
    
    ' Determine if the user has selected more than one row
    If ActiveWindow.RangeSelection.Rows.Count = 1 Then
        
        ' Get the cursor position
        CursorPosition = ActiveWindow.RangeSelection.Row
        
    Else
        
        ' Tell the user to select only one cell
        MsgBox "Select only one cell, please." & vbCrLf & _
               "Choose the first row  in the range?", _
               vbExclamation Or vbYesNo, "Selection Error"
               
        ' Determine if the user selected Yes
        If Result = vbYes Then
        
            ' Modify the selection
            Range("A" & CStr(ActiveWindow.RangeSelection.Row)).Select
            
            ' Try the Check Again
            GoTo RestartCheck
        End If
        
        ' Exit the Sub without further processing
        End
    End If
    
    ' Get the selected bin number
    BinValue = Sheet2.Cells(CursorPosition, 2)
    
    ' Select a choice of storage room based in the bin
    Select Case BinValue
        Case 1
            Output = 1
        Case 2
            Output = 2
        Case 3 To 4
            Output = 1
        Case 5 To 6
            Output = 3
    End Select
    
    ' Store the number in the worksheet
    Sheet2.Cells(CursorPosition, 3) = Output
End Sub

