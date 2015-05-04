Attribute VB_Name = "Listing6_1"
'
' Listing 6-1 Determining the Amount of Free Disk Space
'
Public Sub DriveTest()

    ' Create a variable to hold the free space
    Dim FreeSpace As Double
    
    ' Create a reference to the filesystem
    Dim MyFileSystem As FileSystemObject
    
    ' Create a reference for the target drive
    Dim MyDrive As Drive
    
    ' Create a dialog result variable
    Dim Result As VbMsgBoxResult
    
    ' Provide a jump back point
DoCheckAgain:
    
    ' Fill these two objects with data so they show the available space on C
    Set MyFileSystem = New FileSystemObject
    Set MyDrive = MyFileSystem.GetDrive("C")
    
    ' Determine the amount of free space
    FreeSpace = MyDrive.AvailableSpace

    ' Make the check
    If FreeSpace < 10000000000# Then
    
        ' The drive does not have enough space. Ask what to do
        Result = MsgBox("The drive does not have enough space to hold the data." & _
                        vbCrLf & _
                        "Do you want to correct the error (free some space)?" & _
                        vbCrLf & _
                        Format(FreeSpace, "###, ###") & " bytes available, " & _
                        "1'000'000'000 bytes needed.", vbYesNo Or vbExclamation, _
                        "Drive Space Error")
                        
        ' Determine if the user wants to fix the problem
        If Result = vbYes Then
        
            ' Wait for the user to fix the problem
            MsgBox "Please click Ok when you have freed some disk space.", _
                   vbInformation Or vbOKOnly, "Retry Drive Check"
               
        ' Go to the fallback point.
        GoTo DoCheckAgain
    
        Else
        
            ' The user does not want to fix the error
            MsgBox "The program cannot save your data until the drive has " & _
                   "enough of free space.", _
                   vbInformation Or vbOKOnly, _
                   "Insufficient Drive Space"
    
            ' End the Sub
            Exit Sub
        End If
    End If
                        
End Sub
