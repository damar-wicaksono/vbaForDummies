Attribute VB_Name = "Listing4_1"
'
' Listing 4-1 Examples of Variable and Constant Declarations with Scope
'
Option Explicit

' This variable is visible to other modules
Public MyPublicVariable As String
' This variable is visible only to this module
Private MyPrivateVariable As String
' Using Dim is the same as making the variable private
Dim MyDimVariable As String

' A constant is only used for conditional compilation
#Const MyConditionalConstant = "Hello"
' This constant is visible to other modules
Public Const MyPublicConstant = "Hello"
' This constant is visible only to this module
Private Const MyPrivateConstant = "Hello"

Public Sub DataDeclarations()
    ' Only this Sub can see this variable
    Dim MyDimSubVariable As String
    
    ' Only this sub can see this constant
    Const MySubConstant = "Hello"
End Sub
