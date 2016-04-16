Attribute VB_Name = "ModernDemo"
Option Compare Database
Option Explicit

Public Sub ErrorDemo()

    Dim Test    As Integer
    Dim Message As String
    
    On Error GoTo Err_ErrorDemo
    
    Test = 1 / 0
    
Exit_ErrorDemo:
    Exit Sub
    
Err_ErrorDemo:
    Message = ErrorMox("Short demo")
    Debug.Print Message
    Resume Exit_ErrorDemo
    
End Sub

