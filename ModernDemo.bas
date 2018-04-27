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

Public Sub HelpDemo()

    Const HelpFile = "c:\test\samplehelp.chm"
    Const Context   As Long = 2
    
    Dim Result      As VbMsgBoxResult
    
    Result = MsgMox("Press Help", vbQuestion + vbOKCancel + vbMsgBoxHelpButton, "Help Demo", HelpFile, Context)
    ' Close the help window, should it have been opened.
    CloseHelp
    ' Or call CloseHelp before exiting the application.
    
    Debug.Print Result
    
End Sub

