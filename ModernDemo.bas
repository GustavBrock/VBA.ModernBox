Attribute VB_Name = "ModernDemo"
Option Compare Database
Option Explicit

Public Sub ErrorDemo()

    Const Topic As String = "Example error"
    
    Dim Test    As Integer
    Dim Message As String
    
    On Error GoTo Err_ErrorDemo
    
    Test = 1 / 0
    
Exit_ErrorDemo:
    Exit Sub
    
Err_ErrorDemo:
    Message = ErrorMox(Topic)
    Debug.Print Message
    Resume Exit_ErrorDemo
    
End Sub

Public Sub HelpDemo()

    Const Prompt    As String = "Press Help"
    Const Buttons   As Long = vbQuestion + vbOKCancel + vbMsgBoxHelpButton
    Const Title     As String = "Help Demo"
    Const HelpFile  As String = "c:\test\samplehelp.chm"
    Const Context   As Long = 2
    
    Dim Result      As VbMsgBoxResult
    
    Result = MsgMox(Prompt, Buttons, Title, HelpFile, Context)
    ' Close the help window, should it have been opened.
    CloseHelp
    ' Or call CloseHelp before exiting the application.
    
    Debug.Print Result
    
End Sub

Public Sub ModernDemo()

    Const Prompt1   As String = "What do you wish to tell the World?"
    Const Default   As String = "Hello!"
    Const Buttons   As Long = vbOKCancel + vbInformation + vbDefaultButton2
    Const Title1    As String = "Modern/Metro Input Box"
    Const Title2    As String = "Modern/Metro Message Box"
    
    Dim Message     As String
    
    Message = InputMox(Prompt1, Title1, Default)
    MsgMox Message, Buttons, Title2

End Sub

Public Sub NativeDemo()

    Const Prompt1   As String = "What do you wish to tell the World?"
    Const Default   As String = "Hello!"
    Const Buttons   As Long = vbOKCancel + vbInformation + vbDefaultButton2
    Const Title1    As String = "Boring Input Box"
    Const Title2    As String = "Boring Message Box"
    
    Dim Message     As String
    
    Message = InputBox(Prompt1, Title1, Default)
    MsgBox Message, Buttons, Title2

End Sub

Public Sub ThanksDemo()

    Const Prompt    As String = "No further reading." & vbCrLf & "Proceed on your own, please."
    Const Buttons   As Long = vbCritical + vbOKCancel
    Const Title     As String = "Thank You"
    
    Dim Result      As VbMsgBoxResult
    
    Result = MsgMox(Prompt, Buttons, Title)
    
    Debug.Print Result
    
End Sub

