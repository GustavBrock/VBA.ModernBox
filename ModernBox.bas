Attribute VB_Name = "ModernBox"
Option Compare Database
Option Explicit

' Complete modern/metro styled replacement for MsgBox and InputBox.
' 2016-04-16. Gustav Brock, Cactus Data ApS, CPH.
' Version 1.0.2: ErrorMox added.
'
' License: MIT.

' Requires:
'   Form:
'       ModernBox
'       ModputBox
'   Module:
'       ModernStyle
'       ModernThemeColours


' Global variables for forms ModernBox and ModputBox.
Public mbPrompt             As String
Public mbTitle              As Variant
Public mbHelpFile           As String
Public mbContext            As Long
' Global variables for form ModernBox.
Public mbButtons            As VbMsgBoxStyle
' Global variables for form ModputBox.
Public mbDefault            As String
Public mbXPos               As Variant
Public mbYPos               As Variant

' Global variable set by form ModernBox when closed.
Public mbResult             As VbMsgBoxResult
' Global variable set by form ModputBox when closed.
Public mbInputText          As String

' Form name of the modern message box.
Private Const ModernBoxName As String = "ModernBox"
' Form name of the modern input box.
Private Const ModputBoxName As String = "ModputBox"

' API call for sleep function.
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' API function to open a compiled HTML help file (.chm) with the HTML Help Viewer.
' Note: The help file must reside on a local drive.
' Sample help file for download:
' http://www.innovasys.com/download/examplechmzipfile?ZipFile=%2FStatic%2FHS%2FSamples%2FHelpStudioSample_CHM.zip
Private Declare Function HTMLHelpShowContents Lib "hhctrl.ocx" Alias "HtmlHelpA" ( _
    ByVal hwnd As Long, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Long, _
    ByVal dwData As Long) _
    As Long

Public Function InputMox( _
    Prompt As String, _
    Optional Title As Variant = Null, _
    Optional Default As String, _
    Optional XPos As Variant = Null, _
    Optional YPos As Variant = Null, _
    Optional HelpFile As String, _
    Optional Context As Long, _
    Optional TimeOut As Long) _
    As String
    
' Syntax. As for InputBox with an added parameter, TimeOut:
' InputMox(Prompt, [Title], [Default], [XPos], [YPos], [HelpFile], [Context], [TimeOut]) As VbMsgBoxResult
'
' Note:
'   XPos and YPos are relative to the top-left corner of the
'   application, not the screen as it is for InputBox.
'
' If TimeOut is negative, zero, or missing:
'   InputMox waits forever as InputBox.
' If TimeOut is positive:
'   InputMox exits after TimeOut milliseconds, returning an empty string.
    
    ' Set global variables to be read by form ModernBox.
    mbPrompt = Prompt
    mbTitle = Title
    mbDefault = Default
    mbXPos = XPos
    mbYPos = YPos
    mbHelpFile = HelpFile
    mbContext = Context
    
    Call OpenFormDialog(ModputBoxName, TimeOut)
    
    ' Return return value set by form ModputBoxName.
    InputMox = mbInputText

End Function

Public Function MsgMox( _
    Prompt As String, _
    Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
    Optional Title As Variant = Null, _
    Optional HelpFile As String, _
    Optional Context As Long, _
    Optional TimeOut As Long) _
    As VbMsgBoxResult
    
' Syntax. As for MsgBox with an added parameter, TimeOut:
' MsgMox(Prompt, [Buttons As VbMsgBoxStyle = vbOKOnly], [Title], [HelpFile], [Context], [TimeOut]) As VbMsgBoxResult
'
' If TimeOut is negative, zero, or missing:
'   MsgMox waits forever as MsgBox.
' If TimeOut is positive:
'   MsgMox exits after TimeOut milliseconds, returning the result of the current default button.
    
    ' Set global variables to be read by form ModernBox.
    mbButtons = Buttons
    mbPrompt = Prompt
    mbTitle = Title
    mbHelpFile = HelpFile
    mbContext = Context
    
    Call OpenFormDialog(ModernBoxName, TimeOut)
    
    ' Return result value set by form ModernBoxName.
    MsgMox = mbResult

End Function

Public Function ErrorMox( _
    Optional ByVal Topic As String) _
    As String

' Opens a MsgMox predefined for displaying the error number, source, and description if Err <> 0.
' Also reestablishes the application window if Echo is False, the cursor if Hourglass is True,
' and resets the Status line.

    ' Text to prefix the error number.
    Const Prefix    As String = "Error"
    
    Dim Prompt      As String
    Dim Title       As String
    Dim Buttons     As VbMsgBoxStyle
    Dim Message     As String
    
    If Err = 0 Then
        ' No error. Exit.
    Else
        ' Reestablish display.
        DoCmd.Hourglass False
        DoCmd.Echo True
        
        ' Display error message.
        Title = ApplicationTitle
        Title = Title & ": " & Application.CurrentObjectName
        If Topic <> "" Then
            Title = Title & ", " & Topic
        End If
        
        If Prefix <> "" Then
            Prompt = Prefix & ": "
        End If
        Prompt = Prompt & CStr(Err.Number) & vbCrLf & _
            Err.Description & "."
        
        Buttons = vbOKOnly + vbCritical
        MsgMox Prompt, Buttons, Title
        
        ' Clear status line.
        StatusLineReset
        
        ' Return message lines.
        Message = Title & vbCrLf & Prompt
    End If
    
    ErrorMox = Message

End Function

Public Function OpenFormDialog( _
    ByVal FormName As String, _
    Optional ByVal TimeOut As Long, _
    Optional ByVal OpenArgs As Variant = Null) _
    As Boolean
    
' Open a modal form in non-dialogue mode to prevent dialogue borders to be displayed
' while simulating dialogue behaviour using Sleep.

' If TimeOut is negative, zero, or missing:
'   Form FormName waits forever.
' If TimeOut is positive:
'   Form FormName exits after TimeOut milliseconds.
    
    Const SecondsPerDay     As Single = 86400
    
    Dim LaunchTime          As Date
    Dim CurrentTime         As Date
    Dim TimedOut            As Boolean
    Dim Index               As Integer
    Dim FormExists          As Boolean
    
    ' Check that form FormName exists.
    For Index = 0 To CurrentProject.AllForms.Count - 1
        If CurrentProject.AllForms(Index).Name = FormName Then
            FormExists = True
            Exit For
        End If
    Next
    If FormExists = True Then
        If CurrentProject.AllForms(FormName).IsLoaded = True Then
            ' Don't reopen the form should it already be loaded.
        Else
            ' Open modal form in non-dialogue mode to prevent dialogue borders to be displayed.
            DoCmd.OpenForm FormName, acNormal, , , , acWindowNormal, OpenArgs
        End If
        ' Record launch time and current time with 1/18 second resolution.
        LaunchTime = Date + CDate(Timer / SecondsPerDay)
        Do While CurrentProject.AllForms(FormName).IsLoaded
            ' Form FormName is open.
            ' Make sure form and form actions are rendered.
            DoEvents
            ' Halt Access for 1/20 second.
            ' This will typically cause a CPU load less than 1%.
            ' Looping faster will raise CPU load dramatically.
            Sleep 50
            If TimeOut > 0 Then
                ' Check for time-out.
                CurrentTime = Date + CDate(Timer / SecondsPerDay)
                If (CurrentTime - LaunchTime) * SecondsPerDay > TimeOut / 1000 Then
                    ' Time-out reached.
                    ' Close form FormName and exit.
                    DoCmd.Close acForm, FormName, acSaveNo
                    TimedOut = True
                    Exit Do
                End If
            End If
        Loop
        ' At this point, user or time-out has closed form FormName.
    End If
    
    ' Return True if the form was not found or was closed by user interaction.
    OpenFormDialog = Not TimedOut

End Function

Public Function OpenHelp( _
    ByVal HelpPath As String, _
    Optional ByVal ContextID As Long = 1) _
    As Boolean
    
' Open a help file at context ContextID if found.
    
    Const MinimumContextID  As Long = 1
    
    Dim Success             As Boolean

    ' Adjust invalid context IDs.
    If ContextID < MinimumContextID Then
        ContextID = MinimumContextID
    End If
    
    ' Open help file.
    ' Fails silently if help file or context ID is not found.
    Success = CBool(HTMLHelpShowContents(0, HelpPath, &HF, ContextID))
    
    OpenHelp = Success
    
End Function

Public Function ApplicationTitle() As String

' Returns the CurrentDb property "AppTitle" if set.
' If not found, the sanitised application file name is returned.
'
' Example:
'   AppTitle created as "My Application".
'   Returns: My Application
'
'   No AppTitle. Filename is "super app.accdb".
'   Returns: Super App
    
    Const AppTitle  As String = "AppTitle"
    
    Dim AppProperty As Property
    Dim Title       As String
    
    For Each AppProperty In CurrentDb.Properties
        If AppProperty.Name = AppTitle Then
            Title = AppProperty.Value
            Exit For
        End If
    Next
    If Title = "" Then
        Title = StrConv(StrReverse(Split(StrReverse(CurrentProject.Name), ".", 2)(1)), vbProperCase)
    End If
    
    ApplicationTitle = Title

End Function

Public Sub StatusLineReset()

' Removes a custom status line message and displays the default messages of Access.
'
' SysCmd(acSysCmdClearStatus) cannot be used unconditionally as it will fail if
' SysCmd(acSysCmdSetStatus, "Some message") has not been called, and because
' SysCmd(acSysCmdSetStatus, "") (an empty string) always will fail.

    ' Thus, first set a blank message to make sure a message has been set.
    SysCmd acSysCmdSetStatus, " "
    ' Clear (reset) the status line.
    SysCmd acSysCmdClearStatus

End Sub

