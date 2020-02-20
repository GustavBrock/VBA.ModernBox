Attribute VB_Name = "ModernBox"
Option Compare Database
Option Explicit

' Complete modern/metro styled replacement for MsgBox and InputBox.
' 2020-02-17. Gustav Brock, Cactus Data ApS, CPH.
' Version 1.0.2: ErrorMox added.
' Version 1.0.3: DoCmd.SelectObject inserted to bring form to front of other popup forms.
' Version 1.2.0: Modified API calls to 32/64-bit.
'                HTML Help function and API declarations moved to separate module HtmlHelp.
' Version 1.3.1: Added option to check for Windows 10.
'
' License: MIT.

' Requires:
'   Form:
'       ModernBox
'       ModputBox
'   Module:
'       ModernStyle
'       ModernThemeColours
'       HtmlHelp


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
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
'

' Opens an input box, using form ModputBox, similar to VBA.InputBox.
'
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
'
' 2018-04-26. Gustav Brock, Cactus Data ApS, CPH.
'
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

' Opens a message box, using form ModernBox, similar to VBA.MsgBox.
'
' Syntax. As for MsgBox with an added parameter, TimeOut:
' MsgMox(Prompt, [Buttons As VbMsgBoxStyle = vbOKOnly], [Title], [HelpFile], [Context], [TimeOut]) As VbMsgBoxResult
'
' If TimeOut is negative, zero, or missing:
'   MsgMox waits forever as MsgBox.
' If TimeOut is positive:
'   MsgMox exits after TimeOut milliseconds, returning the result of the current default button.
'
' 2018-04-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function MsgMox( _
    Prompt As String, _
    Optional Buttons As VbMsgBoxStyle = vbOkOnly, _
    Optional Title As Variant = Null, _
    Optional HelpFile As String, _
    Optional Context As Long, _
    Optional TimeOut As Long) _
    As VbMsgBoxResult
    
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

' Opens a MsgMox predefined for displaying the error number, source, and description if Err <> 0.
' Also reestablishes the application window, if Echo is False, and the cursor, if Hourglass is True,
' and resets the Status line.
'
' 2018-04-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ErrorMox( _
    Optional ByVal Topic As String) _
    As String

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
        
        Buttons = vbOkOnly + vbCritical
        MsgMox Prompt, Buttons, Title
        
        ' Clear status line.
        StatusLineReset
        
        ' Return message lines.
        Message = Title & vbCrLf & Prompt
    End If
    
    ErrorMox = Message

End Function

' Opens a modal form in non-dialogue mode to prevent dialogue borders to be displayed
' while simulating dialogue behaviour using Sleep.

' If TimeOut is negative, zero, or missing:
'   Form FormName waits forever.
' If TimeOut is positive:
'   Form FormName exits after TimeOut milliseconds.
'
' 2018-04-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function OpenFormDialog( _
    ByVal FormName As String, _
    Optional ByVal TimeOut As Long, _
    Optional ByVal OpenArgs As Variant = Null) _
    As Boolean
        
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
            ' Bring form to front; it may hide behind a popup form.
            DoCmd.SelectObject acForm, FormName
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

' Open a help file at context ContextID if found.
'
' Note:
'   An opened help viewer window must be closed before exiting the application,
'   or, most likely, Access will chrash.
'
' Requires:
'   HtmlHelp
'
' 2018-04-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function OpenHelp( _
    ByVal HelpFile As String, _
    Optional ByVal ContextID As Long = 1) _
    As Boolean
    
    Const MinimumContextID  As Long = 1
    
    Dim Success             As Boolean

    ' Adjust invalid context IDs.
    If ContextID < MinimumContextID Then
        ContextID = MinimumContextID
    End If
    
    ' Open help file.
    ' Fails silently if help file or context ID is not found.
    Success = HelpControl(OpenContext, HelpFile, ContextID)
    
    OpenHelp = Success
    
End Function

' Close all open HTML Help Viewer windows.
'
' Requires:
'   HtmlHelp
'
' 2018-04-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CloseHelp() As Boolean
    
    Dim Success             As Boolean
    
    ' Close help file.
    ' Fails silently if no Help Viewer windows are open.
    Success = HelpControl(CloseAll)
    
    CloseHelp = Success
    
End Function

' Returns in Access the CurrentDb property "AppTitle" if set.
' If not found, the sanitised application file name is returned.
'
' Example:
'   AppTitle created as "My Application".
'   Returns: My Application
'
'   No AppTitle. Filename is "super app.accdb".
'   Returns: Super App
'
' Requires a reference to:
'   Microsoft Office nn.m Access database engine Object Library
'
' 2018-04-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ApplicationTitle() As String
    
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

' Removes a custom status line message and displays the default messages of Access.
'
' SysCmd(acSysCmdClearStatus) cannot be used unconditionally, as it will fail if
' SysCmd(acSysCmdSetStatus, "Some message") has not been called, and because
' SysCmd(acSysCmdSetStatus, "") (an empty string) always will fail.
' Thus, first set a blank message is set to make sure a message has been set.
'
' 2018-04-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub StatusLineReset()

    ' Set a blank message to make sure a message has been set.
    SysCmd acSysCmdSetStatus, " "
    ' Clear (reset) the status line.
    SysCmd acSysCmdClearStatus

End Sub

' Checks if the primary (current) Windows version is Windows 10.
' Returns True if Windows version is 10, False if not.
'
' The call to WMI takes about 50 ms. Thus, to speed up repeated calls,
' the result is kept in the static variable OsVersion.
'
' 2019-04-07. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsWindows10() As Boolean

    Const NoVersion     As Integer = 0
    Const Version10     As Integer = 10
    
    Static OsVersion    As Integer
    
    Dim OperatingSystem As Object
    Dim Result          As Boolean

    If OsVersion = NoVersion Then
        ' Connect to WMI and obtain instances of Win32_OperatingSystem
        For Each OperatingSystem In GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
            If OperatingSystem.Primary = True Then
                OsVersion = Val(OperatingSystem.Version)
                Exit For
            End If
        Next
    Else
        ' Repeated call. OsVersion has previously been found.
    End If
    Result = (OsVersion = Version10)
    
    IsWindows10 = Result

End Function
