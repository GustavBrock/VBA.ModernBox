VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ModernBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Modern/Metro styled messagebox to replace MsgBox.
' To be called from function MsgMox of module ModernBox.
'
' 2020-02-17. (c) Gustav Brock, Cactus Data ApS, CPH.
' Version 1.3.1
' License: MIT.
'
' Requirements:
'   Function StyleCommandButtons of module ModernStyle.
'
' 2015-04-18: Added replace of & to && to prevent underscores in prompt.
' 2017-09-19: Added limitation of the settings for WindowsLeft and WindowsTop
'             to be within the range of Integer.
' 2018-04-19: Modified API calls to 32/64-bit.
' 2020-02-16: Added Windows 10 design for use under Windows 10.


' Constants.

' Colour of background of close-button picture.
' Border has to be visible to make the picture touch the right border of the form.
Private Const lngBorderColorCloseButton As Long = &H5050C7

' Button positions.
Private Const Position0         As Long = 8503
Private Const PositionDelta     As Long = 1700
' Additional form height for each prompt line exceeding the default count of lines.
Private Const HeightDelta       As Long = 301
' Offset for Windows 10 style.
Private Const Windows10Offset   As Long = 96

' Icon characters.
Private Const CharCritical      As String = "×"
Private Const CharExclamation   As String = "!"
Private Const CharInformation   As String = "i"
Private Const CharQuestion      As String = "?"

' First and last Enum ButtonCaption.
Private Const FirstCaptionId    As Long = 800
Private Const LastCaptionId     As Long = 810


' Enums.

' Button counts of possible button combinations.
Private Enum ButtonCount
    AbortRetryIgnore = 3
    OKCancel = 2
    OKOnly = 1
    RetryCancel = 2
    YesNo = 2
    YesNoCancel = 3
    Help = 1
End Enum

' Index of array to hold variable properties of buttons.
Private Enum ButtonProperty
    Visible = 0
    Caption = 1
    Value = 2
End Enum

' Resource IDs from user32.dll per definition.
Private Enum ButtonCaption
    ButtonOK = 800
    ButtonCancel = 801
    ButtonAbort = 802
    ButtonRetry = 803
    ButtonIgnore = 804
    ButtonYes = 805
    ButtonNo = 806
    ButtonClose = 807
    ButtonHelp = 808
    ButtonTryAgain = 809
    ButtonContinue = 810
End Enum

' Parameters for mouse action.
Private Enum MouseAction
    MouseDown = 1
    MouseMove = 0
    MouseUp = -1
End Enum


' Variables.

' Array to hold localized button captions.
Private Captions( _
    FirstCaptionId To LastCaptionId)    As String

' Array to hold current button settings and result values.
Private Buttons(0 To 3)                 As Variant

' Array to hold button positions.
Private Positions(0 To 3)               As Long
    
' Variables for style and buttons.
Private IconStyle                       As VbMsgBoxStyle
Private DefaultButton                   As VbMsgBoxStyle
Private ButtonSequence                  As VbMsgBoxStyle
Private HelpButton                      As VbMsgBoxStyle
Private PromptRight                     As VbMsgBoxStyle

' Button assigned as help button.
Private HelpButtonIndex                 As Long

' Shall cancel button (close button of form) be hidden.
Private NoCancel                        As Boolean

' Variable for result value to return.
Private Result                          As VbMsgBoxResult


' API calls.

' API functions for retrieval of localized button captions.
#If VBA7 Then
    Private Declare PtrSafe Function LoadString Lib "user32" Alias "LoadStringA" ( _
        ByVal hInstance As LongPtr, _
        ByVal wID As Long, _
        ByVal lpBuffer As String, _
        ByVal nBufferMax As Long) _
        As Long
             
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
        ByVal lpFileName As String) _
        As LongPtr
#Else
    Private Declare Function LoadString Lib "user32" Alias "LoadStringA" ( _
        ByVal hInstance As Long, _
        ByVal wID As Long, _
        ByVal lpBuffer As String, _
        ByVal nBufferMax As Long) _
        As Long
             
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
        ByVal lpFileName As String) _
        As Long
#End If

Private Sub BoxInactive_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call SetCloseButton(False)

End Sub

Private Sub BoxTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Enable dragging of the form.
    Call FormMove(Button, Shift, X, Y, MouseDown)

End Sub

Private Sub BoxTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call SetCloseButton(False)
    
    ' Drag the form if dragging is enabled.
    Call FormMove(Button, Shift, X, Y, MouseMove)

End Sub

Private Sub BoxTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Disable dragging of the form.
    Call FormMove(Button, Shift, X, Y, MouseUp)

End Sub

Private Sub Button0_Click()

    Call ButtonClick(0)

End Sub

Private Sub Button0_GotFocus()

    Call ButtonFocus(0)

End Sub

Private Sub Button1_Click()

    Call ButtonClick(1)

End Sub

Private Sub Button1_GotFocus()

    Call ButtonFocus(1)

End Sub

Private Sub Button2_Click()

    Call ButtonClick(2)

End Sub

Private Sub Button2_GotFocus()

    Call ButtonFocus(2)

End Sub

Private Sub Button3_Click()

    Call ButtonClick(3)

End Sub

Private Sub Button3_GotFocus()

    Call ButtonFocus(3)

End Sub

Private Sub ButtonCancel_Click()

    Call ButtonClick
    
End Sub

Private Sub ButtonCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call SetCloseButton(True)
    
End Sub

Private Sub LabelTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Enable dragging of the form.
    Call FormMove(Button, Shift, X, Y, MouseDown)

End Sub

Private Sub LabelTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call SetCloseButton(False)
    
    ' Drag the form if dragging is enabled.
    Call FormMove(Button, Shift, X, Y, MouseMove)

End Sub

Private Sub LabelTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Disable dragging of the form.
    Call FormMove(Button, Shift, X, Y, MouseUp)

End Sub

Private Sub Form_Load()
    
    ' Adjust top and/or height of some controls.
    Call SetControlSizes
    ' Apply modern colours to form.
    Call SetColours
    ' Show specified icon.
    Call SetIcon
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    Dim NoCancel    As Boolean

    ' Set the messagebox style variables.
    Call SetMsgBoxStyle
    ' Set caption of title bar.
    Call SetTitle
    ' Set prompt.
    Call SetPrompt
    ' Set active buttons and captions and taborder.
    ' Eventually resize form to accommodate buttons and a supersized prompt.
    Call SetButtonSequence
    ' NoCancel has been set by SetButtonSequence.
    If Not NoCancel Then
        ' Set close button status.
        Call SetCloseButton(False)
    End If
    
    ' Set default result value.
    Result = vbCancel

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Return result value to global variable.
    mbResult = Result

End Sub

Private Sub PictureCloseActive_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Happens when mouse leaves ButtonCancel as this is slightly smaller than the close pictures.
    Call SetCloseButton(False)
    
End Sub

Private Sub PictureCloseInactive_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Happens when mouse leaves ButtonCancel as this is slightly smaller than the close pictures.
    Call SetCloseButton(False)

End Sub

Private Sub ButtonClick(Optional ByVal ButtonIndex As Long = -1)

' Close form at click by any button except Help button.

    If ButtonIndex < 0 Then
        ' Cancel by pressing Escape or clicking red close icon.
        ' Set result value to Cancel.
        Result = VbMsgBoxResult.vbCancel
    ElseIf ButtonIndex > 0 Then
        If ButtonIndex = HelpButtonIndex Then
            Call OpenHelp(mbHelpFile, mbContext)
            ' Don't close the form.
            Exit Sub
        End If
    Else
        ' Result value has been set previously in the GotFocus event of the button.
    End If
    
    ' Return value (Result) is set at the OnUnload event.
    DoCmd.Close acForm, Me.Name, acSaveNo

End Sub

Private Sub ButtonFocus(ByVal ButtonIndex As Long)
    
' Style buttons to indicate the new default button.

    ' Set (new) default button.
    Me("Button" & CStr(ButtonIndex)).Default = True
    
    ' Set (new) default result value.
    Result = Buttons(ButtonIndex)(ButtonProperty.Value)
    
    ' Recolour visible buttons.
    Call StyleCommandButtons(Me)

End Sub

Private Sub FillCaptions()

' Retrieve localized button captions into array Captions.

    Const FileName          As String = "user32.dll"
    Const BufferMax         As Long = 256
    
#If VBA7 Then
    Dim Instance            As LongPtr
#Else
    Dim Instance            As Long
#End If
    
    Dim Buffer              As String
    Dim StringLength        As Long
    Dim CaptionId           As Long
    
    Instance = LoadLibrary(FileName)

    ' Read localized captions into static array.
    For CaptionId = FirstCaptionId To LastCaptionId
        Buffer = String(BufferMax, vbNullChar)
        StringLength = LoadString(Instance, CaptionId, Buffer, BufferMax)
        Captions(CaptionId) = Left(Buffer, StringLength)
    Next

End Sub

Private Sub FillPositions()
    
' Fill static array of button positions.
    
    Dim Index               As Integer
    
    For Index = LBound(Positions) To UBound(Positions)
        Positions(Index) = Position0 + Index * PositionDelta
    Next

End Sub

Private Function FormExtend( _
    ByVal LineCount As Integer) _
    As Long

' Extends the form and the prompt area (prompt label caption) to
' accommodate larger (longer) prompts.

    ' Default available count of prompt lines (zero based).
    Const DefaultLineCount  As Integer = 3
    
    Dim LinesDelta          As Integer
    Dim ExtendedHeight      As Long
    
    If LineCount > DefaultLineCount Then
        LinesDelta = LineCount - DefaultLineCount
        ExtendedHeight = HeightDelta * LinesDelta
        ' Extend form.
        Me.Move Me.WindowLeft, , , Me.WindowHeight + ExtendedHeight
        ' Extend or move relevant controls except buttons.
        Me!LabelPrompt.Height = Me!LabelPrompt.Height + ExtendedHeight
        Me!BoxIcon.Height = Me!BoxIcon.Height + ExtendedHeight
        Me!LabelIcon.Top = Me!LabelIcon.Top + Int(ExtendedHeight / 2)
    End If
    
    ' Return value for button positioning.
    FormExtend = ExtendedHeight

End Function

Private Sub FormMove(Button As Integer, Shift As Integer, X As Single, Y As Single, _
    ByVal MouseAction As MouseAction)

' Move the form by dragging the title bar or the label upon it.

    ' WindowLeft and WindowTop must be within the range of Integer.
    Const TopLeftMax        As Single = 2 ^ 15 - 1
    Const TopLeftMin        As Single = -2 ^ 15

    ' Statics to hold the position of the form when mouse is clicked.
    Static PositionX        As Single
    Static PositionY        As Single
    ' Static to hold that a form move is enabled.
    Static MoveEnabled      As Boolean
    
    Dim WindowTop           As Single
    Dim WindowLeft          As Single
    
    ' The value of MoveEnable indicates if the call is from
    ' mouse up, mouse down, or mouse move.
    
    If MouseAction = MouseMove Then
        ' Move form.
        If MoveEnabled = True Then
            ' Form move in progress.
            If Button = acLeftButton Then
                ' Calculate new form position.
                WindowTop = Me.WindowTop + Y - PositionY
                WindowLeft = Me.WindowLeft + X - PositionX
                ' Limit Top and Left.
                If WindowTop > TopLeftMax Then
                    WindowTop = TopLeftMax
                ElseIf WindowTop < TopLeftMin Then
                    WindowTop = TopLeftMax
                End If
                If WindowLeft > TopLeftMax Then
                    WindowLeft = TopLeftMax
                ElseIf WindowLeft < TopLeftMin Then
                    WindowLeft = TopLeftMax
                End If
                Me.Move WindowLeft, WindowTop
            End If
        End If
    Else
        ' Enable/disable form move.
        If Button = acLeftButton Then
            ' Only left-button click accepted.
            'If MoveEnable = True Then
            If MouseAction = MouseDown Then
                ' MouseDown.
                ' Store cursor start position.
                PositionX = X
                PositionY = Y
                MoveEnabled = True
            Else
                ' MouseUp.
                ' Stop form move.
                MoveEnabled = False
            End If
        End If
    End If

End Sub

Private Function PromptLineCount() As Integer

' Calculates the expected count of lines the prompt will occupy.
' Reads from global variable mbPrompt.
' The value is zero based.

    ' Maximum expected character count for a prompt line without line break.
    Const CharsPerLine      As Integer = 64
    
    Dim Prompt              As String
    Dim Lines               As Variant
    Dim Line                As Integer
    Dim LineCount           As Integer
    
    ' Replace individual Cr or Lf with CrLf in prompt.
    Prompt = Replace(Replace(Replace(mbPrompt, vbCrLf, vbNullChar), vbCr, vbNullChar), vbLf, vbNullChar)
    Prompt = Replace(Prompt, vbNullChar, vbCrLf)
    ' Collect prompt lines.
    Lines = Split(Prompt, vbCrLf)
    ' Count lines zero based.
    LineCount = UBound(Lines) - LBound(Lines)
    For Line = LBound(Lines) To UBound(Lines)
        LineCount = LineCount + Int(Len(Lines(Line)) / CharsPerLine)
    Next
    
    PromptLineCount = LineCount
    
End Function

Private Sub SetButtonSequence()

' Arrange from one to four visible buttons and refresh their captions.

    ' Maximum count of enabled (visible) buttons including Help button.
    Const MaxButtonCount    As Long = 3 + 1
    ' First button index.
    Const FirstButton       As Long = 0
    ' Undefined result value for Help button and inactive buttons.
    Const MsgBoxResultNone  As Long = 0
    
    Dim WindowWidth         As Long
    Dim WindowExpand        As Long
    Dim WindowExtend        As Long
    Dim ActiveButtonCount   As Long
    Dim HelpButtonCount     As ButtonCount
    Dim ButtonIndex         As Long
    Dim LineCount           As Integer
    
    ' Fill array of localized captions.
    Call FillCaptions
    ' Fill array of button positions.
    Call FillPositions
    
    If HelpButton = vbMsgBoxHelpButton Then
        ' The Help button shall be displayed.
        HelpButtonCount = ButtonCount.Help
    End If
    
    ' Set captions and return values on active (visible) buttons.
    Select Case ButtonSequence
        Case vbAbortRetryIgnore
            ActiveButtonCount = ButtonCount.AbortRetryIgnore
            Buttons(FirstButton) = Array(True, Captions(ButtonCaption.ButtonAbort), VbMsgBoxResult.vbAbort)
            Buttons(FirstButton + 1) = Array(True, Captions(ButtonCaption.ButtonRetry), VbMsgBoxResult.vbRetry)
            Buttons(FirstButton + 2) = Array(True, Captions(ButtonCaption.ButtonIgnore), VbMsgBoxResult.vbIgnore)
            NoCancel = True
        Case vbOKCancel
            ActiveButtonCount = ButtonCount.OKCancel
            Buttons(FirstButton) = Array(True, Captions(ButtonCaption.ButtonOK), VbMsgBoxResult.vbOK)
            Buttons(FirstButton + 1) = Array(True, Captions(ButtonCaption.ButtonCancel), VbMsgBoxResult.vbCancel)
        Case vbOkOnly
            ' Note: Any click action (except Help) will result in Cancel.
            ActiveButtonCount = ButtonCount.OKOnly
            Buttons(FirstButton) = Array(True, Captions(ButtonCaption.ButtonOK), VbMsgBoxResult.vbCancel)
        Case vbRetryCancel
            ActiveButtonCount = ButtonCount.RetryCancel
            Buttons(FirstButton) = Array(True, Captions(ButtonCaption.ButtonTryAgain), VbMsgBoxResult.vbRetry)
            Buttons(FirstButton + 1) = Array(True, Captions(ButtonCaption.ButtonCancel), VbMsgBoxResult.vbCancel)
        Case vbYesNo
            ActiveButtonCount = ButtonCount.YesNo
            Buttons(FirstButton) = Array(True, Captions(ButtonCaption.ButtonYes), VbMsgBoxResult.vbYes)
            Buttons(FirstButton + 1) = Array(True, Captions(ButtonCaption.ButtonNo), VbMsgBoxResult.vbNo)
            NoCancel = True
        Case vbYesNoCancel
            ActiveButtonCount = ButtonCount.YesNoCancel
            Buttons(FirstButton) = Array(True, Captions(ButtonCaption.ButtonYes), VbMsgBoxResult.vbYes)
            Buttons(FirstButton + 1) = Array(True, Captions(ButtonCaption.ButtonNo), VbMsgBoxResult.vbNo)
            Buttons(FirstButton + 2) = Array(True, Captions(ButtonCaption.ButtonCancel), VbMsgBoxResult.vbCancel)
        Case Else
            ' Identical to OKOnly.
            ' Note: Any click action (except Help) will result in Cancel.
            ActiveButtonCount = ButtonCount.OKOnly
            Buttons(FirstButton) = Array(True, Captions(ButtonCaption.ButtonOK), VbMsgBoxResult.vbCancel)
    End Select
    
    ' Add a Help button at far right if requested.
    If HelpButtonCount = 1 Then
        HelpButtonIndex = ActiveButtonCount
        Buttons(HelpButtonIndex) = Array(True, Captions(ButtonCaption.ButtonHelp), MsgBoxResultNone)
        ActiveButtonCount = ActiveButtonCount + HelpButtonCount
    End If
    ' Reset remaining buttons.
    For ButtonIndex = ActiveButtonCount To MaxButtonCount - 1
        Buttons(ButtonIndex) = Array(False, vbNullString, MsgBoxResultNone)
    Next
    ' Set display status for all buttons.
    For ButtonIndex = FirstButton To MaxButtonCount - 1
        With Me("Button" & CStr(ButtonIndex))
            .Visible = Buttons(ButtonIndex)(ButtonProperty.Visible)
            .Caption = Buttons(ButtonIndex)(ButtonProperty.Caption)
        End With
    Next
    
    ' Expand the form to make room for multiple buttons.
    WindowExpand = Positions(ActiveButtonCount - 1) - Positions(FirstButton)
    WindowWidth = Me.WindowWidth + WindowExpand
    Me.Move Me.WindowLeft, Me.WindowTop, WindowWidth
   
    ' Remove form's close button for certain button combinations.
    If NoCancel = True Then
        Me!ButtonCancel.Enabled = False
    Else
        ' Reposition Cancel button and form's close button (picture).
        Me!ButtonCancel.Left = Me!ButtonCancel.Left + WindowExpand
        Me!PictureClose.Left = Me!PictureClose.Left + WindowExpand
        Me!PictureCloseActive.Left = Me!PictureCloseActive.Left + WindowExpand
        Me!PictureCloseInactive.Left = Me!PictureCloseInactive.Left + WindowExpand
        Me!BoxInactive.Left = Me!BoxInactive.Left + WindowExpand
    End If
    
    ' Extend the form to fit a supersized prompt.
    LineCount = PromptLineCount()
    If LineCount > 0 Then
        ' Extend the form and controls (except buttons) to
        ' make room for multiple prompt lines.
        WindowExtend = FormExtend(LineCount)
    End If
    ' Position active buttons.
    For ButtonIndex = FirstButton To ActiveButtonCount - 1
        With Me("Button" & CStr(ButtonIndex))
            .Left = Positions(ButtonIndex)
            .Top = .Top + WindowExtend
        End With
    Next
    
    ' Apply tab settings.
    Call SetDefaultButton
    
End Sub

Private Sub SetColours()

' Set colours of form and basic controls.

    ' Set modern colours of components.
    ' NB! This colour should match the fill colour of the picture PictureCloseInactive.
    Me!BoxTitle.BackColor = wpThemeColor.Cyan
    Me!BoxInactive.BackColor = wpThemeColor.Steel
    Me!PictureClose.BorderColor = lngBorderColorCloseButton
    
    ' Set modern colours of form.
    Me.Section(acDetail).BackColor = wpThemeColor.Steel

End Sub

Private Sub SetDefaultButton()

' Set functionally and visually the default button.

    Dim ButtonIndex         As Long
    Dim DefaultButtonIndex  As Long
    
    Select Case DefaultButton
        Case vbDefaultButton1
            DefaultButtonIndex = 0
        Case vbDefaultButton2
            DefaultButtonIndex = 1
        Case vbDefaultButton3
            DefaultButtonIndex = 2
        Case vbDefaultButton4
            DefaultButtonIndex = 3
    End Select
    
    ' If vbDefaultButtonX exceeds the count of active buttons, reset to first button.
    If Me("Button" & CStr(DefaultButtonIndex)).Visible = False Then
        DefaultButtonIndex = 0
    End If
    
    ' Set tab stop and index in reverse order.
    For ButtonIndex = 3 To 0 Step -1
        With Me("Button" & CStr((ButtonIndex + DefaultButtonIndex) Mod 4))
            .TabIndex = 1 + ButtonIndex
            .TabStop = True
        End With
    Next
    
    ' Apply colouring of buttons.
    Call ButtonFocus(DefaultButtonIndex)

End Sub

Private Sub SetIcon()

' Set the message icon as specified by opening parameters.

    Dim IconBackColor       As Long
    Dim IconForeColor       As Long
    Dim IconChar            As String
    
    Select Case IconStyle
        Case vbCritical
            IconBackColor = wpThemeColor.Red
            IconForeColor = wpThemeColor.White
            IconChar = CharCritical
        Case vbExclamation
            IconBackColor = wpThemeColor.Yellow
            IconForeColor = wpThemeColor.Black
            IconChar = CharExclamation
        Case vbInformation
            IconBackColor = wpThemeColor.Cobalt
            IconForeColor = wpThemeColor.White
            IconChar = CharInformation
        Case vbQuestion
            IconBackColor = wpThemeColor.Cobalt
            IconForeColor = wpThemeColor.White
            IconChar = CharQuestion
        Case Else
            ' Hide icon area.
            IconBackColor = Me.Section(acDetail).BackColor
            IconChar = vbNullString
    End Select
    
    Me!BoxIcon.BackColor = IconBackColor
    Me!LabelIcon.Caption = IconChar
    Me!LabelIcon.ForeColor = IconForeColor
    Me!BoxIcon.Visible = True
    Me!LabelIcon.Visible = True
    
    ' Audible prompt when icon is shown.
    If IconChar <> "" Then
        DoCmd.Beep
    End If
    
End Sub

Private Sub SetMsgBoxStyle()

' Decode and set the messagebox style variables.
    
    Dim Style               As VbMsgBoxStyle

    ' Get style info from command line.
    Style = mbButtons

    ' Set the style variables of our Modern Box.
    IconStyle = _
        (Style And vbCritical) Or (Style And vbExclamation) Or (Style And vbInformation) Or (Style And vbQuestion)
    DefaultButton = _
        (Style And vbDefaultButton1) Or (Style And vbDefaultButton2) Or (Style And vbDefaultButton3) Or (Style And vbDefaultButton4)
    ButtonSequence = _
        (Style And vbAbortRetryIgnore) Or _
        (Style And vbOKCancel) Or (Style And vbOkOnly) Or (Style And vbRetryCancel) Or _
        (Style And vbYesNo) Or (Style And vbYesNoCancel)
    HelpButton = _
        (Style And vbMsgBoxHelpButton)
    PromptRight = _
        (Style And vbMsgBoxRight)

End Sub

Private Sub SetPrompt()

' Set prompt.

    Const TextAlignRight    As Byte = 3

    Dim Prompt              As String
    
    On Error Resume Next

    If mbPrompt <> "" Then
        ' Will fail if module ModernBox with global variable mbPrompt is missing.
        Prompt = mbPrompt
    End If
    If Prompt = "" Then
        ' Leave empty prompt.
    Else
        ' Double an et "&" to "&&" or the caption display will exchange it with an underscore.
        Prompt = Replace(Prompt, "&", "&&")
    End If
    
    If PromptRight = vbMsgBoxRight Then
        ' Right align prompt.
        Me!LabelPrompt.TextAlign = TextAlignRight
    End If
    
    Me!LabelPrompt.Caption = Prompt

End Sub

Private Sub SetTitle()

' Set title bar caption.

    Dim Title               As String
    
    On Error Resume Next
    
    ' Will fail if module ModernBox with global variable mbTitle is missing.
    If IsNull(mbTitle) Then
        ' Default if parameter of function MsgMox is missing.
        ' Use default application title (Microsoft Access).
        Title = Application.Name
    Else
        Title = mbTitle
    End If
    
    ' Will not fail.
    If Title = "" Then
        ' Use default application title (Microsoft Access).
        Title = Application.Name
    End If
    
    ' Set title.
    Me!LabelTitle.Caption = Title

End Sub

Private Sub SetCloseButton(ByVal Active As Boolean)

    Static Initialized  As Boolean
    Static Status       As Boolean

    If Not NoCancel Then
        If IsWindows10 Then
            If (Not Initialized) Or (Status <> Active) Then
                Me!PictureCloseActive.Visible = Active
                Me!PictureCloseInactive.Visible = Not Active
                Status = Active
            End If
        Else
            If Not Initialized Then
                Me!PictureClose.Visible = True
            End If
        End If
        
        Initialized = True
    End If

End Sub

Private Sub SetControlSizes()

    If IsWindows10 Then
        Me!BoxTitle.Height = Me!BoxTitle.Height - Windows10Offset
        Me!LabelTitle.Top = Me!LabelTitle.Top - (Windows10Offset / 2)
        Me!BoxIcon.Top = Me!BoxIcon.Top - Windows10Offset - 1
        Me!BoxIcon.Height = Me!BoxIcon.Height + Windows10Offset
        Me!BoxInactive.Top = Me!BoxInactive.Top - Windows10Offset
        Me!BoxInactive.Height = Me!BoxInactive.Height + Windows10Offset
    End If
    
End Sub

