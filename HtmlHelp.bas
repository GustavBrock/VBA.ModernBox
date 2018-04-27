Attribute VB_Name = "HtmlHelp"
' Functions for managing the HTML Help Viewer control.
' 2018-04-26. Gustav Brock, Cactus Data ApS, CPH.
' Version 1.0.1
'
' License: MIT.

' Professional tool for creating compressed HTML help files:
'   Innovasys HelpStudio
'   http://www.innovasys.com/product/hs/overview

Option Explicit

' API calls for the HTML Help Viewer control.
' Sample help file for download:
' http://www.innovasys.com/download/examplechmzipfile?ZipFile=%2FStatic%2FHS%2FSamples%2FHelpStudioSample_CHM.zip

' Open a compiled HTML help file (.chm) or close all opened help files.
#If VBA7 Then
    Private Declare PtrSafe Function HTMLHelpShowContents Lib "hhctrl.ocx" _
        Alias "HtmlHelpA" (ByVal hwnd As LongPtr, _
        ByVal lpHelpFile As String, _
        ByVal wCommand As Long, _
        ByVal dwData As Long) As Long
#Else
    Private Declare Function HTMLHelpShowContents Lib "hhctrl.ocx" _
        Alias "HtmlHelpA" (ByVal hwnd As Long, _
        ByVal lpHelpFile As String, _
        ByVal wCommand As Long, _
        ByVal dwData As Long) As Long
#End If

' Open a compiled HTML help file (.chm) with the Search tab active.
#If VBA7 Then
    Private Declare PtrSafe Function HTMLHelpShowSearch Lib "hhctrl.ocx" _
        Alias "HtmlHelpA" (ByVal hwnd As LongPtr, _
        ByVal lpHelpFile As String, _
        ByVal wCommand As Long, _
        ByRef dwData As HhFtsQuery) As Long
#Else
    Private Declare Function HTMLHelpShowSearch Lib "hhctrl.ocx" _
        Alias "HtmlHelpA" (ByVal hwnd As Long, _
        ByVal lpHelpFile As String, _
        ByVal wCommand As Long, _
        ByRef dwData As HhFtsQuery) As Long
#End If


' User Defined Types.
'
' UDT for HTMLHelpShowSearch.
Private Type HhFtsQuery
    cbStruct          As Long       ' Size of structure in bytes.
    fUniCodeStrings   As Long       ' TRUE if all strings are unicode.
    pszSearchQuery    As String     ' String containing the search query.
    iProximity        As Long       ' Word proximity.
    fStemmedSearch    As Long       ' TRUE for StemmedSearch only.
    fTitleOnly        As Long       ' TRUE for Title search only.
    fExecute          As Long       ' TRUE to initiate the search.
    pszWindow         As String     ' Window to display in.
End Type

' Enums.
'
' Commands to control the appearance of the viewer when launched.
Public Enum hhCommand
    ' Select the last opened tab.
    DisplayTopic = &H0
    ' Select the Contents tab.
    DisplayContents = &H1
    ' Select the Index tab.
    DisplayIndex = &H2
    ' Select the Search tab.
    DisplaySearch = &H3
    ' Select the Contents tab and open a topic by its index.
    OpenContext = &HF
    ' Close all windows opened by the viewer.
    CloseAll = &H12
End Enum
'

' Displays a compressed HTML help file using the HTML Help File Viewer.
' Alternatively, closes all windows of the viewer that may be open.
'
' Returns True if the operation was successful.
' Returns False if the operation couldn't be carried out, for example
' if one or more parameter is invalid, or - for closing windows - if
' no windows were open.
'
' Example, display the Contents tab:
'   Success = HelpControl(DisplayContents, "d:\path\HelpStudioSample.chm")
' Example, display the Contents tab and open the page with context id 7:
'   Success = HelpControl(OpenContext, "d:\path\HelpStudioSample.chm", 7)
' Example, display the Index tab:
'   Success = HelpControl(DisplayIndex, "d:\path\HelpStudioSample.chm")
' Example, display the Search tab:
'   Success = HelpControl(DisplaySearch, "d:\path\HelpStudioSample.chm")
' Example, display the last opened tab and the main page:
'   Success = HelpControl(DisplayTopic, "d:\path\HelpStudioSample.chm")
' Example, close all opened Help Viewer windows:
'   Success = HelpControl(CloseAll)
'
' Note:
'   If the help file is stored on a networked folder and will open, but
'   will not display the context pages, move the file to a local folder.
'
' 2018-04-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function HelpControl( _
    ByVal Command As hhCommand, _
    Optional ByVal HelpFile As String, _
    Optional ByVal Context As Long) _
    As Boolean
    
    ' Use default owner handle (Application).
    Const OwnerHandle   As Long = 0
    ' Neutral values.
    Const NoHandle      As Long = 0
    Const NoTopic       As Long = 0
    Const NoFile        As String = ""

    ' Handle of the current Help Viewer window (if any).
    Static OpenHandle   As Long
    
    Dim SearchQuery     As HhFtsQuery
    Dim Handle          As Long
    
    ' Manage the Help Viewer.
    Select Case Command
        ' Open the Help Viewer and display a tab.
        Case hhCommand.DisplayTopic, _
            hhCommand.DisplayContents, _
            hhCommand.DisplayIndex
            Handle = HTMLHelpShowContents(OwnerHandle, HelpFile, Command, NoTopic)
        
        ' Open the Help Viewer and display the topic having the ID of Context.
        Case hhCommand.OpenContext
            ' Reset displayed tab to Contents.
            Handle = HTMLHelpShowContents(OwnerHandle, HelpFile, hhCommand.DisplayContents, NoTopic)
            ' Open help context page.
            Handle = HTMLHelpShowContents(OwnerHandle, HelpFile, Command, Context)
        
        ' Open the Help Viewer and display the Search tab.
        Case hhCommand.DisplaySearch
            SearchQuery.cbStruct = Len(SearchQuery)
            Handle = HTMLHelpShowSearch(OwnerHandle, HelpFile, Command, SearchQuery)
        
        ' Close all windows opened by the Help Viewer.
        Case hhCommand.CloseAll
            If OpenHandle = NoHandle Then
                ' Don't waste time on closing non-existing windows.
            Else
                ' A help file has been opened.
                ' Set Handle to return success.
                Handle = OpenHandle
                ' Make sure, all help windows are closed, and reset OpenHandle.
                OpenHandle = HTMLHelpShowContents(OwnerHandle, NoFile, Command, NoTopic)
            End If
            
        Case Else
            ' Ignore.
    End Select
    
    If Command <> hhCommand.CloseAll Then
        If Handle <> NoHandle Then
            ' Store the handle of the window.
            OpenHandle = Handle
        End If
    End If
    
    ' Return True if success.
    HelpControl = (Handle <> NoHandle)

End Function

