# VBA.ModernBox
Modern/Metro style message box and input box for Microsoft Access 2013+
Version 1.1.1

Modern/Metro styled message box and input box that directly can replace MsgBox() and InputBox()in Microsoft Access 2013 and later.
Also contains a prebuilt error box for use in error handling.

With version 1.1.1 the boxes can not be moved beyond that of an Integer.

' 2017-09-19: Added limitation of the settings for WindowsLeft and WindowsTop
'             to be within the range of Integer.

With version 1.1 a collection of helper functions are included:


' Returns True if the passed colour value is one of the
' Windows Phone Theme Colors.
'
' 2017-04-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsWpThemeColor(ByVal Color As Long) As Boolean


' Returns the literal name of the passed colour value if
' it is one of the Windows Phone Theme Colors.
'
' 2017-04-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function LiteralWpThemeColor( _
    ByVal Color As wpThemeColor) _
    As String


' Loops all(!) possible color values and prints those of the
' Windows Phone Theme Colors.
' This will take nearly 30 seconds.
'
' 2017-04-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ListColors()



Full documentation is found here: https://rdsrc.us/zLJcA9
