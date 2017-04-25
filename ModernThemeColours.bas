Option Explicit

' Adoption of Windows Phone 7.5/8.0 colour theme for VBA.
' 2017-04-19. Gustav Brock, Cactus Data ApS, CPH.
' Version 1.1.0
' License: MIT.

' *

' Windows Phone colour enumeration.
Public Enum wpThemeColor
    ' Official colour names from WP8.
    Lime = &HC4A4&
    Green = &H17A960
    Emerald = &H8A00&
    Teal = &HA9AB00
    Cyan = &HE2A11B
    Cobalt = &HEF5000
    Indigo = &HFF006A
    Violet = &HFF00AA
    Pink = &HD072F4
    Magenta = &H7300D8
    Crimson = &H2500A2
    Red = &H14E5&
    Orange = &H68FA&
    Amber = &HAA3F0
    Yellow = &HC8E3&
    Brown = &H2C5A82
    Olive = &H64876D
    Steel = &H87766D
    Mauve = &H8A6076
    Sienna = &H2D52A0
    ' Colour name aliases from WP7.5
    Viridian = &HA9AB00
    Blue = &HE2A11B
    Purple = &HFF00AA
    Mango = &H68FA&
    ' Used for black in popups.
    Darken = &H1D1D1D
    ' Additional must-have names for grey scale.
    Black = &H0&
    DarkGrey = &H3F3F3F
    Grey = &H7F7F7F
    LightGrey = &HBFBFBF
    White = &HFFFFFF
End Enum

' Variable to hold array of WpThemeColor values.
Private ColorPalette As Variant

' Fill array ColorPalette with the values of wpThemeColor.
'
' 2017-04-21. Gustav Brock, Cactus Data ApS, CPH.
'
Private Sub LoadColors()

    Dim Colors(0 To 29) As Long
    
    Dim Index           As Long
    
    If IsEmpty(ColorPalette) Then
        For Index = LBound(Colors) To UBound(Colors)
            Select Case Index
                Case 0
                    Colors(Index) = wpThemeColor.Lime
                Case 1
                    Colors(Index) = wpThemeColor.Green
                Case 2
                    Colors(Index) = wpThemeColor.Emerald
                Case 3
                    Colors(Index) = wpThemeColor.Teal
                Case 4
                    Colors(Index) = wpThemeColor.Cyan
                Case 5
                    Colors(Index) = wpThemeColor.Cobalt
                Case 6
                    Colors(Index) = wpThemeColor.Indigo
                Case 7
                    Colors(Index) = wpThemeColor.Violet
                Case 8
                    Colors(Index) = wpThemeColor.Pink
                Case 9
                    Colors(Index) = wpThemeColor.Magenta
                Case 10
                    Colors(Index) = wpThemeColor.Crimson
                Case 11
                    Colors(Index) = wpThemeColor.Red
                Case 12
                    Colors(Index) = wpThemeColor.Orange
                Case 13
                    Colors(Index) = wpThemeColor.Amber
                Case 14
                    Colors(Index) = wpThemeColor.Yellow
                Case 15
                    Colors(Index) = wpThemeColor.Brown
                Case 16
                    Colors(Index) = wpThemeColor.Olive
                Case 17
                    Colors(Index) = wpThemeColor.Steel
                Case 18
                    Colors(Index) = wpThemeColor.Mauve
                Case 19
                    Colors(Index) = wpThemeColor.Sienna
                Case 20
                    Colors(Index) = wpThemeColor.Darken
                Case 21
                    Colors(Index) = wpThemeColor.Viridian
                Case 22
                    Colors(Index) = wpThemeColor.Blue
                Case 23
                    Colors(Index) = wpThemeColor.Purple
                Case 24
                    Colors(Index) = wpThemeColor.Mango
                Case 25
                    Colors(Index) = wpThemeColor.Black
                Case 26
                    Colors(Index) = wpThemeColor.DarkGrey
                Case 27
                    Colors(Index) = wpThemeColor.Grey
                Case 28
                    Colors(Index) = wpThemeColor.LightGrey
                Case 29
                    Colors(Index) = wpThemeColor.White
            End Select
        Next
    End If
        
    ColorPalette = Colors()

End Sub

Public Function PaletteColor( _
    ByVal Index As Integer) _
    As Long
    
    If IsEmpty(ColorPalette) Then
        ' Fill array ColorPalette.
        LoadColors
    End If
    
    PaletteColor = ColorPalette(Index)
    
End Function

' Returns True if the passed colour value is one of the
' Windows Phone Theme Colors.
'
' 2017-04-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsWpThemeColor(ByVal Color As Long) As Boolean

    Dim Item            As Integer
    Dim IsColor         As Boolean
        
    If IsEmpty(ColorPalette) Then
        ' Fill public array ColorPalette.
        LoadColors
    End If
    
    For Item = LBound(ColorPalette) To UBound(ColorPalette)
        If Color = ColorPalette(Item) Then
            IsColor = True
            Exit For
        End If
    Next
        
    IsWpThemeColor = IsColor

End Function

' Returns the literal name of the passed colour value if
' it is one of the Windows Phone Theme Colors.
'
' 2017-04-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function LiteralWpThemeColor( _
    ByVal Color As wpThemeColor) _
    As String

    Dim Name    As String
    
    Select Case Color
        Case wpThemeColor.Lime
            Name = "Lime"
        Case wpThemeColor.Green
            Name = "Green"
        Case wpThemeColor.Emerald
            Name = "Emerald"
        Case wpThemeColor.Teal
            Name = "Teal"
        Case wpThemeColor.Cyan
            Name = "Cyan"
        Case wpThemeColor.Cobalt
            Name = "Cobalt"
        Case wpThemeColor.Indigo
            Name = "Indigo"
        Case wpThemeColor.Violet
            Name = "Violet"
        Case wpThemeColor.Pink
            Name = "Pink"
        Case wpThemeColor.Magenta
            Name = "Magenta"
        Case wpThemeColor.Crimson
            Name = "Crimson"
        Case wpThemeColor.Red
            Name = "Red"
        Case wpThemeColor.Orange
            Name = "Orange"
        Case wpThemeColor.Amber
            Name = "Amber"
        Case wpThemeColor.Yellow
            Name = "Yellow"
        Case wpThemeColor.Brown
            Name = "Brown"
        Case wpThemeColor.Olive
            Name = "Olive"
        Case wpThemeColor.Steel
            Name = "Steel"
        Case wpThemeColor.Mauve
            Name = "Mauve"
        Case wpThemeColor.Sienna
            Name = "Sienna"
        Case wpThemeColor.Viridian
            Name = "Viridian"
        Case wpThemeColor.Blue
            Name = "Blue"
        Case wpThemeColor.Purple
            Name = "Purple"
        Case wpThemeColor.Mango
            Name = "Mango"
        Case wpThemeColor.Darken
            Name = "Darken"
        Case wpThemeColor.Black
            Name = "Black"
        Case wpThemeColor.DarkGrey
            Name = "DarkGrey"
        Case wpThemeColor.Grey
            Name = "Grey"
        Case wpThemeColor.LightGrey
            Name = "LightGrey"
        Case wpThemeColor.White
            Name = "White"
    End Select
    
    LiteralWpThemeColor = Name
    
End Function

' Loops all(!) possible color values and prints those of the
' Windows Phone Theme Colors.
' This will take nearly 30 seconds.
'
' 2017-04-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ListColors()

    Dim Color   As wpThemeColor
    
    For Color = wpThemeColor.Black To wpThemeColor.White
        If IsWpThemeColor(Color) Then
            Debug.Print Color, LiteralWpThemeColor(Color)
        End If
    Next

End Function