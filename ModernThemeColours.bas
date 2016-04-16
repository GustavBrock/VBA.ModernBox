Attribute VB_Name = "ModernThemeColours"
Option Compare Database
Option Explicit

' Adoption of Windows Phone 7.5/8.0 colour theme for VBA.
' 2014-10-10. Gustav Brock, Cactus Data ApS, CPH.
' Version 1.0.0
' License: MIT.

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

