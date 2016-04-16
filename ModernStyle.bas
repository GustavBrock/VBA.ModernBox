Attribute VB_Name = "ModernStyle"
Option Compare Database
Option Explicit

Public Sub StyleCommandButtons(ByRef frm As Form)

' Apply a style to all non-transparent command buttons on a form.
' 2014-10-10. Gustav Brock, Cactus Data ApS, CPH.
' Version 1.0.0
' License: MIT.

' Requires:
'   Module:
'       ModernThemeColours

' Typical usage:
'
'   Private Sub Form_Load()
'       Call StyleCommandButtons(Me)
'   End Sub

    Dim ctl                 As Control
    
    For Each ctl In frm.Controls
        If ctl.ControlType = acCommandButton Then
            If ctl.Transparent = True Then
                ' Leave transparent buttons untouched.
            Else
                ctl.Height = 454
                ctl.UseTheme = True
                If ctl.Default = True Then
                    ctl.BackColor = wpThemeColor.Cobalt
                Else
                    ctl.BackColor = ctl.Parent.Section(ctl.Section).BackColor
                End If
                ctl.HoverForeColor = ctl.BackColor
                ctl.HoverColor = wpThemeColor.White
                ctl.PressedColor = wpThemeColor.Darken
                ctl.BorderWidth = 2
                ctl.BorderStyle = 1
                ctl.BorderColor = wpThemeColor.White
                ctl.ForeColor = wpThemeColor.White
                ctl.FontName = "Segoe UI"
                ctl.FontSize = 11
                ctl.FontBold = True
                ctl.FontItalic = False
            End If
        End If
    Next
    
    Set ctl = Nothing

End Sub
