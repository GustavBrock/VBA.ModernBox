# VBA.ModernBox

### Modern/Metro style message box and input box for Microsoft Access 2013+

Version 1.2.0

Modern/Metro styled message box and input box that directly can replace MsgBox() and InputBox()in Microsoft Access 2013 and later.
Also contains a prebuilt error box for use in error handling.

![General](https://raw.githubusercontent.com/GustavBrock/VBA.ModernBox/master/images/ModBox.png)

![General](https://raw.githubusercontent.com/GustavBrock/VBA.ModernBox/master/images/InputMox.png)

![General](https://raw.githubusercontent.com/GustavBrock/VBA.ModernBox/master/images/ErrorMox.png)

With version 1.2.0 support has been added for 64-bit Access.

The functions for calling the *HTML Help Viewer* control have been moved to a separate module.

With version 1.1.1 the boxes can not be moved beyond that of an Integer.

	' 2017-09-19: Added limitation of the settings for WindowsLeft and WindowsTop
	'             to be held within the range of Integer.
	
With version 1.1 a collection of helper functions are included:

	' Returns True if the passed colour value is one of the
	' Windows Phone Theme Colors.
	'
	' 2017-04-21. Gustav Brock, Cactus Data ApS, CPH.
	'
	Public Function IsWpThemeColor(ByVal Color As Long) As Boolean
	

and:

	' Returns the literal name of the passed colour value if
	' it is one of the Windows Phone Theme Colors.
	'
	' 2017-04-21. Gustav Brock, Cactus Data ApS, CPH.
	'
	Public Function LiteralWpThemeColor( _
	    ByVal Color As wpThemeColor) _
	    As String

also:

	' Loops all(!) possible color values and prints those of the
	' Windows Phone Theme Colors.
	' This will take nearly 30 seconds.
	'
	' 2017-04-21. Gustav Brock, Cactus Data ApS, CPH.
	'
	Public Function ListColors()

Full documentation is found here:

![EE Logo](https://raw.githubusercontent.com/GustavBrock/VBA.ModernBox/master/images/EE%20Logo.png)

[Modern/Metro style message box and input box for Microsoft Access 2013+](https://www.experts-exchange.com/articles/17684/Modern-Metro-style-message-box-and-input-box-for-Microsoft-Access-2013.html)

<hr>If you wish to support my work or need extended support or advice, feel free to:

<style>.bmc-button img{width: 27px !important;margin-bottom: 1px !important;box-shadow: none !important;border: none !important;vertical-align: middle !important;}.bmc-button{line-height: 36px !important;height:37px !important;text-decoration: none !important;display:inline-flex !important;color:#ffffff !important;background-color:#FF813F !important;border-radius: 3px !important;border: 1px solid transparent !important;padding: 0px 9px !important;font-size: 17px !important;letter-spacing:-0.08px !important;box-shadow: 0px 1px 2px rgba(190, 190, 190, 0.5) !important;-webkit-box-shadow: 0px 1px 2px 2px rgba(190, 190, 190, 0.5) !important;margin: 0 auto !important;font-family:'Lato', sans-serif !important;-webkit-box-sizing: border-box !important;box-sizing: border-box !important;-o-transition: 0.3s all linear !important;-webkit-transition: 0.3s all linear !important;-moz-transition: 0.3s all linear !important;-ms-transition: 0.3s all linear !important;transition: 0.3s all linear !important;}.bmc-button:hover, .bmc-button:active, .bmc-button:focus {-webkit-box-shadow: 0px 1px 2px 2px rgba(190, 190, 190, 0.5) !important;text-decoration: none !important;box-shadow: 0px 1px 2px 2px rgba(190, 190, 190, 0.5) !important;opacity: 0.85 !important;color:#ffffff !important;}</style><link href="https://fonts.googleapis.com/css?family=Lato&subset=latin,latin-ext" rel="stylesheet"><a class="bmc-button" target="_blank" href="https://www.buymeacoffee.com/gustav"><img src="https://bmc-cdn.nyc3.digitaloceanspaces.com/BMC-button-images/BMC-btn-logo.svg" alt="Buy me a coffee"><span style="margin-left:5px">Buy me a coffee</span></a>