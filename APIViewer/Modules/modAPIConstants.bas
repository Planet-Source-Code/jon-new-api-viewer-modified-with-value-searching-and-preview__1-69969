Attribute VB_Name = "modAPIConstants"
Option Explicit

' **********************************************************************
'
' Module:         modAPIConstants
' Author:         Sharp Dressed Codes
' Web:            http://sharpdressedcodes.com
' Built:          20th January, 2008
' Purpose:        This is used for Constants searching via the textbox.
'                 eg type in &H9 and all the constants with the value of
'                 &H9 will be show in the preview window.
' Comments:       I find this very useful when intercepting Windows Messages.
'                 This helps find what the messages are, alot faster...
' Bugs:           -
'
' **********************************************************************

Public Type APIConstantType
  Name As String
  Value As String 'Variant
  Index As Long
  OtherNames As New Collection
End Type

Public APIConstants() As APIConstantType
