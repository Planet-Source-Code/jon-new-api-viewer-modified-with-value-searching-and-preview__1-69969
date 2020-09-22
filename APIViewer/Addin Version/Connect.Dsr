VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9840
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   10935
   _ExtentX        =   19288
   _ExtentY        =   17357
   _Version        =   393216
   Description     =   "Standard API Viewer replacement application"
   DisplayName     =   "New API Viewer"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "None"
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed          As Boolean                  'flag indicating if form is displayed
Public VBInstance             As VBIDE.VBE                'VB Instance object
Dim mcbMenuCommandBar         As Office.CommandBarControl 'menu bar connection
Dim mfrmAddIn                 As New frmAPIViewer             'select main form
Public WithEvents MenuHandler As CommandBarEvents         'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1

'*******************************************************************************
' Subroutine Name   : Hide
' Purpose           : Hide the form
'modAPIConstants
'*******************************************************************************
Sub Hide()
  On Error Resume Next
  FormDisplayed = False       'indicate form not displayed
  mfrmAddIn.Hide              'hide form
End Sub

'*******************************************************************************
' Subroutine Name   : Show
' Purpose           : Show the form
'*******************************************************************************
Sub Show()
  On Error Resume Next
  If mfrmAddIn Is Nothing Then            'create form if not yet defined
    Set mfrmAddIn = New frmAPIViewer
  End If
  
  Set mfrmAddIn.VBInstance = VBInstance   'set VB instance
  Set mfrmAddIn.Connect = Me              'set connection object
  FormDisplayed = True                    'tag as displayed
  mfrmAddIn.Show                          'then show it
End Sub

'*******************************************************************************
' Subroutine Name   : AddinInstance_OnConnection
' Purpose           : this method adds the Add-In to VB
'*******************************************************************************
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
  'On Error GoTo error_handler
  
  'save the vb instance
  Set VBInstance = Application
  
  'this is a good place to set a breakpoint and
  'test various addin objects, properties and methods
  Debug.Print VBInstance.FullName

  If ConnectMode = ext_cm_External Then
    'Used by the wizard toolbar to start this wizard
    Me.Show
  Else
    'Set mcbMenuCommandBar = AddToAddInCommandBar("New API Viewer")
    'sink the event
    'Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    AddToCommandBar
  End If

'  If ConnectMode = ext_cm_AfterStartup Then
'    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
'      'set this to display the form on connect
'      Me.Show
'    End If
'  End If
  Exit Sub
  
error_handler:
  MsgBox Err.Description
End Sub

'*******************************************************************************
' Subroutine Name   : AddinInstance_OnDisconnection
' Purpose           : this method removes the Add-In from VB
'*******************************************************************************
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
  On Error Resume Next
'
'delete the command bar entry
'
  If Not mcbMenuCommandBar Is Nothing Then mcbMenuCommandBar.Delete
'
'shut down the Add-In
'
'  If FormDisplayed Then
'    SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
'    FormDisplayed = False
'  Else
'    SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
'  End If
  
  Unload mfrmAddIn        'unload form
  Set mfrmAddIn = Nothing 'remove object
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
'  If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
'    'set this to display the form on connect
'    Me.Show
'  End If
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Me.Show
End Sub

Private Sub AddToCommandBar()
  
  On Error GoTo AddToCommandBarErr
  
  ' if you prefer a button on the commandbar uncomment the
  ' next 2 lines, and comment the one after.
  
  'VBInstance.CommandBars(2).Visible = True
  'Set mcbMenuCommandBar = VBInstance.CommandBars(2).Controls.Add(1, , , VBInstance.CommandBars(2).Controls.Count)

  ' this adds a menu item
  
  Set mcbMenuCommandBar = VBInstance.CommandBars("Add-Ins").Controls.Add(1, , , VBInstance.CommandBars("Add-Ins").Controls.Count)
  
  If Not mcbMenuCommandBar Is Nothing Then
    
    mcbMenuCommandBar.Caption = "New API Viewer"
  
    mfrmAddIn.pic.BackColor = vbMenuBar
    Set mfrmAddIn.pic.Picture = mfrmAddIn.imgIcon.Picture
    Set mfrmAddIn.pic.Picture = mfrmAddIn.pic.Image
  
    Clipboard.SetData mfrmAddIn.pic.Picture
    mcbMenuCommandBar.PasteFace
  
    Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
  
  End If
  
  Exit Sub
    
AddToCommandBarErr:
  MsgBox Err.Description
  
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
  Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
  Dim cbMenu As Object

  On Error GoTo AddToAddInCommandBarErr
'
'see if we can find the Add-Ins menu
'
  Set cbMenu = VBInstance.CommandBars("Add-Ins")
  If cbMenu Is Nothing Then
    '
    'not available so we fail
    '
    Exit Function
  End If
'
'add it to the command bar
'
  Set cbMenuCommandBar = cbMenu.Controls.Add(1)
  '
  'set the caption
  '
  cbMenuCommandBar.Caption = sCaption
  
'  mfrmAddIn.pic.BackColor = vbMenuBar
'  Set mfrmAddIn.pic.Picture = mfrmAddIn.imgIcon.Picture
'  Set mfrmAddIn.pic.Picture = mfrmAddIn.pic.Image
'
'  Clipboard.SetData mfrmAddIn.pic.Picture
'  cbMenuCommandBar.PasteFace
  
  Set AddToAddInCommandBar = cbMenuCommandBar
  
AddToAddInCommandBarErr:
End Function

