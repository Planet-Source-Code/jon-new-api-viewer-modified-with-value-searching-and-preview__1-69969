VERSION 5.00
Begin VB.Form frmModify 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify Declaration"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   8670
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstNames 
      Height          =   450
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   1695
   End
   Begin VB.ListBox lstOriginal 
      Height          =   450
      Left            =   120
      TabIndex        =   19
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdResetAll 
      Caption         =   "Reset &All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   15
      ToolTipText     =   "Reset all parameters to original settings"
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton cmdReset1 
      Caption         =   "&Reset Selection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      ToolTipText     =   "Reset current parameter to original setting"
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   17
      ToolTipText     =   "Cancel any changes and close dialog"
      Top             =   4260
      Width           =   2055
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      ToolTipText     =   "Apply changes to local definition only"
      Top             =   4260
      Width           =   2055
   End
   Begin VB.ComboBox cboTypeList 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmModify.frx":0000
      Left            =   5820
      List            =   "frmModify.frx":0002
      TabIndex        =   13
      Top             =   3060
      Width           =   2595
   End
   Begin VB.ComboBox cboStdType 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmModify.frx":0004
      Left            =   5820
      List            =   "frmModify.frx":0020
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2640
      Width           =   2595
   End
   Begin VB.ComboBox cboReference 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmModify.frx":005F
      Left            =   5820
      List            =   "frmModify.frx":006C
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2220
      Width           =   2595
   End
   Begin VB.TextBox txtNewName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2340
      TabIndex        =   5
      ToolTipText     =   "The new name must differ from thedeclaration it is based upon"
      Top             =   1560
      Width           =   6075
   End
   Begin VB.ListBox lstParameters 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   3735
   End
   Begin VB.TextBox txtDeclare 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      HideSelection   =   0   'False
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   8175
   End
   Begin VB.Label lblDecl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Declaration"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: Change the declaration name in order to save any modification."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
      Width           =   5880
   End
   Begin VB.Label lblRefAny 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING: Cannot have Reference verb for type ANY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4080
      TabIndex        =   21
      Top             =   1920
      Width           =   4410
   End
   Begin VB.Shape Shape1 
      Height          =   2595
      Left            =   4080
      Top             =   2160
      Width           =   4395
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000E&
      X1              =   4080
      X2              =   8460
      Y1              =   4140
      Y2              =   4140
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      X1              =   4080
      X2              =   8460
      Y1              =   4140
      Y2              =   4140
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   4080
      X2              =   8460
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      X1              =   4080
      X2              =   8460
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: Passing string types to API functions should normally be provided as 'ByVal'."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   4860
      Width           =   6945
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "...or &TYPE list item:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4200
      TabIndex        =   12
      ToolTipText     =   "Type parameter is declared as"
      Top             =   3120
      Width           =   1380
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New &Standard Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4200
      TabIndex        =   10
      ToolTipText     =   "Type parameter is declared as"
      Top             =   2700
      Width           =   1485
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Referencing &Verb:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4200
      TabIndex        =   8
      ToolTipText     =   "Establish ByRef, ByVal, or none parameter referencing"
      Top             =   2280
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&New Name for Declaration:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1620
      Width           =   1935
   End
   Begin VB.Label lblParms 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Parameter List:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Select a parameter to modify it"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblDeclaration 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Declaration for:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1125
   End
End
Attribute VB_Name = "frmModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------------------------------------------------------------
Private FnSub As String           'Function or Subroutine
Private OriginalDecl As String    'Original declaration
Private BeforeParen As String     'data before parentheses
Private AfterParen As String      'data after parentheses
Private Ignore As Boolean         'ignore flag
Private AsType As String          'storage for type selection
Private CanSaveOrg As Boolean     'True if we can save under original name
'-------------------------------------------------------------------------------

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Initialize form for display
'*******************************************************************************
Private Sub Form_Load()
  Dim S1 As Long, S2 As Long
  Dim Idx As Long, I As Long
  Dim S As String, Ary() As String
  
  Me.lstOriginal.Visible = False                                'hide original definition list
  Me.lstNames.Visible = False                                   'hide name list
  Me.lblRefAny.Caption = vbNullString
'
' first find declaration line
'
  With frmCom
    Me.Icon = .Icon                                             'copy icon
    With .txtSelectedItems
      S1 = InStrRev(.Text, vbCrLf, .SelStart + 1)               'find start of data
      S2 = InStr(.SelStart + 1, .Text, vbCrLf)                  'get end of line
      If S2 = 0 Then Exit Sub                                   'nothing to do
      If S1 = 0 Then
        S = Left$(.Text, S2 - 1)                                'grab text
      Else
        S = Mid$(.Text, S1 + 2, S2 - S1 - 2)                    'grab line
      End If
    End With
    Idx = InStr(1, S, " ")
    OriginalDecl = LTrim$(Mid$(S, Idx + 1))                     'skip public|private
    CanSaveOrg = FindExactMatch(.lstDecl, OriginalDecl) = -1
    Me.txtDeclare.Text = OriginalDecl                           'display original declaration
    Me.txtDeclare.SelStart = 0                                  'ensure at start of text
    
    S = LTrim$(Mid$(OriginalDecl, 9))                           'skip Declare
    Idx = InStr(1, S, " ")
    FnSub = "Declare " & Left$(S, Idx)                          'keep declaration header
    S = LTrim$(Mid$(S, Idx + 1))                                'skip Function|Subroutine
    Idx = InStr(1, S, " ")                                      'find space after declaration name
    DeclName = RTrim$(Left$(S, Idx - 1))                        'keep declaration name
    Me.lblDecl.Caption = DeclName                               'set caption
    Me.txtNewName.Text = DeclName                               'copy as 'new' name
    
    S = Mid$(S, Idx)                                            'grab data after declare name
    Idx = InStr(1, S, "(")                                      'find start of parameter list
    BeforeParen = Left$(S, Idx)                                 'hold data preceding parameters
    S = LTrim$(Mid$(S, Idx + 1))                                'strip left paren and before
    Idx = InStr(1, S, ")")                                      'find right paren
    AfterParen = Mid$(S, Idx)                                   'keep data following parameters
    S = RTrim$(Left$(S, Idx - 1))                               'get list of parameters
'
' build list of parameters
'
    If CBool(Len(S)) Then                                       'if data exists
      Ary = Split(S, ",")                                       'break up
      Me.lblParms.Caption = "Parameters: " & CStr(UBound(Ary) + 1)  'set parameter count
      For Idx = 0 To UBound(Ary)                                'process each parameter
        S = Trim$(Ary(Idx))                                     'grab one
        If CBool(Len(S)) Then                                   'data exists?
          Me.lstParameters.AddItem S                            'add to display list
          Me.lstOriginal.AddItem S                              'add to original list storage
  
          Select Case Left$(S, 5)
            Case "ByRef", "ByVal"                               'strip ByRef or ByVal, if it has one
              S = LTrim$(Mid$(S, 7))
          End Select
          I = InStr(1, S, " ")                                  'find name delimiter
          Me.lstNames.AddItem RTrim$(Left$(S, I - 1))           'keep only parameter name
        End If
      Next Idx
'
' build list of TYPE data
'
      With .lstType
        For Idx = 0 To .ListCount - 1                           'grab from full list of structures
          S = LTrim$(Mid$(.List(Idx), 6))                       'grab an item (Strip 'Type ')
          I = InStr(1, S, vbCrLf)                               'find end of first line
          S = Left$(S, I - 1)
          I = InStr(1, S, "'")                                  'strip any comments on it
          If CBool(I) Then S = RTrim$(Left$(S, I - 1))
          Me.cboTypeList.AddItem S                              'add to list
        Next Idx
      End With
      Me.lstParameters.ListIndex = 0                            'select first entry in list
'
' no parameters, so disable some additional stuff
'
    Else
      Me.cboReference.Enabled = False
      Me.cboStdType.Enabled = False
      Me.cboTypeList.Enabled = False
      Me.lstParameters.Enabled = False
      Me.lblParms.Enabled = False
    End If
  End With
  Me.cmdApply.Enabled = CanSaveOrg
  Me.cmdReset1.Enabled = False
  Me.cmdResetAll.Enabled = False
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCancel_Click
' Purpose           : Cancel and close form
'*******************************************************************************
Private Sub cmdCancel_Click()
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : lstParameters_Click
' Purpose           : A selection was mode on the list
'*******************************************************************************
Private Sub lstParameters_Click()
  Dim S As String
  Dim Idx As Long
  
  With Me.lstParameters
    If .ListIndex = -1 Then Exit Sub        'if nothing selected
    S = .List(.ListIndex)                   'else grab data
    Me.cmdReset1.Enabled = CBool(StrComp(S, Me.lstOriginal.List(.ListIndex)))
  End With
'
' choose reference verb
'
  Ignore = True                             'prevent echoes
  Select Case Left$(S, 5)
    Case "ByRef"
      Me.cboReference.ListIndex = 1
    Case "ByVal"
      Me.cboReference.ListIndex = 2
    Case Else
      Me.cboReference.ListIndex = 0
  End Select
'
' find definition
'
  Idx = InStr(1, S, " As ", vbTextCompare)  'find "AS"
  AsType = LTrim$(Mid$(S, Idx + 4))         'get definition
  Idx = FindMatch(Me.cboStdType, AsType)    'find in standard list
  If Idx <> -1 Then
    Me.cboStdType.ListIndex = Idx           'found in standard, so display selection
  Else
    Idx = FindMatch(Me.cboTypeList, AsType) 'else find in TYPE list
    Me.cboTypeList.ListIndex = Idx          'and make selection
  End If
  Ignore = False                            'allow active changes in combo options
End Sub

'*******************************************************************************
' Subroutine Name   : txtNewName_Change
' Purpose           : Changes made to textbox
'*******************************************************************************
Private Sub txtNewName_Change()
  Dim Bol As Boolean
  
  Bol = False
  With Me.txtNewName
    Bol = CBool(Len(.Text))
    If Bol Then
      Bol = CBool(StrComp(.Text, DeclName, vbTextCompare))
      If Bol Then
        Bol = Not IsNumeric(.Text)
        If Bol Then
          Bol = FindExactMatch(frmAPIViewer.LstItems, .Text) = -1
        End If
      End If
    End If
  End With
  Me.cmdApply.Enabled = Bol Or CanSaveOrg
End Sub

'*******************************************************************************
' Subroutine Name   : cboReference_Click
' Purpose           : When reference verb changes
'*******************************************************************************
Private Sub cboReference_Click()
  If Ignore Then Exit Sub
  Call BuildNewLine
End Sub

'*******************************************************************************
' Subroutine Name   : cboStdType_Click
' Purpose           : when declaration type changes in standard list
'*******************************************************************************
Private Sub cboStdType_Click()
  If Ignore Then Exit Sub
  With Me.cboStdType
    AsType = .List(.ListIndex)
  End With
  Call BuildNewLine
End Sub

'*******************************************************************************
' Subroutine Name   : cboTypeList_Click
' Purpose           : when declaration type changes in TYPE list
'*******************************************************************************
Private Sub cboTypeList_Click()
  If Ignore Then Exit Sub
  With Me.cboTypeList
    AsType = .List(.ListIndex)
    .Text = AsType
  End With
  Call BuildNewLine
End Sub

'*******************************************************************************
' Subroutine Name   : BuildNewLine
' Purpose           : Construct new definition of entry
'*******************************************************************************
Private Sub BuildNewLine()
  Dim S As String, T As String
  Dim Idx As Long, I As Long
  
  Select Case Me.cboReference.ListIndex
    Case 0  'none
      S = vbNullString
      T = "NOTE: Reference verb ByRef will be assumed"
    Case 1  'ByRef
      S = "ByRef "
      If StrComp(AsType, "Any", vbTextCompare) = 0 Then
        T = "NOTE: Reference verb ByRef set for type ANY"
      ElseIf StrComp(AsType, "String", vbTextCompare) = 0 Then
        T = "NOTE: Reference verb ByRef set for type String"
      Else
        T = vbNullString
      End If
    Case 2  'ByVal
      S = "ByVal "
      If StrComp(AsType, "Any", vbTextCompare) = 0 Then
        T = "NOTE: Reference verb ByVal set for type ANY"
      Else
        T = vbNullString
      End If
  End Select
  Me.lblRefAny.Caption = T
'
' construct new parameter definition
'
  S = S & Me.lstNames.List(Me.lstParameters.ListIndex) & " As " & AsType
'
' add to displayed parameters at selected line
'
  With Me.lstParameters
    .List(.ListIndex) = S
'
' now check for parameters changes from original
'
    I = 0
    For Idx = 0 To .ListCount - 1
      If .List(Idx) <> Me.lstOriginal.List(Idx) Then
        I = I + 1
      End If
    Next Idx
    Me.cmdResetAll.Enabled = CBool(I)
    Me.cmdReset1.Enabled = CBool(StrComp(S, Me.lstOriginal.List(.ListIndex)))
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : cmdResetAll_Click
' Purpose           : Reset all parameters to original definitions
'*******************************************************************************
Private Sub cmdResetAll_Click()
  Dim Idx As Long
  
  With Me.lstOriginal
    For Idx = 0 To .ListCount - 1
      Me.lstParameters.List(Idx) = .List(Idx) 'reset a parameter
    Next Idx
  End With
  Call lstParameters_Click                    'force reset of definition data
  Me.cmdResetAll.Enabled = False
  Me.cmdReset1.Enabled = False
End Sub

'*******************************************************************************
' Subroutine Name   : cmdReset1_Click
' Purpose           : Select selected parameter
'*******************************************************************************
Private Sub cmdReset1_Click()
  Dim Idx As Long, I As Long
  
  With Me.lstParameters
    .List(.ListIndex) = Me.lstOriginal.List(.ListIndex)
    Call lstParameters_Click                 'force reset of definition data
'
' now check for parameters changes from original
'
    I = 0
    For Idx = 0 To .ListCount - 1
      If .List(Idx) <> Me.lstOriginal.List(Idx) Then
        I = I + 1
      End If
    Next Idx
    Me.cmdResetAll.Enabled = CBool(I)
  End With
  Me.cmdReset1.Enabled = False
End Sub

'*******************************************************************************
' Subroutine Name   : cmdApply_Click
' Purpose           : Apply changes to Declaration and exit
'*******************************************************************************
Private Sub cmdApply_Click()
  Call ApplyChanges
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : ApplyChanges
' Purpose           : Apply chages to user list of declarations
'*******************************************************************************
Private Sub ApplyChanges()
  Dim S As String
  Dim Idx As Long
  
  S = FnSub & Me.txtNewName.Text & BeforeParen        'init start of declaration
  With Me.lstParameters
    If CBool(.ListCount) Then                         'if parameters exist
      For Idx = 0 To .ListCount - 2
        S = S & .List(Idx) & ", "                     'add parameters
      Next Idx
      S = S & .List(.ListCount - 1) & AfterParen
    End If
  End With
  
  DeclChange = S                                      'stuff new entry
  DeclName = Me.txtNewName.Text                       'new routine name
End Sub

'*******************************************************************************
' Subroutine Name   : txtNewName_KeyPress
' Purpose           : Filter keyboard so that invalid data cannot creep in
'*******************************************************************************
Private Sub txtNewName_KeyPress(KeyAscii As Integer)
  Dim C As String
  
  Select Case KeyAscii
    Case 1 To 31
    Case Else
      C = Chr$(KeyAscii)
      Select Case UCase$(C)
        Case "A" To "Z", "0" To "9", "_"
        Case Else
          KeyAscii = 0
      End Select
  End Select
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************
