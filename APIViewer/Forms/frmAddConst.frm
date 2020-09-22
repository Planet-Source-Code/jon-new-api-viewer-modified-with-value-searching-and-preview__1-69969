VERSION 5.00
Begin VB.Form frmAddConst 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Constant to List"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboConst 
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
      Left            =   4380
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Insert an established constant into the new constant value"
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Timer tmractivate 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3840
      Top             =   1920
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
      Left            =   6300
      TabIndex        =   8
      Top             =   1980
      Width           =   1695
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Default         =   -1  'True
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
      Left            =   4440
      TabIndex        =   7
      Top             =   1980
      Width           =   1695
   End
   Begin VB.TextBox txtValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2340
      TabIndex        =   3
      Top             =   600
      Width           =   5655
   End
   Begin VB.TextBox txtName 
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
      TabIndex        =   1
      Top             =   180
      Width           =   5655
   End
   Begin VB.PictureBox PicPtrBack 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2340
      ScaleHeight     =   195
      ScaleWidth      =   5775
      TabIndex        =   10
      Top             =   900
      Width           =   5775
      Begin VB.PictureBox picptr 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   0
         ScaleHeight     =   555
         ScaleWidth      =   75
         TabIndex        =   11
         ToolTipText     =   "Marker position on line (Click to change)"
         Top             =   0
         Width           =   75
      End
      Begin VB.Label lblRuler 
         BackStyle       =   0  'Transparent
         Caption         =   "....|....|....|....|....|....|....|....|....|....|...."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   12
         ToolTipText     =   "Marker position on line (Click to change)"
         Top             =   0
         Width           =   5655
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Alignment aid)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   840
      TabIndex        =   14
      Top             =   915
      Width           =   1080
   End
   Begin VB.Label lblPosn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2100
      TabIndex        =   13
      ToolTipText     =   "Marker position on line"
      Top             =   900
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: Values can have trailing comments (ie, &&H12 'Offset is at 18)"
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
      TabIndex        =   9
      Top             =   1500
      Width           =   5625
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   240
      X2              =   7980
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label lblInsert 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insert &Defined Constant:"
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
      Left            =   2400
      TabIndex        =   4
      ToolTipText     =   "Insert an established constant into the new constant value"
      Top             =   1140
      Width           =   1800
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: Constants are considered valid if they do not clash with existing constants."
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
      Height          =   435
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   3675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter New Constant &Value:"
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
      TabIndex        =   2
      ToolTipText     =   "This value can be/include other constants, Dec, hex, octal, or binary values and +/- offsets"
      Top             =   660
      Width           =   1950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter New Constant &Name:"
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
      ToolTipText     =   "Name to define new constant as"
      Top             =   240
      Width           =   1965
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      X1              =   240
      X2              =   7980
      Y1              =   1800
      Y2              =   1800
   End
End
Attribute VB_Name = "frmAddConst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
Private NameValid As Boolean  'true when name is not already defined
Private ValueValid As Boolean 'considered valid if it contains anything
Private MouseDown As Boolean  'True if mouse down over control
'-------------------------------------------------------------------------------

'*******************************************************************************
' Subroutine Name   : Form_Activate
' Purpose           : Erase data on form when activated
'*******************************************************************************
Private Sub Form_Activate()
  Me.txtName.Text = vbNullString
  Me.txtValue.Text = vbNullString
  Me.tmractivate.Enabled = True 'let timer set focus on name field
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Initialize form
'*******************************************************************************
Private Sub Form_Load()
  Dim Idx As Integer
  Dim I As Long, J As Long
  Dim S As String
  
  Screen.MousePointer = vbHourglass         'show busy...
  DoEvents
  Me.Icon = frmCom.Icon                     'borrow an icon
'
' start up with certain controls disabled
'
  Me.cmdApply.Enabled = False
  Me.cboConst.Enabled = False
  Me.lblInsert.Enabled = False
'
' build combo list with constant names
'
  With frmCom.lstConst
    For Idx = 0 To .ListCount - 1
      S = .List(Idx)                        'grab a constant
      I = InStr(7, S, " ")                  'find a space following name
      J = InStr(7, S, "=")                  'find = following name
      If J < I Then I = J                   'use lowest index
      Me.cboConst.AddItem Mid$(S, 7, I - 7) 'add name only
    Next Idx
    Me.cboConst.ListIndex = -1              'do not point to anything
  End With
''''--------------------------------------
'''  With colConst
'''    For Idx = 1 To .Count
'''      Me.cboConst.AddItem .Item(Idx)
'''    Next Idx
'''  End With
''''--------------------------------------
  With Me.picptr
    .Left = CLng(GetSetting(App.Title, "Settings", "ConstPtr", "0"))
    Me.lblPosn.Caption = CStr(Fix(.Left / Me.lblRuler.Width * 54) + 1)
  End With
  Screen.MousePointer = vbDefault           'no longer busy
End Sub

'*******************************************************************************
' Subroutine Name   : Form_QueryUnload
' Purpose           : Intercept closing form via X button
'*******************************************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Me.Hide
    Cancel = 1
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Save Pointer position
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  SaveSetting App.Title, "Settings", "ConstPtr", CStr(Me.picptr.Left)
End Sub

'*******************************************************************************
' Routines support moving pointer on ruler
'*******************************************************************************
Private Sub PositionBar(Button As Integer, X As Single)
  Dim Idx As Long, WInc As Long
  
  If Button And vbLeftButton Then
    WInc = Me.lblRuler.Width \ 53
    Idx = Fix(X / Me.lblRuler.Width * 53)
    If Idx < 0 Or Idx > 53 Then Exit Sub
    Me.lblPosn.Caption = CStr(Idx + 1)
    If CBool(Idx) Then Idx = Idx * WInc ' - Me.picptr.Width / 2
    Me.picptr.Left = Idx
  End If
End Sub

Private Sub lblRuler_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MouseDown = True
  Call PositionBar(vbLeftButton, X)
End Sub

Private Sub lblRuler_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button And vbLeftButton Then
    Call PositionBar(vbLeftButton, X)
  End If
End Sub

Private Sub lblRuler_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MouseDown = True
End Sub

Private Sub picptr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MouseDown = True
  Call PositionBar(vbLeftButton, CSng(Me.Left) + X)
End Sub

Private Sub picptr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Lft As Long
  
  If Button And vbLeftButton Then
    With Me.picptr
      Lft = .Left - (.Width \ 2 - CLng(X))
    End With
    Call PositionBar(vbLeftButton, CSng(Lft))
  End If
End Sub

Private Sub picptr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MouseDown = True
End Sub

'*******************************************************************************
' Subroutine Name   : tmractivate_Timer
' Purpose           : Timer used because form could simply be hidden.
'                   : This way we can reset focus to the first line
'*******************************************************************************
Private Sub tmractivate_Timer()
  Me.tmractivate.Enabled = False
  On Error Resume Next
  Me.txtName.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : txtName_Change
' Purpose           : Check if entry is valid
'*******************************************************************************
Private Sub txtName_Change()
  Dim Bol As Boolean
  Dim S As String
  
  Bol = False                                     'init to failure
  With Me.txtName
    Bol = CBool(Len(.Text))                       'initially valid if it contains text
    If Bol Then
      Bol = Not IsNumeric(.Text)                  'do not allow starting with digit
      If Bol Then
        S = "Const " & .Text & " ="               'see if already defined
        Bol = FindMatch(frmCom.lstConst, S) = -1  'valid if nothing found
      End If
    End If
  End With
  NameValid = Bol                                 'mark flag
  Me.cboConst.Enabled = Bol                       'enable combo data if valid
  Me.lblInsert.Enabled = Bol
  
  Me.cmdApply.Enabled = NameValid And ValueValid  'enable/disable apply button
End Sub

'*******************************************************************************
' Subroutine Name   : txtName_KeyPress
' Purpose           : Filter keyboard so that invalid data cannot creep in
'*******************************************************************************
Private Sub txtName_KeyPress(KeyAscii As Integer)
  Dim C As String
  
  Select Case KeyAscii
    Case 1 To 31
    Case Else
      C = UCase$(Chr$(KeyAscii))          'get character from code
      Select Case C
        Case "A" To "Z", "0" To "9", "_"  'range of allowed text
          KeyAscii = Asc(C)
        Case Else
          KeyAscii = 0                    'out of range
      End Select
  End Select
End Sub

'*******************************************************************************
' Subroutine Name   : txtValue_Change
' Purpose           : Check if entry is valid
'*******************************************************************************
Private Sub txtValue_Change()
  Dim S As String
  
  Me.lblRuler.ToolTipText = "column position: " & CStr(Me.txtValue.SelStart + 1)
  S = Trim$(Me.txtValue.Text)
  If Left$(S, 1) = "'" Then S = vbNullString        'if just comment, ignore
  ValueValid = CBool(Len(Trim$(S)))                 'assume OK if ANY data
  Me.cmdApply.Enabled = NameValid And ValueValid    'enable/disable apply button
End Sub

'*******************************************************************************
' Subroutine Name   : txtValue_KeyDown
' Purpose           : Update cursor position
'*******************************************************************************
Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)
  Me.lblRuler.ToolTipText = "column position: " & CStr(Me.txtValue.SelStart + 1)
End Sub

'*******************************************************************************
' Subroutine Name   : txtValue_KeyPress
' Purpose           : Filter keyboard so that invalid data cannot creep in
'*******************************************************************************
Private Sub txtValue_KeyPress(KeyAscii As Integer)
  Dim C As String, S As String
  Dim I As Long, J As Long, K As Long
  Dim AllowLC As Boolean                    'True when lowercase text allowed
  
  J = -1
  With Me.txtValue
    S = .Text                               'grab text
    I = InStr(1, S, "'")                    'comment present?
    If CBool(I) Then J = .SelStart + 1      'if so, check for selection point
  End With
  AllowLC = J >= I                          'set lowercase allowance flag
  
  S = Trim$(Me.txtValue.Text)
  Select Case KeyAscii
    Case 1 To 31
    Case Else
      C = UCase$(Chr$(KeyAscii))            'get text version of code
      If CBool(InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ &1234567890()'", C)) Then
        If Not AllowLC Then
          KeyAscii = Asc(C)
        End If
        Call TestValue
      ElseIf CBool(InStr(1, "+-", C)) Then  'if math involved, ensure enbraced by parens
        S = Trim$(Me.txtValue.Text)
        If Left$(S, 1) <> "(" Then          'no paren?
          S = "(" & S & ")"                 'no, so enclose it
          With Me.txtValue
            .Text = S
            .SelStart = Len(S) - 1          'set insert point immediately before ')'
          End With
        End If
      Else
        KeyAscii = 0                        'invalid data
      End If
  End Select
End Sub

'*******************************************************************************
' Subroutine Name   : TestValue
' Purpose           : Ensure logical entries are properly formatted
'                   : (not necessary, but just for looks)
'*******************************************************************************
Private Sub TestValue()
  Dim S As String
  Dim Idx As Long
  Dim Bol As Boolean
  
  With Me.txtValue
    S = " " & .Text               'prepend space in case we start with "Not"
    
    Idx = InStr(1, S, " OR ")     'found Ucase OR?
    Do While CBool(Idx)
      Mid$(S, Idx, 3) = " Or"     'yes, fix it
      Idx = InStr(Idx + 3, S, " OR ")
    Loop
    
    Idx = InStr(1, S, " AND ")    'found Ucase AND?
    Do While CBool(Idx)
      Mid$(S, Idx, 4) = " And"
      Idx = InStr(Idx + 4, S, " AND ")
    Loop
    
    Idx = InStr(1, S, " XOR ")    'found Ucase XOR?
    Do While CBool(Idx)
      Mid$(S, Idx, 4) = " Xor"
      Idx = InStr(Idx + 4, S, " XOR ")
    Loop
    
    Idx = InStr(1, S, " NOT ")    'found Ucase NOT?
    Do While CBool(Idx)
      Mid$(S, Idx, 4) = " Not"
      Idx = InStr(Idx + 4, S, " NOT ")
    Loop
    '
    ' if we fodund any logical flags, ensure expression embraced
    '
    If CBool(InStr(1, S, " Or ")) Or _
       CBool(InStr(1, S, " And ")) Or _
       CBool(InStr(1, S, " Xor ")) Or _
       CBool(InStr(1, S, " Not ")) Then
      If Left$(S, 2) <> " (" Then   'paren?
        S = " (" & Mid$(S, 2) & ")" 'no, so enclose it
        Idx = .SelStart
        .Text = Mid$(S, 2)
        .SelStart = Idx + 1         'set insert point immediately before ')'
      End If
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : cboConst_Click
' Purpose           : append a defined constant
'*******************************************************************************
Private Sub cboConst_Click()
  Dim Idx As Long
  
  On Error Resume Next
  With Me.txtValue
    Idx = .SelStart + Len(Me.cboConst.Text) 'define insertion point
    .SelText = Me.cboConst.Text             'insert selection
    .SelStart = Idx                         'set insert point after added text
    .SetFocus
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCancel_Click
' Purpose           : Cancel dialog
'*******************************************************************************
Private Sub cmdCancel_Click()
  Me.Hide                                   'hiding helps prevent long rebuild of Const Combobox
End Sub

'*******************************************************************************
' Subroutine Name   : cmdApply_Click
' Purpose           : Apply change
'*******************************************************************************
Private Sub cmdApply_Click()
  Call TestValue                            'check value for logical entries
  Call ApplyChanges                         'add changes user made
  Me.Hide                                   'hiding helps prevent long rebuild of Const Combobox
End Sub

'*******************************************************************************
' Subroutine Name   : ApplyChanges
' Purpose           : Apply chages to user list of declarations
'*******************************************************************************
Private Sub ApplyChanges()
  DeclChange = "Const " & Me.txtName & " = " & Me.txtValue  'set declaration for new entry
  DeclName = Me.txtName.Text                                'new routine name
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

