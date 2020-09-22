VERSION 5.00
Begin VB.Form frmAddEnum 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Enumerator to List"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNewEntry 
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
      Left            =   2580
      TabIndex        =   3
      Top             =   720
      Width           =   3495
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Move &Down in List"
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
      Left            =   6240
      TabIndex        =   7
      ToolTipText     =   "Move enumerator down in the list order (Down Arrow)"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Move &Up in List"
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
      Left            =   6240
      TabIndex        =   6
      ToolTipText     =   "Move enumerator up in the list order (Up Arrow)"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "De&lete"
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
      Left            =   6240
      TabIndex        =   8
      ToolTipText     =   "Delete the selected enumerator (DEL)"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Add/Update &Enum"
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
      Left            =   6240
      TabIndex        =   4
      ToolTipText     =   "Add new enumerator item (ENTER)"
      Top             =   720
      Width           =   1695
   End
   Begin VB.ListBox lstEnums 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   180
      TabIndex        =   5
      Top             =   1200
      Width           =   5895
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
      Left            =   2580
      TabIndex        =   1
      Top             =   180
      Width           =   5355
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
      Left            =   4380
      TabIndex        =   10
      ToolTipText     =   "Accept new enumerator"
      Top             =   3000
      Width           =   1695
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
      Left            =   6240
      TabIndex        =   11
      ToolTipText     =   "Reject or exit without saving"
      Top             =   3000
      Width           =   1695
   End
   Begin VB.PictureBox PicPtrBack 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2580
      ScaleHeight     =   195
      ScaleWidth      =   3675
      TabIndex        =   13
      Top             =   1020
      Width           =   3675
      Begin VB.PictureBox picptr 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   0
         ScaleHeight     =   555
         ScaleWidth      =   75
         TabIndex        =   14
         ToolTipText     =   "Marker position on line (Click to change)"
         Top             =   0
         Width           =   75
      End
      Begin VB.Label lblRuler 
         BackStyle       =   0  'Transparent
         Caption         =   "....|....|....|....|....|....|..."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   15
         ToolTipText     =   "Marker position on line (Click to change)"
         Top             =   0
         Width           =   3405
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
      Left            =   1020
      TabIndex        =   17
      Top             =   1020
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
      Left            =   2295
      TabIndex        =   16
      ToolTipText     =   "Marker position on line"
      Top             =   1020
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: Iintializing values and trailing comments are allowed (ie, MyItem=12 'Start at 12)"
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
      Left            =   180
      TabIndex        =   12
      Top             =   2580
      Width           =   7425
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000E&
      X1              =   180
      X2              =   7920
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      X1              =   180
      X2              =   7920
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter New Enumeration &Member:"
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
      Left            =   180
      TabIndex        =   2
      ToolTipText     =   "Enter new enumeration membe"
      Top             =   840
      Width           =   2370
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter New Enumerator &Name:"
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
      Left            =   180
      TabIndex        =   0
      ToolTipText     =   "Name to declare enumeration grouping as"
      Top             =   240
      Width           =   2145
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: Enumerators are considered valid if they do not clash with existing enumerators."
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
      TabIndex        =   9
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   120
      X2              =   7860
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      X1              =   120
      X2              =   7860
      Y1              =   2880
      Y2              =   2880
   End
End
Attribute VB_Name = "frmAddEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
Private colEntries As Collection    'keep track of unique enumerator entries
Private NameValid As Boolean        'True if selected name is valid
Private Ignore As Boolean           'True if state changes should be ignored
Private MouseDown As Boolean        'True if mouse down over control
'-------------------------------------------------------------------------------

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Init form
'*******************************************************************************
Private Sub Form_Load()
  Me.Icon = frmCom.Icon             'borrow icon
  Set colEntries = New Collection   'keep track of added enumerator entries
  
  Me.cmdUp.Enabled = False          'disable some buttons on startup
  Me.cmdDown.Enabled = False
  Me.cmdDelete.Enabled = False
  Me.cmdAddNew.Enabled = False
  Me.cmdApply.Enabled = False
  NameValid = False
  With Me.picptr
    .Left = CLng(GetSetting(App.Title, "Settings", "EnumPtr", "0"))
    Me.lblPosn.Caption = CStr(Fix(.Left / Me.lblRuler.Width * 33) + 1)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : unload form
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  Set colEntries = Nothing
  SaveSetting App.Title, "Settings", "EnumPtr", CStr(Me.picptr.Left)
End Sub

'*******************************************************************************
' Subroutine Name   : lstEnums_KeyDown
' Purpose           : Allow DEL key to select Delete button
'*******************************************************************************
Private Sub lstEnums_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case 46 'DEL
      If Me.cmdDelete.Enabled Then  'activate delete button if it is enabled
        Call cmdDelete_Click
        KeyCode = 0
      End If
  End Select
End Sub

'*******************************************************************************
' Subroutine Name   : txtName_Change
' Purpose           : Check if entry is valid
'*******************************************************************************
Private Sub txtName_Change()
  Dim Bol As Boolean
  Dim S As String
  
  Bol = False
  With Me.txtName
    Bol = CBool(Len(.Text))                       'initially valid if it contains text
    If Bol Then
      Bol = Not IsNumeric(.Text)                  'do not allow starting with digit
      If Bol Then
        S = "Enum " & .Text & vbCrLf              'see if already defined
        Bol = FindMatch(frmCom.lstEnum, S) = -1   'valid if nothing found
      End If
    End If
  End With
    
  NameValid = Bol                                 'mark flag
'
' enable/disable apply button
'
  Me.cmdApply.Enabled = Me.lstEnums.ListCount > 1 And NameValid
End Sub

'*******************************************************************************
' Subroutine Name   : txtName_KeyPress
' Purpose           : Parse enum name
'*******************************************************************************
Private Sub txtName_KeyPress(KeyAscii As Integer)
  Dim C As String
  
  Select Case KeyAscii
    Case 1 To 31
    Case Else
      C = Chr$(KeyAscii)
      Select Case UCase(C)
        Case "A" To "Z", "_", "0" To "9"
        Case Else
          KeyAscii = 0
      End Select
  End Select
End Sub

'*******************************************************************************
' Subroutine Name   : cmdAddNew_Click
' Purpose           : Add a new enumerator
'*******************************************************************************
Private Sub cmdAddNew_Click()
  Dim S As String, T As String, TT As String
  Dim Idx As Long
  
  If Ignore Then Exit Sub
  Ignore = True
  S = Trim$(Me.txtNewEntry.Text)          'grab text
  T = ExtractName(S)                      'grab just name, incase comment and/or initializer added
  On Error Resume Next
  colEntries.Add T, UCase$(T)             'try adding to local collection
  With Me.lstEnums
    If CBool(Err.Number) Then
      Idx = FindMatch(Me.lstEnums, T)
      .List(Idx) = S
    Else
      .AddItem Me.txtNewEntry.Text        'add to list
      Idx = .ListCount - 1                'index of new data
    End If
    .ListIndex = Idx                      'mark new/old selection
    Call lstEnums_Click
    Me.txtNewEntry.Text = vbNullString  'remove text data
    Me.cmdAddNew.Enabled = False
  End With
'
' enable/disable apply button
'
  Me.cmdApply.Enabled = CBool(Me.lstEnums.ListCount) And NameValid
  Me.txtNewEntry.SetFocus
  Ignore = False
End Sub

'*******************************************************************************
' Function Name     : ExtractName
' Purpose           : Extract enumerator name from definition
'*******************************************************************************
Private Function ExtractName(Txt As String) As String
  Dim S As String
  Dim I As Long, J As Long, K As Long
  
  S = Trim$(Txt)
  I = InStr(1, Txt, "'")
  J = InStr(1, Txt, "=")
  K = InStr(1, Txt, " ")
  If CBool(I) Then
    If CBool(J) And J < I Then I = J
    If CBool(K) And K < I Then I = K
  ElseIf CBool(J) Then
    I = J
    If CBool(K) And K < I Then I = K
  ElseIf CBool(K) Then
    I = K
  End If
  
  If CBool(I) Then
    ExtractName = RTrim$(Left$(Txt, I - 1))
  Else
    ExtractName = Txt
  End If
End Function

'*******************************************************************************
' Subroutine Name   : cmdCancel_Click
' Purpose           : Cancel changes and leave
'*******************************************************************************
Private Sub cmdCancel_Click()
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdDelete_Click
' Purpose           : Delete selected entry
'*******************************************************************************
Private Sub cmdDelete_Click()
  Dim Idx As Integer
  With Me.lstEnums
    Idx = .ListIndex
    .RemoveItem Idx
    If Idx = .ListCount Then Idx = .ListCount - 1
    .ListIndex = Idx
    Me.cmdApply.Enabled = .ListCount > 1 And NameValid
  End With
  Call RebuildCol
  Call lstEnums_Click
  Me.txtNewEntry.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : RebuildCol
' Purpose           : Rebuild collection
'*******************************************************************************
Private Sub RebuildCol()
  Dim Idx As Integer
  Dim S As String
  
  With colEntries
    Do While .Count
      .Remove 1
    Loop
  End With
  
  With Me.lstEnums
    For Idx = 0 To .ListCount - 1
      S = .List(Idx)
      colEntries.Add S, UCase$(S)
    Next Idx
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : cmdDown_Click
' Purpose           : Move entry down in list
'*******************************************************************************
Private Sub cmdDown_Click()
  Dim Idx As Integer
  Dim S As String
  
  With Me.lstEnums
    Idx = .ListIndex              'get current index
    S = .List(Idx)                'get text there
    .RemoveItem Idx               'remove from list
    Idx = Idx + 1                 'nove down one more in line
    If Idx < .ListCount Then      'if still below listcount, simply isnert it
      .AddItem S, Idx
    Else
      Idx = .ListCount            'else get new index for item (after add)
      .AddItem S                  'add to bottom of list
    End If
    .ListIndex = Idx              'set the selection point
  End With
  
  Call RebuildCol                 'rebuild unique collection
  Call lstEnums_Click             'select current item to set up display of its options
End Sub

'*******************************************************************************
' Subroutine Name   : cmdUp_Click
' Purpose           : Move entry up in list
'*******************************************************************************
Private Sub cmdUp_Click()
  Dim Idx As Integer
  Dim S As String
  
  With Me.lstEnums
    Idx = .ListIndex              'get current index
    S = .List(Idx)                'get text there
    .RemoveItem Idx               'remove from list
    Idx = Idx - 1                 'move up in list
    .AddItem S, Idx               'insert at new point
    .ListIndex = Idx              'mark selection
  End With
  Call RebuildCol                 'rebuild unique collection
  Call lstEnums_Click             'select current item to set up display of its options
End Sub

'*******************************************************************************
' Subroutine Name   : lstEnums_Click
' Purpose           : User selected an enumerator
'*******************************************************************************
Private Sub lstEnums_Click()
  Me.cmdUp.Enabled = False
  Me.cmdDown.Enabled = False
  
  With Me.lstEnums
    Me.cmdDelete.Enabled = CBool(.ListCount)
    If CBool(.ListCount) Then
      If .ListCount > 1 Then
        Me.cmdUp.Enabled = .ListIndex > 0
        Me.cmdDown.Enabled = .ListIndex < .ListCount - 1
      End If
      If Not Ignore Then
        Ignore = True
        Me.txtNewEntry.Text = .List(.ListIndex)
        With Me.txtNewEntry
          .SetFocus
          .SelStart = Len(.Text)
        End With
        Me.cmdAddNew.Enabled = True
        Ignore = False
      End If
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : txtNewEntry_Change
' Purpose           : Check for enabling Add button when text changes
'*******************************************************************************
Private Sub txtNewEntry_Change()
  Dim S As String, T As String
  Dim I As Long, J As Long
  
  If Ignore Then Exit Sub
  
  Me.lblRuler.ToolTipText = "column position: " & CStr(Me.txtNewEntry.SelStart + 1)
  
  S = Trim$(Me.txtNewEntry.Text)                                'grab text, trimmed
  If CBool(Len(S)) Then                                         'if data exists
    I = InStr(1, S, "'")                                        'check for a comment
    J = InStr(1, S, "=")
    If CBool(J) And J < I Then I = J
    If CBool(I) Then                                            'found one?
      T = Trim$(Left$(S, I - 1))                                'get data before it
      I = InStr(1, T, " ")                                      'check for an embedded space
      If CBool(I) Then T = vbNullString                         'if embedded, then illegal name
      Me.cmdAddNew.Enabled = CBool(Len(T)) And Not IsNumeric(S) 'set add button as required
    Else
      Me.cmdAddNew.Enabled = Not IsNumeric(S)
    End If
  Else
    Me.cmdAddNew.Enabled = False
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : txtNewEntry_GotFocus
' Purpose           : Select all text with this control gets focus
'*******************************************************************************
Private Sub txtNewEntry_GotFocus()
  With Me.txtNewEntry
    .SelStart = 0
    .SelLength = Len(.Text)
    Me.lblRuler.ToolTipText = "column position: " & CStr(.SelStart + 1)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : txtNewEntry_KeyDown
' Purpose           : Update Cursor position
'*******************************************************************************
Private Sub txtNewEntry_KeyDown(KeyCode As Integer, Shift As Integer)
  Me.lblRuler.ToolTipText = "column position: " & CStr(Me.txtNewEntry.SelStart + 1)
End Sub

'*******************************************************************************
' Subroutine Name   : txtNewEntry_KeyPress
' Purpose           : Parse new enumerator
'*******************************************************************************
Private Sub txtNewEntry_KeyPress(KeyAscii As Integer)
  Dim C As String
  
  Select Case KeyAscii
    Case 1 To 12, 14 To 31
    Case 13
    If Me.cmdAddNew.Enabled Then
      Call cmdAddNew_Click
      KeyAscii = 0
    End If
    Case Else
      C = Chr$(KeyAscii)
      Select Case UCase(C)
        Case "A" To "Z", "_", "0" To "9"
        Case "&", "=", "'", " "
        Case Else
          KeyAscii = 0
      End Select
  End Select
End Sub

'*******************************************************************************
' Subroutine Name   : cmdApply_Click
' Purpose           : Apply change
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
  Dim Idx As Integer
  
  S = "Enum " & Me.txtName & vbCrLf
  With Me.lstEnums
    For Idx = 0 To .ListCount - 1
      S = S & "  " & .List(Idx) & vbCrLf
    Next Idx
  End With
  S = S & "End Enum"
  
  DeclChange = S                                  'stuff new entry
  DeclName = Me.txtName.Text                      'new routine name
End Sub

'*******************************************************************************
' Routines support moving pointer on ruler
'*******************************************************************************
Private Sub PositionBar(Button As Integer, X As Single)
  Dim Idx As Long, WInc As Long
  
  If Button And vbLeftButton Then
    WInc = Me.lblRuler.Width \ 32
    Idx = Fix(X / Me.lblRuler.Width * 32)
    If Idx < 0 Or Idx > 32 Then Exit Sub
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

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************
