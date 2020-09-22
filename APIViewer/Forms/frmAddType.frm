VERSION 5.00
Begin VB.Form frmAddType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New User-Defined Type"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Referencing and Of Type"
      Height          =   1875
      Left            =   4080
      TabIndex        =   5
      Top             =   1200
      Width           =   4395
      Begin VB.TextBox txtSize 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1740
         TabIndex        =   13
         Text            =   "0"
         Top             =   1440
         Width           =   2595
      End
      Begin VB.CheckBox chkSize 
         Caption         =   "Ubound/Size:"
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
         Left            =   240
         TabIndex        =   11
         ToolTipText     =   "String size or Array Dim This can also be a Constant)"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkArray 
         Caption         =   "Define as an Array"
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
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "Select to define type item as an array"
         Top             =   1140
         Width           =   1695
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
         ItemData        =   "frmAddType.frx":0000
         Left            =   1740
         List            =   "frmAddType.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   300
         Width           =   2595
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
         ItemData        =   "frmAddType.frx":0053
         Left            =   1740
         List            =   "frmAddType.frx":0055
         TabIndex        =   9
         Top             =   720
         Width           =   2595
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(You can enter a constant)"
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
         Left            =   2400
         TabIndex        =   12
         Top             =   1260
         Width           =   1935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Standard Type:"
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
         ToolTipText     =   "Type to assign to new member"
         Top             =   360
         Width           =   1125
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
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "Type to assign to new member"
         Top             =   780
         Width           =   1380
      End
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
      Left            =   6720
      TabIndex        =   20
      ToolTipText     =   "Reject or exit without saving"
      Top             =   3480
      Width           =   1695
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
      Left            =   4920
      TabIndex        =   19
      ToolTipText     =   "Accept new User-Defined Type"
      Top             =   3480
      Width           =   1695
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
      Left            =   1980
      TabIndex        =   1
      Top             =   180
      Width           =   6435
   End
   Begin VB.ListBox lstTypeItems 
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
      Height          =   1425
      ItemData        =   "frmAddType.frx":0057
      Left            =   180
      List            =   "frmAddType.frx":0059
      TabIndex        =   14
      Top             =   1200
      Width           =   3795
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Add New &Type Item"
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
      Left            =   6720
      TabIndex        =   4
      ToolTipText     =   "Add new enumerator item (ENTER)"
      Top             =   660
      Width           =   1695
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
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
      Left            =   2820
      TabIndex        =   17
      ToolTipText     =   "Delete the selected Type Item (DEL)"
      Top             =   2700
      Width           =   1155
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Move &Up"
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
      Left            =   180
      TabIndex        =   15
      ToolTipText     =   "Move Type Item up in the list order (Up Arrow)"
      Top             =   2700
      Width           =   1155
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Move &Down"
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
      Left            =   1500
      TabIndex        =   16
      ToolTipText     =   "Move Type Item down in the list order (Down Arrow)"
      Top             =   2700
      Width           =   1155
   End
   Begin VB.TextBox txtNewEntry 
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
      Left            =   1980
      TabIndex        =   3
      Top             =   720
      Width           =   4635
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000E&
      X1              =   180
      X2              =   8400
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      X1              =   180
      X2              =   8400
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter New Type &Item:"
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
      ToolTipText     =   "New member of Type list"
      Top             =   780
      Width           =   1590
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   180
      X2              =   8340
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: User-Defined Types are considered valid if they do not clash with existing Types."
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
      Left            =   180
      TabIndex        =   18
      Top             =   3420
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter New Type &Name:"
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
      ToolTipText     =   "Name to assign to user-defined type"
      Top             =   240
      Width           =   1665
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      X1              =   180
      X2              =   8340
      Y1              =   3240
      Y2              =   3240
   End
End
Attribute VB_Name = "frmAddType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
Private colEntries As Collection    'keep track of unique Type Item entries
Private NameValid As Boolean        'True if selected name is valid
Private AsType As String          'storage for type selection
Private Ignore As Boolean         'ignore flag
'-------------------------------------------------------------------------------

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Init form
'*******************************************************************************
Private Sub Form_Load()
  Dim Idx As Integer
  Dim I As Long
  Dim S As String
  
  Me.Icon = frmCom.Icon             'borrow icon
  Set colEntries = New Collection   'keep track of added Type Item entries
  
  Me.cmdUp.Enabled = False          'disable some buttons on startup
  Me.cmdDown.Enabled = False
  Me.cmdDelete.Enabled = False
  Me.cmdAddNew.Enabled = False
  Me.cmdApply.Enabled = False
  Me.lstTypeItems.Clear
  NameValid = False
'
' build list of TYPE data
'
  With frmCom.lstType
    For Idx = 0 To .ListCount - 1                   'grab from full list of structures
      S = LTrim$(Mid$(.List(Idx), 6))               'grab an item (Strip 'Type ')
      I = InStr(1, S, vbCrLf)                       'find end of first line
      S = Left$(S, I - 1)
      I = InStr(1, S, "'")                          'strip any comments on it
      If CBool(I) Then S = RTrim$(Left$(S, I - 1))
      Me.cboTypeList.AddItem S                      'add to list
    Next Idx
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : unload form
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  Set colEntries = Nothing
End Sub

'*******************************************************************************
' Subroutine Name   : lstTypeItems_KeyDown
' Purpose           : Allow DEL key to select Delete button
'*******************************************************************************
Private Sub lstTypeItems_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case 46 'DEL
      If Me.cmdDelete.Enabled Then
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
        Bol = FindMatch(frmCom.lstType, S) = -1   'valid if nothing found
      End If
    End If
  End With
    
  NameValid = Bol                                 'mark flag
'
' enable/disable apply button
'
  Me.cmdApply.Enabled = CBool(Me.lstTypeItems.ListCount) And NameValid
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
      C = Chr$(KeyAscii)                  'get text version of code
      Select Case UCase(C)
        Case "A" To "Z", "_", "0" To "9"  'parse allowed characters
        Case Else
          KeyAscii = 0                    'out of range
      End Select
  End Select
End Sub

'*******************************************************************************
' Subroutine Name   : cmdAddNew_Click
' Purpose           : Add a new enumerator
'*******************************************************************************
Private Sub cmdAddNew_Click()
  Dim S As String, T As String
  
  S = Me.txtNewEntry.Text                     'grab text
  On Error Resume Next
  colEntries.Add S, UCase$(S)                 'try adding to local collection
  If Not CBool(Err.Number) Then               'if it is not a dumplication
    On Error GoTo 0
    Ignore = True
    T = Me.txtSize.Text
    If Not CBool(Len(T)) Then T = "0"
    If Right$(T, 1) = "," Then T = Left$(T, Len(T) - 1)
    Me.txtSize.Text = T
    If Me.chkArray.Value = vbChecked Then     'if array option is checked
      If Me.chkSize.Value = vbChecked Then    'if sizing option is checked
        S = S & "(" & T & ")"                 'set Dim on array
      Else
        S = S & "()"                          'else assume empty array
      End If
    End If
    
    With Me.lstTypeItems
      If CBool(Len(AsType)) Then              'if type has been defined...
        S = S & " As " & AsType               'apply type
      Else
        S = S & " As Long"                    'otherwise, apply default
      End If
      
      If Me.chkSize.Value = vbChecked And Me.chkArray.Value = vbUnchecked And AsType = "String" Then
        S = S & " * " & Me.txtSize.Text       'allow special sizing for non-array strings
      End If
      .AddItem S                              'add new text to display list
      .ListIndex = .ListCount - 1             'mark new selection
      Me.txtNewEntry.Text = vbNullString      'remove text data
      Ignore = False
      Call lstTypeItems_Click
    End With
  End If
'
' enable/disable apply button
'
  Me.cmdApply.Enabled = CBool(Me.lstTypeItems.ListCount) And NameValid
  Me.txtNewEntry.SetFocus
End Sub

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
  
  With Me.lstTypeItems
    Idx = .ListIndex                                      'get index to target
    .RemoveItem Idx                                       'remove item
    If Idx = .ListCount Then Idx = .ListCount - 1         'adjust index
    .ListIndex = Idx                                      'adjust selection to drop on next or last
    Me.cmdApply.Enabled = CBool(.ListCount) And NameValid 'set apply button enablement
  End With
  
  Call RebuildCol                                         'rebuild unique collection
  Call lstTypeItems_Click                                 'refresh options
  Me.txtNewEntry.SetFocus                                 'set focus to new entry field
End Sub

'*******************************************************************************
' Subroutine Name   : cmdDown_Click
' Purpose           : Move entry down in list
'*******************************************************************************
Private Sub cmdDown_Click()
  Dim Idx As Integer
  Dim S As String
  
  With Me.lstTypeItems
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
  Call lstTypeItems_Click         'select current item to set up display of its options
  Me.txtNewEntry.SetFocus         'and set focus for next new entry
End Sub

'*******************************************************************************
' Subroutine Name   : cmdUp_Click
' Purpose           : Move entry up in list
'*******************************************************************************
Private Sub cmdUp_Click()
  Dim Idx As Integer
  Dim S As String
  
  With Me.lstTypeItems
    Idx = .ListIndex              'get current index
    S = .List(Idx)                'get text there
    .RemoveItem Idx               'remove from list
    Idx = Idx - 1                 'move up in list
    .AddItem S, Idx               'insert at new point
    .ListIndex = Idx              'mark selection
  End With
  
  Call RebuildCol                 'rebuild unique collection
  Call lstTypeItems_Click         'select current item to set up display of its options
  Me.txtNewEntry.SetFocus         'and set focus for next new entry
End Sub

'*******************************************************************************
' Subroutine Name   : RebuildCol
' Purpose           : Rebuild collection
'*******************************************************************************
Private Sub RebuildCol()
  Dim Idx As Integer
  Dim I As Long, J As Long
  Dim S As String
'
' first erase collection
'
  With colEntries
    Do While .Count
      .Remove 1
    Loop
  End With
'
' now rebuild from accepted list
'
  With Me.lstTypeItems
    For Idx = 0 To .ListCount - 1
      S = .List(Idx)              'grab an item
      I = InStr(1, S, " ")
      J = InStr(1, S, "(")
      If CBool(J) And J < I Then I = J
      S = Left$(S, I - 1)
      colEntries.Add S, UCase$(S) 'add just name
    Next Idx
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : lstTypeItems_Click
' Purpose           : User selected a Type Item
'*******************************************************************************
Private Sub lstTypeItems_Click()
  Dim S As String, T As String
  Dim Idx As Long, I As Long
  
  If Ignore Then Exit Sub
  With Me.lstTypeItems
    If .ListIndex = -1 Then Exit Sub        'if nothing selected
    S = .List(.ListIndex)                   'else grab data
  End With
'
' find definition
'
  Ignore = True                             'prevent endless loops
  I = InStr(1, S, "(")                      'array dimensioning?
  If CBool(I) Then
    T = Mid$(S, I + 1)                      'grab array sizing
    I = InStr(1, T, ")")
    T = Left$(T, I - 1)                     'T contains dimensioning
'
' do array definition checking
'
    Me.chkArray.Value = vbChecked           'array set
    If CBool(Len(T)) Then                   'dimensioning also set?
      Me.chkSize.Value = vbChecked          'yes, to tag sizing
      Me.txtSize.Text = T                   'stuff size
    Else
      Me.chkSize.Value = vbUnchecked        'else untag sizing (at least for now)
    End If
  Else
    Me.chkArray.Value = vbUnchecked         'no array
    Me.chkSize.Value = vbUnchecked          'init to no sizing
  End If
'
' check for definition type
'
  Idx = InStr(1, S, " As ", vbTextCompare)  'find "AS"
  If CBool(Idx) Then
    AsType = LTrim$(Mid$(S, Idx + 4))         'get definition
'
' also check for special string fixed sizing
'
    I = InStr(1, AsType, "*")
    If CBool(I) Then
      Me.txtSize.Text = Mid$(AsType, I + 2)
      Me.chkSize.Value = vbChecked
      AsType = Left$(AsType, I - 2)
    End If
    
    Idx = FindMatch(Me.cboStdType, AsType)    'find in standard list
    If Idx <> -1 Then
      Me.cboStdType.ListIndex = Idx           'found in standard, so display selection
    Else
      Idx = FindMatch(Me.cboTypeList, AsType) 'else find in TYPE list
      Me.cboTypeList.ListIndex = Idx          'and make selection
    End If
  End If
  Ignore = False                              'allow active changes in combo options
  If Me.chkSize.Value = vbUnchecked Then
    Me.txtSize.Text = "0"                     'if sizing not set, then force 0 size
  End If
'
' now set options based upon position in the list
'
  With Me.lstTypeItems
    Me.cmdDelete.Enabled = CBool(.ListCount)
    If CBool(.ListCount) Then
      If .ListCount > 1 Then
        Me.cmdUp.Enabled = .ListIndex > 0
        Me.cmdDown.Enabled = .ListIndex < .ListCount - 1
        Exit Sub
      End If
    End If
  End With
  Me.cmdUp.Enabled = False
  Me.cmdDown.Enabled = False
End Sub

'*******************************************************************************
' Subroutine Name   : txtNewEntry_Change
' Purpose           : Check for enabling Add button when text changes
'*******************************************************************************
Private Sub txtNewEntry_Change()
  Dim S As String
  
  S = Me.txtNewEntry.Text
  Me.cmdAddNew.Enabled = CBool(Len(S)) And Not IsNumeric(S)
End Sub

'*******************************************************************************
' Subroutine Name   : txtNewEntry_GotFocus
' Purpose           : Select all text with this control gets focus
'*******************************************************************************
Private Sub txtNewEntry_GotFocus()
  With Me.txtNewEntry
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : txtNewEntry_KeyPress
' Purpose           : Parse new enumerator
'*******************************************************************************
Private Sub txtNewEntry_KeyPress(KeyAscii As Integer)
  Dim C As String
  
  Select Case KeyAscii
    Case 1 To 12, 14 To 31
    Case 13                               'CR
    If Me.cmdAddNew.Enabled Then
      Call cmdAddNew_Click                'force add New button
      KeyAscii = 0
    End If
    Case Else
      C = Chr$(KeyAscii)                  'get text version
      Select Case UCase(C)
        Case "A" To "Z", "_", "0" To "9"  'check allowances
        Case Else
          KeyAscii = 0                    'else invalid
      End Select
  End Select
End Sub

'*******************************************************************************
' Subroutine Name   : cmdApply_Click
' Purpose           : Apply change
'*******************************************************************************
Private Sub cmdApply_Click()
  Call ApplyChanges                       'apply user-defined changes
  Unload Me                               'exit form
End Sub

'*******************************************************************************
' Subroutine Name   : ApplyChanges
' Purpose           : Apply chages to user list of declarations
'*******************************************************************************
Private Sub ApplyChanges()
  Dim S As String
  Dim Idx As Integer
'
'
' build definition of type
  S = "Type " & Me.txtName & vbCrLf
  With Me.lstTypeItems
    For Idx = 0 To .ListCount - 1
      S = S & "  " & .List(Idx) & vbCrLf
    Next Idx
  End With
  S = S & "End Type"
  
  DeclChange = S                                  'stuff new entry
  DeclName = Me.txtName.Text                      'new routine name
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
  Dim S As String
  Dim Idx As Long, I As Long, J As Long
'
' add to displayed parameters at selected line
'
  With Me.lstTypeItems
    If .ListIndex = -1 Then Exit Sub            'if nothing selected
    S = .List(.ListIndex)                       'else get line
    I = InStr(1, S, " ")                        'find space or (
    J = InStr(1, S, "(")
    If CBool(J) And J < I Then I = J
    S = Left$(S, I - 1)                         'get just name
'
' check array dimensioning
'
    Ignore = True
    If Me.chkArray.Value = vbChecked Then
      If Me.chkSize.Value = vbChecked Then      'array sizing set
        S = S & "(" & Me.txtSize.Text & ")"
      Else
        S = S & "()"                            'no array sizing
      End If
    End If
'
' apply type
'
    S = S & " As " & AsType
'
' check special fixed-length string sizing
'
    If Me.chkSize.Value = vbChecked And Me.chkArray.Value = vbUnchecked And AsType = "String" Then
      If CBool(Len(Me.txtSize.Text)) Then
        If Me.txtSize.Text <> "0" Then S = S & " * " & Me.txtSize.Text
      End If
    End If
    Ignore = False
'
' apply to current line in list
'
    .List(.ListIndex) = S
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : txtSize_Change
' Purpose           : Sizing text changed
'*******************************************************************************
Private Sub txtSize_Change()
  Dim I As Long
  
  With Me.txtSize
    I = .SelStart
    If Len(.Text) = 0 Or Left$(.Text, 1) = "," Then 'check for blank or line starting with comma
      .Text = "0"                                   'ignore it
      .SelStart = 0
      .SelLength = 1
      Call BuildNewLine                             'rebuild line
    Else
      Call BuildNewLine                             'else rebuild line
      .SelStart = I                                 'and reset selection start
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : txtSize_GotFocus
' Purpose           : Select all text when sizing field gets focus
'*******************************************************************************
Private Sub txtSize_GotFocus()
  With Me.txtSize
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : txtSize_KeyPress
' Purpose           : Parse user entry on sizing fields
'*******************************************************************************
Private Sub txtSize_KeyPress(KeyAscii As Integer)
  Dim C As String
    
  Select Case KeyAscii
    Case 1 To 31
    Case Else
      C = UCase$(Chr$(KeyAscii))                    'get text version of code
      Select Case C
        Case "0" To "9"                             'allow 0-9
        Case "A" To "Z", "_"                        'allow A-Z and "_"
          KeyAscii = Asc(C)                         'ensure uppercase
        Case ","
          If Me.chkArray.Value = vbUnchecked Then   'if array dim not checked, so not allow multi-D
            KeyAscii = 0
          End If
        Case Else
          KeyAscii = 0
      End Select
  End Select
End Sub

'*******************************************************************************
' Subroutine Name   : txtSize_LostFocus
' Purpose           : Ensure sizing if correctly set when it loses focus
'*******************************************************************************
Private Sub txtSize_LostFocus()
  Dim S As String
  
  With Me.txtSize
    S = .Text                                           'get original text
    If Not CBool(Len(S)) Then S = "0"                   'make 0 if null
    If Right$(S, 1) = "," Then S = Left$(S, Len(S) - 1) 'if trailing comma, remove it
    If .Text <> S Then                                  'data changed?
      .Text = S                                         'yes, so set new text
      Call BuildNewLine                                 'and rebuild line
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : chkArray_Click
' Purpose           : Set/Reset Array definition
'*******************************************************************************
Private Sub chkArray_Click()
  Dim S As String
  Dim I As Long
  
  If Ignore Then Exit Sub
  
  If Me.chkArray.Value = vbUnchecked Then 'if array not checked, do not allow Multi-D
    With Me.txtSize
      S = .Text
       If Len(S) = 0 Then S = "0"
      I = InStr(1, S, ",")
      If CBool(I) Then
        S = Left$(S, I - 1)               'strip comma
      End If
      If .Text <> S Then                  'if data changed
        .Text = S                         'set new data
      End If
    End With
  End If
  Call BuildNewLine                       'and rebuild line
End Sub

'*******************************************************************************
' Subroutine Name   : chkSize_Click
' Purpose           : Set array size or string size
'*******************************************************************************
Private Sub chkSize_Click()
  If Ignore Then Exit Sub
  Call BuildNewLine
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************
