VERSION 5.00
Begin VB.Form frmAddDecl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Declare Statement"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   8595
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboRtnType 
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
      ItemData        =   "frmAddDecl.frx":0000
      Left            =   5700
      List            =   "frmAddDecl.frx":0002
      TabIndex        =   11
      Top             =   1440
      Width           =   2715
   End
   Begin VB.ComboBox cboReturnStd 
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
      ItemData        =   "frmAddDecl.frx":0004
      Left            =   1980
      List            =   "frmAddDecl.frx":0020
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1440
      Width           =   2055
   End
   Begin VB.ComboBox cboFnSub 
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
      ItemData        =   "frmAddDecl.frx":0062
      Left            =   4380
      List            =   "frmAddDecl.frx":006C
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1020
      Width           =   4035
   End
   Begin VB.TextBox txtLIB 
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
      TabIndex        =   3
      Top             =   420
      Width           =   4035
   End
   Begin VB.TextBox txtAlias 
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
      TabIndex        =   5
      Top             =   1020
      Width           =   4035
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
      TabIndex        =   29
      ToolTipText     =   "Move Parameter down in the list order (Down Arrow)"
      Top             =   4320
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
      TabIndex        =   28
      ToolTipText     =   "Move Parameter up in the list order (Up Arrow)"
      Top             =   4320
      Width           =   1155
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
      Left            =   2820
      TabIndex        =   30
      ToolTipText     =   "Delete the selected Parameter (DEL)"
      Top             =   4320
      Width           =   1155
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
      Height          =   2010
      ItemData        =   "frmAddDecl.frx":007F
      Left            =   180
      List            =   "frmAddDecl.frx":0081
      TabIndex        =   27
      Top             =   2220
      Width           =   3795
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
      Left            =   240
      TabIndex        =   1
      Top             =   420
      Width           =   4035
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
      TabIndex        =   32
      ToolTipText     =   "Accept new Declaration"
      Top             =   5040
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
      Left            =   6720
      TabIndex        =   33
      ToolTipText     =   "Reject or exit without saving"
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add New Prameter / Edit Selected Parameter properties"
      Height          =   2715
      Left            =   4020
      TabIndex        =   12
      Top             =   1980
      Width           =   4395
      Begin VB.PictureBox picFixBorder 
         Height          =   675
         Left            =   60
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1455
         Begin VB.CommandButton cmdAddNew 
            Caption         =   "Add Para&meter"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   60
            TabIndex        =   25
            ToolTipText     =   "Add new declare parameter (ENTER)"
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.TextBox txtSize 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2880
         TabIndex        =   23
         Text            =   "0"
         Top             =   2220
         Width           =   1335
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
         Left            =   1620
         TabIndex        =   14
         Top             =   240
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
         ItemData        =   "frmAddDecl.frx":0083
         Left            =   1620
         List            =   "frmAddDecl.frx":0090
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   660
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
         ItemData        =   "frmAddDecl.frx":00AA
         Left            =   1620
         List            =   "frmAddDecl.frx":00AC
         TabIndex        =   20
         Top             =   1500
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
         ItemData        =   "frmAddDecl.frx":00AE
         Left            =   1620
         List            =   "frmAddDecl.frx":00CA
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1080
         Width           =   2595
      End
      Begin VB.CheckBox chkArray 
         Caption         =   "Define as an A&rray"
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
         Left            =   1620
         TabIndex        =   21
         ToolTipText     =   "Select to define parameter as an array"
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox chkSize 
         Caption         =   "Ubound/Si&ze:"
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
         Left            =   1620
         TabIndex        =   22
         ToolTipText     =   "String size or Array Dim This can also be a Constant)"
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&New Parametrer:"
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
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "When this field is blank, you can edit existing parameter properties"
         Top             =   300
         Width           =   1230
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
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Establish ByRef, ByVal, or none parameter referencing"
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "...or TY&PE list item:"
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
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Type parameter is declared as"
         Top             =   1560
         Width           =   1380
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Standard T&ype:"
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
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Type parameter is declared as"
         Top             =   1140
         Width           =   1125
      End
   End
   Begin VB.Label lblTemplate 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(Template)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   180
      TabIndex        =   34
      Top             =   5520
      Width           =   8235
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000E&
      X1              =   240
      X2              =   8400
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      X1              =   240
      X2              =   8400
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Label Label12 
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
      TabIndex        =   10
      ToolTipText     =   "Return type for Function declarations"
      Top             =   1500
      Width           =   1380
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Standard Return Type:"
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
      ToolTipText     =   "Return type for Function declarations"
      Top             =   1500
      Width           =   1665
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Establish &Function or Subroutine Definition"
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
      Left            =   4440
      TabIndex        =   6
      ToolTipText     =   "Declaration is defined as a Function or a Sub(routine)"
      Top             =   780
      Width           =   3030
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parameter list, in se&quential order:"
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
      TabIndex        =   26
      ToolTipText     =   "Select a parameter to edit its properties or move it"
      Top             =   1980
      Width           =   2490
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Library (LIB) Name used by Declaration:"
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
      Left            =   4380
      TabIndex        =   2
      ToolTipText     =   "The name of the DLL library, less '.DLL'"
      Top             =   180
      Width           =   2865
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Alias Name (if different):"
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
      ToolTipText     =   "ACTUAL name of member definition in the DLL Library (LIB)"
      Top             =   780
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter New Declare &Name:"
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
      ToolTipText     =   "Name for VB to use to access this declaration"
      Top             =   180
      Width           =   1845
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: Declares are considered valid if they do not clash with other declarations."
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
      TabIndex        =   31
      Top             =   4980
      Width           =   3975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   180
      X2              =   8340
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      X1              =   180
      X2              =   8340
      Y1              =   4800
      Y2              =   4800
   End
End
Attribute VB_Name = "frmAddDecl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
Private colEntries As Collection  'keep track of unique Type Item entries
Private NameValid As Boolean      'True if selected name is valid
Private AliasValid As Boolean     'True if Alias is valid
Private LibValid As Boolean       'True if LIB valid
Private ReturnType As String      'function return type
Private AsType As String          'parameter storage type
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
  
  Me.picFixBorder.BorderStyle = 0
  Me.cmdUp.Enabled = False          'disable some buttons on startup
  Me.cmdDown.Enabled = False
  Me.cmdDelete.Enabled = False
  Me.cmdAddNew.Enabled = False
  Me.cmdApply.Enabled = False
  Me.lstTypeItems.Clear
  Me.txtNewEntry.Enabled = False
  Me.txtSize.Text = "0"
  Me.cboFnSub.ListIndex = 0
  Me.cboReference.ListIndex = 0
  Me.cboStdType.ListIndex = 4
  Me.cboReturnStd.ListIndex = 4
  NameValid = False
  AliasValid = True
  LibValid = False
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
      Me.cboTypeList.AddItem S                      'add to list two lists on form
      Me.cboRtnType.AddItem S
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
      If Me.cmdDelete.Enabled Then  'if delete button enabled, then select it
        Call cmdDelete_Click
        KeyCode = 0                 'remove possibility for hiccup
      End If
  End Select
End Sub

'*******************************************************************************
' Subroutine Name   : txtAlias_Change
' Purpose           : Check if entry is valid
'*******************************************************************************
Private Sub txtAlias_Change()
  Dim Bol As Boolean
  Dim S As String
'
' build definition of type
'
  Call UpdateTemplate
  
  Bol = False
  With Me.txtAlias
    Bol = Not CBool(Len(.Text))                   'initially valid if it contains text
    If CBool(Len(.Text)) Then
      Bol = Not IsNumeric(.Text)                  'do not allow starting with digit
    End If
  End With
    
  AliasValid = Bol                                 'mark flag
'
' enable/disable apply button
'
  EnableParm NameValid And LibValid And AliasValid
  Me.cmdApply.Enabled = CBool(Me.lstTypeItems.ListCount) And NameValid And LibValid And AliasValid
End Sub

'*******************************************************************************
' Subroutine Name   : txtAlias_KeyPress
' Purpose           : Parse enum name
'*******************************************************************************
Private Sub txtAlias_KeyPress(KeyAscii As Integer)
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
' Subroutine Name   : txtLIB_Change
' Purpose           : Check if entry is valid
'*******************************************************************************
Private Sub txtLIB_Change()
  Dim Bol As Boolean
  Dim S As String
'
' build definition of type
'
  Call UpdateTemplate
  
  Bol = False
  With Me.txtLIB
    Bol = CBool(Len(.Text))                       'initially valid if it contains text
    If Bol Then
      Bol = Not IsNumeric(.Text)                  'do not allow starting with digit
    End If
  End With
    
  LibValid = Bol                                  'mark flag
'
' enable/disable apply button
'
  EnableParm NameValid And LibValid And AliasValid
  Me.cmdApply.Enabled = CBool(Me.lstTypeItems.ListCount) And NameValid And LibValid And AliasValid
End Sub

'*******************************************************************************
' Subroutine Name   : txtLIB_KeyPress
' Purpose           : Parse enum name
'*******************************************************************************
Private Sub txtLIB_KeyPress(KeyAscii As Integer)
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
' Subroutine Name   : txtName_Change
' Purpose           : Check if entry is valid
'*******************************************************************************
Private Sub txtName_Change()
  Dim Bol As Boolean
  Dim S As String
'
' build definition of type
'
  Call UpdateTemplate
'
' check for valid data
'
  Bol = False
  With Me.txtName
    Bol = CBool(Len(.Text))                       'initially valid if it contains text
    If Bol Then
      Bol = Not IsNumeric(.Text)                  'do not allow starting with digit
      If Bol Then
        S = "Declare " & Me.cboFnSub.Text & .Text & " " 'see if already defined
        Bol = FindMatch(frmCom.lstDecl, S) = -1   'valid if nothing found
      End If
    End If
  End With
    
  NameValid = Bol                                 'mark flag
'
' enable/disable apply button
'
  EnableParm NameValid And LibValid And AliasValid
  Me.cmdApply.Enabled = CBool(Me.lstTypeItems.ListCount) And NameValid And LibValid And AliasValid
End Sub

'*******************************************************************************
' Subroutine Name   : UpdateTemplate
' Purpose           : Set up and display the declaration template
'*******************************************************************************
Private Sub UpdateTemplate()
  Dim S As String, Tn As String, Tl As String, Ta As String, Tp As String
  Dim Idx As Integer
'
' build definition of type
'
  Tn = Me.txtName.Text                            'grab name for declaration
  Tl = Me.txtLIB.Text                             'grab library name
  If Not CBool(Len(Tl)) Then Tl = "<NOT DEFINED>" 'if library not yet defined
  Ta = Me.txtAlias.Text                           'get alias name
  
  If CBool(Len(Tn)) Then                          'if declaration name defined
    S = "Declare " & Me.cboFnSub.Text & " " & Tn & " LIB """ & Tl & """"
    If CBool(Len(Me.txtAlias.Text)) Then          'add alias if present
      S = S & " Alias """ & Ta & """ ()"
    Else
      S = S & " ()"
    End If
    If CBool(Len(ReturnType)) And Me.cboFnSub.Text <> "Sub" Then
      S = S & " As " & ReturnType                 'add return type if not subroutine
    End If
    Tp = vbNullString
    With Me.lstTypeItems
      For Idx = 0 To .ListCount - 1
        Tp = Tp & .List(Idx) & ", "
      Next Idx
      If CBool(Len(Tp)) Then
        Idx = InStr(1, S, "(")
        S = Left$(S, Idx) & Left$(Tp, Len(Tp) - 2) & Mid$(S, Idx + 1)
      End If
    End With
  Else
    S = "(Template)"                              'no declaration name present
  End If
  Me.lblTemplate.Caption = S                      'stuff result
End Sub

'*******************************************************************************
' Subroutine Name   : EnableParm
' Purpose           : Enable parameters as needed
'*******************************************************************************
Private Sub EnableParm(Flag As Boolean)
  Me.txtNewEntry.Enabled = Flag
  Me.cmdAddNew.Enabled = CBool(Len(Me.txtNewEntry.Text))
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
    Select Case Me.cboReference.ListIndex
      Case 1
        S = "ByRef " & S
      Case 2
        S = "ByVal " & S
    End Select
    T = Me.txtSize.Text                       'get array/string sizing
    If Not CBool(Len(T)) Then T = "0"         'default to null if no data
    If Right$(T, 1) = "," Then T = Left$(T, Len(T) - 1) 'remove trailing coma
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
      Me.txtNewEntry.Text = vbNullString      'remove text data in prep for next prep
      Ignore = False                          'now allow processing
      Call lstTypeItems_Click                 'and process it
    End With
  End If
'
' enable/disable apply button
'
  Me.cmdApply.Enabled = CBool(Me.lstTypeItems.ListCount) And NameValid
  Call UpdateTemplate
  Me.txtNewEntry.SetFocus                     'and set focus for next parameter entry
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
  Call UpdateTemplate
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
  Call UpdateTemplate
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
  Call UpdateTemplate
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
      S = .List(Idx)                    'grab an item
      Select Case Left$(S, 6)           'skip any ByRef/ByVal entry
        Case "ByRef ", "ByVal "
          S = Mid$(S, 7)
      End Select
      I = InStr(1, S, " ")              'find space after name
      J = InStr(1, S, "(")              'find parameter list
      If CBool(J) And J < I Then I = J  'set to lower item
      S = Left$(S, I - 1)               'grab name
      colEntries.Add S, UCase$(S)       'add just name to collection
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
  
  If Ignore Then Exit Sub                   'exit if we are to ignore any of this for now
  With Me.lstTypeItems
    If .ListIndex = -1 Then Exit Sub        'if nothing selected
    S = .List(.ListIndex)                   'else grab data
  End With
'
' find definition
'
  Ignore = True                             'prevent endless loops
  Select Case Left$(S, 6)                   'set type of referencing
    Case "ByRef "
      I = 1
    Case "ByVal "
      I = 2
    Case Else
      I = 0
  End Select
  Me.cboReference.ListIndex = I             'set combobox for reference verb
  I = InStr(1, S, "(")                      'array dimensioning?
  If CBool(I) Then
    T = Mid$(S, I + 1)                      'grab array sizing
    I = InStr(1, T, ")")
    T = Left$(T, I - 1)                     'T contains dimensioning without parems
'
' do array definition checking
'
    Me.chkArray.Value = vbChecked           'array set
    If CBool(Len(T)) Then                   'dimensioning also set?
      Me.chkSize.Value = vbChecked          'yes, so also tag sizing
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
  Idx = InStr(1, S, " As ", vbTextCompare)  'find "As"
  If CBool(Idx) Then
    AsType = LTrim$(Mid$(S, Idx + 4))       'get definition
'
' also check for special string fixed sizing
'
    I = InStr(1, AsType, "*")               'find "*"
    If CBool(I) Then
      Me.txtSize.Text = Mid$(AsType, I + 2) 'found special sizing
      Me.chkSize.Value = vbChecked          'check sizing bheckbox
      AsType = Left$(AsType, I - 2)         'define type (String)
    End If
    
    Idx = FindMatch(Me.cboStdType, AsType)    'find type in standard list
    If Idx <> -1 Then
      Me.cboStdType.ListIndex = Idx           'found in standard, so display selection
    Else
      Idx = FindMatch(Me.cboTypeList, AsType) 'else find in TYPE list
      If Idx = -1 Then
        Me.cboStdType.ListIndex = 4           'default to Long if nothing found (very unlikely)
      Else
        Me.cboTypeList.ListIndex = Idx        'and make selection
      End If
    End If
  End If
  Ignore = False                              'allow active changes in combo options
  If Me.chkSize.Value = vbUnchecked Then
    Me.txtSize.Text = "0"                     'if sizing not set, then force 0 size
  End If
'
' now set options based upon position in the list
'
  Me.cmdUp.Enabled = False                    'init to no scrolling possible
  Me.cmdDown.Enabled = False
  With Me.lstTypeItems
    Me.cmdDelete.Enabled = CBool(.ListCount)  'delete possible if something exists
    If Me.cmdDelete.Enabled Then
      If .ListCount > 1 Then                  'if more than one, movement possible
        Me.cmdUp.Enabled = .ListIndex > 0
        Me.cmdDown.Enabled = .ListIndex < .ListCount - 1
      End If
    End If
  End With
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
' build definition of type
'
  S = "Declare " & Me.cboFnSub.Text & " " & Me.txtName & " LIB """ & Me.txtLIB & """"
  If CBool(Len(Me.txtAlias.Text)) Then
    S = S & " Alias """ & Me.txtAlias.Text & """ ("
  Else
    S = S & " ("
  End If
'
' build list of parameters
'
  With Me.lstTypeItems
    For Idx = 0 To .ListCount - 1
      S = S & .List(Idx) & ", "
    Next Idx
  End With
'
' strip final CR/LF add append ')'
'
  S = Left$(S, Len(S) - 2) & ")"
'
' if a return type declared and not a subroutine, then append return type
'
  If CBool(Len(ReturnType)) And Me.cboFnSub.Text <> "Sub" Then
    S = S & " As " & ReturnType
  End If
  DeclChange = S                                  'stuff new entry
  DeclName = Me.txtName.Text                      'new routine name
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
' Subroutine Name   : cboReference_KeyPress
' Purpose           : Check for ENTER
'*******************************************************************************
Private Sub cboReference_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 13                               'CR
    If Me.cmdAddNew.Enabled Then
      Call cmdAddNew_Click                'force add New button
      KeyAscii = 0
    End If
  End Select
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
' Subroutine Name   : cboStdType_KeyPress
' Purpose           : Check for ENTER
'*******************************************************************************
Private Sub cboStdType_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 13                               'CR
    If Me.cmdAddNew.Enabled Then
      Call cmdAddNew_Click                'force add New button
      KeyAscii = 0
    End If
  End Select
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
' Subroutine Name   : cboTypeList_KeyPress
' Purpose           : Check for ENTER
'*******************************************************************************
Private Sub cboTypeList_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 13                               'CR
    If Me.cmdAddNew.Enabled Then
      Call cmdAddNew_Click                'force add New button
      KeyAscii = 0
    End If
  End Select
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
    If CBool(Len(Me.txtNewEntry.Text)) Then Exit Sub
    S = .List(.ListIndex)                       'else get line
    Select Case Left$(S, 6)
      Case "ByRef ", "ByVal "
        S = Mid$(S, 7)
    End Select
    I = InStr(1, S, " ")                        'find space or (
    J = InStr(1, S, "(")
    If CBool(J) And J < I Then I = J
    S = Left$(S, I - 1)                         'get just name
    Select Case Me.cboReference.ListIndex
      Case 0
      Case 1
        S = "ByRef " & S
      Case 2
        S = "ByVal " & S
    End Select
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
  Call UpdateTemplate
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
        Case "A" To "Z", "_"
          KeyAscii = Asc(C)
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

'*******************************************************************************
' Subroutine Name   : cboFnSub_Click
' Purpose           : Choose Function or Sub
'*******************************************************************************
Private Sub cboFnSub_Click()
  Select Case Me.cboFnSub.ListIndex
    Case 0
      Me.cboReturnStd.Enabled = True
      Me.cboRtnType.Enabled = True
      If Me.cboReturnStd.ListIndex < 1 And Me.cboRtnType.ListIndex = -1 Then
        Me.cboReturnStd.ListIndex = 4
      End If
    Case 1
      Me.cboReturnStd.Enabled = False
      Me.cboRtnType.Enabled = False
      Me.cboReturnStd.ListIndex = 0
  End Select
'
' build definition of type
'
  Call UpdateTemplate
End Sub

'*******************************************************************************
' Subroutine Name   : cboReturnStd_Click
' Purpose           : Use Standard Return type
'*******************************************************************************
Private Sub cboReturnStd_Click()
  With Me.cboReturnStd
    If CBool(.ListIndex) Then
      ReturnType = .List(.ListIndex)
    Else
      ReturnType = vbNullString
    End If
  End With
'
' build definition of type
'
  Call UpdateTemplate
End Sub

'*******************************************************************************
' Subroutine Name   : cboRtnType_Click
' Purpose           : Use Special Return type
'*******************************************************************************
Private Sub cboRtnType_Click()
  With Me.cboRtnType
    ReturnType = .List(.ListIndex)
  End With
'
' build definition of type
'
  Call UpdateTemplate
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

