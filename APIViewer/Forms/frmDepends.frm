VERSION 5.00
Begin VB.Form frmDepends 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Check Selected Item Type Dependencies Not Yet Added to Current Select List"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   7200
      TabIndex        =   7
      ToolTipText     =   "Do nothing and leave this dialog"
      Top             =   4410
      Width           =   1395
   End
   Begin VB.CommandButton cmdOK 
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
      Left            =   5490
      TabIndex        =   6
      ToolTipText     =   "Add only checked items"
      Top             =   4410
      Width           =   1395
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select &All"
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
      Left            =   3780
      TabIndex        =   5
      ToolTipText     =   "Select (check) all items in the list"
      Top             =   4410
      Width           =   1395
   End
   Begin VB.TextBox Text1 
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
      Height          =   3975
      Left            =   3780
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   4875
   End
   Begin VB.ListBox List1 
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
      Height          =   4110
      ItemData        =   "frmDepends.frx":0000
      Left            =   180
      List            =   "frmDepends.frx":0002
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   360
      Width           =   3435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: Check items you want to add"
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
      TabIndex        =   4
      Top             =   4560
      Width           =   2955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Item Layout:"
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
      Left            =   3780
      TabIndex        =   2
      Top             =   120
      Width           =   1590
   End
   Begin VB.Label lblList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List of Items depended on by Main List:"
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
      Width           =   2820
   End
End
Attribute VB_Name = "frmDepends"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Initialize form
'*******************************************************************************
Private Sub Form_Load()
  Me.Icon = frmCom.Icon
  With colDepnd
    Do While .Count
      Me.List1.AddItem .Item(1) 'add list of unresolved dependencies
      .Remove 1
    Loop
  End With
  Me.List1.ListIndex = -1       'point to top of list-1
  Me.cmdOK.Enabled = False      'disable Apply button to start
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCancel_Click
' Purpose           : User chose to cancel adding additional entries
'*******************************************************************************
Private Sub cmdCancel_Click()
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdOK_Click
' Purpose           : User chose to apply selected entries for inclusion
'*******************************************************************************
Private Sub cmdOK_Click()
  Dim Idx As Integer
  
  With colDepnd
    For Idx = 0 To Me.List1.ListCount - 1 'add only selected items to dependency list
      If Me.List1.Selected(Idx) Then
        .Add Me.List1.List(Idx)
      End If
    Next Idx
  End With
  Unload Me                               'then return to caller
End Sub

'*******************************************************************************
' Subroutine Name   : cmdSelectAll_Click
' Purpose           : Select all entries for inclusion
'*******************************************************************************
Private Sub cmdSelectAll_Click()
  Dim Idx As Integer
  
  With Me.List1
    For Idx = 0 To .ListCount - 1
      .Selected(Idx) = True                   'mark all items in list as selected
    Next Idx
    .ListIndex = -1                           'do not choose any item
  End With
  Me.Text1.Text = vbNullString                'erase any formatting text
  Me.cmdOK.Enabled = CBool(Me.List1.SelCount) 'enable apply button if anything in list
End Sub

'*******************************************************************************
' Subroutine Name   : List1_Click
' Purpose           : User selected an entry of unserolved items
'*******************************************************************************
Private Sub List1_Click()
  Dim Idx As Long
  Dim S As String
  
  With Me.List1
    Me.cmdOK.Enabled = CBool(.SelCount)             'enable Apply if something selected
    S = .List(.ListIndex)                           'get text of item
  End With
  If Left$(S, 6) = "Const " Then
    Idx = FindMatch(frmCom.lstConst, S & " ")
    If Idx <> -1 Then                               'current item has data?
      Me.Text1.Text = frmCom.lstConst.List(Idx)     'stuff textbox with contents if so
      Me.Text1.SelStart = 0                         'point to top
    End If
  Else
    Idx = FindMatch(frmCom.lstType, "Type " & S & vbCrLf)
    If Idx <> -1 Then                               'current item has data?
      Me.Text1.Text = frmCom.lstType.List(Idx)      'stuff textbox with contents if so
      Me.Text1.SelStart = 0                         'point to top
    End If
  End If
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

