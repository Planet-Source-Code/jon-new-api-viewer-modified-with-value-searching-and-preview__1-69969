VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAPIViewer 
   Caption         =   "New API Viewer"
   ClientHeight    =   7440
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8250
   Icon            =   "frmAPIViewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   4920
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1320
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   27
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
   End
   Begin RichTextLib.RichTextBox rtbPreview 
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   3840
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1508
      _Version        =   393217
      BackColor       =   -2147483624
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAPIViewer.frx":0D42
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrResize 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2280
      Top             =   5640
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Create &New Entry..."
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
      Left            =   2580
      TabIndex        =   1
      ToolTipText     =   "Add a new entry of current type"
      Top             =   300
      Width           =   2235
   End
   Begin VB.Timer tmrAutoLoad 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1740
      Top             =   5640
   End
   Begin VB.Timer tmrMsg 
      Enabled         =   0   'False
      Interval        =   7000
      Left            =   1200
      Top             =   5640
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   26
      Top             =   7155
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14464
            MinWidth        =   14464
            Picture         =   "frmAPIViewer.frx":0DC2
            Key             =   "Msg"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrAlert 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   660
      Top             =   5640
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtSelectedItems 
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
      Height          =   915
      HideSelection   =   0   'False
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   5040
      Width           =   5535
   End
   Begin VB.ListBox LstItems 
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
      Height          =   2205
      ItemData        =   "frmAPIViewer.frx":0F1C
      Left            =   240
      List            =   "frmAPIViewer.frx":0F1E
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   5535
   End
   Begin VB.TextBox txtSrch 
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
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   5535
   End
   Begin VB.ComboBox cboAPIType 
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
      ItemData        =   "frmAPIViewer.frx":0F20
      Left            =   240
      List            =   "frmAPIViewer.frx":0F30
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   2235
   End
   Begin VB.PictureBox picAdd 
      BorderStyle     =   0  'None
      Height          =   5595
      Left            =   5940
      ScaleHeight     =   5595
      ScaleWidth      =   2115
      TabIndex        =   24
      Top             =   360
      Width           =   2115
      Begin VB.CheckBox chkAdd 
         Caption         =   "Add Constants && Types"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   2850
         Width           =   2055
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   1500
         Picture         =   "frmAPIViewer.frx":0F57
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Delete selected item from list"
         Top             =   1260
         Width           =   375
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "Insert into V&B Code"
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
         Left            =   0
         TabIndex        =   14
         ToolTipText     =   "Insert the selection into the current source file"
         Top             =   3960
         Width           =   1875
      End
      Begin VB.CommandButton cmdCvtDeclare 
         Caption         =   "&Modify Declare..."
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
         Left            =   0
         TabIndex        =   17
         ToolTipText     =   "Edit current entry and parameters"
         Top             =   5220
         Width           =   1875
      End
      Begin VB.CommandButton cmdDepends 
         Caption         =   "Chec&k Dependencies..."
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
         Left            =   0
         TabIndex        =   16
         ToolTipText     =   "List (still) contains unresolved Type entries"
         Top             =   4800
         Width           =   1875
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "&Copy to Clipboard"
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
         Left            =   0
         TabIndex        =   15
         ToolTipText     =   "Copy selection list to the clipboard"
         Top             =   4380
         Width           =   1875
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove Entry"
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
         Left            =   0
         TabIndex        =   12
         ToolTipText     =   "Remove entry the cursor is located within"
         Top             =   3120
         Width           =   1875
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "C&learAll Selections"
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
         Left            =   0
         TabIndex        =   13
         ToolTipText     =   "Erase all entries in the selection list"
         Top             =   3540
         Width           =   1875
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
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
         Left            =   0
         TabIndex        =   6
         ToolTipText     =   "Add the selected item to the list"
         Top             =   1260
         Width           =   1515
      End
      Begin VB.Frame frameScope 
         Caption         =   "Declare Scope"
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
         Left            =   0
         TabIndex        =   21
         Top             =   1680
         Width           =   1875
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   60
            ScaleHeight     =   495
            ScaleWidth      =   1755
            TabIndex        =   25
            Top             =   240
            Width           =   1755
            Begin VB.OptionButton optPub 
               Caption         =   "P&ublic"
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
               Left            =   0
               TabIndex        =   8
               Top             =   -60
               Value           =   -1  'True
               Width           =   915
            End
            Begin VB.OptionButton optPvt 
               Caption         =   "Priva&te"
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
               Left            =   0
               TabIndex        =   9
               Top             =   240
               Width           =   1035
            End
         End
      End
      Begin VB.CheckBox chkLong 
         Caption         =   "Make Constants Long"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   10
         ToolTipText     =   "Add ""As Long"" to constant with undeclared variable types"
         Top             =   2580
         Width           =   1875
      End
      Begin VB.Image Image1 
         Height          =   1080
         Left            =   420
         Picture         =   "frmAPIViewer.frx":1641
         Top             =   0
         Width           =   1080
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T&ype the first few letters of the word OR the constant value you require:"
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
      TabIndex        =   19
      ToolTipText     =   "Either type letters in, or type a value (eg &H200) in"
      Top             =   720
      Width           =   5280
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   840
      Picture         =   "frmAPIViewer.frx":322B
      Top             =   6360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgOn 
      Height          =   240
      Left            =   480
      Picture         =   "frmAPIViewer.frx":35B5
      Top             =   6360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgOff 
      Height          =   240
      Left            =   120
      Picture         =   "frmAPIViewer.frx":36FF
      Top             =   6360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblSelectedItems 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Selected Items: 0"
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
      TabIndex        =   22
      Top             =   4815
      Width           =   1260
   End
   Begin VB.Label lblClick 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " NOTE: Select text for more options."
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
      Left            =   2790
      TabIndex        =   23
      Top             =   4815
      Width           =   2985
   End
   Begin VB.Label lblAvailItems 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Available &Items: 0"
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
      TabIndex        =   20
      Top             =   1320
      Width           =   1290
   End
   Begin VB.Label lblAPIType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A&PI Type:"
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
      TabIndex        =   18
      Top             =   60
      Width           =   720
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileLoadText 
         Caption         =   "Load API &Text File..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileList 
         Caption         =   "FileList"
         Index           =   0
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save changes to current API file"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileResave 
         Caption         =   "&Resave API file in sorted order"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "&Add item to selection list"
      End
      Begin VB.Menu mnuEditDeleteItem 
         Caption         =   "&Delete item from list"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditAddNew 
         Caption         =   "Create New &Item of Selected Type..."
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuEditSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditRemove 
         Caption         =   "&Remove Entry from selection list"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "C&lear All Selections"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy Selections to Clipboard"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditDepends 
         Caption         =   "Chec&k Selection Dependencies"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "&Modify Declaration Definition"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewLine 
         Caption         =   "View Selections as &Line Items"
      End
      Begin VB.Menu mnuViewFull 
         Caption         =   "View &Full Text of Selections"
      End
      Begin VB.Menu mnuViewSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewLoad 
         Caption         =   "L&oad Last API File on Startup"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpUsing 
         Caption         =   "Using the New API Viewer..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About New API Viewer..."
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuURL 
         Caption         =   "Visit Sharp Dressed Codes..."
      End
   End
End
Attribute VB_Name = "frmAPIViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
' API Declarations
'-------------------------------------------------------------------------------
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" _
       (ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

Private Const EM_LINEFROMCHAR As Long = &HC9
Private Const EM_LINEINDEX As Long = &HBB
Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_LINESCROLL As Long = &HB6

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Const VK_LBUTTON = &H1
'-------------------------------------------------------------------------------
Private Typ As Integer          'current selected type (Const, Declare, Type)
Private LastX As Single         'mouse movement storage for LstIntems listbox
Private LastY As Single
Private Resizing As Boolean     'flag used to over-ride resizing
Private Saved As Boolean        'True when items have been saved to clipboard
Private PrevInst As Boolean     'True when previous instance found
Private Loading As Boolean      'True when for is loading

'flashing timout

Private Counter As Byte
Private Const COUNTER_TIMEOUT As Byte = 30

'
' NOTE, the ISADDIN variable is defined in the ADD-IN version of the producr in the Project Properties,
' on the Make tab, in the Conditional Compilation Properties. In the Add-In, simly set "ISADDED=1".
' For Stand-Alone, set "ISADDIN=0".
'
#If ISADDIN = 1 Then
  Public VBInstance As VBIDE.VBE
  Public Connect As Connect
#End If
'-------------------------------------------------------------------------------

Private Sub SetImage(ByVal boolOn As Boolean)
  
  Set StatusBar1.Panels(1).Picture = IIf(boolOn, imgOn.Picture, imgOff.Picture)
  
End Sub

Private Sub Command1_Click()
Dim I
Debug.Print Space$(50000)
For I = 1 To Extras.Count
Debug.Print "EXTRA: " & Extras(I)
Next
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Initialize
' Purpose           : Prepare app for XP-style, if available
'*******************************************************************************
Private Sub Form_Initialize()
  #If ISADDIN = 0 Then
    Call FormInitialize
  #End If
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Initialize main form
'*******************************************************************************
Private Sub Form_Load()
  Dim Path As String    'API file path
  Dim TS As TextStream  'File I/O text stream object
  Dim Ary() As String   'file line storage
  Dim Idx As Long, UB As Long, Cnt As Long, I As Long
  Dim S As String, T As String
  Dim AutoLoad As Boolean
'
' check for previous instance, and activate it
'
  PrevInst = CBool(App.PrevInstance)
  If PrevInst Then              'previous instance exists?
    If ActivatePrv(Me) Then     'activate previous instance if found
      #If ISADDIN = 1 Then
        Connect.Hide
      #End If
      Unload Me                 'unload current instance
      Exit Sub                  'leave and let previous instance assume control
    End If
  End If
  Me.mnuFileExit.Caption = "E&xit" & Chr$(9) & "Alt-F4"
  Loading = True
  
  #If ISADDIN = 0 Then
    Me.cmdInsert.Visible = False              'hide Insert button if not an addin
    Me.cmdCvtDeclare.Top = Me.cmdDepends.Top  'move buttons up
    Me.cmdDepends.Top = Me.cmdCopy.Top
    Me.cmdCopy.Top = Me.cmdInsert.Top
  #End If
  
  Load frmCom                     'load communications form
  Set Fso = New FileSystemObject  'create main file I/O object
  
  Set colAdded = New Collection   'create all collection objects
  Set colAddFL = New Collection
  Set colNew = New Collection
  Set colDepnd = New Collection
  Set colDelete = New Collection
  
''''-------------------
'''  Set colConst = New Collection
'''  Set colDecl = New Collection
'''  Set colType = New Collection
'''  Set ColEnum = New Collection
''''-------------------
'
' get recent file list and add to File menu
'
  Cnt = CLng(GetSetting$(App.Title, "Settings", "FileCnt", "0"))  'number of entries
  If CBool(Cnt) Then                                              'anything?
    I = 0                                                         'init loading index
    For Idx = 0 To Cnt - 1                                        'yes, so add to File menu
      S = GetSetting(App.Title, "Settings", "File" & CStr(Idx), vbNullString)
      If CBool(Len(Dir(S))) Then                                  'if api file exists
        If CBool(I) Then Load Me.mnuFileList(I)                   'if not index 0, load new storage entry
        With Me.mnuFileList(I)
          .Caption = S                                            'set menu caption
          .Visible = True                                         'always ensure it can be seen
        End With
        I = I + 1                                                 'bump index
      End If
    Next Idx
    If I <> Cnt Then                                              'if differences, save new list
      SaveSetting App.Title, "Settings", "FileCnt", CStr(I)       'save new count
      For Idx = 0 To I - 1                                        'save new list
        SaveSetting App.Title, "Settings", "File" & CStr(Idx), Me.mnuFileList(Idx).Caption
      Next Idx
    End If
  Else
    Me.mnuFileList(0).Visible = False                             'nothing, so hide default entry
    Me.mnuFileSep0.Visible = False
  End If
  
  chkAdd.Value = Int(GetSetting(App.Title, "Settings", "AddConstants", "0"))
'
' flag to convert constants to Long type
'
  Me.chkLong.Value = CInt(GetSetting(App.Title, "Settings", "LongConsts", "0"))
'
' flag to cause last file to be loaded
'
  Me.mnuViewLoad.Checked = CBool(GetSetting(App.Title, "Settings", "LoadLastFile", "1"))
'
' see if help is available
'
  Me.mnuHelpUsing.Enabled = Fso.FileExists(App.Path & "\UsingAPIVwr.htm")
'
' size form
'
  Resizing = True
  Me.Width = CLng(GetSetting(App.Title, "Settings", "FormWid", CStr(WinMinW)))
  Me.Height = CLng(GetSetting(App.Title, "Settings", "FormHit", CStr(WinMinH)))
  Idx = (Screen.Height - Me.Height) \ 2
  Me.Top = CLng(GetSetting(App.Title, "Settings", "FormTop", CStr(Idx)))
  Idx = (Screen.Width - Me.Width) \ 2
  Me.Left = CLng(GetSetting(App.Title, "Settings", "FormLft", CStr(Idx)))
  Me.WindowState = CInt(GetSetting(App.Title, "Settings", "WinState", CStr(vbNormal)))
  Resizing = False
  Call Form_Resize
'
' check for viewing line item or full definition in LstItem listbox
'
  Idx = CLng(GetSetting(App.Title, "Settings", "ViewStyle", "1"))
  If Idx = 1 Then
    Me.mnuViewLine.Checked = False
    Me.mnuViewFull.Checked = True
  Else
    Me.mnuViewLine.Checked = True
    Me.mnuViewFull.Checked = False
  End If
'
' set up API file path as last-loaded file, if option selected
'
  Path = GetSetting(App.Title, "Settings", "LastFile", vbNullString)
  AutoLoad = Not CBool(Len(Path))
  If Not Me.mnuViewLoad.Checked Then Path = vbNullString
  
'  Caption = "Loading API File, one moment..."
'  Counter = 0
'  tmrAlert.Enabled = True
'  Show
'  DoEvents
  
  
'
' load API file if it exists
'
  If CBool(Len(Path)) Then
    If Not OpenFile(Path) Then
      DeleteSetting App.Title, "Settings", "LastFile"
      Me.cboAPIType.Enabled = False 'if no file opened, then disable certain items
      Me.LstItems.Enabled = False
      Me.txtSrch.Enabled = False
    End If
  End If
'
' initialize display, pretty-up face
'
  
  Me.picAdd.BorderStyle = 0
  Me.cmdAdd.Enabled = False       'Add button
  Me.mnuEditAdd.Enabled = False
  Me.cmdDelete.Enabled = False    'Delete button
  Me.mnuEditDeleteItem.Enabled = False
  Me.cmdClear.Enabled = False     'Clear All Selections Item button
  Me.mnuEditClear.Enabled = False
  Me.cmdCopy.Enabled = False      'Copy to Clipboard button
  Me.mnuEditCopy.Enabled = False  'menu Copy option
  Me.cmdInsert.Enabled = False
  Me.cmdDepends.Enabled = False   'Check Dependencies button
  Me.mnuEditDepends.Enabled = False
  Me.cmdRemove.Enabled = False    'Remove Entry button
  Me.mnuEditRemove.Enabled = False
  Me.cmdCvtDeclare.Enabled = False 'Convert Declare button
  Me.mnuEditModify.Enabled = False
  Me.optPub.Enabled = False       'Public option
  Me.optPvt.Enabled = False       'Private option
  Me.frameScope.Enabled = False   'Option frame
  Me.chkLong.Enabled = False      'Convert Const to Long checkbox
  Me.chkAdd.Enabled = False
  Me.mnuFileSave.Enabled = False
'
' start the ball rolling by selecting the Declare types
'
  Me.cboAPIType.ListIndex = 1     'combo type list
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' hook window for sizing control
' Disable the following line if
' you will be debugging the form.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#If ISADDIN = 0 Then
  'Call HookWin(Me.hwnd, m_hWnd)
#End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' autostart load dialog if AutoLoad is True
'
  Loading = False
  Me.tmrAutoLoad.Enabled = AutoLoad
  
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Resize
' Purpose           : Resize form
'*******************************************************************************
Private Sub Form_Resize()
  Dim Bol As Boolean
  
  If Resizing Then Exit Sub                           'do nothing if Resizing flag set
  Select Case Me.WindowState
    Case vbMinimized                                  'do nothing special if minimized
      Exit Sub
    Case vbNormal
      
      If GetKeyState(VK_LBUTTON) < 0 Then             'if left mouse button down...
        
        With Me.tmrResize                             'let timer handle fix
          .Enabled = False                            'disable timer
          DoEvents                                    'let screen catch up
          .Enabled = True                             're-enable timer
        End With
        Exit Sub
      End If
      
      If Me.Width < WinMinW Or Me.Height < WinMinH Then 'if too small...
        With Me.tmrResize                             'smooth w/timer
          .Enabled = False                            'turn timer off
          DoEvents                                    'screen catch up
          .Enabled = True                             'restart timer
        End With
        Exit Sub                                      'let timer do work
      End If
    Case vbMaximized
      Resizing = True                                 'turn on resizing flag
      If Me.Width < WinMinW Then Me.Width = WinMinW   'do not go below minimum size
      If Me.Height < WinMinH Then Me.Height = WinMinH
  End Select
  
  Resizing = True                                     'turn on resizing flag
  With Me.picAdd                                      'place items on panel in proportion
    .Left = Me.ScaleWidth - .Width
  End With
  With Me.txtSrch
    .Width = Me.picAdd.Left - .Left - 120
    Me.txtSelectedItems.Width = .Width
    Me.LstItems.Width = .Width
    rtbPreview.Width = .Width
  End With
  With Me.lblClick
    .Left = Me.LstItems.Left + Me.LstItems.Width - .Width
  End With
  StatusBar1.Panels(1).Width = Me.ScaleWidth
  ''Me.txtSelectedItems.Height = Me.ScaleHeight - Me.StatusBar1.Height - Me.txtSelectedItems.Top - 30
  
  LstItems.Height = (Me.ScaleHeight - StatusBar1.Height) / 3 '4
  rtbPreview.Top = LstItems.Top + LstItems.Height + 30
  rtbPreview.Height = (Me.ScaleHeight - StatusBar1.Height - rtbPreview.Top) / 2
  lblSelectedItems.Top = rtbPreview.Top + rtbPreview.Height + 30
  lblClick.Top = lblSelectedItems.Top
  txtSelectedItems.Top = lblClick.Top + lblClick.Height + 50
  txtSelectedItems.Height = Me.ScaleHeight - StatusBar1.Height - txtSelectedItems.Top
  
  Resizing = False
'
' now save window state
'
  Call SaveSetting(App.Title, "Settings", "WinState", CStr(Me.WindowState))
'
' save form sizing only if normal
'
  If Me.WindowState = vbNormal Then
    Call SaveSetting(App.Title, "Settings", "FormWid", CStr(Me.Width))
    Call SaveSetting(App.Title, "Settings", "FormHit", CStr(Me.Height))
    Call SaveSetting(App.Title, "Settings", "FormTop", CStr(Me.Top))
    Call SaveSetting(App.Title, "Settings", "FormLft", CStr(Me.Left))
  End If
  
End Sub

Private Sub mnuURL_Click()

  Const URL As String = "http://sharpdressedcodes.com"
  
  If Not HyperJump(URL) Then
    MsgBox "Error launching web browser." & vbCrLf & _
           "Try launching it manually." & vbCrLf & _
           URL, _
           vbExclamation, App.Title
  End If

End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
  
  Me.tmrAlert.Enabled = False
  SetImage False
  StatusBar1.Panels(1).ToolTipText = vbNullString
  
End Sub

'*******************************************************************************
' Subroutine Name   : tmrResize_Timer
' Purpose           : Check Win for too small
'*******************************************************************************
Private Sub tmrResize_Timer()
  '
  ' Exit if Mouse pick button still down
  '
  If GetKeyState(VK_LBUTTON) < 0 Then Exit Sub
  '
  'turn off timer
  '
  Me.tmrResize.Enabled = False
  '
  'do nothing if minimized
  '
  If Me.WindowState = vbMinimized Then Exit Sub
  '
  'resize to minimum dims
  '
  Resizing = True     'block resize envent
  If Me.Width < WinMinW Then Me.Width = WinMinW
  If Me.Height < WinMinH Then Me.Height = WinMinH
  Resizing = False    'unblock resize event
  Call Form_Resize    'now process all resizing
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Unload main form
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  Dim S As String, T As String, Tname As String
  Dim Idx As Long
  Dim TS As TextStream
'
' do nothing if we are closing due to previous instance
'
  If PrevInst Then Exit Sub
'
' check to see if anything in selection buffer
'
  If CBool(Len(Me.txtSelectedItems.Text)) And Not Saved Then
    If CenterMsgBoxOnForm(Me, "Data buffer is not clear. Go ahead and continue?", vbYesNo Or vbQuestion, "Data Buffer Not Clear") = vbNo Then
      Cancel = 1
      Exit Sub
    End If
  End If
'
' get the file path and filename for the current API file
'
  T = GetSetting(App.Title, "Settings", "LastFile", vbNullString) 'get path to API file
  Idx = InStrRev(T, "\")
  Tname = Mid$(T, Idx + 1)
'
' check for deleted items
'
  If CBool(colDelete.Count) Then
    If colDelete.Count > 1 Then
      S = CStr(colDelete.Count) & " items have"
    Else
      S = "1 item has"
    End If
    If CenterMsgBoxOnForm(Me, S & " have been marked for deletion. Go ahead and 'Delete' by" & vbCrLf & _
                         "appending a non-inclusion tag at the end of the API file (" & Tname & ")?", _
                          vbYesNo Or vbQuestion, "Delete from API List") = vbYes Then
        S = "'" & vbCrLf & "'Deleted entries tagged: " & CStr(Now) & vbCrLf
      With colDelete
        Do While .Count
          S = S & "Delete " & .Item(1) & vbCrLf
          .Remove 1
        Loop
      End With
      Set TS = Fso.OpenTextFile(T, ForAppending, False)               'open for append
      TS.Write S                                                      'append new data
      TS.Close                                                        'close file
    End If
  End If
'
' check to see if new entries have been added
'
  If CBool(colNew.Count) Then
    If colNew.Count > 1 Then
      S = CStr(colNew.Count) & " new items have"
    Else
      S = "1 new item has"
    End If
    Select Case CenterMsgBoxOnForm(Me, S & " been added to the API list." & vbCrLf & _
                                       "Append new data to the current API file (" & Tname & ")?", _
                                       vbYesNoCancel Or vbQuestion, "New Data Found")
'
' build list of new entries
'
      Case vbYes
        S = "'" & vbCrLf & "'New entries added: " & CStr(Now) & vbCrLf
        With colNew
          Do While CBool(.Count)
            S = S & .Item(1) & vbCrLf
            .Remove 1
          Loop
        End With
        Set TS = Fso.OpenTextFile(T, ForAppending, False)               'open for append
        TS.Write S                                                      'append new data
        TS.Close                                                        'close file
'
' exit without saving
'
      Case vbNo
'
' cancel exiting
'
      Case vbCancel
        Cancel = 1
        Exit Sub
    End Select
  End If
'
' continue unloading
'
  Unload frmDepends       'ensure secondary forms are removed
  Unload frmModify
  Unload frmAddConst
  Unload frmAddEnum
  Unload frmAddType
  Unload frmCom
  
  Set colAdded = Nothing  'clear created objects
  Set colAddFL = Nothing
  Set colNew = Nothing
  Set colDepnd = Nothing
  Set colDelete = Nothing

''''-------------------
'''  Set colConst = Nothing
'''  Set colDecl = Nothing
'''  Set colType = Nothing
'''  Set ColEnum = Nothing
''''-------------------
'
' save window positioning of display is normal window
'
  If Me.WindowState = vbNormal Then
    Call SaveSetting(App.Title, "Settings", "FormTop", CStr(Me.Top))
    Call SaveSetting(App.Title, "Settings", "FormLft", CStr(Me.Left))
  End If
  
#If ISADDIN = 0 Then
  'Call UnhookWin(Me.hwnd, m_hWnd)
#End If

End Sub

'*******************************************************************************
' Subroutine Name   : SaveData
' Purpose           : Save new data to the API file
'*******************************************************************************
Private Sub SaveData()
  Dim S As String, T As String, Tname As String
  Dim Idx As Long
  Dim TS As TextStream
'
' get path to file to save
'
  T = GetSetting(App.Title, "Settings", "LastFile", vbNullString) 'get path to API file
  Idx = InStrRev(T, "\")
  Tname = Mid$(T, Idx + 1)
'
' build list of new entries
'
  S = "'" & vbCrLf & "'New entries added: " & CStr(Now) & vbCrLf
  With colNew
    Do While CBool(.Count)
      S = S & .Item(1) & vbCrLf
      .Remove 1
    Loop
  End With
'
' save file
'
  Set TS = Fso.OpenTextFile(T, ForAppending, False)               'open for append
  TS.Write S                                                      'append new data
  TS.Close                                                        'close file
  SetMsg "New data appended to " & Tname
  Me.mnuFileSave.Enabled = False
End Sub

'*******************************************************************************
' Subroutine Name   : cboAPIType_Click
' Purpose           : Make a selection from the item type combobox
'*******************************************************************************
Private Sub cboAPIType_Click()
  Dim S As String, Ary() As String
  Dim I As Long, Idx As Long
  
  If Not Me.cboAPIType.Enabled Then Exit Sub            'ignore if disabled
  If Typ = Me.cboAPIType.ListIndex Then Exit Sub        'already there
  Screen.MousePointer = vbHourglass
  DoEvents
  Typ = Me.cboAPIType.ListIndex                         'get type 0-2 (Const, Declare, Type)
  Me.chkLong.Enabled = (Typ = Constants)                'allow Long conversion on Constants
  chkAdd.Enabled = (Typ = Declares)
  Me.cmdAddNew.Enabled = True
  Me.mnuEditAddNew.Enabled = True
  Me.lblAvailItems.Caption = "Available &Items: 0"
  With Me.LstItems
    .Clear                                              'clear visible selection list
    Select Case Typ
      Case Constants                                    'constants
        For Idx = 0 To frmCom.lstConst.ListCount - 1    'grab from hidden full list of constants
          S = LTrim$(Mid$(frmCom.lstConst.List(Idx), 7)) 'get an item (Strip 'Const ')
          I = InStr(1, S, "'")                          'strip any comment
          If CBool(I) Then S = RTrim$(Left$(S, I - 1))
          .AddItem S                                    'add to visible list
        Next Idx
      Case Declares                                     'function/sub declarations
        For Idx = 0 To frmCom.lstDecl.ListCount - 1     'grab from full list of declarations
          S = LTrim$(Mid$(frmCom.lstDecl.List(Idx), 9)) 'get an item, strip 'Declare '
          I = InStr(1, S, " ")                          'Skip function or Sub
          S = LTrim$(Mid$(S, I + 1))
          I = InStr(1, S, " ")                          'grab just name
          .AddItem Left$(S, I - 1)                      'add just name
        Next Idx
      Case Types                                        'structures
        For Idx = 0 To frmCom.lstType.ListCount - 1     'grab from full list of structures
          S = LTrim$(Mid$(frmCom.lstType.List(Idx), 6)) 'grab an item (Strip 'Type ')
          I = InStr(1, S, vbCrLf)                       'find end of first line
          S = Left$(S, I - 1)
          I = InStr(1, S, "'")                          'strip any comments on it
          If CBool(I) Then S = RTrim$(Left$(S, I - 1))
          .AddItem S                                    'add to list
        Next Idx
      Case Enums
        For Idx = 0 To frmCom.lstEnum.ListCount - 1     'grab from full list of enumerators
          S = LTrim$(Mid$(frmCom.lstEnum.List(Idx), 6)) 'grab an item (Strip 'Enum ')
          I = InStr(1, S, vbCrLf)                       'find end of first line
          S = Left$(S, I - 1)
          I = InStr(1, S, "'")                          'strip any comments on it
          If CBool(I) Then S = RTrim$(Left$(S, I - 1))
          .AddItem S                                    'add to list
        Next Idx
    End Select
    .ListIndex = -1                                     'point to top with no selection
    Me.lblAvailItems.Caption = "Available &Items: " & CStr(.ListCount)
  End With
  Screen.MousePointer = vbDefault
End Sub

'*******************************************************************************
' Subroutine Name   : cmdAdd_Click
' Purpose           : Add selected item to the selection list
'*******************************************************************************
Private Sub cmdAdd_Click()
  Dim S As String, T As String
  Dim Idx As Integer
  
  With Me.LstItems
    Idx = .ListIndex
    S = .List(Idx)                              'grab item to process
  End With
  On Error Resume Next
  colAdded.Add S, S
  If CBool(Err.Number) Then Exit Sub            'if repeat, then do not bother with
  On Error GoTo 0
  
  Dim abc As String, Pos As Long
  
  Select Case Typ
    Case Constants                                          'is a constant?
      T = frmCom.lstConst.List(Idx)                         'yes, grab constant
      'If Not CBool(InStr(1, T, " As ")) Then                'check for 'As' declaration
        'Idx = InStr(1, T, "=")
        'T = RTrim$(Left$(T, Idx - 1)) & " As Long = " & LTrim$(Mid$(T, Idx + 1))
      'End If
      If InStr(LCase$(T), " as ") = 0 Then
        Pos = InStr(T, "= ")
        If Pos Then
          abc = Mid$(T, Pos + 2)
          Pos = InStr(abc, Space$(1))
          If Pos Then abc = Left$(abc, Pos - 1)
          Pos = InStr(abc, vbCrLf)
          If Pos Then abc = Left$(abc, Pos - 1)
        End If
        If LenB(abc) Then
          If (IsNumeric(abc)) Or (Right(abc, 1) = "&") Then
            T = Replace$(T, " = ", " As Long = ")
          ElseIf Left$(abc, 1) = Chr$(34) Then
            T = Replace$(T, " = ", " As String = ")
          End If
        End If
      End If
      
    Case Declares
      Idx = FindMatch(frmCom.lstDecl, "Declare Function " & S) 'check for function
      If Idx = -1 Then
        Idx = FindMatch(frmCom.lstDecl, "Declare Sub " & S)   'if not function, check sub
      End If
      T = frmCom.lstDecl.List(Idx)                          'get entry
    Case Types
      Idx = FindMatch(frmCom.lstType, "Type " & S)          'find user-defined type
      T = frmCom.lstType.List(Idx)                          'get entry
    Case Enums
      Idx = FindMatch(frmCom.lstEnum, "Enum " & S)          'find enumerator
      T = frmCom.lstEnum.List(Idx)                          'get entry
  End Select
  
  Idx = FindInsertIndex()                                   'find insertion index
  If Idx = -1 Then                                          'append flag?
    colAddFL.Add T, S                                       'yes, so add to full list at end of list
  Else
    With colAdded
      .Remove .Count                                        'not appending, so remove from end
      .Add S, S, Idx                                        'and insert as specified point
    End With
    colAddFL.Add T, S, Idx                                  'and insert full data before selection
  End If
  Call CheckStyle                                           'now update text display
End Sub

'*******************************************************************************
' Function Name     : FindInsertIndex
' Purpose           : Return Insert Before index in ColAddFL. If -1, then
'                   : simply append to collection
'*******************************************************************************
Private Function FindInsertIndex() As Integer
  Dim Ptr As Long, I As Long
  Dim Idx As Integer
  Dim Txt As String, S As String
  
  FindInsertIndex = -1                                          'init to append flag
  With Me.txtSelectedItems
    Ptr = Len(.Text)                                            'get length of text
    If CBool(Ptr) Then                                          'if data exists...
      If .SelStart = Ptr Then Exit Function                     'if end of file, append
      Ptr = SendMessageByNum(.hwnd, EM_LINEINDEX, -1&, 0&) + 1  'get start of current line
      Txt = .Text                                               'grab text
      If Mid$(Txt, Ptr, 1) = vbCr Then                          'CR?
        Ptr = Ptr + 2                                           'yes, so skip over CR/LF
        If Ptr >= Len(Txt) Then Exit Function
      Else
        Do
          Ptr = InStrRev(Txt, vbCrLf, Ptr)                      'find start of line
          If Ptr = 0 Then                                       'if at very start...
            Ptr = 1                                             'set to first character
            Exit Do
          Else
            Ptr = Ptr + 2                                       'else skip CR/LF
            I = InStr(Ptr, Txt, " ")
            If I = 0 Then Exit Function
            Select Case Mid(Txt, Ptr, I - Ptr)
              Case "Sub", "Public"
                Exit Do
            End Select
            Ptr = Ptr - 3
          End If
        Loop
      End If
      
      Txt = Mid$(Txt, Ptr)                                      'get data, starting with Public}Private
      If Me.mnuViewFull.Checked Then
        Ptr = InStr(1, Txt, " ") + 1                            'skip Public|Private
        Txt = Mid$(Txt, Ptr)
        Select Case Left$(Txt, 4)
          Case "Cons", "Decl"
            Ptr = InStr(1, Txt, vbCr) - 1                       'find terminating CR
            Txt = Left$(Txt, Ptr)                               'get declaration
          Case Else                                             'Handle Type and Enum
            Ptr = 1
            Do While Mid$(Txt, Ptr, 4) <> "End "
              Ptr = InStr(Ptr, Txt, vbCrLf) + 2
            Loop
            Ptr = InStr(Ptr, Txt, vbCr) - 1                     'find terminating CR
            Txt = Left$(Txt, Ptr)                               'get declaration
        End Select
        With colAddFL                                           'find entry in ADD FULL collection
          For Idx = 1 To .Count
            If .Item(Idx) = Txt Then Exit For                   'found a match
          Next Idx
          If Idx > .Count Then Idx = -1                         'this is a net; it should never happen
          FindInsertIndex = Idx                                 'set INSERT BEFORE index
        End With
      Else
        Ptr = InStr(1, Txt, vbCr) - 1                           'find terminating CR
        Txt = Left$(Txt, Ptr)                                   'get declaration
        With colAdded                                           'find entry in ADD FULL collection
          For Idx = 1 To .Count
            If .Item(Idx) = Txt Then Exit For                   'found a match
          Next Idx
          If Idx > .Count Then Idx = -1                         'this is a net; it should never happen
          FindInsertIndex = Idx                                 'set INSERT BEFORE index
        End With
      End If
    End If
  End With
End Function

'*******************************************************************************
' Subroutine Name   : LstItems_DblClick
' Purpose           : Add item to selected list when double-clicked
'*******************************************************************************
Private Sub LstItems_DblClick()
  Call cmdAdd_Click
End Sub

'*******************************************************************************
' Subroutine Name   : cmdRemove_Click
' Purpose           : Remove item from selected item list
'*******************************************************************************
Private Sub cmdRemove_Click()
  Dim S1 As Long, S2 As Long
  Dim S As String
  Dim Idx As Integer, I As Integer
  Dim Removed As Boolean
  
  Me.tmrAlert.Enabled = False                       'disable altert flag
  SetImage False
  
  Do
    With Me.txtSelectedItems
      S1 = InStrRev(.Text, vbCrLf, .SelStart + 1)   'find start of data
      S2 = InStr(.SelStart + 1, .Text, vbCrLf)      'get end of line
      If S2 = 0 Then Exit Sub                       'nothing to do
      If S1 = 0 Then
        S = Left$(.Text, S2 - 1)                    'grab text
        Exit Do
      End If
      S = Mid$(.Text, S1 + 2, S2 - S1 - 2)
      If Me.mnuViewLine.Checked Then Exit Do
      If Left$(S, 7) = "Public " Then Exit Do
      If Left$(S, 8) = "Private " Then Exit Do
      .SelStart = S1 - 2
    End With
  Loop
  
  Removed = False
  If Me.mnuViewLine.Checked Then                  'just line entry
    With colAdded
      For Idx = 1 To .Count
        If .Item(Idx) = S Then
          .Remove Idx                             'remove from all 3 storage collections
          colAddFL.Remove Idx
          Removed = True
          Exit For
        End If
      Next Idx
    End With
  Else                                            'doing full text
    S1 = InStr(1, S, " ")                         'strip public/private
    S = Mid$(S, S1 + 1)                           'grab data
    If Left$(S, 5) = "Type " Then S = S & vbCrLf  'build search text
    I = Len(S)                                    'length of data
    With colAddFL
      For Idx = 1 To .Count
        If Left$(.Item(Idx), I) = S Then          'found a match?
          .Remove Idx                             'yes, so remove from all 3 storage collections
          colAdded.Remove Idx
          Removed = True
          Exit For
        End If
      Next Idx
    End With
  End If
'
' done, so disable buttons for now and rebuild selected item list
'
  Me.cmdRemove.Enabled = False                    'disable remove command for now
  Me.mnuEditRemove.Enabled = False
  Me.cmdCvtDeclare.Enabled = False
  Me.mnuEditModify.Enabled = False
  MsgBeep beepSystemQuestion
  Call CheckStyle                                 'refresh selection list display
End Sub

'*******************************************************************************
' Subroutine Name   : cmdClear_Click
' Purpose           : Clear selected items list
'*******************************************************************************
Private Sub cmdClear_Click()
  With colAdded                                   'use master list's count
    Do While .Count
      .Remove 1                                   'remove from master
      colAddFL.Remove 1                           'and equal-sized counterparts
    Loop
  End With
  
  With colDepnd                                   'then clear dependency list, if it has data
    Do While .Count
      .Remove 1
    Loop
  End With
  Me.txtSelectedItems.Text = vbNullString
  Me.lblSelectedItems.Caption = "&Selected Items: 0"
  MsgBeep beepSystemQuestion
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCopy_Click
' Purpose           : Copy selected items to the clipboard
'*******************************************************************************
Private Sub cmdCopy_Click()
  Dim S As String
  
  S = GetSelection()                        'get selection to transmit
'
' now copy the data to the clipboard
'
  With Clipboard
    .Clear
    .SetText S, vbCFText
  End With
'
' report data copied
'
  SetMsg "All items in list copied to clipboard"
  MsgBeep beepSystemQuestion
  Saved = True                                'indicate data saved
End Sub

'*******************************************************************************
' Function Name     : GetSelection
' Purpose           : Build selection to send to user
'*******************************************************************************
Private Function GetSelection() As String
  Dim S As String, T As String
  Dim Idx As Long
  Dim I As Integer
  
  S = vbNullString                          'accumulator init
  With colAddFL                             'use the ID list to gather information
    For I = 1 To .Count
      T = .Item(I)                          'grab data
      '
      ' strip "As Long" from Constants if Long conversion is not set
      '
      If Left$(T, 5) = "Const" Then
        If Me.chkLong.Value = vbUnchecked Then
          Idx = InStr(1, T, " As Long")
          If CBool(Idx) Then
            T = Left$(T, Idx - 1) & Mid$(T, Idx + 8)
          End If
        End If
      End If
      If Me.optPub.Value Then               'add Public|Private parm and build accum
        S = S & "Public " & T & vbCrLf
      Else
        S = S & "Private " & T & vbCrLf
      End If
    Next I
  End With
  GetSelection = S                          'return result
End Function

'*******************************************************************************
' Subroutine Name   : mnuEditAdd_Click
' Purpose           : Add selection to list
'*******************************************************************************
Private Sub mnuEditAdd_Click()
  Call cmdAdd_Click
End Sub

'*******************************************************************************
' Subroutine Name   : mnuEditAddNew_Click
' Purpose           : Add a new entry of the current type
'*******************************************************************************
Private Sub mnuEditAddNew_Click()
  Call cmdAddNew_Click
End Sub

'*******************************************************************************
' Subroutine Name   : mnuEditClear_Click
' Purpose           : Clear all selections
'*******************************************************************************
Private Sub mnuEditClear_Click()
  Call cmdClear_Click
End Sub

'*******************************************************************************
' Subroutine Name   : mnuEditCopy_Click
' Purpose           : Invoke Copy button from menu entry
'*******************************************************************************
Private Sub mnuEditCopy_Click()
  Call cmdCopy_Click
End Sub

'*******************************************************************************
' Subroutine Name   : cmdDepends_Click
' Purpose           : Check Unresoved Dependencies
'*******************************************************************************
Private Sub cmdDepends_Click()
  Dim S As String, T As String
  Dim I As Long
  
  Me.tmrAlert.Enabled = False               'disable altert flag
  SetImage False
  If Not TestDepends() Then                 'if no dependencies found
    CenterMsgBoxOnForm Me, "No unresolved dependency Types found.", vbOKOnly Or vbInformation, "Information"
    Exit Sub
  End If
'
' invoke dependencies dialog
'
  frmDepends.Show vbModal, Me
'
' now add dependencies the user has selected
'
  With colDepnd
    If Not CBool(.Count) Then Exit Sub      'if nothing to do, then just leave
    Do While .Count                         'else grab selections
      S = .Item(1)                          'get a selection
      If Left$(S, 6) = "Const " Then
        I = FindMatch(frmCom.lstConst, S & " ") 'find match in Type list
        T = frmCom.lstConst.List(I)             'grab structure
        S = LTrim$(Mid$(S, 7))                  'strip "Const" header from name
      Else
        T = "Type " & S & vbCrLf              'set up for searching
        I = FindMatch(frmCom.lstType, T)      'find match in Type list
        T = frmCom.lstType.List(I)            'grab structure
      End If
      On Error Resume Next
      colAdded.Add S, S, 1                  'add to selection list
      If Not CBool(Err.Number) Then         'if it does not already exist
        colAddFL.Add T, S, 1                'add full reference
      End If
      On Error GoTo 0
      .Remove 1                             'remove processed selection
    Loop                                    'process all others
  End With
  Call CheckStyle                           'update selection list
  If Me.tmrAlert.Enabled Then
    SetMsg "MORE unresolved dependencies were found, maybe due to added types"
    MsgBeep beepSystemExclamation
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCvtDeclare_Click
' Purpose           : Modify entry
'*******************************************************************************
Private Sub cmdCvtDeclare_Click()
  With frmCom
    DeclChange = vbNullString             'init result buffer
    .txtSelectedItems.Text = Me.txtSelectedItems.Text
    .txtSelectedItems.SelStart = Me.txtSelectedItems.SelStart
    frmModify.Show vbModal, Me            'show modification form
    If CBool(Len(DeclChange)) Then        'changes?
      .lstDecl.AddItem DeclChange         'yes, so add new entry to declaration list
      On Error Resume Next
      colNew.Add DeclChange, DeclName     'add new entry to new list
      Call cmdRemove_Click                'remove old entry
      On Error GoTo 0
      Typ = -1                            'allow Declare selection to reinitialize
      Call cboAPIType_Click               're-select Declare option
      Me.LstItems.ListIndex = FindExactMatch(Me.LstItems, DeclName)  'select new entry
      Call cmdAdd_Click                   'add new entry
      Me.mnuFileSave.Enabled = True
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : LstItems_Click
' Purpose           : When an entry clicked, ensure certain properties are
'                   : made available to the user
'*******************************************************************************
Private Sub LstItems_Click()
  Dim S As String, T As String
  Dim Idx As Integer
  
  With Me.LstItems
    S = .List(.ListIndex)                                   'grab item to process
  End With
      
  If LenB(S) = 0 Then Exit Sub
  
'
' enable some options
'
  Me.optPub.Enabled = True
  Me.optPvt.Enabled = True
  Me.frameScope.Enabled = True
  Me.cmdAdd.Enabled = True
  Me.mnuEditAdd.Enabled = True
  Me.cmdDelete.Enabled = True
  Me.mnuEditDeleteItem.Enabled = True
  
  Select Case Typ
    Case Constants
      Call LstItems_MouseMove(0, 0, LastX, LastY)
      Idx = FindMatch(frmCom.lstConst, "Const " & S)
      If Idx > -1 Then T = frmCom.lstConst.List(Idx)
    Case Declares
      '
      ' special handling for Declarations. We will check for presence of "As Any" declarations
      '
      Idx = FindMatch(frmCom.lstDecl, "Declare Function " & S)  'check for function
      If Idx = -1 Then
        'Idx = FindMatch(frmCom.lstConst, "Declare Sub " & s)    'if not function, check sub
        Idx = FindMatch(frmCom.lstDecl, "Declare Sub " & S)    'if not function, check sub
      End If
      If Idx > -1 Then T = frmCom.lstDecl.List(Idx)                              'get entry
    Case Types
      Idx = FindMatch(frmCom.lstType, "Type " & S)
      If Idx > -1 Then T = frmCom.lstType.List(Idx)
    Case Enums
      Idx = FindMatch(frmCom.lstEnum, "Enum " & S)
      If Idx > -1 Then T = frmCom.lstEnum.List(Idx)
  End Select
  
  If Idx = -1 Then Exit Sub
  
  If CBool(InStr(1, T, " As Any")) Then
    Counter = 0
    tmrAlert.Enabled = True
    SetMsg "This Declaration contains at least one " & Chr$(34) & "As Any" & Chr$(34) & " clause"
  Else
    SetMsg vbNullString
  End If
  
  Debug.Print Space$(50000)
  ParseVBKeyWords rtbPreview, T, CBool(chkLong.Value), CBool(chkAdd.Value)

End Sub

'*******************************************************************************
' Subroutine Name   : LstItems_MouseMove
' Purpose           : Offer additional help on Const types
'                   : (display embedded comment if present)
'*******************************************************************************
Private Sub LstItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Idx As Integer
  Dim S As String
  
  LastX = X                                           'save current X/Y location
  LastY = Y
  If CBool(Typ) Then Exit Sub                         'if not Const, then ignore
  'Idx = CInt(ListItemByCoordinate(Me.LstItems, X, Y)) 'else grab data
  Idx = IsMouseOverItem(LstItems)
  If Idx > -1 Then
    S = frmCom.lstConst.List(Idx)                       'grab constant data (1 for 1)
    Idx = InStr(1, S, "'")                              'find a comment
    If CBool(Idx) Then                                  'found one?
      S = LTrim$(Mid$(S, Idx + 1))                      'yes, so grab it
      Counter = COUNTER_TIMEOUT / 2
      tmrAlert.Enabled = True
    Else
      S = vbNullString                                  'else no comment
    End If
    SetMsg S, IIf(LenB(S), True, False)                 'send to status bar
    If LenB(S) = 0 Then Counter = COUNTER_TIMEOUT
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuEditDepends_Click
' Purpose           : Check dependencies
'*******************************************************************************
Private Sub mnuEditDepends_Click()
  Call cmdDepends_Click
End Sub

'*******************************************************************************
' Subroutine Name   : mnuEditModify_Click
' Purpose           : Modify entry
'*******************************************************************************
Private Sub mnuEditModify_Click()
  Call cmdCvtDeclare_Click
End Sub

'*******************************************************************************
' Subroutine Name   : mnuEditRemove_Click
' Purpose           : Remove selecterd entry
'*******************************************************************************
Private Sub mnuEditRemove_Click()
  Call cmdRemove_Click
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileExit_Click
' Purpose           : Exit application
'*******************************************************************************
Private Sub mnuFileExit_Click()
  #If ISADDIN = 1 Then
    Connect.Hide
  #End If
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileList_Click
' Purpose           : Select a previously opened itext file
'*******************************************************************************
Private Sub mnuFileList_Click(Index As Integer)
  Dim Idx As Integer, Cnt As Integer
  Dim S As String
  
  If Not OpenFile(Me.mnuFileList(Index).Caption) Then
    Cnt = CLng(GetSetting$(App.Title, "Settings", "FileCnt", "0"))  'number of entries
    For Idx = Index + 1 To Cnt - 1                                  'yes, so add to File menu
      Me.mnuFileList(Idx - 1).Caption = Me.mnuFileList(Idx).Caption
    Next Idx
    Me.mnuFileList(Idx - 1).Visible = False
    Unload Me.mnuFileList(Idx - 1)
  End If
  Call cboAPIType_Click
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileLoadText_Click
' Purpose           : Select a file to load
'*******************************************************************************
Private Sub mnuFileLoadText_Click()
  Dim Path As String
  Dim Idx As Long
  
  Path = Environ$("MsDevDir")                   'grab MS development directory, if there is one
  If CBool(Len(Path)) Then
    Idx = InStrRev(Path, "\")                   'strip Dev directory
    Path = Left$(Path, Idx) & "Tools\Winapi"    'and add WinAPI folder in Tools directory
    If CBool(Len(Dir$(Path, vbDirectory))) Then 'if it exists
      Path = Path & "\*.txt"                    'allow all files
    Else
      Path = vbNullString
    End If
  End If
  If Not CBool(Len(Path)) Then Path = App.Path & "\*.txt"  'if nothing found, use current folder
'
' open dialog box
'
  With Me.CommonDialog1
    .DialogTitle = "Open API Text File"         'title for dialog box
    .DefaultExt = ".txt"                        'set default extension
    .FileName = Path                            'set target
    .Filter = "Text Files|*.txt"                'define filter and flags
    .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNLongNames Or cdlOFNHideReadOnly
    .CancelError = True                         'allow user to cancel out of dialog
    On Error Resume Next
    .ShowOpen                                   'show OPEN dialog box
    If CBool(Err.Number) Then Exit Sub          'user hit Cancel (no actual error)
    On Error GoTo 0
    Path = .FileName                            'else get selecton
  End With
  OpenFile Path                                 'load file
  Call cboAPIType_Click
End Sub

'*******************************************************************************
' Function Name     : OpenFile
' Purpose           : Open a selected API text file
'*******************************************************************************
Private Function OpenFile(Path As String) As Boolean
  Dim TS As TextStream
  Dim Ary() As String, S As String, T As String, U As String
  Dim Idx As Long, UB As Long, Idy As Long, I As Long, J As Long
  
  T = GetSetting(App.Title, "Settings", "LastFile", vbNullString) 'get path to API file
  If T = Path And Not Loading Then
    OpenFile = True
    Exit Function
  End If
  
  Call SaveOld                                          'save any updated data
  If Fso.FileExists(Path) Then
    Counter = 0
    tmrAlert.Enabled = True
    Screen.MousePointer = vbHourglass                   'show busy
    DoEvents
    SetMsg "Loading " & Path & "...", True                    'report loading
''''-------------------
''''
'''' clean out old collections
''''
'''    With colConst
'''      Do While .Count
'''        .Remove 1
'''      Loop
'''    End With
'''    With colDecl
'''      Do While .Count
'''        .Remove 1
'''      Loop
'''    End With
'''    With colType
'''      Do While .Count
'''        .Remove 1
'''      Loop
'''    End With
'''    With ColEnum
'''      Do While .Count
'''        .Remove 1
'''      Loop
'''    End With
''''-------------------
'
' make a backup file if one does not exist
'
    Idx = InStrRev(Path, ".")
    T = Left$(Path, Idx) & "bak"
    If Not Fso.FileExists(T) Then                       'does it exist?
      Fso.CopyFile Path, T                              'no, so make a backup copy
    End If
    Set TS = Fso.OpenTextFile(Path, ForReading, False)  'open file
    Ary = Split(TS.ReadAll, vbCrLf)                     'load file
    TS.Close                                            'close file
  Else
    CenterMsgBoxOnForm Me, "Cannot find selected file:" & vbCrLf & _
                            Path, vbOKOnly Or vbExclamation, "File Not Found"
    SetMsg vbNullString
    Exit Function
  End If
 
  If IsDimmed(Ary) Then                                 'if array has data
    UB = UBound(Ary)                                    'get upper bounds of array
    Idx = 0                                             'init index
    frmCom.lstConst.Clear                               'init list buffers
    frmCom.lstDecl.Clear
    frmCom.lstType.Clear
    frmCom.lstEnum.Clear
    frmCom.lstHex.Clear
    ReDim APIConstants(0) As APIConstantType
    Unload frmAddConst                                  'unload form for fresh constant data
    
    Me.LstItems.Clear
    
    Dim FF As Integer, abc1 As String, abc2 As String, Pos As Long
    
    'FF = FreeFile
    'Open App.Path & "\fuckyouman.txt" For Binary Access Write As #FF
     
    Do While Idx <= UB                                  'while we are not beyond the data limit
      S = Trim$(Ary(Idx))                               'grab a line
      If CBool(Len(S)) Then                             'if data present...
        If Left$(S, 1) <> "'" Then                      'and not a comment...
          Select Case Left$(S, 4)                       'check type
            
            Case "Cons"                                 'Const
              Idy = InStr(1, S, "=")                    'find '=' and add spaces
              S = RTrim$(Left$(S, Idy - 1)) & " = " & LTrim$(Mid$(S, Idy + 1))
              frmCom.lstConst.AddItem S                 'then add to list ***
              
              abc1 = RTrim$(Left$(S, Idy - 1))
              abc2 = LTrim$(Mid$(S, Idy + 1))
                
              'tidy up the shit
              Pos = InStrRev(abc1, Space$(1))
              If Pos > 0 Then abc1 = Mid$(abc1, Pos + 1)

              Pos = InStr(abc2, "'")
              If Pos > 0 Then abc2 = Left$(abc2, Pos - 1)
                
              Pos = InStr(abc2, "(")
              If Pos > 0 Then Pos = InStr(Pos, abc2, ")")
                
              If Pos > 0 Then
                abc2 = Left$(abc2, Pos)
              Else
                Pos = InStr(abc2, Space$(1))
                If Pos > 0 Then abc2 = Left$(abc2, Pos - 1)
              End If
                
              'Put #FF, LOF(FF) + 1, abc1 & "=" & abc2 & vbCrLf
              
              Dim OldIndex As Long
              
              OldIndex = FindExactMatch(frmCom.lstHex, abc2)
              
              If OldIndex > -1 Then
                APIConstants(OldIndex).OtherNames.Add abc1 ', abc2
              Else
                frmCom.lstHex.AddItem abc2
                If LenB(APIConstants(UBound(APIConstants)).Name) Then ReDim Preserve APIConstants(UBound(APIConstants) + 1) As APIConstantType
                With APIConstants(UBound(APIConstants))
                  .Index = frmCom.lstHex.ListCount - 1
                  .Name = abc1
                  .Value = abc2
                End With
              End If
              
''''-------------------
'''              T = LTrim$(Mid$(S, 7))                    'skip Const
'''              I = InStr(1, T, " ")                      'find a space following name
'''              J = InStr(1, T, "=")                      'find = following name
'''              If J < I Then I = J                       'use lowest index
'''              T = RTrim$(Left$(T, I - 1))               'get just constant name
'''              On Error Resume Next
'''              colConst.Add T, UCase$(T)                 'add name only
'''              If Not CBool(Err.Number) Then
'''                frmCom.lstConst.AddItem S               'then add to list
'''              End If
'''              On Error GoTo 0
''''-------------------
              
            Case "Decl"                                 'Declare
              frmCom.lstDecl.AddItem S                  'add declaration ***
''''-------------------
'''              I = InStr(1, S, " ")
'''              T = LTrim$(Mid$(S, I + 1))                'skip declare
'''              I = InStr(1, T, " ")
'''              T = LTrim$(Mid$(T, I + 1))                'skip Function|Sub
'''              I = InStr(1, T, " ")
'''              T = Left$(T, I - 1)                       'get just name
'''              On Error Resume Next
'''              colDecl.Add T, UCase$(T)                  'add name only
'''              If Not CBool(Err.Number) Then
'''                frmCom.lstDecl.AddItem S                'add declaration
'''              End If
'''              On Error GoTo 0
''''-------------------
            
            Case "Type"                                 'Type
              I = InStr(1, S, "'")
              If CBool(I) Then
                S = RTrim$(Left$(S, I - 1))             'strip comment from header
              End If
              T = S & vbCrLf                            'init accumulator
              Idx = Idx + 1                             'now scan for End Type
              Do While Trim$(Ary(Idx)) <> "End Type"    'find definition
                U = Trim$(Ary(Idx))                     'grab a line
                If CBool(Len(U)) Then                   'if data present...
                  If Left$(U, 1) <> "'" Then            'and not a comment
                    T = T & "  " & U & vbCrLf           'add line with 2 leading spaces
                  End If
                End If
                Idx = Idx + 1                           'bump to next line
              Loop                                      'grab all definition lines for Type
              frmCom.lstType.AddItem T & Trim$(Ary(Idx)) 'then add all data plus End Type line ***
              'If LCase$(S) = "type context" Then
              'Debug.Print T & Trim$(Ary(Idx))
''''-------------------
'''              I = InStr(1, T, " ")
'''              S = LTrim$(Mid$(T, I + 1))                'skip Type
'''              I = InStr(1, S, vbCr)
'''              S = RTrim$(Left$(S, I - 1))               'get just type name
'''              On Error Resume Next
'''              colType.Add S, UCase$(S)                  'add name only
'''              If Not CBool(Err.Number) Then
'''                frmCom.lstType.AddItem T & Trim$(Ary(Idx)) 'then add all data plus End Type line
'''              End If
'''              On Error GoTo 0
''''-------------------
            
            Case "Enum"                                 'Enum
              T = S & vbCrLf                            'init accumulator
              Idx = Idx + 1                             'now scan for End Type
              Do While Trim$(Ary(Idx)) <> "End Enum"    'find definition
                U = Trim$(Ary(Idx))                     'grab a line
                If CBool(Len(U)) Then                   'if data present...
                  If Left$(U, 1) <> "'" Then            'and not a comment
                    T = T & "  " & U & vbCrLf           'add line with 2 leading spaces
                  End If
                End If
                Idx = Idx + 1                           'bump to next line
              Loop                                      'grab all definition lines for Type
              frmCom.lstEnum.AddItem T & Trim$(Ary(Idx)) 'then add all data plus End Enum line ***
''''-------------------
'''              I = InStr(1, T, " ")
'''              S = LTrim$(Mid$(T, I + 1))                'skip Type
'''              I = InStr(1, S, vbCr)
'''              S = RTrim$(Left$(S, I - 1))               'get just type name
'''              On Error Resume Next
'''              ColEnum.Add S, UCase$(S)                  'add name only
'''              If Not CBool(Err.Number) Then
'''                frmCom.lstEnum.AddItem T & Trim$(Ary(Idx)) 'then add all data plus End Enum line
'''              End If
'''              On Error GoTo 0
''''-------------------
            
            Case "Dele"                                 'Delete
              S = Mid$(S, 8)                            'strip "Delete "
              Select Case Left$(S, 4)                   'now find which list to remove from
                Case "Cons"
                  Idy = FindMatch(frmCom.lstConst, S)
                  If Idy <> -1 Then
                    frmCom.lstConst.RemoveItem Idy
''''-------------------
'''                    T = LTrim$(Mid$(S, 7))                    'skip Const
'''                    I = InStr(1, T, " ")                      'find a space following name
'''                    J = InStr(1, T, "=")                      'find = following name
'''                    If J < I Then I = J                       'use lowest index
'''                    T = RTrim$(Left$(T, I - 1))               'get just constant name
'''                    On Error Resume Next
'''                    colConst.Remove UCase$(T)                 'add name only
'''                    On Error GoTo 0
''''-------------------
                  End If
                
                Case "Decl"
                  Idy = FindMatch(frmCom.lstDecl, S)
                  If Idy <> -1 Then
                    frmCom.lstDecl.RemoveItem Idy
''''-------------------
'''                    I = InStr(1, S, " ")
'''                    T = LTrim$(Mid$(S, I + 1))                'skip declare
'''                    I = InStr(1, T, " ")
'''                    T = LTrim$(Mid$(T, I + 1))                'skip Function|Sub
'''                    I = InStr(1, T, " ")
'''                    T = Left$(T, I - 1)                       'get just name
'''                    On Error Resume Next
'''                    colDecl.Remove UCase$(T)
'''                    On Error GoTo 0
''''-------------------
                  End If
                
                Case "Type"
                  Idy = FindMatch(frmCom.lstType, S)
                  If Idy <> -1 Then
                    frmCom.lstType.RemoveItem Idy
''''-------------------
'''                    I = InStr(1, T, " ")
'''                    S = LTrim$(Mid$(T, I + 1))                'skip Type
'''                    I = InStr(1, S, vbCr)
'''                    S = RTrim$(Left$(S, I - 1))               'get just type name
'''                    On Error Resume Next
'''                    colType.Remove UCase$(S)
'''                    On Error GoTo 0
''''-------------------
                  End If
                
                Case "Enum"
                  Idy = FindMatch(frmCom.lstEnum, S)
                  If Idy <> -1 Then
                    frmCom.lstEnum.RemoveItem Idy
''''-------------------
'''                    I = InStr(1, T, " ")
'''                    S = LTrim$(Mid$(T, I + 1))                   'skip Type
'''                    I = InStr(1, S, vbCr)
'''                    S = RTrim$(Left$(S, I - 1))                  'get just type name
'''                    On Error Resume Next
'''                    ColEnum.Remove UCase$(S)
'''                    If Not CBool(Err.Number) Then
'''                      frmCom.lstEnum.AddItem T & Trim$(Ary(Idx)) 'then add all data plus End Enum line
'''                    End If
'''                    On Error GoTo 0
''''-------------------
                  End If
              End Select
          End Select
        End If
      End If
      Idx = Idx + 1                                     'bump to next line
      
      Counter = 0
      DoEvents
      
    Loop
    
    'Close #FF
    Counter = COUNTER_TIMEOUT
'
' all loaded, so add to main form caption
'
    Me.Caption = "New API Viewer - " & Path
    SetMsg "Successfully loaded " & Idx + 1 & " item" & IIf(Idx + 1 = 1, vbNullString, "s") & " from " & Mid$(Path, InStrRev(Path, "\") + 1), True
  Else
    Me.Caption = "New API Viewer"                       'nothing loaded
    Me.cboAPIType.Enabled = False
    Me.LstItems.Enabled = False
    Me.txtSrch.Enabled = False
  End If
    
  If frmCom.lstConst.ListCount = 0 And _
     frmCom.lstDecl.ListCount = 0 And _
     frmCom.lstType.ListCount = 0 And _
     Me.cboAPIType.Enabled Then
    CenterMsgBoxOnForm Me, "The selected file is not an API code file. Not accepting it.", _
                           vbOKOnly Or vbExclamation, "Invalid API Code File"
    Me.Caption = "New API Viewer"                       'nothing loaded
    Me.cboAPIType.Enabled = False
    Me.LstItems.Enabled = False
    Me.txtSrch.Enabled = False
    Me.cmdAddNew.Enabled = False
  Else
    SaveSetting App.Title, "Settings", "LastFile", Path           'save valid path
    UB = GetSetting(App.Title, "Settings", "FileCnt", "0")        'get count of items
    For Idx = 0 To UB - 1                                         'see if already in list
      If StrComp(Me.mnuFileList(Idx).Caption, Path) = 0 Then Exit For 'found match
    Next Idx
    If Idx = UB Then                                              'did not find?
      SaveSetting App.Title, "Settings", "FileCnt", CStr(UB + 1)  'no, so save new count
      SaveSetting App.Title, "Settings", "File" & CStr(UB), Path  'save new entry
      If CBool(UB) Then Load Me.mnuFileList(UB)                   'add menu entry if not offset 0
      With Me.mnuFileList(UB)
        .Caption = Path                                           'add path to it
        .Visible = True                                           'ensure we can see it
        Me.mnuFileSep0.Visible = True                             'ensure separator visible
      End With
    End If
    Me.LstItems.Enabled = True
    Me.txtSrch.Enabled = True
    Me.cboAPIType.Enabled = True
  End If
'
' init buttons and such for new file
'
  Me.mnuFileResave.Enabled = Me.cboAPIType.Enabled
  Me.cmdAdd.Enabled = False
  Me.mnuEditAdd.Enabled = False
  Me.cmdDelete.Enabled = False
  Me.mnuEditDeleteItem.Enabled = False
  Me.cmdDepends.Enabled = False
  Me.mnuEditDepends.Enabled = False
  Me.chkLong.Enabled = False
  Me.chkAdd.Enabled = False
  
  If Me.cboAPIType.Enabled Then
    Call cboAPIType_Click
  End If
  
  Screen.MousePointer = vbDefault
  Me.lblAvailItems.Caption = "Available &Items: 0"
  'SetMsg vbNullString
  Typ = -1
  On Error Resume Next
  Me.cboAPIType.SetFocus
  Me.mnuFileSave.Enabled = CBool(colNew.Count) And Me.cboAPIType.Enabled
  OpenFile = Me.cboAPIType.Enabled
End Function

'*******************************************************************************
' Subroutine Name   : SaveOld
' Purpose           : Save changes when changing API files
'*******************************************************************************
Private Sub SaveOld()
  Dim S As String, T As String, Tname As String
  Dim Idx As Long
  Dim TS As TextStream
'
' get the file path and filename for the current API file
'
  T = GetSetting(App.Title, "Settings", "LastFile", vbNullString) 'get path to API file
  Idx = InStrRev(T, "\")
  Tname = Mid$(T, Idx + 1)
'
' check for deleted items
'
  If CBool(colDelete.Count) Then
    If colDelete.Count > 1 Then
      S = CStr(colDelete.Count) & " items have"
    Else
      S = "1 item has"
    End If
    If CenterMsgBoxOnForm(Me, S & " have been marked for deletion. Go ahead and 'Delete' by" & vbCrLf & _
                         "appending a non-inclusion tag at the end of the API file (" & Tname & ")?", _
                          vbYesNo Or vbQuestion, "Delete from API List") = vbYes Then
        S = "'" & vbCrLf & "'Deleted entries tagged: " & CStr(Now) & vbCrLf
      With colDelete
        Do While .Count
          S = S & "Delete " & .Item(1) & vbCrLf
          .Remove 1
        Loop
      End With
      Set TS = Fso.OpenTextFile(T, ForAppending, False)               'open for append
      TS.Write S                                                      'append new data
      TS.Close                                                        'close file
    Else
      With colDelete
        Do While .Count
          .Remove 1
        Loop
      End With
    End If
  End If
'
' check to see if new entries have been added
'
  If CBool(colNew.Count) Then
    If colNew.Count > 1 Then
      S = CStr(colNew.Count) & " new items have"
    Else
      S = "1 new item has"
    End If
    Select Case CenterMsgBoxOnForm(Me, S & " been added to the API list." & vbCrLf & _
                                       "Append new data to the current API file (" & Tname & ")?", _
                                       vbYesNo Or vbQuestion, "New Data Found")
'
' build list of new entries
'
      Case vbYes
        S = "'" & vbCrLf & "'New entries added: " & CStr(Now) & vbCrLf
        With colNew
          Do While CBool(.Count)
            S = S & .Item(1) & vbCrLf
            .Remove 1
          Loop
        End With
        Set TS = Fso.OpenTextFile(T, ForAppending, False)               'open for append
        TS.Write S                                                      'append new data
        TS.Close                                                        'close file
'
' flush New data without saving
'
      Case vbNo
        With colNew
          Do While CBool(.Count)
            .Remove 1
          Loop
        End With
    End Select
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileResave_Click
' Purpose           : Resave API file in sorted order
'*******************************************************************************
Private Sub mnuFileResave_Click()
  Dim S As String, Path As String
  Dim TS As TextStream
  Dim Idx As Long
  
  Path = GetSetting(App.Title, "Settings", "LastFile", vbNullString)
  If CBool(Len(Path)) Then
    Set TS = Fso.OpenTextFile(Path, ForWriting, True)
    With frmCom
      With .lstConst
        For Idx = 0 To .ListCount - 1
          TS.WriteLine .List(Idx)
        Next Idx
      End With
      
      With .lstDecl
        For Idx = 0 To .ListCount - 1
          TS.WriteLine .List(Idx)
        Next Idx
      End With
      
      With .lstType
        For Idx = 0 To .ListCount - 1
          TS.WriteLine .List(Idx)
        Next Idx
      End With
      
      With .lstEnum
        For Idx = 0 To .ListCount - 1
          TS.WriteLine .List(Idx)
        Next Idx
      End With
    End With
    TS.Close
    CenterMsgBoxOnForm Me, "Updated: " & Path, vbOKOnly Or vbInformation, "Update Complete"
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : mnuFileSave_Click
' Purpose           : Save new data to the API file
'*******************************************************************************
Private Sub mnuFileSave_Click()
  Call SaveData
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHelpAbout_Click
' Purpose           : Show informationa bout API Viewer
'*******************************************************************************
Private Sub mnuHelpAbout_Click()
  CenterMsgBoxOnForm Me, "New API Viewer " & GetAppVersion & vbCrLf & _
    String$(110, "-") & vbCrLf & _
    "A utility for viewing Win32 API Constants, Declarations, User-Defined Types, and Enums." & vbCrLf & _
    "Selected items can be copied from the viewer to your VB application via the Clipboard." & vbCrLf & vbCrLf & _
    "Copyright 2007 © by David Ross Goben." & vbCrLf & vbCrLf & _
    "Copyright 2008 © modified by Sharp Dressed Codes." & vbCrLf & vbCrLf & _
    "Based upon visual functionality of the API Viewer copyright 1998 © Microsoft Corporation." & vbCrLf & vbCrLf & _
    "Special thanks to Dan Appleman of Desaware, Inc. for his much expanded API32.TXT file.", _
    vbOKOnly Or vbInformation, "About New API Viewer"
End Sub

'*******************************************************************************
' Subroutine Name   : mnuHelpUsing_Click
' Purpose           : View help file
'*******************************************************************************
Private Sub mnuHelpUsing_Click()
  OpenFilePath Me.hwnd, App.Path & "\UsingAPIVwr.htm"
End Sub

'*******************************************************************************
' Subroutine Name   : mnuViewFull_Click
' Purpose           : View fiull defintions in selected item list
'*******************************************************************************
Private Sub mnuViewFull_Click()
  If Me.mnuViewFull.Checked Then Exit Sub
  Me.mnuViewFull.Checked = True
  Me.mnuViewLine.Checked = False
  SaveSetting App.Title, "Settings", "ViewStyle", "1"
  Call CheckStyle
End Sub

'*******************************************************************************
' Subroutine Name   : mnuViewLine_Click
' Purpose           : View simple definition in selected item list
'*******************************************************************************
Private Sub mnuViewLine_Click()
  If Me.mnuViewLine.Checked Then Exit Sub
  Me.mnuViewLine.Checked = True
  Me.mnuViewFull.Checked = False
  SaveSetting App.Title, "Settings", "ViewStyle", "0"
  Call CheckStyle
End Sub

'*******************************************************************************
' Subroutine Name   : CheckStyle
' Purpose           : Support style selection
'*******************************************************************************
Private Sub CheckStyle()
  Dim S As String, T As String, Ary() As String
  Dim Idx As Long, I As Long, J As Long
  
  S = vbNullString                                          'init accumulator
  If Me.mnuViewLine.Checked Then                            'view line items?
    With colAdded                                           'yes, so add just main entries
      For Idx = 1 To .Count
        S = S & .Item(Idx) & vbCrLf                         'add a line
      Next Idx
    End With
  Else
    With colAddFL                                           'else add full entries
      For Idx = 1 To .Count
        T = .Item(Idx)
        '
        ' strip "As Long" from Constants if Long conversion is not set
        ' strip comments from Constants, Types, and Enums
        '
        Select Case Left$(T, 4)
          Case "Cons"
            I = InStr(1, T, "'")                            'strip any comment
            If CBool(I) Then T = RTrim$(Left$(T, I - 1))
            If Me.chkLong.Value = vbUnchecked Then
              I = InStr(1, T, " As Long")
              If CBool(I) Then
                T = Left$(T, I - 1) & Mid$(T, I + 8)
              End If
            End If
          
          Case "Type", "Enum"
            Ary = Split(T, vbCrLf)                          'bream Enum/Type into array
            For J = 0 To UBound(Ary)                        'scan each line
              T = Ary(J)                                    'grab a line
              I = InStr(1, T, "'")                          'contains a comment?
              If CBool(I) Then
                Ary(J) = RTrim$(Left$(T, I - 1))            'yes, so strip it
              End If
            Next J
            T = Join(Ary, vbCrLf)                           'rebuild string
        End Select
        '
        ' apply Private|Public declaration
        '
        If Me.optPub.Value Then                             'add data with Private/Public header
          S = S & "Public " & T & vbCrLf & vbCrLf
        Else
          S = S & "Private " & T & vbCrLf & vbCrLf
        End If
      Next Idx
      If CBool(Len(S)) Then S = Left$(S, Len(S) - 2)        'strip final trailing CR/LF
    End With
  End If
'
' display selections is textbox
'
  With Me.txtSelectedItems
    .Text = S               'add text
    .SelStart = Len(.Text)  'point to end
    Saved = False           'indicate not saved
  End With
  Me.lblSelectedItems.Caption = "&Selected Items: " & CStr(colAdded.Count)
End Sub

'*******************************************************************************
' Subroutine Name   : mnuViewLoad_Click
' Purpose           : Toggle check on Load Last File option
'*******************************************************************************
Private Sub mnuViewLoad_Click()
  Dim Bol As Boolean
  
  Bol = Not Me.mnuViewLoad.Checked
  Me.mnuViewLoad.Checked = Bol
  SaveSetting App.Title, "Settings", "LoadLastFile", CStr(Bol)
End Sub

'*******************************************************************************
' Subroutine Name   : optPub_Click
' Purpose           : set state on Public
'*******************************************************************************
Private Sub optPub_Click()
  If CBool(Len(Me.txtSelectedItems.Text)) Then
    Call CheckStyle
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : optPvt_Click
' Purpose           : set state on Private
'*******************************************************************************
Private Sub optPvt_Click()
  Call optPub_Click
End Sub

'*******************************************************************************
' Subroutine Name   : tmrAlert_Timer
' Purpose           : Toggle alert button image
'*******************************************************************************
Private Sub tmrAlert_Timer()
  
  If Counter >= COUNTER_TIMEOUT Then
    Counter = 0
    tmrAlert.Enabled = False
    SetImage False
    StatusBar1.Panels(1).ToolTipText = vbNullString
    Exit Sub
  End If
  
  StatusBar1.Panels(1).ToolTipText = "Click me to stop flashing"
  
  If tmrAlert.Tag = vbNullString Then tmrAlert.Tag = False
  tmrAlert.Tag = IIf(tmrAlert.Tag = True, False, True)
  SetImage tmrAlert.Tag
  
  Counter = Counter + 1
  
End Sub

'*******************************************************************************
' Subroutine Name   : tmrAutoLoad_Timer
' Purpose           : Auto-Invoke Open Dialog
'*******************************************************************************
Private Sub tmrAutoLoad_Timer()
  Me.tmrAutoLoad.Enabled = False
  Call mnuFileLoadText_Click
End Sub

'*******************************************************************************
' Subroutine Name   : tmrMsg_Timer
' Purpose           : Erase msg after a moment
'*******************************************************************************
Private Sub tmrMsg_Timer()
  Me.tmrMsg.Enabled = False
  Me.StatusBar1.Panels("Msg").Text = vbNullString
End Sub

'*******************************************************************************
' Subroutine Name   : txtSelectedItems_Change
' Purpose           : When text for selected items chage, check enablements
'*******************************************************************************
Private Sub txtSelectedItems_Change()
  Dim Cnt As Integer, NewCnt As Integer
  Dim Bol As Boolean
  
  Me.cmdClear.Enabled = CBool(Len(Me.txtSelectedItems.Text))  'enable Clear key as needed
  Me.mnuEditClear.Enabled = Me.cmdClear.Enabled
  Me.cmdCopy.Enabled = Me.cmdClear.Enabled
  Me.mnuEditCopy.Enabled = Me.cmdClear.Enabled
  Me.cmdInsert.Enabled = Me.cmdClear.Enabled
'
' if clear key enabled, then check for unresolved dependencies generated by selections
'
  If Me.cmdClear.Enabled Then
    Bol = Me.cmdDepends.Enabled           'save button state
    Cnt = colDepnd.Count                  'save old dependency count
    Me.cmdDepends.Enabled = TestDepends() 'check for more
    Me.mnuEditDepends.Enabled = Me.cmdDepends.Enabled
    If Cnt <> colDepnd.Count Or Not Bol And CBool(colDepnd.Count) Then
      Counter = 0
      Me.tmrAlert.Enabled = True          'we will need to alert user to new unresolved dependencies
      SetMsg "Click flashing dot to turn if off if you know about this issue"
      MsgBeep beepSystemExclamation
    End If
  Else
    Me.tmrAlert.Enabled = False           'else ensure dependency alert and button disabled
    SetImage False
    Me.cmdDepends.Enabled = False
    Me.mnuEditDepends.Enabled = False
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : txtSelectedItems_Click
' Purpose           : When the user clicks into the list
'*******************************************************************************
Private Sub txtSelectedItems_Click()
  Dim S1 As Long, S2 As Long
  Dim Bol As Boolean
  Dim S As String
  
  With Me.txtSelectedItems
    S1 = InStrRev(.Text, vbCrLf, .SelStart + 1)   'find leading CR/LF
    S2 = InStr(.SelStart + 1, .Text, vbCrLf)      'find following CR/LF
    If S1 = 0 Then S1 = 1                         'if at start, then set 1
    Bol = CBool(S2 - S1 - 2)                      'set bol to True if data found
    If Bol Then
      Bol = (.SelStart <> Len(.Text))             'check for being at end of text
    End If
    Me.cmdRemove.Enabled = Bol                    'enable button as needed
    Me.mnuEditRemove.Enabled = Bol
    If Bol Then                                   'if something there
      S = Mid$(.Text, S1, S2 - S1)                'get text there
      If Left$(S, 2) = vbCrLf Then S = Mid$(S, 3) 'skip over leaing CR/LF
      If Left$(S, 8) = "Private " Then            'skip header
        S = Mid$(S, 9)
      ElseIf Left$(S, 7) = "Public " Then
        S = Mid$(S, 8)
      End If
      If Left$(S, 8) <> "Declare " Then Bol = False 'Declare statement?
    End If
    Me.cmdCvtDeclare.Enabled = Bol                'enable Declare editing as required
    Me.mnuEditModify.Enabled = Bol
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : txtSrch_Change
' Purpose           : Something typed on search box
'*******************************************************************************
Private Sub txtSrch_Change()
  
  Dim I As Long, Idx As Long, S As String
  
  S = Trim$(Me.txtSrch.Text)                              'grab typeed data
  
  If CBool(Len(S)) Then
    Idx = FindMatch(Me.LstItems, S)                       'find a match
    If Idx <> -1 Then
      Me.LstItems.ListIndex = Idx                         'select entry in main list
    Else
      If S = "&" Then Exit Sub
      Idx = FindExactMatch(frmCom.lstHex, S)
      If Idx > -1 Then
        With APIConstants(Idx)
          If LCase$(.Value) <> LCase$(S) Then Exit Sub
          rtbPreview.TextRTF = vbNullString
          rtbPreview.SelColor = vbBlack
          rtbPreview.SelText = .Name
          If .OtherNames.Count > 0 Then
            For I = 1 To .OtherNames.Count
              rtbPreview.SelText = vbCrLf & .OtherNames.Item(I)
            Next
          End If
        End With
      End If
    End If
  End If
  
End Sub

'*******************************************************************************
' Subroutine Name   : txtSrch_GotFocus
' Purpose           : Select entire entry when user selects search box
'*******************************************************************************
Private Sub txtSrch_GotFocus()
  With Me.txtSrch
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : chkLong_Click
' Purpose           : User chose to check/uncheck convert const to long option
'*******************************************************************************
Private Sub chkLong_Click()
  SaveSetting App.Title, "Settings", "LongConsts", CStr(Me.chkLong.Value)
  If CBool(Len(Me.txtSelectedItems.Text)) Then
    Call CheckStyle
  End If
End Sub

Private Sub chkAdd_Click()
  
  SaveSetting App.Title, "Settings", "AddConstants", CStr(chkAdd.Value)
  If Not Loading Then LstItems_Click
  
End Sub

'*******************************************************************************
' Function Name     : TestDepends
' Purpose           : Test for unresoved dependencies
'*******************************************************************************
Private Function TestDepends() As Boolean
  Dim Idx As Integer, Idy As Integer
  Dim I As Long, J As Long, K As Long, L As Long, M As Long
  Dim S As String, T As String, U As String, Ary() As String
  
  With colDepnd
    Do While .Count
      .Remove 1
    Loop
  End With
  
  With colAddFL                               'check full entries
    For Idx = 1 To .Count
      S = .Item(Idx)                          'get an entry
      Select Case Left$(S, 4)                 'check type
        
        Case "Cons"
          I = InStr(1, S, "=")                'get data after "="
          S = LTrim$(Mid$(S, I + 1))
          I = InStr(1, S, "'")                'strip any comment
          If CBool(I) Then S = RTrim$(Left$(S, I - 1))
          For I = 1 To Len(S)                 'strip special characters from string
            Select Case UCase$(Mid$(S, I, 1))
              Case "A" To "Z", "_", "&", "0" To "9"  'valid character list
              Case Else
                Mid$(S, I, 1) = " "           'convert all others to spaces
            End Select
          Next I
          Ary = Split(S, " ")                 'now build an array from data
          
          For I = 0 To UBound(Ary)
            S = Ary(I)                              'grab an item
            If CBool(Len(S)) Then                   'if there is data...
              If Not IsNumeric(S) Then              'and is not numeric?
                
                Select Case UCase$(S)               'ignore binary operators
                  Case "AND", "OR", "NOT", "XOR"
                    S = vbNullString
                End Select
                If CBool(Len(S)) Then               'if we still have data...
                  U = "Const " & S & " "            'init for Const entry
                  J = FindMatch(frmCom.lstConst, U) 'search for match
                  If J <> -1 Then
                    K = Len(U)                      'found match, so set length of search data
                    With colAddFL
                      For Idy = 1 To .Count         'see if defined in list
                        If Left$(.Item(Idy), K) = U Then Exit For 'found a match
                      Next Idy
                      If Idy > .Count Then          'did not find match
                        U = RTrim$(U)               'so strip space at end
                        On Error Resume Next
                        colDepnd.Add U, U           'add to dependency list
                        On Error GoTo 0
                      End If
                    End With
                  End If
                End If
              End If
            End If
          Next I
          S = vbNullString                          'then set ingmore flag for furtern testing
        
        Case "Decl"
        
        Case "Type"
          Ary = Split(S, vbCrLf)                    'break type declaration up
          For K = 1 To UBound(Ary) - 1              'check data between heaer and End Type
            T = Ary(K)                              'get a line
            I = InStr(1, T, "'")                    'strip comment
            If CBool(I) Then T = RTrim$(Left$(T, I - 1))
            I = InStrRev(T, "*")                    'check for fixed-length specifier
            If CBool(I) Then
              U = Trim$(Mid$(T, I + 1))
              If Not IsNumeric(U) Then
                U = "Const " & U & " "
                T = RTrim$(Left$(T, I - 1))
                J = FindMatch(frmCom.lstConst, U)     'search for match
                If J <> -1 Then
                  K = Len(U)                          'found match, so set length of search data
                  With colAddFL
                    For Idy = 1 To .Count             'see if defined in list
                      If Left$(.Item(Idy), K) = U Then Exit For 'found a match
                    Next Idy
                    If Idy > .Count Then              'did not find match
                      U = RTrim$(U)                   'so strip space at end
                      On Error Resume Next
                      colDepnd.Add U, U               'add to dependency list
                      On Error GoTo 0
                    End If
                  End With
                End If
              End If
            End If
            I = InStr(1, T, "(")                    'see if array size set
            If CBool(I) Then                        'found '("
              J = InStr(I + 1, T, ")")              'find ")"
              If CBool(J) Then                      'found it?
                U = Mid$(T, I + 1, J - I - 1)       'yes, get data between parens
                If Not IsNumeric(U) Then            'not numeric, so check for constant
                  U = "Const " & U & " "
                  J = FindMatch(frmCom.lstConst, U) 'search for match
                  If J <> -1 Then
                    K = Len(U)                      'found match, so set length of search data
                    With colAddFL
                      For Idy = 1 To .Count         'see if defined in list
                        If Left$(.Item(Idy), K) = U Then Exit For 'found a match
                      Next Idy
                      If Idy > .Count Then          'did not find match
                        U = RTrim$(U)               'so strip space at end
                        On Error Resume Next
                        colDepnd.Add U, U           'add to dependency list
                        On Error GoTo 0
                      End If
                    End With
                  End If
                End If
              End If
            End If
          Next K
          
        Case "Enum"
          S = vbNullString                          'ignore enumerator
      End Select
      
      If CBool(Len(S)) Then                   'if not constant
        I = InStr(1, S, " As ", vbTextCompare) 'contains an "AS" clause?
        Do While CBool(I)                     'process all matches
          J = InStr(I + 4, S, " ")            'find next space
          K = InStr(I + 4, S, ",")            'find next comma
          L = InStr(I + 4, S, ")")            'find next paren
          M = InStr(I + 4, S, vbCrLf)         'find next CR/LF
          If J = 0 Then J = K                 'set J to a minimum match
          If J = 0 Then J = L
          If J = 0 Then J = M
          If J > I And K > I Then             'ensure J set to lowest match
            If K < J Then J = K
          End If
          If J > I And L > I Then
            If L < J Then J = L
          End If
          If J > I And M > I Then
            If M < J Then J = M
          End If
          If J > I Then                       'If J greater than index
            T = Mid$(S, I + 4, J - I - 4)     'point after " AS "
            U = "Type " & T & vbCrLf          'find Type entry
            J = FindMatch(frmCom.lstType, U)  'search for match
            If J <> -1 Then
              K = Len(U)                      'found 1, set length of search data
              With colAddFL
                For Idy = 1 To .Count         'see if defined in list
                  If Left$(.Item(Idy), K) = U Then Exit For 'found a match
                Next Idy
                If Idy > .Count Then          'did not find match
                  On Error Resume Next
                  colDepnd.Add T, T           'add to dependency list
                  On Error GoTo 0
                End If
              End With
            End If
          End If
          I = InStr(I + 4, S, " As ")         'then scan for next match
        Loop
      End If
    Next Idx
  End With
  TestDepends = CBool(colDepnd.Count)         'set True if any unresolved dependencies found
End Function

'*******************************************************************************
' Subroutine Name   : SetMsg
' Purpose           : Display a message on the status line
'*******************************************************************************
Private Sub SetMsg(Text As String, Optional ByVal FixedText As Boolean)
  Me.tmrMsg.Enabled = False
  With Me.StatusBar1.Panels("Msg")
    If CBool(Len(Text)) Then
      Text = UCase$(Left$(Text, 1)) & Mid$(Text, 2)
    End If
    If .Text <> Text Then .Text = Text
  End With
  If Not FixedText Then Me.tmrMsg.Enabled = CBool(Len(Text))
End Sub

'*******************************************************************************
' Subroutine Name   : cmdInsert_Click
' Purpose           : Insert selection into VB code
'*******************************************************************************
Private Sub cmdInsert_Click()
  #If ISADDIN = 1 Then
    Dim StL As Long, StC As Long, EnL As Long, EnC As Long
    Dim S As String, SelectData As String
    Dim W As Window
    
    SelectData = GetSelection                           'get data
'
' first copy the data to the clipboard
'
    With Clipboard
      .Clear
      .SetText S, vbCFText
    End With
'
' now strip terminating data from end (CR/LF)
'
    SelectData = Left$(SelectData, Len(SelectData) - 2)
'
' now find the currently active code window to insert the data into
'
    For Each W In VBInstance.Windows
      If W.Visible = True And W.Type = vbext_wt_CodeWindow And W.Caption = VBInstance.ActiveCodePane.Window.Caption Then
        With VBInstance.ActiveCodePane
          .GetSelection StL, StC, EnL, EnC              'get current cursor position
          With .CodeModule
            If CBool(Len(SelectData)) Then              'if data to insert
              .InsertLines StL, SelectData              'insert data
              SelectData = vbNullString                 'erase local copy
            End If
          End With
        End With
        Exit For                                        'no need to continue looping
      End If
    Next W
    Saved = True                                        'indicate data saved
    Connect.Hide
  #End If
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdAddNew_Click
' Purpose           : Add a new entry of the current type
'*******************************************************************************
Private Sub cmdAddNew_Click()
  Dim Idx As Long
  Dim S As String
  
  Select Case Typ
    
    Case Constants
      DeclChange = vbNullString             'init result buffer
      frmAddConst.Show vbModal, Me          'show add constant form
      If CBool(Len(DeclChange)) Then        'changes?
        frmCom.lstConst.AddItem DeclChange  'yes, so add new entry to declaration list
        On Error Resume Next
'''        colConst.Add DeclName, UCase$(DeclName)
        colNew.Add DeclChange, DeclName     'add new entry to new list
        On Error GoTo 0
        Typ = -1                            'allow Declare selection to reinitialize
        Call cboAPIType_Click               're-select Declare option
        Idx = InStr(1, DeclChange, "'")
        If CBool(Idx) Then
          S = RTrim$(Left$(DeclChange, Idx - 1))
        Else
          S = DeclChange
        End If
        Me.LstItems.ListIndex = FindExactMatch(Me.LstItems, Mid$(S, 7)) 'select new entry
        Me.mnuFileSave.Enabled = True
      End If
    
    Case Declares
      DeclChange = vbNullString             'init result buffer
      frmAddDecl.Show vbModal, Me           'add new Declaration
      If CBool(Len(DeclChange)) Then        'changes?
        frmCom.lstDecl.AddItem DeclChange   'yes, so add new entry to declaration list
        On Error Resume Next
'''        colDecl.Add DeclName, UCase$(DeclName)
        colNew.Add DeclChange, DeclName     'add new entry to new list
        On Error GoTo 0
        Typ = -1                            'allow Declare selection to reinitialize
        Call cboAPIType_Click               're-select Declare option
        Me.LstItems.ListIndex = FindExactMatch(Me.LstItems, DeclName) 'select new entry
        Me.mnuFileSave.Enabled = True
      End If
    
    Case Types
      DeclChange = vbNullString             'init result buffer
      frmAddType.Show vbModal, Me           'add new Type
      If CBool(Len(DeclChange)) Then        'changes?
        frmCom.lstType.AddItem DeclChange   'yes, so add new entry to declaration list
        On Error Resume Next
'''        colType.Add DeclName, UCase$(DeclName)
        colNew.Add DeclChange, DeclName     'add new entry to new list
        On Error GoTo 0
        Typ = -1                            'allow Declare selection to reinitialize
        Call cboAPIType_Click               're-select Declare option
        Me.LstItems.ListIndex = FindExactMatch(Me.LstItems, DeclName) 'select new entry
        Me.mnuFileSave.Enabled = True
      End If
    
    Case Enums
      DeclChange = vbNullString             'init result buffer
      frmAddEnum.Show vbModal, Me           'add new enumerator
      If CBool(Len(DeclChange)) Then        'changes?
        frmCom.lstEnum.AddItem DeclChange   'yes, so add new entry to declaration list
        On Error Resume Next
'''        ColEnum.Add DeclName, UCase$(DeclName)
        colNew.Add DeclChange, DeclName     'add new entry to new list
        On Error GoTo 0
        Typ = -1                            'allow Declare selection to reinitialize
        Call cboAPIType_Click               're-select Declare option
        Me.LstItems.ListIndex = FindExactMatch(Me.LstItems, DeclName) 'select new entry
        Me.mnuFileSave.Enabled = True
      End If
  End Select
End Sub

'*******************************************************************************
' Subroutine Name   : cmdDelete_Click
' Purpose           : Remove item from list and place is a 'delete' pool
'*******************************************************************************
Private Sub cmdDelete_Click()
  Dim S As String
  Dim Idx As Integer
  
  With Me.LstItems
    Idx = .ListIndex
    S = .List(Idx)                                        'grab item to process
  End With
  If InStr(1, S, "=") Then S = RTrim$(Left$(S, InStr(1, S, "=") - 1))
  
  If CenterMsgBoxOnForm(Me, "Verify deletion of '" & S & "'." & vbCrLf & _
                            "(Actually, just append a Deleted" & vbCrLf & _
                            "Tag at the end of the API file)", _
                             vbYesNo Or vbQuestion Or vbDefaultButton2, _
                            "Verify Delete") = vbNo Then Exit Sub
  Select Case Typ
    Case Constants
      S = frmCom.lstConst.List(Idx)
      frmCom.lstConst.RemoveItem Idx
    Case Declares
      Idx = FindMatch(frmCom.lstDecl, "Declare Function " & S) 'check for function
      If Idx = -1 Then
        Idx = FindMatch(frmCom.lstConst, "Declare Sub " & S)   'if not function, check sub
      End If
      S = frmCom.lstDecl.List(Idx)                          'get entry
      frmCom.lstDecl.RemoveItem Idx
    Case Types
      Idx = FindMatch(frmCom.lstType, "Type " & S)          'find user-defined type
      S = frmCom.lstType.List(Idx)                          'get entry
      frmCom.lstType.RemoveItem Idx
    Case Enums
      Idx = FindMatch(frmCom.lstEnum, "Enum " & S)          'find enumerator
      S = frmCom.lstEnum.List(Idx)                          'get entry
      frmCom.lstEnum.RemoveItem Idx
  End Select
'
' now determine if we should modify save list
' check for current delete being in created list
' (we created something we now want to delete)
'
  With colNew
    For Idx = 1 To .Count
      If .Item(Idx) = S Then Exit For                       'the delete item matches the created item?
    Next Idx
    If Idx > .Count Then                                    'if no match was found
      colDelete.Add S                                       'then simly add item to the delete list
    Else
     .Remove Idx                                            'else remove from created list
     'enable save option based upon something still to save
     Me.mnuFileSave.Enabled = CBool(.Count + colDelete.Count)
    End If
  End With
  
  With Me.LstItems
    Idx = .ListIndex                                      'get index to target
    .RemoveItem Idx                                       'remove item
    If Idx = .ListCount Then Idx = .ListCount - 1         'adjust index
    .ListIndex = Idx                                      'adjust selection to drop on next or last
  End With
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

