VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6648
   ClientLeft      =   192
   ClientTop       =   864
   ClientWidth     =   10260
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6648
   ScaleWidth      =   10260
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RTBDetails 
      Height          =   4455
      Left            =   7080
      TabIndex        =   28
      Top             =   1440
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5313
      _ExtentY        =   7853
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   3
      RightMargin     =   2.50000e5
      TextRTF         =   $"frmMain.frx":57E2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   3960
      ScaleHeight     =   2508
      ScaleWidth      =   3228
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CheckBox chkAutoFind 
         Caption         =   "Auto find as you type"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   3015
      End
      Begin VB.CheckBox chkMatch 
         Caption         =   "Match Case"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   3015
      End
      Begin VB.CheckBox chkCurrent 
         Caption         =   "Check Current Only"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   3015
      End
      Begin VB.CheckBox chkSearchCode 
         Caption         =   "Search in Code"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   3015
      End
      Begin CS.lvButtons_H lvBSearch 
         Height          =   375
         Left            =   1920
         TabIndex        =   20
         ToolTipText     =   "Delete code"
         Top             =   2040
         Width           =   1215
         _ExtentX        =   2138
         _ExtentY        =   656
         Caption         =   "Search"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16744576
         cGradient       =   16744576
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmMain.frx":5863
         cBack           =   -2147483633
      End
      Begin CS.lvButtons_H lvBReset 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Save and Update code"
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1926
         _ExtentY        =   656
         Caption         =   "Reset"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16744576
         cGradient       =   16744576
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmMain.frx":5DFD
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin CS.lvButtons_H lvBHide 
         Height          =   375
         Left            =   840
         TabIndex        =   27
         ToolTipText     =   "Delete code"
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   656
         Caption         =   "Hide"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16744576
         cGradient       =   16744576
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmMain.frx":6397
         cBack           =   -2147483633
      End
   End
   Begin CS.lvButtons_H lvBSummary 
      Height          =   375
      Left            =   7080
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5313
      _ExtentY        =   656
      Caption         =   "Code Summary"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   16761024
      cGradient       =   16761024
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":71E9
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ImageList IL_Menu 
      Left            =   4080
      Top             =   4080
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7783
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtCodeUnderType 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   6120
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CDMain 
      Left            =   5340
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Left            =   6720
      TabIndex        =   3
      Top             =   120
      Width           =   2235
   End
   Begin VB.TextBox txtCurrLang 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6120
      Width           =   2655
   End
   Begin VB.TextBox txtAddNew 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox txtHasChange 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox txtCurrType 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   6120
      Width           =   2655
   End
   Begin VB.TextBox txtCodeName 
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   600
      Width           =   5655
   End
   Begin MSComctlLib.ImageList IL_Types 
      Left            =   4740
      Top             =   4080
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":85D5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin CS.lvButtons_H lvBLanguage 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "List, Add, Edit, Delete Language"
      Top             =   120
      Width           =   1455
      _ExtentX        =   2561
      _ExtentY        =   656
      Caption         =   "Language"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   16744576
      cGradient       =   16744576
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmMain.frx":872A
      cBack           =   -2147483633
   End
   Begin VB.CommandButton cmdStretch 
      Height          =   2655
      Left            =   2760
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   120
   End
   Begin MSComctlLib.TreeView tvCode 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   2535
      _ExtentX        =   4466
      _ExtentY        =   9335
      _Version        =   393217
      Indentation     =   176
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "IL_Types"
      Appearance      =   1
   End
   Begin CS.lvButtons_H lvBTypes 
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      ToolTipText     =   "List, Add, Edit, Delete Types"
      Top             =   120
      Width           =   1455
      _ExtentX        =   2561
      _ExtentY        =   656
      Caption         =   "Types"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   16744576
      cGradient       =   16744576
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmMain.frx":957C
      cBack           =   -2147483633
   End
   Begin CS.lvButtons_H lvBNewCode 
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      ToolTipText     =   "Create new code"
      Top             =   120
      Width           =   1455
      _ExtentX        =   2985
      _ExtentY        =   656
      Caption         =   "NewCode"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   16744576
      cGradient       =   16744576
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmMain.frx":9E56
      cBack           =   -2147483633
   End
   Begin CS.lvButtons_H lvBAddSave 
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      ToolTipText     =   "Save and Update code"
      Top             =   120
      Width           =   1695
      _ExtentX        =   2985
      _ExtentY        =   656
      Caption         =   "Add\Save"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   16744576
      cGradient       =   16744576
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmMain.frx":A730
      cBack           =   -2147483633
   End
   Begin CS.lvButtons_H lvBDelete 
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      ToolTipText     =   "Delete code"
      Top             =   120
      Width           =   1455
      _ExtentX        =   2561
      _ExtentY        =   656
      Caption         =   "Delete"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   16744576
      cGradient       =   16744576
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmMain.frx":B00A
      cBack           =   -2147483633
   End
   Begin CS.lvButtons_H lvBMove 
      Height          =   375
      Left            =   9060
      TabIndex        =   14
      ToolTipText     =   "Move code"
      Top             =   120
      Width           =   1035
      _ExtentX        =   1820
      _ExtentY        =   656
      Caption         =   "Move Code"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   16744576
      cGradient       =   16744576
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      cBack           =   -2147483633
   End
   Begin CS.lvButtons_H lvBStatus 
      Height          =   375
      Left            =   3960
      TabIndex        =   26
      Top             =   1080
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5736
      _ExtentY        =   656
      Caption         =   "Ready to Search..."
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   16761024
      cGradient       =   16761024
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":B8E4
      cBack           =   -2147483633
   End
   Begin RichTextLib.RichTextBox RTBSummary 
      Height          =   4455
      Left            =   6600
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5313
      _ExtentY        =   7853
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   3
      RightMargin     =   2.50000e5
      TextRTF         =   $"frmMain.frx":BE7E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   4815
      Left            =   3960
      TabIndex        =   1
      Top             =   1080
      Width           =   5655
      _ExtentX        =   9970
      _ExtentY        =   8488
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   3
      RightMargin     =   2.50000e5
      TextRTF         =   $"frmMain.frx":BF03
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CS.McToolBar McTB1 
      Height          =   2004
      Left            =   3000
      TabIndex        =   29
      Top             =   600
      Width           =   528
      _ExtentX        =   931
      _ExtentY        =   3535
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Button_Count    =   5
      ButtonsPerRow   =   1
      HoverColor      =   8438015
      TooTipStyle     =   0
      BackGradientCol =   16761024
      ButtonsMode     =   5
      ButtonsBackColor=   14807794
      ButtonsGradientCol=   16761024
      ButtonsGradient =   6
      ButtonCaption1  =   ""
      ButtonIcon1     =   "frmMain.frx":BF84
      ButtonToolTipText1=   "Copy Code"
      ButtonToolTipIcon1=   1
      ButtonCaption2  =   ""
      ButtonIcon2     =   "frmMain.frx":C51E
      ButtonToolTipText2=   "Paste from clipboard"
      ButtonToolTipIcon2=   1
      ButtonCaption3  =   ""
      ButtonIcon3     =   "frmMain.frx":CAB8
      ButtonToolTipText3=   "Undo"
      ButtonToolTipIcon3=   1
      ButtonCaption4  =   ""
      ButtonIcon4     =   "frmMain.frx":D052
      ButtonToolTipText4=   "Cut Selected"
      ButtonToolTipIcon4=   1
      ButtonCaption5  =   ""
      ButtonIcon5     =   "frmMain.frx":D5EC
      ButtonToolTipText5=   "Search Code"
      ButtonToolTipIcon5=   1
   End
   Begin CS.McToolBar McTB2 
      Height          =   2508
      Left            =   3000
      TabIndex        =   30
      Top             =   2700
      Width           =   528
      _ExtentX        =   931
      _ExtentY        =   3535
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Button_Count    =   5
      ButtonsPerRow   =   1
      HoverColor      =   8438015
      TooTipStyle     =   0
      BackGradientCol =   16761024
      ButtonsMode     =   5
      ButtonsBackColor=   14807794
      ButtonsGradientCol=   16761024
      ButtonsGradient =   6
      ButtonCaption1  =   ""
      ButtonIcon1     =   "frmMain.frx":DB86
      ButtonToolTipText1=   "Color Code"
      ButtonToolTipIcon1=   1
      ButtonCaption2  =   ""
      ButtonIcon2     =   "frmMain.frx":E120
      ButtonToolTipText2=   "Show Code Summary"
      ButtonToolTipIcon2=   1
      ButtonCaption3  =   ""
      ButtonIcon3     =   "frmMain.frx":E6BA
      ButtonToolTipText3=   "Show Code Details"
      ButtonToolTipIcon3=   1
      ButtonCaption4  =   ""
      ButtonIcon4     =   "frmMain.frx":EC54
      ButtonToolTipText4=   "Enable\Disable Checkboxes"
      ButtonToolTipIcon4=   1
      ButtonCaption5  =   "Fmt"
      ButtonToolTipText5=   "Close Format Code"
      ButtonToolTipIcon5=   1
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "Search"
         Index           =   0
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Exit"
         Index           =   1
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MenuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "List Language"
         Index           =   0
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "List Type"
         Index           =   1
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Close Format of Code"
         Index           =   3
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Delete Number/Dots (vbforums)"
         Index           =   4
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu menuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuTools 
         Caption         =   "BackUp Code and Settings"
         Index           =   0
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Restore BackUp"
         Index           =   1
      End
      Begin VB.Menu mnuTools 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Compact and Repair Database"
         Index           =   3
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuTools 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuT 
         Caption         =   "Clean All Codes"
         Index           =   0
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuT 
         Caption         =   "Open Notepad"
         Index           =   1
      End
   End
   Begin VB.Menu menuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptions 
         Caption         =   "Word Wrap"
         Index           =   0
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Select Font"
         Index           =   1
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Gradient Backgorund"
         Index           =   2
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Disable Color Code"
         Index           =   4
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Enable Back to Last Code after Exit"
         Index           =   5
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Enable Indented View"
         Index           =   6
      End
   End
   Begin VB.Menu MenuLang 
      Caption         =   "&Language"
      Begin VB.Menu mnuLang 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Documentation"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   1
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Owner Info"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------APPLICATION----------------------------------
'App name: CS (Code Storage)
'App desc: Utility to store codes, types from different language,
'   or even notes from batch, pc registry, access or excel programming, etc.
'App author: Daniel A. Cadsawan Jr.
'Author info: xavierjohn22@yahoo.com, +639209190715
'http:\\cad3dmdd.ucoz.com
'http:\\www.cad3dmdd.com

'--------------------------ACKNOWLEDGEMENT---------------------------------
'Special thanks to the following, I modified most on these owner's code
'LGS Static's Code on colorcode, wordlist, find/search codes
'Static, I redo the database but made the samples from Codebank appear herein
'Napparan Philip's style on database connecting and handling
'User Controls acknowledgement as follows:
'McToolBar 2.3 by Jim Jose
'La Volpe Buttons vH.1

'-------------------------------HISTORY------------------------------------
'CS (Code Storage) was inspired by my Windows 7 laptop where i experienced
'a lot of error and also registering ocx from CodeBank
'I was using CodeBank Version: 4.0.0, Geoff Goldsmith -aka [LGS]Static
'Co-Author is me, Daniel A. Cadsawan Jr. -aka xavierjohn22 only
'bumping it several revisions to get away from the pc registry settings
'I needed to remake this for Windows 7, got rid of the OCX dependency,
'and completely redo pretty much everything, only the connection is active
'For this i connect and get records set the recordset to nothing
'in the future i will cut the connection to the database as well

'---------NOTE ONLY IMPT LIST WHERE THIS PROJECT IS REFERENCED TO----------
'msjro.dll, msbind.dll, msado25.tlb

Option Explicit

Private lngOldX As Long
Private lngOldY As Long
Private blnIsMoving As Boolean

Private hasChanged As Boolean
Private loadingCode As Boolean

Private currCodeName As String
Private currCodeNo As String
Private currCodeTypeName As String
Private currCodeTypeNo As String
Private newCodeTypeNo As String

Private currTVCodePos As Long
Private currTVTypePos As Long

Private countTypes As Long
Private countCodes As Long

Private addNewCode As Boolean

Private fontColor As Long

Private only_hiding_summary As Boolean
Private only_hiding_details As Boolean

Dim MiceDown As Boolean
Dim IsNodeCheck As Boolean
Dim A_NODE As Node

Private Const Provider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
Private Const JetVersion = ";Jet OLEDB:Engine Type=5"

Dim nextNoFind As Integer
Dim newFind As Boolean

Private Sub Form_Initialize()
    m_hMod = LoadLibrary("shell32.dll")
    InitCommonControls
End Sub

Private Sub Form_Paint()
On Error GoTo ErrorHandler:
    'MsgBox RTop
    PaintGradient Me, RTop, GTop, BTop, RBtm, GBtm, BBtm
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler:
    'variables
    str_WLFolder = App.Path & "\" & App.Title & "WordList"
    str_SetFolder = App.Path & "\" & App.Title & "Settings"
    str_IconFolder = App.Path & "\" & App.Title & "Icons"
    str_BackUpsFolder = App.Path & "\" & App.Title & "BackUps"
    str_databaseFile = App.Path & "\Data.cs"
    
    str_iniSet = str_SetFolder & "\" & App.Title & "_Settings.ini"
    
    'This is the only connection active as of 7/14/2010
    Call Get_Connected(cn, str_databaseFile, True, "dani")

    'Just put the current type selected to 1, for loading only
    currTVCodePos = 1
            
    pbarCaption = "Loading database... Please wait..."
    frmProgressBar.Seconds = 2
    frmProgressBar.Show 1, Me
            
    Load_Settings
    
    Me.MousePointer = vbHourglass
    LockWindowUpdate Me.hwnd
    
    Load_OwnerInfo
    Load_LangInMenu
    Load_TypeNames
    Load_CodeNames
        
    If mnuOptions(5).Checked = True Then GoBack_PrevNode
    
    LockWindowUpdate 0&
    Me.MousePointer = vbDefault
    
    ' do this for now to fire enabling these controls
    RTB1.Enabled = True
    txtCodeName.Enabled = True
    
    'eto naman sa load ng search, set to true agad
    newFind = True

    Call Hook(Me.hwnd)
Exit Sub
ErrorHandler:
MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
    vbOKOnly + vbInformation, App.Title
End Sub

Private Sub GoBack_PrevNode()
    Dim prevNKey As String
    Dim nCount As Long
    If Not ReadIni(str_iniSet, "Options", "LastKey") = vbNullString Then
        prevNKey = ReadIni(str_iniSet, "Options", "LastKey")
    Else
        prevNKey = ""
    End If
    For nCount = 1 To tvCode.Nodes.Count
        If tvCode.Nodes(nCount).Key = prevNKey Then
            tvCode.Nodes(prevNKey).Selected = True
            tvCode.Nodes(prevNKey).EnsureVisible
            tvCode_Click
            Exit For
        End If
    Next nCount
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveAndCheckCodeChange
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Save_Settings
    'So here i cut the connection, but in the future
    'i will cut the connection once i have loaded the database
    Set cn = Nothing
    FreeLibrary m_hMod
    Call Unhook(Me.hwnd)
End Sub

Private Sub Load_OwnerInfo()
On Error GoTo ErrorHandler:
    Dim rs_owner As New ADODB.Recordset
    Call Get_Records(rs_owner, cn, _
    "Select TableOwner.* From TableOwner")
    If rs_owner.RecordCount > 0 Then
        With rs_owner
            .MoveFirst
            owner_name = .Fields(0)
            owner_address = .Fields(1)
        End With
    End If
    Set rs_owner = Nothing
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
        Set rs_owner = Nothing
End Sub

Public Sub Load_LangInMenu()
On Error GoTo ErrorHandler:
    Dim X As Long
    'clear menu muna
    If mnuLang.Count > 1 Then
    'zero is separator
    For X = 1 To mnuLang.Count - 1
        Unload mnuLang(X)
    Next X
    End If
    'connect to table then load sa menu
    Call Get_Records(rs_lang, cn, _
    "Select TableLang.* From TableLang Order by LangName ASC")
    If rs_lang.RecordCount > 0 Then
        With rs_lang
            .MoveFirst
            For X = 1 To .RecordCount
                Load mnuLang(X)
                mnuLang(X).Caption = "Show " & """" & .Fields(0).Value & """" & " Codes"
                mnuLang(X).Visible = True
                mnuLang(X).Checked = False
                mnuLang(X).Tag = "Lang_" & .Fields(0).Value
                .MoveNext
            Next X
        End With
        Set rs_lang = Nothing
    Else
        Set rs_lang = Nothing
    End If
    
    'read menu, count it, then set CurrLangNo
    Select Case mnuLang.Count
        Case 0, 1
            '0 = separator
        Case 2
            CurrLangNo = 1
            mnuLang(CurrLangNo).Checked = True
            CurrLangText = mnuLang(CurrLangNo).Caption
            CurrLangText = Right$(CurrLangText, Len(mnuLang(CurrLangNo).Caption) - 6)
            CurrLangText = Left$(CurrLangText, Len(CurrLangText) - 7)
        Case Is > 2
            If ReadIni(str_iniSet, "Settings", "CurrentLanguage") = vbNullString Then
                CurrLangNo = 1
            Else
                CurrLangNo = ReadIni(str_iniSet, "Settings", "CurrentLanguage")
            End If
            mnuLang(CurrLangNo).Checked = True
            CurrLangText = mnuLang(CurrLangNo).Caption
            CurrLangText = Right$(CurrLangText, Len(mnuLang(CurrLangNo).Caption) - 6)
            CurrLangText = Left$(CurrLangText, Len(CurrLangText) - 7)
    End Select
    'including '0' index
    'MsgBox mnuLang.Count & CurrLangNo & currlangtext
    txtCurrLang.Text = " Current Language = " & CurrLangText
    
    LoadWordsAndColors
    
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
    Set rs_lang = Nothing
End Sub

'Called by Load_LangInMenu after loading the languages
Private Sub LoadWordsAndColors()
On Error GoTo ErrorHandler:
    Call Get_Records(rs_lang, cn, _
    "Select Tablelang.* From TableLang Where LangName ='" & CurrLangText & "'")
    If rs_lang.RecordCount > 0 Then
        With rs_lang
            'assign colors read from database
            'Wordlist 1 to 3()
            str_WLOne = .Fields(2)
            WLOneColor = .Fields(3)
            str_WLTwo = .Fields(4)
            WLTwoColor = .Fields(5)
            str_WLThree = .Fields(6)
            WLThreeColor = .Fields(7)
            
            str_CMarker = .Fields(8)
            CMarkerColor = .Fields(9)
        End With
    End If
    Set rs_lang = Nothing
    'Now load the wordlist in Memory
    LoadWordLists
    'MsgBox str_WLOne & ", " & str_WLTwo & ", " & str_WLThree & vbCrLf _
        & WLOneColor & ", " & WLTwoColor & ", " & WLThreeColor & vbCrLf _
        & str_CMarker & ", " & CMarkerColor
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

'LOAD TYPES
Public Sub Load_TypeNames()
On Error GoTo ErrorHandler:
    'Clear treeview muna, pati combo type
    tvCode.Nodes.Clear
    cboType.Clear
    'Clear other imagelist, other 2 default remains
    Dim i As Long
    For i = IL_Types.ListImages.Count To 2 Step -1
        IL_Types.ListImages.Remove i
    Next i

    Call Get_Records(rs_type, cn, _
    "Select TableType.* From TableType Where LangName ='" & CurrLangText & _
    "' Order by TypeNo ASC")
    
    'record number of types in memory
    countTypes = rs_type.RecordCount
    
    If rs_type.RecordCount > 0 Then
        rs_type.MoveFirst
        Do While Not rs_type.EOF
            'Load available icons in imagelist
            'IL_Types.ListImages.Add , CurrLangText & rs_type.Fields(1), _
            '   LoadPicture(str_IconFolder & "\" & rs_type.Fields(3))
            'Load in tvcode with its icon image
            If FileExists(str_IconFolder & "\" & rs_type.Fields(3)) Then
                IL_Types.ListImages.Add , CurrLangText & rs_type.Fields(1), _
                LoadPicture(str_IconFolder & "\" & rs_type.Fields(3))
                tvCode.Nodes.Add , , _
                "ROOT||" & rs_type.Fields(1) & "||CHILD" & rs_type.Fields(0), _
                rs_type.Fields(1), _
                CurrLangText & rs_type.Fields(1) 'icon loaded in imagelist
                'MsgBox CurrLangText & rs_type.Fields(1)
            Else
                tvCode.Nodes.Add , , _
                "ROOT||" & rs_type.Fields(1) & "||CHILD" & rs_type.Fields(0), _
                rs_type.Fields(1), _
                1 'icon default in imagelist
            End If
            tvCode.Nodes.Item("ROOT||" & rs_type.Fields(1) & "||CHILD" & rs_type.Fields(0)).ForeColor = vbBlue
           'load in cbotype
            cboType.AddItem rs_type.Fields(1)
            rs_type.MoveNext
        Loop
    End If
    
    Set rs_type = Nothing
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
    Set rs_type = Nothing
End Sub

'LOAD CODE NAMES
Public Sub Load_CodeNames()
On Error GoTo ErrorHandler:
    Dim X As Long
    'Loop through the TYPES
    For X = 1 To tvCode.Nodes.Count
        
        'GET the "record name" only,
        'KEY is "ROOT||" prefix + RecName + "||CHILD" + RecNo suffix
        'NEED TO DO THIS WHEN LOADING TYPE, FEW FILTER
        Dim strTypeName As String
        Dim strTypeno As String
        strTypeName = tvCode.Nodes(X).Key
        strTypeName = SubstringAfter(strTypeName, "ROOT||")
        strTypeName = SubstringBefore(strTypeName, "||CHILD")
        
        strTypeno = tvCode.Nodes(X).Key
        strTypeno = SubstringAfter(strTypeno, "||CHILD")

        'MsgBox strTypeName
        'MsgBox strTypeNo
        
        'need to Open then Close for each tvcode key
        Call Get_Records(rs_qry_code, cn, _
        "Select qryCodes.* From qryCodes Order by CodeName ASC")
        
        rs_qry_code.Filter = "LangName ='" & CurrLangText & "'" _
            & " And TypeName = '" & strTypeName & "'" _
            & " And TypeNo = " & CLng(strTypeno)
        
        'record number of types in memory
        'MsgBox rs_qry_code.RecordCount
        
        If rs_qry_code.RecordCount > 0 Then
            rs_qry_code.MoveFirst
            'MsgBox tvCode.Nodes(x).Key
            Do While Not rs_qry_code.EOF
                'If icon image file exists use this
                If FileExists(str_IconFolder & "\" & rs_qry_code.Fields(6)) Then
                    tvCode.Nodes.Add tvCode.Nodes(X).Key, tvwChild, _
                    rs_qry_code.Fields(0) & "||CHILD" & rs_qry_code.Fields(5), _
                    rs_qry_code.Fields(0), CurrLangText & strTypeName 'iconimage from types
                    'MsgBox CurrLangText & strTypeName
               Else
                    tvCode.Nodes.Add tvCode.Nodes(X).Key, tvwChild, _
                    rs_qry_code.Fields(0) & "||CHILD" & rs_qry_code.Fields(5), _
                    rs_qry_code.Fields(0), 1 'iconimage from default
                End If
                rs_qry_code.MoveNext
            Loop
        End If
        'count all codes
        countCodes = countCodes + rs_qry_code.RecordCount

        Set rs_qry_code = Nothing
    Next X
    
    'Set up how to select in tv code
    Dim strRoot As String
    Dim cnt As Long
    Dim cntRoot As Long
    For cnt = 1 To tvCode.Nodes.Count
        strRoot = tvCode.Nodes.Item(cnt).Key
        strRoot = Left$(strRoot, Len(strRoot) - (Len(strRoot) - 6))
        If strRoot = "ROOT||" Then
            cntRoot = cntRoot + 1
        End If
    Next cnt
    'MsgBox cntRoot
    Select Case cntRoot
        Case Is <= 0
            Exit Sub
        Case Is = 1
            tvCode.Nodes(1).Selected = True
            tvCode.Nodes(1).EnsureVisible
            tvCode_Click
        Case Else
            'MsgBox currTVCodePos
            tvCode.Nodes(CLng(currTVCodePos)).Selected = True
            tvCode.Nodes(CLng(currTVCodePos)).EnsureVisible
            tvCode_Click
    End Select
    
    LoadFormHeading
    
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
Set rs_qry_code = Nothing
End Sub

Private Sub Load_CodeContent()
On Error GoTo ErrorHandler:
    
    RTB1.Text = ""
    
    Call Get_Records(rs_qry_code, cn, _
    "Select qryCodes.* From qryCodes Order by CodeName ASC")
        
    'Code name is currCodeName
    rs_qry_code.Filter = "LangName ='" & CurrLangText & "'" _
        & " And Typename = '" & currCodeTypeName & "'" _
        & " And TypeNo = " & CLng(currCodeTypeNo) _
        & " And CodeName = '" & currCodeName & "'" _
        & " And CodeNo = " & CLng(currCodeNo)
        
    Me.MousePointer = vbHourglass
    LockWindowUpdate Me.hwnd
        
    If rs_qry_code.RecordCount > 0 Then
        'rs_qry_code.MoveFirst
        txtCodeName.Text = rs_qry_code.Fields(0)
        RTB1.Text = rs_qry_code.Fields(1)
        RTBSummary.Text = rs_qry_code.Fields(2)
    End If
        
    'close recordset
    Set rs_qry_code = Nothing
                
    ColorTheCode
            
    LockWindowUpdate 0&
    Me.MousePointer = vbDefault
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
Set rs_qry_code = Nothing
        LockWindowUpdate 0&
        Me.MousePointer = vbDefault
End Sub

Private Sub LoadFormHeading()
    Me.Caption = App.Title _
        & " Version " & App.Major & "." _
        & App.Revision & "." & App.Minor & " - " _
        & "(" & countCodes & ") Codes in " _
        & "(" & countTypes & ") Types of " _
        & CurrLangText & " Language"
    countTypes = 0
    countCodes = 0
End Sub

Private Sub ColorTheCode()
    Me.MousePointer = vbHourglass
    LockWindowUpdate RTB1.hwnd
    If mnuOptions(4).Checked = False Then
        Call Color_Code(RTB1, WLOneColor, WLTwoColor, WLThreeColor, _
                CMarkerColor, str_CMarker)
    Else
        RTB1.SelStart = 0
        RTB1.SelLength = Len(RTB1.Text)
        RTB1.SelColor = fontColor
        RTB1.SelLength = 0
    End If
    LockWindowUpdate 0&
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    'heights of control
On Error GoTo ErrorHandler:
    If Me.WindowState = vbMinimized Then Exit Sub
    tvCode.height = Me.ScaleHeight - lvBLanguage.height - 240 - txtCurrType.height
    RTB1.height = tvCode.height - txtCodeName.height - 120
    RTBSummary.Top = RTB1.Top + lvBSummary.height
    RTBSummary.height = RTB1.height - lvBSummary.height
    RTBDetails.Top = RTBSummary.Top
    RTBDetails.height = RTBSummary.height
    
    txtCurrLang.Top = Me.ScaleHeight - txtCurrType.height
    txtCurrType.Top = txtCurrLang.Top
    txtCodeUnderType.Top = txtCurrLang.Top
    
    txtHasChange.Top = txtCurrLang.Top
    txtAddNew.Top = txtCurrLang.Top
    
    'always make the find pic appear
    If lvBStatus.Top > RTB1.Top + RTB1.height - picSearch.height - lvBStatus.height Then _
        lvBStatus.Top = RTB1.Top + RTB1.height - picSearch.height - lvBStatus.height: _

    cmdStretch.height = tvCode.height
    
    StretchSize
    FollowFind
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub StretchSize()
On Error GoTo ErrorHandler:
    'widths of control
    tvCode.width = cmdStretch.Left - 180
    McTB1.Left = cmdStretch.Left + cmdStretch.width + 60
    McTB2.Left = McTB1.Left
    txtCodeName.Left = McTB1.Left + McTB1.width + 60
    txtCodeName.width = Me.ScaleWidth - cmdStretch.Left _
    - cmdStretch.width - McTB1.width - 180
    RTB1.width = txtCodeName.width
    RTB1.Left = txtCodeName.Left
    RTBSummary.width = RTB1.width / 3
    RTBSummary.Left = RTB1.Left + RTB1.width - RTBSummary.width
    lvBSummary.Left = RTBSummary.Left
    lvBSummary.width = RTBSummary.width
    RTBDetails.Left = RTBSummary.Left
    RTBDetails.width = RTBSummary.width
    
    lvBStatus.Left = RTB1.Left
    picSearch.Left = lvBStatus.Left
    
    DoEvents: DoEvents: DoEvents
    
    Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title

End Sub

Private Sub Load_Settings()
On Error GoTo ErrorHandler:
    'FOLDERS
    If Len(Dir(str_WLFolder, vbDirectory)) = 0 Then
        MkDir str_WLFolder
    End If
    If Len(Dir(str_SetFolder, vbDirectory)) = 0 Then
        MkDir str_SetFolder
    End If
    If Len(Dir(str_IconFolder, vbDirectory)) = 0 Then
        MkDir str_IconFolder
    End If
    If Len(Dir(str_BackUpsFolder, vbDirectory)) = 0 Then
        MkDir str_BackUpsFolder
    End If
    'POSITIONS
    If ReadIni(str_iniSet, "Positioner", "Left") = vbNullString Then
        WriteIni str_iniSet, "Positioner", "Left", 2760
    Else
        cmdStretch.Left = ReadIni(str_iniSet, "Positioner", "Left")
    End If
    If ReadIni(str_iniSet, "FormPosition", "Top") = vbNullString Then
        WriteIni str_iniSet, "FormPosition", "Top", 120
    Else
        Me.Top = ReadIni(str_iniSet, "FormPosition", "Top")
    End If
    If ReadIni(str_iniSet, "FormPosition", "Left") = vbNullString Then
        WriteIni str_iniSet, "FormPosition", "Left", 120
    Else
        Me.Left = ReadIni(str_iniSet, "FormPosition", "Left")
    End If
    If ReadIni(str_iniSet, "FormPosition", "Width") = vbNullString Then
        WriteIni str_iniSet, "FormPosition", "Width", 10000
    Else
        Me.width = ReadIni(str_iniSet, "FormPosition", "Width")
    End If
    If ReadIni(str_iniSet, "FormPosition", "Height") = vbNullString Then
        WriteIni str_iniSet, "FormPosition", "Height", 7000
    Else
        Me.height = ReadIni(str_iniSet, "FormPosition", "Height")
    End If
    'FIND FORM POSITION
    If ReadIni(str_iniSet, "FindForm", "Top") = vbNullString Then
        WriteIni str_iniSet, "FindForm", "Top", RTB1.Top
    Else
        lvBStatus.Top = ReadIni(str_iniSet, "FindForm", "Top")
    End If
    'OPTIONS
    If ReadIni(str_iniSet, "Options", "WordWrap") = vbNullString Then
        WriteIni str_iniSet, "Options", "WordWrap", 0
    Else
        mnuOptions(0).Checked = ReadIni(str_iniSet, "Options", "WordWrap")
        RTB1.RightMargin = IIf(mnuOptions(0).Checked, 0, 200000)
    End If
    If ReadIni(str_iniSet, "Options", "ColorCode") = vbNullString Then
        WriteIni str_iniSet, "Options", "ColorCode", mnuOptions(4).Checked
    Else
        mnuOptions(4).Checked = ReadIni(str_iniSet, "Options", "ColorCode")
    End If
    If ReadIni(str_iniSet, "Options", "LastKeyEnabled") = vbNullString Then
        WriteIni str_iniSet, "Options", "LastKeyEnabled", mnuOptions(5).Checked
    Else
        mnuOptions(5).Checked = ReadIni(str_iniSet, "Options", "LastKeyEnabled")
    End If
    If ReadIni(str_iniSet, "Options", "EnableIndent") = vbNullString Then
        WriteIni str_iniSet, "Options", "EnableIndent", mnuOptions(6).Checked
    Else
        mnuOptions(6).Checked = ReadIni(str_iniSet, "Options", "EnableIndent")
    End If
    'FONTS
    If ReadIni(str_iniSet, "Fonts", "Name") = vbNullString Then
        WriteIni str_iniSet, "Fonts", "Name", RTB1.Font.Name
    Else
        RTB1.Font.Name = ReadIni(str_iniSet, "Fonts", "Name")
    End If
    If ReadIni(str_iniSet, "Fonts", "Size") = vbNullString Then
        WriteIni str_iniSet, "Fonts", "Size", RTB1.Font.Size
    Else
        RTB1.Font.Size = ReadIni(str_iniSet, "Fonts", "Size")
    End If
    If ReadIni(str_iniSet, "Fonts", "Italic") = vbNullString Then
        WriteIni str_iniSet, "Fonts", "Italic", RTB1.Font.Italic
    Else
        RTB1.Font.Italic = ReadIni(str_iniSet, "Fonts", "Italic")
    End If
    If ReadIni(str_iniSet, "Fonts", "Bold") = vbNullString Then
        WriteIni str_iniSet, "Fonts", "Bold", RTB1.Font.Bold
    Else
        RTB1.Font.Bold = ReadIni(str_iniSet, "Fonts", "Bold")
    End If
    If ReadIni(str_iniSet, "Fonts", "Strikethru") = vbNullString Then
        WriteIni str_iniSet, "Fonts", "Strikethru", RTB1.Font.Strikethrough
    Else
        RTB1.Font.Strikethrough = ReadIni(str_iniSet, "Fonts", "Strikethru")
    End If
    If ReadIni(str_iniSet, "Fonts", "Underline") = vbNullString Then
        WriteIni str_iniSet, "Fonts", "Underline", RTB1.Font.Underline
    Else
        RTB1.Font.Underline = ReadIni(str_iniSet, "Fonts", "Underline")
    End If
    If ReadIni(str_iniSet, "Fonts", "Color") = vbNullString Then
        WriteIni str_iniSet, "Fonts", "Color", vbBlack
    Else
        fontColor = ReadIni(str_iniSet, "Fonts", "Color")
    End If
    
    Read_SaveGradient
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Public Sub Read_SaveGradient()
    'GRADIENT
    Dim i As Integer
    Dim skey As String
    For i = 1 To 6
        skey = "Color" & CStr(i)
        If ReadIni(str_iniSet, "Gradient", skey) = vbNullString Then
            WriteIni str_iniSet, "Gradient", skey, 255
        End If
    Next i
    RTop = ReadIni(str_iniSet, "Gradient", "Color1")
    GTop = ReadIni(str_iniSet, "Gradient", "Color2")
    BTop = ReadIni(str_iniSet, "Gradient", "Color3")
    RBtm = ReadIni(str_iniSet, "Gradient", "Color4")
    GBtm = ReadIni(str_iniSet, "Gradient", "Color5")
    BBtm = ReadIni(str_iniSet, "Gradient", "Color6")
End Sub

Private Sub Save_Settings()
On Error GoTo ErrorHandler:
    'Form
    WriteIni str_iniSet, "Positioner", "Left", cmdStretch.Left
    WriteIni str_iniSet, "FormPosition", "Top", Me.Top
    WriteIni str_iniSet, "FormPosition", "Left", Me.Left
    WriteIni str_iniSet, "FormPosition", "Width", Me.width
    WriteIni str_iniSet, "FormPosition", "Height", Me.height
    'Find pic
    WriteIni str_iniSet, "FindForm", "Top", lvBStatus.Top
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub cmdStretch_MouseDown(Button As Integer, Shift As Integer, _
                                X As Single, Y As Single)
    lngOldX = X
    lngOldY = Y
    blnIsMoving = True
End Sub

Private Sub cmdStretch_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    On Error Resume Next
    Me.MousePointer = vbSizeWE
    If cmdStretch.Left < 1000 Then cmdStretch.Left = 1000: _
        blnIsMoving = False: Me.MousePointer = vbDefault
    If cmdStretch.Left > 5000 Then cmdStretch.Left = 5000: _
        blnIsMoving = False: Me.MousePointer = vbDefault
    If blnIsMoving Then
        cmdStretch.Left = cmdStretch.Left - (lngOldX - X)
        StretchSize
    End If
End Sub

Private Sub cmdStretch_MouseUp(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    On Error Resume Next
    
    tvCode.SetFocus
    blnIsMoving = False
    If cmdStretch.Left < 1000 Then cmdStretch.Left = 1000: _
        blnIsMoving = False: Me.MousePointer = vbDefault
    If cmdStretch.Left > 5000 Then cmdStretch.Left = 5000: _
        blnIsMoving = False: Me.MousePointer = vbDefault
    StretchSize
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
    Me.MousePointer = vbDefault
End Sub

Private Sub ShowListLang()
    frmListLang.Show 1, Me
End Sub

Private Sub ShowListType()
    frmListType.Show 1, Me
End Sub

Private Sub lvBLanguage_Click()
    ShowListLang
End Sub

Private Sub cboType_Click()
On Error GoTo ErrorHandler:

    Call Get_Records(rs_type, cn, _
    "Select TableType.* From TableType Where LangName ='" & CurrLangText & _
    "' Order by TypeNo ASC")

    'WEAK FILTER for now, NEED TO CHANGE KEY OF TYPE
    'ADD AUTONUMBER FIELD AND ADD THAT TO FIELD might be an option
    rs_type.Filter = "TypeName = '" & cboType.Text & "'"
    rs_type.MoveFirst
    newCodeTypeNo = rs_type.Fields(0).Value
    'MsgBox newCodeTypeNo
    
    Set rs_type = Nothing
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
    Set rs_type = Nothing
End Sub

Private Sub Move_Single()
    'exit muna pag empty, conflict sa lost focus ng rtb at txtcodename
    If currCodeTypeNo = "" Then Exit Sub
    If currCodeName = "" Then Exit Sub
        
    Dim strMoveToType As String
    strMoveToType = cboType.Text
    Dim currKey As String
    currKey = tvCode.SelectedItem.Key
        
    Dim rep As Integer
    rep = MsgBox("Do you want to move the selected code" & vbCrLf _
        & "to " & """" & strMoveToType & """" & " Type?", _
        vbQuestion + vbYesNo, App.Title)
    If rep = vbYes Then
        Call Get_Records(rs_code, cn, _
        "Select TableCode.* From TableCode Order by CodeName ASC")
            
        rs_code.Filter = "CodeName = " & "'" & currCodeName & "'" _
            & " And TypeNo = " & CLng(currCodeTypeNo)
            
        With rs_code
            .MoveFirst
            '.Fields(0) = txtCodeName.Text   'CodeName
            .Fields(1) = CLng(newCodeTypeNo) 'TypeNo
            '.Fields(2) = RTB1.Text          'CodeContent
            '.Fields(3)                      'CodeSummary
            '.Fields(4)                      'CodeNo
            .Update
        End With
        Set rs_code = Nothing
            
        hasChanged = False
        txtHasChange.Text = " Code Change = " & hasChanged
        txtAddNew.Text = " Adding New = " & addNewCode
        Load_TypeNames
        Load_CodeNames
            
        'return
        tvCode.Nodes(currKey).Selected = True
        tvCode.Nodes(currKey).EnsureVisible
        tvCode_Click
            
        MsgBox "Code has been transfered to " & """" & strMoveToType & """" & " Type.", _
            vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub Move_Selected()
    If tvCode.Checkboxes = False Then Exit Sub
    Dim strMoveToType As String
    strMoveToType = cboType.Text

    Dim rep As Integer
    rep = MsgBox("Do you want to move ALL selected code" & vbCrLf _
        & "to " & """" & strMoveToType & """" & " Type?", _
        vbQuestion + vbYesNo, App.Title)
    If rep = vbYes Then
    
    Call Get_Records(rs_code, cn, _
    "Select TableCode.* From TableCode Order by CodeName ASC")
        
    Dim X As Long
    Dim strNo As String
    For X = 1 To tvCode.Nodes.Count
        If tvCode.Nodes(X).Checked = True Then
            strNo = tvCode.Nodes(X).Key
            strNo = SubstringAfter(strNo, "||CHILD")
        
            With rs_code
                .Requery
                .Filter = "CodeName = " & "'" & tvCode.Nodes(X).Text & "'" _
                    & " AND CodeNo = " & CLng(strNo)
                .MoveFirst
                '.Fields(0) = txtCodeName.Text   'CodeName
                .Fields(1) = CLng(newCodeTypeNo) 'TypeNo
                '.Fields(2) = RTB1.Text          'CodeContent
                '.Fields(3)                      'CodeSummary
                '.Fields(4)                      'CodeNo
                .Update
            End With
            Debug.Print rs_code.Fields(0) & vbTab & rs_code.Fields(1)
        End If
    Next X
    
    Set rs_code = Nothing
    
    hasChanged = False
    txtHasChange.Text = " Code Change = " & hasChanged
    txtAddNew.Text = " Adding New = " & addNewCode
    Load_TypeNames
    Load_CodeNames
    
    'not yet done
    'go to the trabsferred key
    'tvCode.Nodes(cboType.Text).Selected = True
    'tvCode.Nodes(cboType.Text).EnsureVisible
    'tvCode_Click

    MsgBox "Code has been transfered to " & """" & strMoveToType & """" & " Type.", _
        vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub lvBMove_Click()
On Error GoTo ErrorHandler:
    If cboType.ListIndex < 0 Then
        MsgBox "Please choose a valid 'Code Type' to move the code to!", _
            vbOKOnly + vbExclamation, App.Title
    Exit Sub
    End If
    
    If tvCode.Checkboxes = True Then
        Move_Selected
    Else
        Move_Single
    End If

Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
    Set rs_code = Nothing
End Sub

Private Sub lvBTypes_Click()
    ShowListType
End Sub

Private Sub McTB1_Click(ByVal ButtonIndex As Long)
On Error GoTo ErrorHandler:
    Select Case ButtonIndex
    
        Case 1
            CSCopyCode
        Case 2
            CSPasteCode
        Case 3
            CSUndo
        Case 4
            CSCutCode
        Case 5
            FindPic_HideUnHide
    End Select
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub lvBSummary_Click()
    'MsgBox only_hiding_details
    If lvBSummary.Caption = "Code Details" Or lvBSummary.Caption = "Hiding Code Details" Then
    If only_hiding_details = True Then
        RTBDetails.Visible = True
        lvBSummary.Caption = "Code Details"
        only_hiding_details = False
        'Exit Sub
    Else
        RTBDetails.Visible = False
        lvBSummary.Caption = "Hiding Code Details"
        only_hiding_details = True
        'Exit Sub
    End If
    End If
    'MsgBox only_hiding_summary
    If lvBSummary.Caption = "Code Summary" Or lvBSummary.Caption = "Hiding Summary Window" Then
    If only_hiding_summary = True Then
        RTBSummary.Visible = True
        lvBSummary.Caption = "Code Summary"
        only_hiding_summary = False
    Else
        RTBSummary.Visible = False
        lvBSummary.Caption = "Hiding Summary Window"
        only_hiding_summary = True
    End If
    End If
End Sub

Private Sub DisplayCodeInfo()
    '---------------------------------------------
    Dim strSource As String
    Dim strDigits As String
    Dim strLetters As String
    Dim strOtherChars As String
    strSource = RTB1.Text
    Call FilterChars(strSource, strDigits, strLetters, strOtherChars)
    '---------------------------------------------
  
    RTBDetails.Text = _
        "-----------------------" & vbCrLf _
        & Count_Lines(RTB1.Text) & vbTab & "Lines" & vbCrLf _
        & Count_Spaces(RTB1.Text) & vbTab & "Spaces" & vbCrLf _
        & Len(RTB1.Text) & vbTab & "Chars/Len" & vbCrLf _
        & Vowel_Count(RTB1.Text) & vbTab & "Vowels" & vbCrLf _
        & Word_Count(RTB1.Text) & vbTab & "Approx.Words" & vbCrLf _
        & TextWidth(RTB1.Text) & vbTab & "TextWidth" & vbCrLf _
        & "-----------------------" & vbCrLf _
        & Len(strDigits) & vbTab & "No of Digits" & vbCrLf _
        & Len(strLetters) & vbTab & "No of Letters" & vbCrLf _
        & Len(strOtherChars) & vbTab & "No of Other Chars"
End Sub

Private Sub CSCutCode()
    SendMessage RTB1.hwnd, WM_CUT, 0&, 0&
End Sub

Private Sub CSUndo()
    SendMessage RTB1.hwnd, EM_UNDO, 0&, 0&
End Sub

Private Sub CSCopyCode()
    Clipboard.Clear
    If RTB1.SelText = "" Then
        Clipboard.SetText Replace(RTB1.Text, vbTab, "    ")
    Else
        Clipboard.SetText Replace(RTB1.SelText, vbTab, "    ")
    End If
    RTB1.SetFocus
    'SetEditMenu
End Sub

Private Sub CSPasteCode()
On Error GoTo ErrorHandler:
    Dim tmp As Long
    If RTB1.Text = "" Then
        RTB1.SelColor = vbBlack
        RTB1.Text = Clipboard.GetText & vbCrLf
        ColorTheCode
        RTB1.SelStart = 0
    ElseIf RTB1.SelText <> "" Then
        tmp = RTB1.SelStart
        RTB1.SelText = Clipboard.GetText & vbCrLf
        RTB1.SelStart = 0
        RTB1.SelLength = Len(RTB1.Text)
        RTB1.SelColor = vbBlack
        ColorTheCode
        RTB1.SelStart = tmp + Len(Clipboard.GetText)
    Else
        tmp = RTB1.SelStart
        RTB1.Text = Left(RTB1.Text, RTB1.SelStart) & Clipboard.GetText & Right(RTB1.Text, Len(RTB1.Text) - RTB1.SelStart)
        RTB1.SelStart = 0
        RTB1.SelLength = Len(RTB1.Text)
        RTB1.SelColor = vbBlack
        ColorTheCode
        RTB1.SelStart = tmp + Len(Clipboard.GetText)
    End If
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
    LockWindowUpdate 0&
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuOptions_Click(Index As Integer)
    Select Case Index
        Case 0
            mnuOptions(0).Checked = Not mnuOptions(0).Checked
            RTB1.RightMargin = IIf(mnuOptions(0).Checked, 0, 200000)
            WriteIni str_iniSet, "Options", "WordWrap", mnuOptions(0).Checked
        Case 1
            loadingCode = True
            SelectFont
            loadingCode = False
        Case 2
            frmGradient.Show 1, Me
        Case 4
            mnuOptions(4).Checked = Not mnuOptions(4).Checked
            WriteIni str_iniSet, "Options", "ColorCode", mnuOptions(4).Checked
            loadingCode = True
            ColorTheCode
            loadingCode = False
        Case 5
            mnuOptions(5).Checked = Not mnuOptions(5).Checked
            WriteIni str_iniSet, "Options", "LastKeyEnabled", mnuOptions(5).Checked
        Case 6
            mnuOptions(6).Checked = Not mnuOptions(6).Checked
            WriteIni str_iniSet, "Options", "EnableIndent", mnuOptions(6).Checked
            loadingCode = True
            If mnuOptions(6).Checked = True Then
                Set_Indent
            Else
                UnSet_Indent
            End If
            loadingCode = False
    End Select
End Sub

Private Sub SelectFont()
On Error GoTo ErrorHandler:

    With CDMain
        .CancelError = True
        .Flags = cdlCFBoth Or cdlCFEffects
        .FontName = RTB1.Font.Name
        .FontSize = RTB1.Font.Size
        .FontItalic = RTB1.Font.Italic
        .FontBold = RTB1.Font.Bold
        .FontStrikethru = RTB1.Font.Strikethrough
        .FontUnderline = RTB1.Font.Underline
        .Color = fontColor
        .ShowFont
    End With
    
    With RTB1
        .Font.Name = CDMain.FontName
        .Font.Size = CDMain.FontSize
        .Font.Italic = CDMain.FontItalic
        .Font.Bold = CDMain.FontBold
        .Font.Strikethrough = CDMain.FontStrikethru
        .Font.Underline = CDMain.FontUnderline
        fontColor = CDMain.Color
    End With
    
    WriteIni str_iniSet, "Fonts", "Name", CDMain.FontName
    WriteIni str_iniSet, "Fonts", "Size", CDMain.FontSize
    WriteIni str_iniSet, "Fonts", "Italic", CDMain.FontItalic
    WriteIni str_iniSet, "Fonts", "Bold", CDMain.FontBold
    WriteIni str_iniSet, "Fonts", "Strikethru", CDMain.FontStrikethru
    WriteIni str_iniSet, "Fonts", "Underline", CDMain.FontUnderline
    WriteIni str_iniSet, "Fonts", "Color", CDMain.Color
        
    ColorTheCode
    
Exit Sub
ErrorHandler:
    If Err <> cdlCancel Then
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub mnuT_Click(Index As Integer)
    Select Case Index
        Case 0
            CleanAllCodes
        Case 1
            Shell "Notepad", vbNormalFocus
    End Select
End Sub

Private Sub RTB1_GotFocus()
    Dim Cntrl As Control
    For Each Cntrl In Me.Controls
        If TypeOf Cntrl Is TextBox Or TypeOf Cntrl Is ComboBox _
        Or TypeOf Cntrl Is lvButtons_H Or TypeOf Cntrl Is McToolBar Then
            Cntrl.TabStop = False
        End If
    Next
End Sub

Private Sub McTB2_Click(ByVal ButtonIndex As Long)
On Error GoTo ErrorHandler:
    Select Case ButtonIndex
        Case 1  'COLOR CODE
            If RTB1.Text = "" Then Exit Sub
            If lvBSummary.Visible = True Then Exit Sub
            Me.MousePointer = vbHourglass
            LockWindowUpdate Me.hwnd
            'ColorTheCode
            Call Color_Code(RTB1, WLOneColor, WLTwoColor, WLThreeColor, _
                CMarkerColor, str_CMarker)
            LockWindowUpdate 0&
            Me.MousePointer = vbDefault
        Case 2 'SHOW CODE SUMMARY
            If RTB1.Text = "" Then Exit Sub
            If lvBSummary.Caption = "Code Details" Or _
                lvBSummary.Caption = "Hiding Code Details" Then Exit Sub
            If only_hiding_summary = True Then
                RTBSummary.Visible = True
                lvBSummary.Caption = "Code Summary"
                only_hiding_summary = False
            Else
                RTB1.Enabled = Not RTB1.Enabled
                txtCodeName.Enabled = Not txtCodeName.Enabled
                RTBSummary.Visible = Not RTBSummary.Visible
                lvBSummary.Visible = Not lvBSummary.Visible
                lvBSummary.Caption = "Code Summary"
                'MsgBox only_hiding_summary
            End If
        Case 3 'SHOW CODE DETAILS
            If RTB1.Text = "" Then Exit Sub
            If RTB1.Enabled = False Then Exit Sub
            
            If only_hiding_details = True Then
                RTBDetails.Visible = True
                lvBSummary.Caption = "Code Details"
                only_hiding_details = False
            Else
                RTBDetails.Visible = Not RTBDetails.Visible
                lvBSummary.Visible = Not lvBSummary.Visible
                If RTBDetails.Visible = True Then
                    RTBDetails.Locked = True
                    lvBSummary.Caption = "Code Details"
                    DisplayCodeInfo
                Else
                    lvBSummary.Caption = "Code Summary"
                End If
                'MsgBox only_hiding_details
            End If
        Case 4 'checkboxes not checkboxes
            tvCode.Checkboxes = Not tvCode.Checkboxes
        Case 5 'closeformat code
            CloseFormatCode
    End Select
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub tvCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tvCode.Checkboxes = False Then Exit Sub
    MiceDown = True
End Sub

Private Sub tvCode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tvCode.Checkboxes = False Then Exit Sub
    MiceDown = False
    If IsNodeCheck = True Then
        If A_NODE.ForeColor = vbBlue Then
            A_NODE.Checked = False
        Else
            A_NODE.Selected = True
            'Debug.Print A_NODE.Key & vbTab & A_NODE.Text
        End If
        IsNodeCheck = False
    End If
        'If Node.ForeColor = vbBlue Then Node.Checked = False
End Sub

Private Sub tvCode_NodeCheck(ByVal Node As MSComctlLib.Node)
   If MiceDown = True Then
        Set A_NODE = Node
        IsNodeCheck = True
    End If
End Sub

Private Sub tvCode_Click()
On Error GoTo ErrorHandler:
    If tvCode.Nodes.Count = 0 Then Exit Sub
    If tvCode.Checkboxes = True Then Exit Sub
    
    Dim strRoot As String
    strRoot = tvCode.SelectedItem.Key
    strRoot = Left$(strRoot, Len(strRoot) - (Len(strRoot) - 6))
    
    If strRoot = "ROOT||" Then
        loadingCode = True
        'MsgBox "Assigned Root ko na to"
        ClearControls
        
        currCodeTypeNo = tvCode.SelectedItem.Key
        currCodeTypeName = tvCode.SelectedItem.Key
        
        'Get RECORD TYPE NO
        currCodeTypeNo = SubstringAfter(currCodeTypeNo, "||CHILD")
        'Get Codename in between
        currCodeTypeName = SubstringAfter(currCodeTypeName, "ROOT||")
        currCodeTypeName = SubstringBefore(currCodeTypeName, "||CHILD")
        
        txtCurrType.Text = " Current Type = " & currCodeTypeName
        txtCodeUnderType.Text = " Code/s Under Type = " & tvCode.SelectedItem.Children
    
        'you can add since root is selected
        addNewCode = True
        txtAddNew.Text = " Adding New = " & addNewCode
        loadingCode = False
        'MsgBox "here"
    Else
        'MsgBox tvCode.SelectedItem.Key
        loadingCode = True  'code is loading
        currCodeName = tvCode.SelectedItem.Key
        currCodeNo = tvCode.SelectedItem.Key
        currCodeTypeNo = tvCode.SelectedItem.Parent.Key
        currCodeTypeName = tvCode.SelectedItem.Parent.Key
        currTVCodePos = tvCode.SelectedItem.Index
        currTVTypePos = tvCode.SelectedItem.Parent.Index
        
        'Write key for remembering
        WriteIni str_iniSet, "Options", "LastKey", currCodeName

        'Get RECORD NAME
        currCodeName = SubstringBefore(currCodeName, "||CHILD")
        'Get COde Number
        currCodeNo = SubstringAfter(currCodeNo, "||CHILD")
        'Get RECORD TYPE NO
        currCodeTypeNo = SubstringAfter(currCodeTypeNo, "||CHILD")
        'Get TYPE NAME
        currCodeTypeName = SubstringAfter(currCodeTypeName, "ROOT||")
        currCodeTypeName = SubstringBefore(currCodeTypeName, "||CHILD")
        'Debug.Print currCodeTypeName

        txtCurrType.Text = " Current Type = " & currCodeTypeName
        txtCodeUnderType.Text = " Code is Under Current Type"
        
        'LOAD CONTENT
        Load_CodeContent
        
        'Display code detailed info when the rtb is visible
        If RTBDetails.Visible = True Then DisplayCodeInfo
        If mnuOptions(6).Checked = True Then Set_Indent
        
        addNewCode = False
        txtAddNew.Text = " Adding New = " & addNewCode
        loadingCode = False 'code is now loaded
        'MsgBox "nakarating sa load"
        
    End If
    
        hasChanged = False
        txtHasChange.Text = " Code Change = " & hasChanged
        
        txtHasChange.Refresh
        txtCurrType.Refresh
        txtCodeUnderType.Refresh
        txtAddNew.Refresh

'Debug.Print tvCode.SelectedItem.Key
'Debug.Print "Code has change = " & hasChanged
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub Set_Indent()
    With RTB1
        LockWindowUpdate .hwnd
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelIndent = 713 '715
        .SelLength = 0
        LockWindowUpdate 0
    End With
End Sub
Private Sub UnSet_Indent()
    With RTB1
        LockWindowUpdate .hwnd
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelIndent = 0 '715
        .SelLength = 0
        LockWindowUpdate 0
    End With
End Sub

'CLICK ADD BUTTON
Private Sub lvBNewCode_Click()
    'add new code
    RTB1.Text = ""
    txtCodeName.Text = ""
    addNewCode = True
    txtAddNew.Text = " Adding New = " & addNewCode
    hasChanged = False
    txtHasChange.Text = " Code Change = " & hasChanged
    'MsgBox hasChanged
End Sub

'CLICK SAVE BUTTON
Private Sub lvBAddSave_Click()
On Error GoTo ErrorHandler:
    SaveAndCheckCodeChange
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
'Set rs_code = Nothing
End Sub

'DELETE CODE
Private Sub lvBDelete_Click()
On Error GoTo ErrorHandler:
If tvCode.Nodes.Count = 0 Then Exit Sub
    'Okay there's no code below it
    If tvCode.SelectedItem.Children = 0 Then
        'see if it is a "TYPE" code
        Dim strRoot As String
        strRoot = tvCode.SelectedItem.Key
        strRoot = Left$(strRoot, Len(strRoot) - (Len(strRoot) - 6))
        'MsgBox strRoot
        If strRoot = "ROOT||" Then
            MsgBox "This item is a 'TYPE' Code." & vbCrLf _
                & "Please use the 'Types' button to delete it.", _
                vbInformation, App.Title
        Else
            
            Call Get_Records(rs_code, cn, _
            "Select TableCode.* From TableCode Order by CodeName ASC")
            
            rs_code.Filter = "CodeName = " & "'" & currCodeName _
                & "' And CodeNo = " & CLng(currCodeNo) & " And TypeNo =" _
                & CLng(currCodeTypeNo)
            
            With rs_code
                'Check if there is no record
                If .RecordCount < 1 Then MsgBox "No Item in the list!", _
                vbExclamation, App.Title: Exit Sub
                'Confirm deletion of record
                Dim ans As Integer
                ans = MsgBox("Are you sure you want to delete the selected code?", _
                vbCritical + vbYesNo, "Confirm Code Delete")
                Screen.MousePointer = vbHourglass
                If ans = vbYes Then
                    'Delete the record
                    cn.Execute "Delete * From TableCode Where TypeNo =" _
                        & CLng(currCodeTypeNo) & " And CodeName ='" _
                        & currCodeName & "' And CodeNo = " & CLng(currCodeNo)
                    
                    'pagbalik neto sa load codenames, it must select the type
                    currTVCodePos = currTVTypePos
                    
                    Load_TypeNames
                    Load_CodeNames
                    ClearControls
                    
                    tvCode.SelectedItem.Expanded = True
                    
                    MsgBox "Code has been successfully deleted.", _
                    vbInformation, "Confirm"
                End If
                ans = 0
                Screen.MousePointer = vbDefault
            End With
        End If
    End If

Set rs_code = Nothing
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
Set rs_code = Nothing
End Sub

Private Sub ClearControls()
    txtCodeName.Text = ""
    RTB1.Text = ""
End Sub

Private Sub RTB1_Change()
    If loadingCode = True Then Exit Sub
    DetectChanged
End Sub

Private Sub txtCodeName_Change()
    DetectChanged
End Sub

Private Sub RTB1_LostFocus()
On Error GoTo ErrorHandler:
    If loadingCode = True Then Exit Sub
    If frmMain.ActiveControl Is txtCodeName Then
    Else
    SaveAndCheckCodeChange
    End If

    Dim Cntrl As Control
    For Each Cntrl In Me.Controls
        If TypeOf Cntrl Is TextBox Or TypeOf Cntrl Is ComboBox _
        Or TypeOf Cntrl Is lvButtons_H Or TypeOf Cntrl Is McToolBar Then
            Cntrl.TabStop = True
        End If
    Next

Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub txtCodeName_LostFocus()
On Error GoTo ErrorHandler:
    If loadingCode = True Then Exit Sub
        If frmMain.ActiveControl Is RTB1 Then
    Else
    SaveAndCheckCodeChange
    End If
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub DetectChanged()
On Error GoTo ErrorHandler:
    If tvCode.Nodes.Count = 0 Then Exit Sub
    
    Dim strRoot As String
    strRoot = tvCode.SelectedItem.Key
    strRoot = Left$(strRoot, Len(strRoot) - (Len(strRoot) - 6))
    If strRoot = "ROOT||" Then
        'Get RECORD TYPE NO
        currCodeTypeNo = tvCode.SelectedItem.Key
        currCodeTypeNo = SubstringAfter(currCodeTypeNo, "||CHILD")
        currTVCodePos = tvCode.SelectedItem.Index
        'show all status
        hasChanged = False
        txtHasChange.Text = " Code Change = " & hasChanged
        txtAddNew.Text = " Adding New = " & addNewCode
    Else
        'THIS DETECTS CHANGE AND ASSIGN VARIABLES IF EXISTING CODE
        If Not loadingCode And Not addNewCode Then
            currTVCodePos = tvCode.SelectedItem.Index
            currCodeName = tvCode.SelectedItem.Key
            currCodeTypeNo = tvCode.SelectedItem.Parent.Key
            'Get RECORD TYPE NO
            currCodeTypeNo = SubstringAfter(currCodeTypeNo, "||CHILD")
            'Get RECORD NAME
            currCodeName = SubstringBefore(currCodeName, "||CHILD")
            'MsgBox "eto nagload lang at false ang add new" 'show all status
            hasChanged = True
            txtHasChange.Text = " Code Change = " & hasChanged
            txtAddNew.Text = " Adding New = " & addNewCode
       End If
    End If
    'Debug.Print "Code has change = " & hasChanged
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub SaveAndCheckCodeChange()
On Error GoTo ErrorHandler:
If tvCode.Nodes.Count = 0 Then Exit Sub

'SAVING EXISTING RECORD
    If hasChanged = True Then
        'exit muna for now pag empty
        If txtCodeName.Text = "" Then Exit Sub
        If RTB1.Text = "" Then Exit Sub
            'MsgBox "Please provide a code name!", vbOKOnly + vbExclamation, App.Title
            'Exit Sub
        'End If
        If Len(txtCodeName.Text) > 255 Then
            MsgBox "Code Name should only be 255 characters maximum", _
            vbOKOnly + vbExclamation, App.Title
            Exit Sub
        End If
        Dim rep As Integer
        rep = MsgBox("Do you want to save the changes" & vbCrLf _
        & "to the CURRENT Code?", _
        vbQuestion + vbYesNo, App.Title)
        If rep = vbYes Then
                                    
            Call Get_Records(rs_code, cn, _
            "Select TableCode.* From TableCode Order by CodeName ASC")
            
            rs_code.Filter = "CodeName = " & "'" & currCodeName & "'" _
                & " And TypeNo = " & CLng(currCodeTypeNo) _
                & " And CodeNo = " & CLng(currCodeNo)
            
            Me.MousePointer = vbHourglass
            LockWindowUpdate Me.hwnd
            
            With rs_code
                .MoveFirst
                .Fields(0) = txtCodeName.Text   'CodeName
                '.Fields(1)                     'TypeNo
                .Fields(2) = RTB1.Text          'CodeContent
                .Fields(3) = RTBSummary.Text    'CodeSummary
                .Update
            End With
            Set rs_code = Nothing
            hasChanged = False
            txtHasChange.Text = " Code Change = " & hasChanged
            txtAddNew.Text = " Adding New = " & addNewCode
            
            'Only do loading typenames and codenames in add mode
            Load_TypeNames
            Load_CodeNames
            
            pbarCaption = "Updating database... Please wait..."
            frmProgressBar.Seconds = 2
            frmProgressBar.Show 1, Me
            
        Else
            'user napindot o kaya pinindot ang NO
            hasChanged = False
            txtHasChange.Text = " Code Change = " & hasChanged
            txtAddNew.Text = " Adding New = " & addNewCode
        End If
        
        rep = 0
        'return to the saved code, i highlight ulit
        tvCode.Nodes(currTVCodePos).Selected = True
        tvCode.Nodes(currTVCodePos).EnsureVisible
        tvCode_Click

'NEW RECORD
    ElseIf addNewCode = True Then
        'see if there is title and code
        If txtCodeName.Text = "" Or RTB1.Text = "" Then Exit Sub
        If Len(txtCodeName.Text) > 255 Then
            MsgBox "Code Name should only be 255 characters maximum", _
            vbOKOnly + vbExclamation, App.Title
            Exit Sub
        End If

        Dim resp As Integer
        resp = MsgBox("Do you want to save this NEW Code?", _
        vbQuestion + vbYesNo, App.Title)
        If resp = vbYes Then
        
        Call Get_Records(rs_code, cn, _
        "Select TableCode.* From TableCode Order by CodeName ASC")
            
        Me.MousePointer = vbHourglass
        LockWindowUpdate Me.hwnd
           
            With rs_code
                .AddNew
                .Fields(0) = txtCodeName.Text       'CodeName
                .Fields(1) = CLng(currCodeTypeNo)   'TypeNo
                .Fields(2) = RTB1.Text              'CodeContent
                .Fields(3) = RTBSummary.Text        'CodeSummary
                .Update
            End With
            resp = 0
            Set rs_code = Nothing
            Load_TypeNames
            Load_CodeNames
        
        pbarCaption = "Updating database... Please wait..."
        frmProgressBar.Seconds = 2
        frmProgressBar.Show 1, Me
        
        Else
            'user canceled
        End If
        
        addNewCode = True
        txtAddNew.Text = " Adding New = " & addNewCode
        ClearControls
                    
    End If
    
    LockWindowUpdate 0&
    Me.MousePointer = vbDefault
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
Set rs_code = Nothing
    hasChanged = False
    txtHasChange.Text = " Code Change = " & hasChanged
    txtAddNew.Text = " Adding New = " & addNewCode
    LockWindowUpdate 0&
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Select Case Index
        Case 0
            FindPic_HideUnHide
        Case 1
            Unload Me
    End Select
End Sub

Private Sub mnuEdit_Click(Index As Integer)
    Select Case Index
        Case 0
            ShowListLang
        Case 1
            ShowListType
        Case 3
            CloseFormatCode
        Case 4
            Call RemoveNumDotVbForums(RTB1)
            CloseFormatCode
    End Select
End Sub

Private Sub CloseFormatCode()
On Error GoTo ErrorHandler:
    Dim strTemp As String
    Me.MousePointer = vbHourglass
    LockWindowUpdate RTB1.hwnd
    
    With RTB1
        strTemp = .Text
        .SetFocus
        .Text = ""
        .SelFontName = ReadIni(str_iniSet, "Fonts", "Name")
        .SelFontSize = ReadIni(str_iniSet, "Fonts", "Size")
        .SelAlignment = rtfLeft
        .Text = strTemp
    End With
    
    ''''''''''''''''''''''''''''''''
    'TESTING FOR NOW
    Call AddRemove_Indent(RTB1)
    ''''''''''''''''''''''''''''''''
    ColorTheCode
    
    LockWindowUpdate 0&
    Me.MousePointer = vbDefault
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub RemoveNumDotVbForums(ByRef RTBox As RichTextBox)
On Error GoTo ErrorHandler:
    Dim strLines() As String, lonLoop As Long
    Dim strRet As String, intPos As Integer
    Dim strNum As String
    
    'Split text into lines.
    strLines = Split(RTBox.Text, vbCrLf)
    'Loop through all lines.
    For lonLoop = 0 To UBound(strLines)
        'Only add lines that are not line numbers to return value.
        If Len(strLines(lonLoop)) > 0 Then
            'Clean the tabs first
            strLines(lonLoop) = Clean_Tabs(strLines(lonLoop))
            intPos = InStr(1, strLines(lonLoop), ".")
            If intPos > 0 Then
                'MsgBox strLines(lonLoop)
                strNum = Left$(strLines(lonLoop), intPos - 1)
                If IsNumeric(strNum) And intPos = Len(strLines(lonLoop)) Then
                    'do nothing
                Else
                    strRet = strRet & strLines(lonLoop) & vbCrLf
                End If
            Else
                strRet = strRet & strLines(lonLoop) & vbCrLf
            End If
        Else
            'do not remove blanklines for now
            strRet = strRet & strLines(lonLoop) & vbCrLf
        End If
    Next lonLoop
    RTBox.Text = strRet
    Erase strLines

Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub AddRemove_Indent(ByRef RTBox As RichTextBox)
On Error GoTo ErrorHandler:
    Dim strLines() As String
    Dim lineCount As Long
    Dim strRes As String
    Dim strOneLine As String
    Dim strRTBText As String
    strRTBText = RTBox.Text
        
    'Split text into lines.
    strLines = Split(strRTBText, vbCrLf)
    'Loop through all lines.
    For lineCount = 0 To UBound(strLines)
        'len of line is greater than 0
        strOneLine = Clean_Tabs(strLines(lineCount))
        If Len(strLines(lineCount)) > 0 Then
            'MsgBox Left$(strRes, 7)
            'Exit Sub
            If UCase$(Left$(strOneLine, 1)) = str_CMarker Then 'THIS IS THE MARKER
                strRes = strRes & strOneLine & vbCrLf
            ElseIf UCase$(Left$(strOneLine, 6)) = UCase$("Const ") Then
                strRes = strRes & strOneLine & vbCrLf
            ElseIf UCase$(Left$(strOneLine, 8)) = UCase$("Private ") Then
                strRes = strRes & strOneLine & vbCrLf
            ElseIf UCase$(Left$(strOneLine, 7)) = UCase$("Public ") Then
                strRes = strRes & strOneLine & vbCrLf
            ElseIf UCase$(Left$(strOneLine, 7)) = UCase$("End Sub") Then
                strRes = strRes & strOneLine & vbCrLf
            ElseIf UCase$(Left$(strOneLine, 8)) = UCase$("End Type") Then
                strRes = strRes & strOneLine & vbCrLf
            ElseIf UCase$(Left$(strOneLine, 8)) = UCase$("End Enum") Then
                strRes = strRes & strOneLine & vbCrLf
            ElseIf UCase$(Left$(strOneLine, 9)) = UCase$("Function ") Then
                strRes = strRes & strOneLine & vbCrLf
            ElseIf UCase$(Left$(strOneLine, 12)) = UCase$("End Function") Then
                strRes = strRes & strOneLine & vbCrLf
            ElseIf UCase$(Left$(strOneLine, 8)) = UCase$("On Error") Then
                strRes = strRes & strOneLine & vbCrLf
            ElseIf UCase$(Left$(strOneLine, 7)) = UCase$("Option ") Then
                strRes = strRes & strOneLine & vbCrLf
            ElseIf Right$(strOneLine, 1) = ":" Then
                strRes = strRes & strOneLine & vbCrLf
            Else
                strRes = strRes & vbTab & strOneLine & vbCrLf
            End If
        Else
            'do not remove blanklines for now
            'MsgBox Len(strOneLine)
            strRes = strRes & strOneLine & vbCrLf
        End If
    Next lineCount
    
    RTBox.Text = TrimSpaceTABCRLF(strRes)
    
    strRTBText = ""
    strOneLine = ""
    strRes = ""
    Erase strLines

Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub
Private Function TrimSpaceTABCRLF(Text As String) As String
    Dim L As Long
    L = Len(Text)
    Do While L
        Select Case AscW(Mid$(Text, L, 1))
        Case 32, 9, 13, 10
            L = L - 1
        Case Else
            Exit Do
        End Select
    Loop
    If L Then TrimSpaceTABCRLF = Left$(Text, L)
End Function

Private Function Clean_Tabs(strSource As String) As String
'convert tabs to spaces first
    strSource = Replace$(strSource, vbTab, " ")
'Find and replace any occurences of multiple spaces
    Do While (InStr(strSource, "  "))
        strSource = Replace$(strSource, "  ", " ")
    Loop
'Remove any leading or training spaces and return result
    Clean_Tabs = Trim$(strSource)
End Function

Private Sub mnuTools_Click(Index As Integer)
On Error GoTo ErrorHandler:
    'close connection
    Set cn = Nothing
    Select Case Index
    
        Case 0
            Dim newBackUpFile As String
            newBackUpFile = Format(Now, "dd-mmmm-yyyy-h-mm-ss-AM/PM") & "_Data.cs"
            If Len(Dir(str_BackUpsFolder & "\" & newBackUpFile)) = 0 Then
                FileCopy str_databaseFile, str_BackUpsFolder & "\" & newBackUpFile
            End If
            MsgBox "Back up file" & vbCrLf _
                & "'" & str_BackUpsFolder & "\" & newBackUpFile & "'" & vbCrLf _
                & "successfully created. Refresh will continue...", vbInformation, "BackUp " & App.Title & " database"
            
        Case 1
            currCodeTypeNo = 0
            currCodeNo = 0
            
            With CDMain
                '.CancelError = True
                .Filter = "Code Data Files(*.cs)|*.cs"
                .DialogTitle = "Choose BackUp File"
                .InitDir = str_BackUpsFolder & "\"
                .ShowOpen
            End With
    
            If CDMain.FileName = vbNullString Then
                'if user did not choose just use old replaced file
                
            Else
                Dim strReplacedFile As String
                Dim oldBackUpFile   As String
                'when user choose, use that file
                oldBackUpFile = CDMain.FileName
                strReplacedFile = "Replaced_" & Format(Now, "dd-mmmm-yyyy-h-mm-ss-AM/PM") & "_Data.cs"
                
                'i think this will always be true with the Data.cs existing
                'first copy the present with replaced name
                'If Dir$(App.Path & "\Data.cs") <> vbNullString Then
                FileCopy App.Path & "\Data.cs", App.Path & "\" & strReplacedFile
                'End If
                'then restore backup code in "backups" folder
                FileCopy oldBackUpFile, App.Path & "\Data.cs"
            
                MsgBox "Restore of" & vbCrLf _
                    & "'" & oldBackUpFile & "'" & vbCrLf _
                    & "successful. Loading will continue...", vbInformation, "Restore " & App.Title & " database"
            End If
        Case 3
            CompactDatabase str_databaseFile, "dani"
            MsgBox App.Title & " Database compacted," & vbCrLf _
                & "Reloading will continue...", vbInformation, App.Title
    End Select
    
    'reconnect and reload
    Call Get_Connected(cn, str_databaseFile, True, "dani")
    
    Me.MousePointer = vbHourglass
    LockWindowUpdate Me.hwnd
    
    Load_LangInMenu
    Load_TypeNames
    Load_CodeNames
    
    LockWindowUpdate 0&
    Me.MousePointer = vbDefault

Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub mnuLang_Click(Index As Integer)
On Error GoTo ErrorHandler:
    Dim X As Long
    For X = 0 To mnuLang.Count - 1
        mnuLang(X).Checked = False
    Next
    mnuLang(Index).Checked = True
    WriteIni str_iniSet, "Settings", "CurrentLanguage", mnuLang(Index).Index
    
    Me.MousePointer = vbHourglass
    LockWindowUpdate Me.hwnd
    
    Load_LangInMenu
    Load_TypeNames
    Load_CodeNames
    
    LockWindowUpdate 0&
    Me.MousePointer = vbDefault
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
    'MsgBox mnuLang(X).Name & " : " & mnuLang(X).Caption
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    Select Case Index
        Case 0
            Dim lResult As Long
            Dim docFile As String
            docFile = App.Path & "\Docs.htm"
            If FileExists(docFile) Then
            lResult = ShellExecute(Me.hwnd, "Open", docFile, _
                    vbNullString, "C:\", SW_SHOWNORMAL)
            End If
        Case 1
            frmAbout.Show 1, Me
        Case 2
            frmEditOwner.Show 1, Me
    End Select
End Sub

Private Sub RTBSummary_Change()
    If loadingCode = True Then Exit Sub
    DetectChanged
End Sub

Private Sub RTBSummary_LostFocus()
    SaveAndCheckCodeChange
End Sub

Private Sub txtCodeName_GotFocus()
    Call Highlight_Focus(txtCodeName)
End Sub

Private Sub RTB1_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
    Me.MousePointer = vbDefault
End Sub

Private Sub tvCode_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
    Me.MousePointer = vbDefault
End Sub

Private Sub chkSearchCode_Click()
    'If Loading Then Exit Sub
    chkMatch.Enabled = chkSearchCode.Value
    chkCurrent.Enabled = chkSearchCode.Value
    If Trim(RTB1.Text) = "" Or chkSearchCode.Value = 0 Then _
        chkCurrent.Enabled = False
    If chkMatch.Enabled = False Then chkMatch.Value = 0
    If chkCurrent.Enabled = False Then chkCurrent.Value = 0
    'txtFind.SetFocus
    txtSearch_Change
End Sub

Private Sub FindPic_HideUnHide()
On Error GoTo ErrorHandler:
    FollowFind
    picSearch.Visible = Not picSearch.Visible
    lvBStatus.Visible = Not lvBStatus.Visible
    If picSearch.Visible = True Then
        txtSearch.SetFocus
        If Len(RTB1.SelText) > 0 Then
            txtSearch.Text = RTB1.SelText
        Else
            If Not ReadIni(str_iniSet, "FindForm", "LastSearch") = vbNullString Then _
            txtSearch.Text = ReadIni(str_iniSet, "FindForm", "LastSearch")
        End If
        If Not ReadIni(str_iniSet, "FindForm", "SearchInCode") = vbNullString Then _
        chkSearchCode.Value = ReadIni(str_iniSet, "FindForm", "SearchInCode")
        If Not ReadIni(str_iniSet, "FindForm", "AutoFind") = vbNullString Then _
        chkAutoFind.Value = ReadIni(str_iniSet, "FindForm", "AutoFind")
    Else
        WriteIni str_iniSet, "FindForm", "LastSearch", txtSearch.Text
        WriteIni str_iniSet, "FindForm", "SearchInCode", chkSearchCode.Value
        WriteIni str_iniSet, "FindForm", "AutoFind", chkAutoFind.Value
    End If
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub
Private Sub lvBHide_Click()
    FindPic_HideUnHide
End Sub

Private Sub lvBReset_Click()
    lvBStatus.Caption = "Ready to Search..."
    lvBSearch.Enabled = True
    lvBSearch.Caption = "Search"
    lvBReset.Enabled = False
    nextNoFind = 0
    newFind = True
End Sub

Private Sub lvBSearch_Click()
On Error GoTo ErrorHandler:
    If txtSearch.Text = "" Then Exit Sub
    If lvBSearch.Caption = "Next" Then
        newFind = False
        nextNoFind = nextNoFind + 1
        SearchCodes
    Else
        SearchCodes
        lvBSearch.Caption = "Next"
    End If
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lvBSearch_Click
End Sub

Private Sub txtSearch_Change()
    'If Loading Then Exit Sub
    lvBSearch.Caption = "Search"
    'cmdReset.Visible = False
    lvBSearch.Enabled = True
    'Me.Caption = " Find '" & txtFind & "'"
    nextNoFind = 0
    newFind = True
    If txtSearch.Text = "" Then RTB1.SelLength = 0: Exit Sub
    If chkAutoFind.Value = 0 Then Exit Sub
    If chkAutoFind.Value <> 0 Then lvBSearch.Caption = "Next"
    SearchCodes
End Sub

Private Sub SearchCodes()
    Dim mCase  As Integer
    Dim CConly As Boolean
    'frmMain.CodeChanged = False
    If chkMatch.Value <> 0 Then mCase = 0 Else mCase = 1
    If chkCurrent.Value <> 0 Then CConly = True Else CConly = False
    If chkSearchCode.Value = 1 Then
        'MsgBox nextNoFind
        FindInCode txtSearch.Text, nextNoFind, mCase, newFind, CConly
    Else
        FindInTV txtSearch.Text, nextNoFind, newFind
    End If
    'frmMain.CodeChanged = False
    'txtFind.SetFocus
    lvBReset.Enabled = True
    If RTB1.Text <> "" And chkSearchCode.Value = 1 Then chkCurrent.Enabled = True
End Sub

Public Sub FindInTV(searchText As String, Skip As Integer, _
Optional newSearch As Boolean = False)
On Error GoTo ErrorHandler:
    Static fNode As Node
    Dim X        As Long
    Dim skp      As Long
    
    If newSearch Then Set fNode = Nothing
    
    'make all nodes not expanded
    LockWindowUpdate Me.hwnd
    For X = 1 To tvCode.Nodes.Count
        tvCode.Nodes(X).Expanded = False
    Next
    ClearControls
    
    'if text to search is empty
    'If searchText = "" Then LockWindowUpdate 0&: Exit Sub
    
    For X = 1 To tvCode.Nodes.Count
        DoEvents: DoEvents
        If InStr(LCase(tvCode.Nodes(X).Text), LCase(searchText)) _
        And Not tvCode.Nodes(X).Parent Is Nothing Then
            If skp = Skip Then
                'cboType.Text = tvCode.Nodes(X).Parent.Text
                tvCode.Nodes(X).Parent.Expanded = True
                tvCode.Nodes(X).Selected = True
                Set fNode = tvCode.Nodes(X)
                'hasChanged = False
                tvCode_Click
                GoTo NodeFound
            Else
                skp = skp + 1
            End If
        End If
    Next
    RTB1.SelLength = 0
    lvBSearch.Enabled = False
    lvBStatus.Caption = "No Records!"
    'lvBReset.Enabled = True
    If Not fNode Is Nothing Then
        If fNode.Parent Is Nothing Then
            '
        Else
    '        cboType.Text = fNode.Parent.Text
            fNode.Parent.Expanded = True
        End If
        fNode.Selected = True
        'hasChanged = False
        tvCode_Click
    End If
    LockWindowUpdate 0&
    Exit Sub
NodeFound:
    lvBStatus.Caption = "Record Found!"
    LockWindowUpdate 0&
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Public Sub FindInCode(searchText As String, _
                    Skip As Integer, _
                    MatchCase As Integer, _
                    Optional newSearch As Boolean = False, _
                    Optional CheckCurrentOnly As Boolean = False)
On Error GoTo ErrorHandler:
                                        
    Dim fLoc       As Long
    Static fNoMore As Long
    Static lastLoc As Long
    Dim cnt        As Integer
    Dim sSpace     As Long
    Dim sch        As String
    Dim skp        As Integer
    'Dim imSearching As Boolean
    
''''''''''''''''''''''
    Dim rs_find_qry As New ADODB.Recordset
    
    'If Not imSearching Then
    Call Get_Records(rs_find_qry, cn, _
    "Select qryCodes.* From qryCodes Order by CodeNo ASC")
'''''''''''''''''''
    
    cnt = 0
    'MsgBox newSearch
    'MsgBox fNoMore
    If newSearch Then fNoMore = -1
    If CheckCurrentOnly Or fNoMore <> -1 Then
        rs_find_qry.Filter = "LangName = '" & CurrLangText & "' And CodeName = '" & currCodeName & "'"
    Else
        rs_find_qry.Filter = "LangName = '" & CurrLangText & "' And CodeContent Like '*" & searchText & "*'"
        'rs_find_qry.Sort = "CodeType, CodeName"
        If rs_find_qry.RecordCount <> 0 Then rs_find_qry.MoveFirst
    End If

    'MsgBox rs_find_qry.Filter
    'MsgBox rs_find_qry.RecordCount

    If rs_find_qry.RecordCount = 0 Then
        MsgBox "'" & searchText & "' does not exist."
        rs_find_qry.Filter = ""
        Exit Sub
    ElseIf rs_find_qry.RecordCount > 1 And CheckCurrentOnly Then
        MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
        rs_find_qry.Filter = ""
        Exit Sub
    End If
        
    'MsgBox Skip
    'MsgBox fNoMore
    If fNoMore <> -1 Then
        fLoc = fNoMore
        skp = Skip
    Else
        fLoc = 1
        skp = 0
    End If
    
    Do While Not rs_find_qry.EOF
Redo:
        cnt = cnt + 1
        'Status " Searching..."
        lvBStatus.Caption = "Searching..."
        If sSpace > Len(sch) Then sSpace = 1
        DoEvents: DoEvents
        If InStr(fLoc, rs_find_qry.Fields(1), searchText, MatchCase) Then
            If skp = Skip Then
                'bCodeChanged False
                If fLoc = 1 Then
                    If Not CheckCurrentOnly Then
                    '''''''''''''''''''''''''''''''''''''''''''
                        'imSearching = True
                    '''''''''''''''''''''''''''''''''''''''''''
                        tvCode.Nodes(rs_find_qry.Fields(0) & "||CHILD" & rs_find_qry.Fields(5)).Parent.Expanded = True
                        tvCode.Nodes(rs_find_qry.Fields(0) & "||CHILD" & rs_find_qry.Fields(5)).Selected = True
                    End If
                    'bCodeChanged False
                    If Not CheckCurrentOnly Then tvCode_Click
                End If
                RTB1.Find searchText, fLoc
                fLoc = InStr(fLoc, rs_find_qry.Fields(1), searchText, MatchCase) + 1
                fNoMore = InStr(fLoc, rs_find_qry.Fields(1), searchText, MatchCase) - 1
                Exit Do
            Else
                fLoc = InStr(fLoc, rs_find_qry.Fields(1), searchText, MatchCase) + 1
                fNoMore = InStr(fLoc, rs_find_qry.Fields(1), searchText, MatchCase) - 1
                If fLoc = 0 Then fLoc = 1
                lastLoc = fLoc
                skp = skp + 1
                If fLoc <> 1 Then GoTo Redo
                MsgBox "Redo"
            End If
        Else
            fLoc = lastLoc
        End If
        rs_find_qry.MoveNext
        fLoc = 1
    Loop
    'MsgBox rs_find_qry.RecordCount
    
    If rs_find_qry.EOF Then
        'Status "'" & searchText & "' does not exist!"
        lvBStatus.Caption = "'" & searchText & "' does not exist!"
        lvBSearch.Enabled = False
        'lvBReset.Enabled = True
        RTB1.Find searchText, lastLoc
    Else
        'Status " Found '" & searchText & "'"
        lvBStatus.Caption = " Found '" & searchText & "'"
        lvBSearch.Enabled = True
        'lvBReset.Enabled = False
    End If
    rs_find_qry.Filter = ""
    
'If Not imSearching Then
Set rs_find_qry = Nothing
'MsgBox "here"

Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
Set rs_find_qry = Nothing
End Sub




Private Sub lvBStatus_MouseDown(Button As Integer, Shift As Integer, _
                                X As Single, Y As Single)
    lngOldX = X
    lngOldY = Y
    blnIsMoving = True
End Sub

Private Sub lvBStatus_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    On Error Resume Next
    Me.MousePointer = vbSizeNS
    If lvBStatus.Top < RTB1.Top Then lvBStatus.Top = RTB1.Top: _
        blnIsMoving = False: Me.MousePointer = vbDefault
    FollowFind
    If lvBStatus.Top > RTB1.Top + RTB1.height - _
    picSearch.height - lvBStatus.height Then _
        lvBStatus.Top = RTB1.Top + RTB1.height - _
        picSearch.height - lvBStatus.height: _
        blnIsMoving = False: Me.MousePointer = vbDefault
    FollowFind
    If blnIsMoving Then
        lvBStatus.Top = lvBStatus.Top - (lngOldY - Y)
        FollowFind
    End If
End Sub

Private Sub lvBStatus_MouseUp(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    On Error Resume Next
    tvCode.SetFocus
    blnIsMoving = False
    If lvBStatus.Top < RTB1.Top Then lvBStatus.Top = RTB1.Top: _
        blnIsMoving = False: Me.MousePointer = vbDefault
    FollowFind
    If lvBStatus.Top > RTB1.Top + RTB1.height - _
    picSearch.height - lvBStatus.height Then _
        lvBStatus.Top = RTB1.Top + RTB1.height - _
        picSearch.height - lvBStatus.height: _
        blnIsMoving = False: Me.MousePointer = vbDefault
    FollowFind
End Sub

Private Sub FollowFind()
    picSearch.Top = lvBStatus.Top + lvBStatus.height
    DoEvents: DoEvents: DoEvents
End Sub

Private Sub CompactDatabase(pstrDatabase As String, _
Optional pstrPassword As String)
On Error GoTo ErrorHandler:
    Dim JRO As JetEngine
    Dim strPassword As String
    Dim strTemp As String
    
    ' Generate temporary file name
    strTemp = Left$(pstrDatabase, InStrRev(pstrDatabase, "\")) & "Compact.cs"
    
    If Len(Dir(strTemp)) <> 0 Then Kill strTemp
    ' Create password string
    If Len(pstrPassword) <> 0 Then strPassword = _
        ";Jet OLEDB:Database Password=" & pstrPassword
    ' Compact database
    Set JRO = New JetEngine
    JRO.CompactDatabase Provider & pstrDatabase & strPassword, _
        Provider & strTemp & JetVersion & strPassword
    Set JRO = Nothing
    ' Copy compacted version over old one
    Kill pstrDatabase
    Name strTemp As pstrDatabase
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub


















''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'THIS is not yet final, just experimenting on it for added features
Private Sub RTB1_KeyDown(KeyCode As Integer, Shift As Integer)
    If RTB1.SelLength > 0 Then
        If KeyCode = 9 And Shift = 1 Then
            Call GoBack_Indent
            'ColorTheCode
            KeyCode = 27
        End If
    End If
End Sub

Private Sub RTB1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Then
        If Not RTB1.SelLength = 0 Then
            Call Tab_Indent
            'ColorTheCode
            KeyAscii = 27
        End If
    End If
End Sub

Private Sub GoBack_Indent()
Dim strORIGINAL As String, i As Integer
Dim strNEW As String, intSTART As Integer
Dim arrLINE() As String, LN As String

    intSTART = RTB1.SelStart
    strORIGINAL = RTB1.SelText
    arrLINE = Split(strORIGINAL, vbCrLf)
    For i = 0 To UBound(arrLINE)
        If Mid(arrLINE(i), 1, 5) = Space(5) Then
            LN = Mid(arrLINE(i), 5, Len(arrLINE(i)))
        Else
            LN = arrLINE(i)
        End If
        If i = UBound(arrLINE) Then
          strNEW = strNEW & LN
        Else
          strNEW = strNEW & LN & vbCrLf
        End If
    Next i
    RTB1.SelText = strNEW
    RTB1.SelStart = intSTART
    RTB1.SelLength = Len(strNEW)
End Sub

Private Sub Tab_Indent()
Dim intSTART As Integer, intLENGTH As Integer, txtNEW As String
Dim txtORIGINAL As String, SPL_TXT() As String
Dim i As Long
    intSTART = RTB1.SelStart
    txtORIGINAL = RTB1.SelText
    txtNEW = ""
    SPL_TXT = Split(RTB1.SelText, vbCrLf)
        For i = 0 To UBound(SPL_TXT)
            If Not i = UBound(SPL_TXT) Then
                txtNEW = txtNEW & Space(5) & SPL_TXT(i) & vbCrLf
            Else
                txtNEW = txtNEW & Space(5) & SPL_TXT(i)
            End If
        Next i
    RTB1.SelText = txtNEW
    RTB1.SelStart = intSTART
    RTB1.SelLength = Len(txtNEW)
End Sub

Private Sub CleanAllCodes()
On Error GoTo ErrorHandler:
Dim X As Long
Dim cntType As Long
Dim cntCode As Long
Dim totCode As Long

Dim strTemp As String
Dim strRoot As String

Call Get_Records(rs_qry_code, cn, _
    "Select qryCodes.* From qryCodes WHERE LangName ='" & CurrLangText & "' ORDER BY CodeName ASC")

totCode = rs_qry_code.RecordCount
frmUpdating.pbar.Min = 0
frmUpdating.pbar.Max = totCode
frmUpdating.Show vbModeless, Me

Me.MousePointer = vbHourglass
LockWindowUpdate Me.hwnd
For X = 1 To tvCode.Nodes.Count
    strRoot = tvCode.Nodes(X).Key
    strRoot = Left$(strRoot, Len(strRoot) - (Len(strRoot) - 6))
    
    If strRoot = "ROOT||" Then
        currCodeTypeNo = tvCode.Nodes(X).Key
        currCodeTypeName = tvCode.Nodes(X).Key
        'Get RECORD TYPE NO
        currCodeTypeNo = SubstringAfter(currCodeTypeNo, "||CHILD")
        'Get Codename in between
        currCodeTypeName = SubstringAfter(currCodeTypeName, "ROOT||")
        currCodeTypeName = SubstringBefore(currCodeTypeName, "||CHILD")

        cntType = cntType + 1
        'Debug.Print "TYPE :" & currCodeTypeNo, currCodeTypeName
    Else
        currCodeName = tvCode.Nodes(X).Key
        currCodeNo = tvCode.Nodes(X).Key
        currCodeTypeNo = tvCode.Nodes(X).Parent.Key
        currCodeTypeName = tvCode.Nodes(X).Parent.Key

        'Get RECORD NAME
        currCodeName = SubstringBefore(currCodeName, "||CHILD")
        'Get COde Number
        currCodeNo = SubstringAfter(currCodeNo, "||CHILD")
        'Get RECORD TYPE NO
        currCodeTypeNo = SubstringAfter(currCodeTypeNo, "||CHILD")
        'Get TYPE NAME
        currCodeTypeName = SubstringAfter(currCodeTypeName, "ROOT||")
        currCodeTypeName = SubstringBefore(currCodeTypeName, "||CHILD")

        cntCode = cntCode + 1
        frmUpdating.pbar.Value = cntCode
        frmUpdating.lblClean.Caption = "Cleaning " & cntCode & " of " & totCode & " Codes..."
        frmUpdating.Refresh
        
        'if i use this without requery is it faster?
        'cn.Execute "UPDATE TableCode SET CodeContent ='" _
            & XXXX & "' WHERE CodeNo='" & currCodeNo & "'"

        rs_qry_code.Requery
        rs_qry_code.Filter = "CodeName = " & "'" & currCodeName & "'" _
                    & " And TypeNo = " & CLng(currCodeTypeNo) _
                    & " And CodeNo = " & CLng(currCodeNo)
        If rs_qry_code.RecordCount > 0 Then rs_qry_code.MoveFirst
        strTemp = rs_qry_code.Fields(1)
        strTemp = TrimSpaceTABCRLF(strTemp)
        With rs_qry_code
            .Fields(1) = strTemp             'CodeContent
            .Update
        'Debug.Print "CODE :" & .Fields(0), currCodeNo, .Fields(7), .Fields(4)
        End With
        strTemp = ""
    End If
Next X
'Debug.Print "Type: " & cntType, "Code: " & cntCode
Set rs_qry_code = Nothing

If cntCode = totCode Then MsgBox "Cleaning of Codes completed!" & vbCrLf _
            & cntCode & " Codes in " & CurrLangText, vbInformation, "Code Cleaner"

'frmUpdating.lblClean.Caption = "Reloading Codes... Please wait ..."
'Load_TypeNames
'Load_CodeNames

LockWindowUpdate 0&
Me.MousePointer = vbDefault
Unload frmUpdating

Exit Sub
ErrorHandler:
MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
    vbOKOnly + vbInformation, App.Title
End Sub

