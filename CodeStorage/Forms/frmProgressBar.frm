VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgressBar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Satus..."
   ClientHeight    =   735
   ClientLeft      =   255
   ClientTop       =   1740
   ClientWidth     =   4080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmProgressBar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timTwo 
      Left            =   1080
      Top             =   180
   End
   Begin VB.Timer timOne 
      Enabled         =   0   'False
      Left            =   540
      Top             =   180
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblLoad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Database... Please wait..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   3030
   End
End
Attribute VB_Name = "frmProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This progress bar does not function in sync
'with the application performing the operation,
'i just use it after and before messages
Option Explicit

Public Seconds As Byte
Dim Counter As Integer

Private Sub Form_Load()
lblLoad.Caption = pbarCaption
timOne.Interval = Seconds * 1000
timTwo.Interval = 1000
Counter = 1
pbar.Min = 0
pbar.Max = Seconds
timOne.Enabled = True
timTwo.Enabled = True
End Sub

Private Sub timOne_Timer()
timTwo.Enabled = False
Unload Me
End Sub

Private Sub timTwo_Timer()
On Error Resume Next
Counter = Counter + 1
pbar.Value = Counter
End Sub
