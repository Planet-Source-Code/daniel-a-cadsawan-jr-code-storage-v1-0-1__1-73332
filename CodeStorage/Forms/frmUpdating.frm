VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdating 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Status..."
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pbar 
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
   Begin VB.Label lblClean 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reloading Codes... Please wait ..."
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
      Width           =   2985
   End
End
Attribute VB_Name = "frmUpdating"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

