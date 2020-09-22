VERSION 5.00
Begin VB.Form frmGradient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGradient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   2550
   StartUpPosition =   1  'CenterOwner
   Begin VB.VScrollBar VSColor 
      Height          =   1815
      Index           =   6
      Left            =   2040
      Max             =   255
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.VScrollBar VSColor 
      Height          =   1815
      Index           =   5
      Left            =   1680
      Max             =   255
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.VScrollBar VSColor 
      Height          =   1815
      Index           =   4
      Left            =   1320
      Max             =   255
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.VScrollBar VSColor 
      Height          =   1815
      Index           =   3
      Left            =   840
      Max             =   255
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.VScrollBar VSColor 
      Height          =   1815
      Index           =   2
      Left            =   480
      Max             =   255
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.VScrollBar VSColor 
      Height          =   1815
      Index           =   1
      Left            =   120
      Max             =   255
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin CS.lvButtons_H lvBClose 
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Close"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      cBack           =   -2147483633
   End
   Begin CS.lvButtons_H lvBApply 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Apply"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      cBack           =   -2147483633
   End
   Begin VB.Label lblBtm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BOTTOM:"
      Height          =   195
      Left            =   1440
      TabIndex        =   9
      Top             =   2040
      Width           =   840
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TOP:"
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   2040
      Width           =   360
   End
End
Attribute VB_Name = "frmGradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private gradChanged As Boolean
Private loadingGrad As Boolean

Private Sub Form_Load()
    Me.Caption = App.Title & " Gradient"
    loadingGrad = True
    Read_Gradient
    loadingGrad = False
End Sub

Private Sub Read_Gradient()
    Dim i As Integer
    Dim skey As String
    For i = 1 To 6
        skey = "Color" & CStr(i)
        VSColor(i).Value = ReadIni(str_iniSet, "Gradient", skey)
    Next i
End Sub

Private Sub Save_Gradient()
'save to ini
    Dim i As Integer
    Dim skey As String
    For i = 1 To 6
        skey = "Color" & CStr(i)
        WriteIni str_iniSet, "Gradient", skey, VSColor(i).Value
    Next i
    RTop = VSColor(1).Value
    GTop = VSColor(2).Value
    BTop = VSColor(3).Value
    RBtm = VSColor(4).Value
    GBtm = VSColor(5).Value
    BBtm = VSColor(6).Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
'here is saving
    If gradChanged = True Then
        Dim ans As Integer
        ans = MsgBox("Save Changes to Gradient Background?", _
        vbCritical + vbYesNo, "Confirm Save")
        If ans = vbYes Then
            Save_Gradient
        Else
            'get last values
            frmMain.Read_SaveGradient
            PaintGradient frmMain, RTop, GTop, BTop, _
            RBtm, GBtm, BBtm

            frmMain.Refresh
        End If
    End If
    gradChanged = False
End Sub

Private Sub lvBApply_Click()
    Save_Gradient
    gradChanged = False
End Sub

Private Sub lvBClose_Click()
    Unload Me
End Sub

Private Sub VSColor_Change(Index As Integer)
If loadingGrad = True Then Exit Sub
'No saving yet, just preview
Select Case Index
    Case 1
        RTop = VSColor(1).Value
    Case 2
        GTop = VSColor(2).Value
    Case 3
        BTop = VSColor(3).Value
    Case 4
        RBtm = VSColor(4).Value
    Case 5
        GBtm = VSColor(5).Value
    Case 6
        BBtm = VSColor(6).Value
End Select
    PaintGradient frmMain, RTop, GTop, BTop, _
        RBtm, GBtm, BBtm
    frmMain.Refresh
    gradChanged = True
End Sub
