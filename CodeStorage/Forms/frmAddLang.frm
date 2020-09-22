VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddLang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Language"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddLang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtWordList 
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   11
      Text            =   "'"
      Top             =   2280
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog CDLang 
      Left            =   3720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtLang 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox txtDesc 
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin CS.lvButtons_H lvBBrowse 
      Height          =   255
      Index           =   0
      Left            =   3960
      TabIndex        =   3
      Top             =   1200
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "..."
      CapAlign        =   2
      BackStyle       =   2
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
   Begin VB.TextBox txtColor 
      BackColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox txtColor 
      BackColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox txtColor 
      BackColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox txtColor 
      BackColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1200
      Width           =   255
   End
   Begin CS.lvButtons_H lvBSave 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Save"
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
   Begin VB.TextBox txtWordList 
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   10
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txtWordList 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   9
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtWordList 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   8
      Top             =   1200
      Width           =   2415
   End
   Begin CS.lvButtons_H lvBCancel 
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Cancel"
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
   Begin CS.lvButtons_H lvBBrowse 
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   4
      Top             =   1560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "..."
      CapAlign        =   2
      BackStyle       =   2
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
   Begin CS.lvButtons_H lvBBrowse 
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   5
      Top             =   1920
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   ".."
      CapAlign        =   2
      BackStyle       =   2
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
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Brief Description:"
      Height          =   435
      Left            =   120
      TabIndex        =   20
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comment:"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblWordList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Word List 3:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblWordList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Word List 2:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblWordList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Word List 1:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblLang 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Language:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "frmAddLang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public add_lang As Boolean

Private Sub Form_Load()
On Error GoTo ErrorHandler:
    If add_lang = True Then
        Me.Caption = "Add New Language"
        lvBSave.Caption = "Save"
    Else
        Me.Caption = "Edit Existing Language"
        lvBSave.Caption = "Update"
        With rs_lang
            txtLang.Text = .Fields(0)
            txtDesc.Text = .Fields(1)
            txtWordList(0).Text = .Fields(2)
            txtColor(0).BackColor = .Fields(3)
            txtWordList(1).Text = .Fields(4)
            txtColor(1).BackColor = .Fields(5)
            txtWordList(2).Text = .Fields(6)
            txtColor(2).BackColor = .Fields(7)
            txtWordList(3).Text = .Fields(8)
            txtColor(3).BackColor = .Fields(9)
        End With
        Dim Y As Integer
        For Y = 0 To 3
        txtWordList(Y).ForeColor = txtColor(Y).BackColor
        Next Y
    End If
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub Form_Unload(Cancel As Integer)
    add_lang = False
End Sub

Private Sub lvBBrowse_Click(Index As Integer)
On Error GoTo ErrorHandler:
    'MsgBox str_WLFolder
    With CDLang
        .CancelError = True
        .DialogTitle = "Choose WordList Files"
        .InitDir = str_WLFolder & "\"
        .ShowOpen
    End With
    If CDLang.FileName <> "" Then
        If Len(Dir(str_WLFolder & "\" & CDLang.FileTitle)) = 0 Then
            FileCopy CDLang.FileName, str_WLFolder & "\" & CDLang.FileTitle
        End If
        txtWordList(Index).Text = CDLang.FileTitle
    End If
Exit Sub
ErrorHandler:
    If Err.Number <> cdlCancel Then
        MsgBox Err.Number & " - " & Err.Description, _
        vbOKOnly + vbExclamation, App.Title
    End If
    'MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
    vbOKOnly + vbInformation, App.Title
End Sub

Private Sub lvBCancel_Click()
    Unload Me
End Sub

Private Sub lvBSave_Click()
On Error GoTo ErrorHandler:
    If Is_Empty(txtLang, lblLang) Then Exit Sub
    If Is_Empty(txtDesc, lblDesc) Then Exit Sub
    
    With rs_lang
        If add_lang = True Then .AddNew
        .Fields(0) = txtLang.Text
        .Fields(1) = txtDesc.Text
        .Fields(2) = txtWordList(0).Text
        .Fields(3) = txtColor(0).BackColor
        .Fields(4) = txtWordList(1).Text
        .Fields(5) = txtColor(1).BackColor
        .Fields(6) = txtWordList(2).Text
        .Fields(7) = txtColor(2).BackColor
        .Fields(8) = txtWordList(3).Text
        .Fields(9) = txtColor(3).BackColor
        .Update
    End With

'Inform updates
If add_lang = True Then
    MsgBox "Adding of new language has been successful.", _
    vbInformation, App.Title
    Dim rep As Integer
    rep = MsgBox("Do you want to add another language?", _
    vbQuestion + vbYesNo, App.Title)
    If rep = vbYes Then
        txtLang.Text = ""
        txtDesc.Text = ""
        txtLang.SetFocus
        rs_lang.Requery
        frmListLang.Lang_ReLoad_Rec
    Else
        rs_lang.Requery
        frmListLang.Lang_ReLoad_Rec
        Unload Me
    End If
    rep = 0
Else
    MsgBox "Changes in language has been successfully saved.", _
    vbInformation, App.Title
    Dim pos As Long
    
    pos = rs_lang.AbsolutePosition
    rs_lang.Requery
    frmListLang.Lang_ReLoad_Rec
    rs_lang.AbsolutePosition = pos
    frmListLang.LVLang.ListItems.Item(pos).EnsureVisible
    frmListLang.LVLang.ListItems.Item(pos).Selected = True
    pos = 0
    Unload Me
End If

Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
    vbOKOnly + vbInformation, App.Title
End Sub

Private Sub txtColor_Click(Index As Integer)
On Error GoTo ErrorHandler
    With CDLang
        .Flags = cdlCCRGBInit
        .Color = txtColor(Index).BackColor
        .CancelError = True
        .ShowColor
    End With
        txtColor(Index).BackColor = CDLang.Color
        txtWordList(Index).ForeColor = CDLang.Color
Exit Sub
ErrorHandler:
    If Err.Number <> cdlCancel Then
        MsgBox Err.Number & " - " & Err.Description, _
        vbOKOnly + vbExclamation, App.Title
    End If
End Sub

Private Sub txtDesc_Validate(Cancel As Boolean)
    If Len(txtDesc.Text) > 255 Then
        MsgBox "Description should only be 255 characters maximum", _
        vbInformation, App.Title
        Cancel = True
    End If
End Sub
