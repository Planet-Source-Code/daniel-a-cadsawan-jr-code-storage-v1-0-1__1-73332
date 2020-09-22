VERSION 5.00
Begin VB.Form frmEditOwner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Owner"
   ClientHeight    =   1695
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
   Icon            =   "frmEditOwner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtOwner 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox txtAddress 
      Height          =   615
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin CS.lvButtons_H lvBSave 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
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
   Begin CS.lvButtons_H lvBCancel 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1200
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
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblOwner 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Owner:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "frmEditOwner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_owner As New ADODB.Recordset
Dim no_ownerinfo As Boolean

Private Sub Form_Load()
On Error GoTo ErrorHandler:
    Call Get_Records(rs_owner, cn, _
    "Select TableOwner.* From TableOwner")
    If rs_owner.RecordCount > 0 Then
        With rs_owner
            .MoveFirst
            txtOwner.Text = .Fields(0)
            txtAddress.Text = .Fields(1)
        End With
    Else
        no_ownerinfo = True
    End If
    Me.Caption = "Edit " & App.Title & " Owner"
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs_owner = Nothing
End Sub

Private Sub lvBCancel_Click()
    Unload Me
End Sub

Private Sub lvBSave_Click()
On Error GoTo ErrorHandler:
    With rs_owner
        .MoveFirst
        If no_ownerinfo Then .AddNew
        .Fields(0) = txtOwner.Text
        .Fields(1) = txtAddress.Text
        .Update
    End With
    MsgBox "Information successfully saved.", vbInformation, App.Title
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub
