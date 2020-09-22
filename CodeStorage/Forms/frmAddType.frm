VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Type"
   ClientHeight    =   2055
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
   Icon            =   "frmAddType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtIconPath 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog CDIcon 
      Left            =   3720
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtType 
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
   Begin CS.lvButtons_H lvBSave 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
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
      Left            =   2280
      TabIndex        =   4
      Top             =   1560
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
   Begin CS.lvButtons_H lvBIcon 
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Add Icon"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
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
   Begin VB.Label lblIcon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IconFile:"
      Height          =   195
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblIconPath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icon File:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   810
   End
   Begin VB.Image imgIcon 
      Height          =   615
      Left            =   3600
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Brief Description:"
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lang Type:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   810
   End
End
Attribute VB_Name = "frmAddType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public add_type As Boolean
Private iconPath As String

Private Sub Form_Load()
On Error GoTo ErrorHandler:
    If add_type = True Then
        Me.Caption = "Add New Type"
        lvBSave.Caption = "Save"
    Else
        Me.Caption = "Edit Existing Type"
        lvBSave.Caption = "Update"
        lvBIcon.Caption = "Edit Icon"
        With rs_type
            'fields 0 = number
            'fields 1 = language
            txtType.Text = .Fields(1)
            txtDesc.Text = .Fields(2)
            txtIconPath.Text = .Fields(3)
            imgIcon.Picture = LoadPicture(str_IconFolder & "\" & .Fields(3))
        End With
    End If
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
    vbOKOnly + vbInformation, App.Title
End Sub

Private Sub Form_Unload(Cancel As Integer)
    add_type = False
End Sub

Private Sub lvBCancel_Click()
    Unload Me
End Sub

Private Sub lvBIcon_Click()
    On Error GoTo ErrorHandler
    With CDIcon
        .CancelError = True
        .Filter = "All Files (*.*)|*.*|GIF Files (*.gif)|*.gif|Bitmap Files (*.bmp)|*.bmp|JPEG Files (*.jpg)|*.jpg"
        .DialogTitle = "Choose 16 x 16 Images Only"
        .InitDir = str_IconFolder & "\"
        .ShowOpen
    End With
    If getImgDim(CDIcon.FileName, ImgSize) = False Then
        Exit Sub 'MsgBox "Not a Valid file!!"
    Else
        If ImgSize.height = 16 And ImgSize.width = 16 Then
            Me.imgIcon.Picture = LoadPicture(CDIcon.FileName)
            txtIconPath.Text = CDIcon.FileTitle
            iconPath = CDIcon.FileName
        Else
            MsgBox "Image Size is not 16x16!"
        End If
    End If
    Exit Sub
ErrorHandler:
    If Err.Number <> cdlCancel Then
        MsgBox Err.Number & " - " & Err.Description, _
        vbOKOnly + vbExclamation, App.Title
    End If
End Sub

Private Sub lvBSave_Click()
On Error GoTo ErrorHandler:
    If Is_Empty(txtType, lblType) Then Exit Sub
    If Is_Empty(txtDesc, lblDesc) Then Exit Sub
    If Is_Empty(txtIconPath, lblIconPath) Then Exit Sub

    'Verify if existing type already, check whether same item
    Dim cnt As Long
    Dim strNewType As String
    Dim strTypeno As String
        For cnt = 1 To frmListType.LVType.ListItems.Count
            strNewType = frmListType.LVType.ListItems(cnt).ListSubItems(2).Text
            strTypeno = frmListType.LVType.ListItems(cnt).Index
            If UCase(txtType.Text) = UCase(strNewType) Then
                'If UCase(frmListType.LVType.SelectedItem.ListSubItems(2).Text) = UCase(rs_type.Fields(1)) Then
                If CLng(strTypeno) = rs_type.AbsolutePosition Then
                    'MsgBox strTypeno & "=" & rs_type.AbsolutePosition
                Else
                    MsgBox "Type '" & txtType.Text & "' already exist." & vbCrLf _
                    & "Please choose another Type Name.", vbOKOnly + vbExclamation, App.Title
                Exit Sub
                End If
            End If
        Next cnt
       'just making sure filename is valid
    'naka lock na yung textbox
    'MsgBox txtIconPath.Text
    If Not FileExists(str_IconFolder & "\" & txtIconPath.Text) Then
        MsgBox "Please choose a valid icon file that exists!", _
        vbOKOnly + vbInformation, App.Title
        Exit Sub
    End If
    With rs_type
        If add_type = True Then
            Dim c_no As Long
            c_no = Get_Next_Num("TableType", "TypeNo", cn)
            .AddNew
            .Fields(0) = c_no
            'MsgBox c_no
        End If
        .Fields(1) = txtType.Text
        .Fields(2) = txtDesc.Text
        .Fields(3) = txtIconPath.Text
        .Fields(4) = CurrLangText
        .Update
    End With
    c_no = 0
    
    If Len(Dir(str_IconFolder & "\" & txtIconPath.Text)) = 0 Then
        FileCopy iconPath, str_IconFolder & "\" & txtIconPath.Text
    End If

'ADDING NEW
If add_type = True Then
    MsgBox "Adding of new type has been successful.", _
    vbInformation, App.Title
    Dim rep As Integer
    rep = MsgBox("Do you want to add another type?", _
    vbQuestion + vbYesNo, App.Title)
    If rep = vbYes Then
        txtType.Text = ""
        txtDesc.Text = ""
        txtIconPath.Text = ""
        txtType.SetFocus
        'rs_type.Requery
        Set rs_type = Nothing
        frmMain.Load_TypeNames
        frmMain.Load_CodeNames
        
        Call Get_Records(rs_type, cn, _
        "Select TableType.* From TableType Where LangName ='" & CurrLangText & _
        "' Order by TypeNo ASC")
        
        frmListType.Type_ReLoad_Rec
    Else
        'rs_type.Requery
        Set rs_type = Nothing
        frmMain.Load_TypeNames
        frmMain.Load_CodeNames
        
        Call Get_Records(rs_type, cn, _
        "Select TableType.* From TableType Where LangName ='" & CurrLangText & _
        "' Order by TypeNo ASC")

        frmListType.Type_ReLoad_Rec
        Unload Me
    End If
    rep = 0

'SAVING CHANGES ONLY
Else
    MsgBox "Changes in type has been successfully saved.", _
    vbInformation, App.Title
    Dim pos As Long
    
    pos = rs_type.AbsolutePosition
    'rs_type.Requery
    Set rs_type = Nothing
    frmMain.Load_TypeNames
    frmMain.Load_CodeNames
    
    Call Get_Records(rs_type, cn, _
    "Select TableType.* From TableType Where LangName ='" & CurrLangText & _
    "' Order by TypeNo ASC")
    
    frmListType.Type_ReLoad_Rec
    rs_type.AbsolutePosition = pos
    frmListType.LVType.ListItems.Item(pos).EnsureVisible
    frmListType.LVType.ListItems.Item(pos).Selected = True
    pos = 0
    Unload Me
End If

Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
    vbOKOnly + vbInformation, App.Title
End Sub

Private Sub txtDesc_Validate(Cancel As Boolean)
    If Len(txtDesc.Text) > 255 Then
        MsgBox "Description should only be 255 characters maximum", _
        vbInformation, App.Title
        Cancel = True
    End If
End Sub

