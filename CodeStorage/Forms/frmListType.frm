VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4320
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
   Icon            =   "frmListType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   4095
      Begin VB.TextBox txtSearchType 
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter text to search:"
         Height          =   435
         Left            =   240
         TabIndex        =   2
         Top             =   225
         Width           =   795
      End
   End
   Begin MSComctlLib.ImageList IL_Types 
      Left            =   3600
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListType.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LVType 
      Height          =   2850
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5027
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
   End
   Begin CS.lvButtons_H lvBAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Add"
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
   Begin CS.lvButtons_H lvBDelete 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Delete"
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
   Begin CS.lvButtons_H lvBClose 
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   3840
      Width           =   975
      _ExtentX        =   1720
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
   Begin CS.lvButtons_H lvBEdit 
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Edit"
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
   Begin CS.lvButtons_H lvBReload 
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Reload"
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
End
Attribute VB_Name = "frmListType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Error GoTo ErrorHandler:
    
    Call Get_Records(rs_type, cn, _
    "Select TableType.* From TableType Where LangName ='" & CurrLangText & _
    "' Order by TypeNo ASC")
    
    Type_ReLoad_Rec

Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub LoadIcons()
On Error GoTo ErrorHandler:
    'Destroy imagelist Icon assigned to lv
    LVType.SmallIcons = Nothing
    LVType.Icons = Nothing
    Dim cnt As Long
    For cnt = IL_Types.ListImages.Count To 2 Step -1
        IL_Types.ListImages.Remove cnt
    Next cnt

    If rs_type.RecordCount > 0 Then
        With rs_type
            .MoveFirst
            For cnt = 1 To .RecordCount
            If FileExists(str_IconFolder & "\" & rs_type.Fields(3)) Then
                IL_Types.ListImages.Add , CurrLangText & rs_type.Fields(1), _
                    LoadPicture(str_IconFolder & "\" & rs_type.Fields(3))
            Else
                'icon file not existing
            End If
            'MsgBox CurrLangText & rs_type.Fields(1)
            .MoveNext
            Next cnt
        End With
    End If
    'MsgBox IL_Types.ListImages.Count
    
    'assign imagelist
    LVType.SmallIcons = IL_Types
    LVType.Icons = IL_Types
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub AssignIcon()
    Dim X As Long
    Dim imgcnt As Long
        For X = 1 To LVType.ListItems.Count
            For imgcnt = 1 To IL_Types.ListImages.Count
                'MsgBox IL_Types.ListImages.Item(imgcnt).Key
                If IL_Types.ListImages.Item(imgcnt).Key = CurrLangText _
                & LVType.ListItems(X).ListSubItems(2).Text Then
                    LVType.ListItems.Item(X).SmallIcon = CurrLangText _
                    & LVType.ListItems(X).ListSubItems(2).Text
                End If
            Next imgcnt
            'MsgBox CurrLangText & LVType.ListItems(x).ListSubItems(2).Text
            'MsgBox LVType.ListItems.Item(x).Text
        Next X
End Sub

Public Sub Type_ReLoad_Rec()
On Error GoTo ErrorHandler:

    LoadIcons
    
    Screen.MousePointer = vbHourglass
    Call LoadColumnHeaders(rs_type, LVType, 5)
    Call FillListView(LVType, rs_type, 6, 1, True, True)
    Call ResizeListView(LVType)
    Me.Caption = rs_type.RecordCount & " Available Type/s"
    Screen.MousePointer = vbDefault
    
    AssignIcon
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs_type = Nothing
    'fire yung reloading ng menu sa main form baka nabago
    'because of delete or update
    'frmMain.Load_TypeNames
    'frmMain.Load_CodeNames
End Sub

Private Sub lvBAdd_Click()
    frmAddType.add_type = True
    frmAddType.Show 1, Me
End Sub

Private Sub lvBClose_Click()
    'Me.Hide
    Unload Me
End Sub

Private Sub lvBDelete_Click()
On Error GoTo ErrorHandler:
With rs_type
    'Check if there is no record
    If .RecordCount < 1 Then MsgBox "No Item in the list!", _
    vbExclamation, App.Title: Exit Sub
    'Confirm deletion of record
    Dim ans As Integer
    'Dim pos As Long
    ans = MsgBox("Are you sure you want to delete the selected type?", _
    vbCritical + vbYesNo, "Confirm Type Delete")
    Screen.MousePointer = vbHourglass
    If ans = vbYes Then
        'Delete the record
        'pos = Val(LVType.SelectedItem)
        Call Delete_Record(cn, "TableType", "TypeNo", _
        "", True, Val(LVType.SelectedItem.ListSubItems(1)))
        '.Requery
        If .RecordCount > 0 Then
            '.AbsolutePosition = pos
            If .EOF Then .MoveFirst
            'pos = .AbsolutePosition
            
            Set rs_type = Nothing
            frmMain.Load_TypeNames
            frmMain.Load_CodeNames
        
            Call Get_Records(rs_type, cn, _
            "Select TableType.* From TableType Where LangName ='" & CurrLangText & _
            "' Order by TypeNo ASC")
            
            Type_ReLoad_Rec
            'LVType.ListItems.Item(pos).EnsureVisible
            'LVType.ListItems.Item(pos).Selected = True
            '.AbsolutePosition = LVType.SelectedItem
        Else
            LVType.ListItems.Clear
        End If
        MsgBox "Type has been successfully deleted.", _
        vbInformation, "Confirm"
    End If
    ans = 0
    'pos = 0
    Screen.MousePointer = vbDefault
End With
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
    Screen.MousePointer = vbDefault
End Sub

Private Sub lvBEdit_Click()
    If rs_type.AbsolutePosition < 0 Then
        MsgBox "Nothing Selected!", vbInformation, App.Title
        Exit Sub
    End If
    frmAddType.Show 1, Me
    frmAddType.add_type = False
End Sub

Private Sub lvBReload_Click()
    rs_type.Requery
    Type_ReLoad_Rec
End Sub

Private Sub LVType_Click()
On Error GoTo ErrorHandler:
    If Not rs_type.RecordCount < 1 Then _
    rs_type.AbsolutePosition = LVType.SelectedItem
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub LVType_DblClick()
    lvBEdit_Click
End Sub
