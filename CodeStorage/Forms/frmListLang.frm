VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListLang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6705
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListLang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6705
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtColor 
      Height          =   285
      Index           =   3
      Left            =   5760
      TabIndex        =   16
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtColor 
      Height          =   285
      Index           =   2
      Left            =   5280
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtColor 
      Height          =   285
      Index           =   1
      Left            =   4800
      TabIndex        =   14
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtColor 
      Height          =   285
      Index           =   0
      Left            =   4320
      TabIndex        =   13
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtWordList 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox txtWordList 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtWordList 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtWordList 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   4095
      Begin VB.TextBox txtSearchLang 
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
   Begin MSComctlLib.ImageList IL_Lang 
      Left            =   4320
      Top             =   3240
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
            Picture         =   "frmListLang.frx":0E42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LVLang 
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
      Icons           =   "IL_Lang"
      SmallIcons      =   "IL_Lang"
      ForeColor       =   8388608
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   4
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
   Begin VB.Label lblWord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comment available is:"
      Height          =   195
      Index           =   3
      Left            =   4320
      TabIndex        =   20
      Top             =   1920
      Width           =   1560
   End
   Begin VB.Label lblWord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WordList is available:"
      Height          =   195
      Index           =   2
      Left            =   4320
      TabIndex        =   19
      Top             =   1320
      Width           =   2235
   End
   Begin VB.Label lblWord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WordList is available:"
      Height          =   195
      Index           =   1
      Left            =   4320
      TabIndex        =   18
      Top             =   720
      Width           =   2235
   End
   Begin VB.Label lblWord 
      BackStyle       =   0  'Transparent
      Caption         =   "WordList is available:"
      Height          =   195
      Index           =   0
      Left            =   4320
      TabIndex        =   17
      Top             =   120
      Width           =   2235
   End
End
Attribute VB_Name = "frmListLang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LangChanged As Boolean

Private Sub Form_Load()
On Error GoTo ErrorHandler:
    Call Get_Records(rs_lang, cn, _
    "Select TableLang.* From TableLang Order by LangName ASC")
    Call LoadColumnHeaders(rs_lang, LVLang, 2)
    Lang_ReLoad_Rec
    Lang_BindControls
    LangChanged = False
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, vbOKOnly + vbInformation, App.Title
End Sub

Public Sub Lang_ReLoad_Rec()
On Error GoTo ErrorHandler:
    Screen.MousePointer = vbHourglass
    Call FillListView(LVLang, rs_lang, 3, 1, True, True)
    Call ResizeListView(LVLang)
    Me.Caption = rs_lang.RecordCount & " Available Language/s"
    Screen.MousePointer = vbDefault
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, vbOKOnly + vbInformation, App.Title
End Sub

Private Sub Lang_BindControls()
    'Set the datasource
    Dim n As Integer
    For n = 0 To 3
        Set txtWordList(n).DataSource = rs_lang
        Set txtColor(n).DataSource = rs_lang
    Next n
    txtWordList(0).DataField = "WLOne"
    txtWordList(1).DataField = "WLTwo"
    txtWordList(2).DataField = "WLThree"
    txtWordList(3).DataField = "Comment"
    
    txtColor(0).DataField = "WLOneColor"
    txtColor(1).DataField = "WLTwoColor"
    txtColor(2).DataField = "WLThreeColor"
    txtColor(3).DataField = "CommentColor"
End Sub

Private Sub Lang_UnbindControls()
    Dim n As Integer
    For n = 0 To 3
    'Set the datasource
    Set txtWordList(n).DataSource = Nothing
    Set txtColor(n).DataSource = Nothing
   'Set the datafield
    txtWordList(n).DataField = ""
    txtColor(n).DataField = ""
    Next n
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Lang_UnbindControls
    Set rs_lang = Nothing
    'fire yung reloading ng menu sa main form baka nabago
    'because of delete or update
    If LangChanged = True Then
    Me.MousePointer = vbHourglass
    LockWindowUpdate Me.hwnd
    
    frmMain.Load_LangInMenu
    frmMain.Load_TypeNames
    frmMain.Load_CodeNames
    
    LangChanged = False
    LockWindowUpdate 0&
    Me.MousePointer = vbDefault
    End If
End Sub

Private Sub lvBAdd_Click()
    LangChanged = True
    frmAddLang.add_lang = True
    frmAddLang.Show 1, Me
End Sub

Private Sub lvBClose_Click()
    Unload Me
End Sub

Private Sub lvBDelete_Click()
On Error GoTo ErrorHandler:
With rs_lang
    'Check if there is no record
    If .RecordCount < 1 Then MsgBox "No Item in the list!", _
    vbExclamation, App.Title: Exit Sub
    'Confirm deletion of record
    Dim ans As Integer
    Dim pos As Long
    ans = MsgBox("Are you sure you want to delete the selected language?", _
    vbCritical + vbYesNo, "Confirm Language Delete")
    Screen.MousePointer = vbHourglass
    If ans = vbYes Then
        'Delete the record
        pos = Val(LVLang.SelectedItem)
        Call Delete_Record(cn, "TableLang", "LangName", _
        LVLang.SelectedItem.ListSubItems(1), False, 0)
        .Requery
        If .RecordCount > 0 Then
            .AbsolutePosition = pos
            If .EOF Then .MoveFirst
            pos = .AbsolutePosition
            Lang_ReLoad_Rec
            LVLang.ListItems.Item(pos).EnsureVisible
            LVLang.ListItems.Item(pos).Selected = True
            .AbsolutePosition = LVLang.SelectedItem
        Else
            LVLang.ListItems.Clear
        End If
        MsgBox "Language has been successfully deleted.", _
        vbInformation, "Confirm"
        LangChanged = True
    End If
    ans = 0
    pos = 0
    Screen.MousePointer = vbDefault
End With
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, vbOKOnly + vbInformation, App.Title
    Screen.MousePointer = vbDefault
End Sub

Private Sub lvBEdit_Click()
    If rs_lang.AbsolutePosition < 0 Then
        MsgBox "Nothing Selected!", vbInformation, App.Title
        Exit Sub
    End If
    LangChanged = True
    frmAddLang.Show 1, Me
    frmAddLang.add_lang = False
End Sub

Private Sub lvBReload_Click()
    rs_lang.Requery
    Lang_ReLoad_Rec
End Sub

Private Sub LVLang_Click()
On Error GoTo ErrorHandler:
    If Not rs_lang.RecordCount < 1 Then _
    rs_lang.AbsolutePosition = LVLang.SelectedItem
    If LVLang.ListItems.Count > 0 Then
        Dim X As Integer
        For X = 0 To 2
            txtWordList(X).ForeColor = CLng(txtColor(X).Text)
            If txtWordList(X).Text = "" Then
                lblWord(X).Caption = "WordList" & CStr(X + 1) & " is NOT Available:"
            Else
                lblWord(X).Caption = "WordList" & CStr(X + 1) & " is Available:"
            End If
        Next X
        txtWordList(3).ForeColor = CLng(txtColor(3).Text)
    End If
Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, _
        vbOKOnly + vbInformation, App.Title
End Sub

Private Sub LVLang_DblClick()
    lvBEdit_Click
End Sub
