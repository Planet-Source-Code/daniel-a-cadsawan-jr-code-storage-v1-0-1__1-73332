Attribute VB_Name = "modGlobalPublic"
Option Explicit
     
'CUT, UNDO
Public Const WM_CUT = &H300 'Message Too Cut
Public Const EM_UNDO = &HC7

'LISTVIEW Resize
Public Const LVM_FIRST = &H1000
Public Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Public Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

'shellexecute
Public Const SW_SHOWNORMAL = 1

Public Declare Function SendMessage _
               Lib "user32.dll" _
               Alias "SendMessageA" (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     lParam As Any) As Long
                                     
Public Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Philip Napparan's listview fill
Public Sub FillListView(ByRef sListView As ListView, _
ByRef sRecordsource As ADODB.Recordset, _
ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, _
ByVal with_num As Boolean, ByVal show_first_rec As Boolean)

    Dim X As Variant '|Optional to be declare as variant|
    Dim i As Byte
    On Error Resume Next
    sRecordsource.MoveFirst
    sListView.ListItems.Clear
    Do While Not sRecordsource.EOF
        If with_num = True Then
            Set X = sListView.ListItems.Add(, , sRecordsource.AbsolutePosition, _
            sNumIco, sNumIco)
        Else
            Set X = sListView.ListItems.Add(, , sRecordsource.Fields(0), _
            sNumIco, sNumIco)
        End If
            For i = 1 To sNumOfFields - 1
            'this does not put value to listview when null string, bypass, dani
                'If Not sRecordSource.Fields(Val(i)) = "" Then
                    If show_first_rec = True Then
                        X.SubItems(i) = sRecordsource.Fields(Val(i) - 1)
                    Else
                        X.SubItems(i) = sRecordsource.Fields(Val(i))
                    End If
                'End If
            Next i
        sRecordsource.MoveNext
    Loop
    i = 0
    Set X = Nothing
End Sub

Public Sub LoadColumnHeaders(ByRef sRecordsource As ADODB.Recordset, _
                             ByRef LV As ListView, _
                             ByVal Fieldnum As Integer)
    Dim X As Integer
    With sRecordsource
        For X = 1 To Fieldnum
            LV.ColumnHeaders(X + 1).Text = sRecordsource.Fields(X - 1).Name
        Next X
    End With
End Sub

Public Function ResizeListView(ByRef LV As ListView)
    Dim colCnt As Integer
    Dim intcol As Integer
    Dim i      As Integer
    With LV
        colCnt = .ColumnHeaders.Count
        ReDim myCols(colCnt)
        intcol = 0
        For i = 1 To colCnt
            Call SendMessage(LV.hwnd, LVM_SETCOLUMNWIDTH, intcol, ByVal LVSCW_AUTOSIZE_USEHEADER)
            intcol = intcol + 1
        Next i
    End With
End Function

Public Sub Highlight_Focus(ByRef sText As TextBox)
    With sText
        .SelStart = 0
        .SelLength = Len(sText.Text)
    End With
End Sub

'This is LGS Static's color code
Public Sub Color_Code(ByRef RTBox As RichTextBox, _
    ByVal Color1 As Long, ByVal Color2 As Long, _
    ByVal Color3 As Long, ByVal Color4 As Long, ByVal CMarker As String)
    On Error GoTo ErrorHandler:
    'Dim Color1        As Long
    'Dim Color2        As Long
    'Dim Color3        As Long
    'Dim Color4        As Long
    'Dim CMarker As String
    Dim CMarkerColor    As Long
    
    'Color1 = CLng(ReadIni(ini_CD, CurrentLanguage & "_CodeData", "Color1"))
    'Color2 = CLng(ReadIni(ini_CD, CurrentLanguage & "_CodeData", "Color2"))
    'Color3 = CLng(ReadIni(ini_CD, CurrentLanguage & "_CodeData", "Color3"))
    'Color4 = CLng(ReadIni(ini_CD, CurrentLanguage & "_CodeData", "Color4"))
    'CMarker = ReadIni(ini_CD, CurrentLanguage & "_CodeData", "Comment")
    CMarkerColor = 65793 '= RGB(1,1,1)
    
    'Dim rtb As RichTextBox
    Dim SS  As Long
    
    'If HLP Then
    '    Set rtb = RTBh
    '    tmp = RTBh.SelStart
    'Else
    Dim tmp As String
    'Set rtb = RTBox
    tmp = RTBox.SelStart
    'End If
    
    'If mnuAIndent.Checked = True And PopMenu = False Then
    'AutoIndent frmMain.RTB1
    'End If
    
    SS = RTBox.SelStart
    RTBox.SelLength = Len(RTBox.Text)
    RTBox.SelColor = 0
    RTBox.SelStart = SS
    
    Dim wrd  As Integer
    Dim Z    As Integer
    Dim nxt  As Integer
    Dim eol  As Integer
    Dim tmpL As Integer
    'FIND QUOTED WORDS
    Do Until wrd = -1
        wrd = RTBox.Find(Chr(34), Z)
        If wrd <> -1 Then
            nxt = RTBox.Find(Chr(34), wrd + 1)
            eol = RTBox.Find(vbCrLf, wrd)
            If nxt < eol And nxt > wrd Then
                RTBox.SelStart = wrd
                RTBox.SelLength = nxt - wrd
                tmpL = nxt - wrd
            Else
                RTBox.SelStart = wrd
                RTBox.SelLength = eol - wrd
                tmpL = eol - wrd
            End If
            RTBox.SelColor = CMarkerColor
            Z = wrd + tmpL + 1
        End If
    Loop
    
    wrd = 0
    Z = 0
    eol = 0
    'FIND COMMENT MARKER
    Do Until wrd = -1
        wrd = RTBox.Find(CMarker, Z)
        If wrd <> -1 And RTBox.SelColor <> CMarkerColor Then
            eol = RTBox.Find(vbCrLf, wrd)
            RTBox.SelStart = wrd
            RTBox.SelLength = eol - wrd
            RTBox.SelColor = Color4: Z = wrd + 1
        Else
            Z = wrd + 1
        End If
    Loop
    wrd = 0
    Z = 0
    'COLOR WORDS
    Dim c        As Integer
    Dim WrdLst() As String
    Dim tColor   As Long
    Dim X        As Long
    Dim tmpW     As String
    For c = 1 To 3
        If c = 1 Then WrdLst() = WordList1(): tColor = Color1
        If c = 2 Then WrdLst() = WordList2(): tColor = Color2
        If c = 3 Then WrdLst() = WordList3(): tColor = Color3
        If ArrayHasData(WrdLst) Then
            For X = 0 To UBound(WrdLst)
                Do Until wrd = -1
                    wrd = RTBox.Find(WrdLst(X), Z, , rtfWholeWord)
                    If wrd <> -1 Then
                        If RTBox.SelColor <> Color4 And RTBox.SelColor <> CMarkerColor Then
                            RTBox.SelColor = tColor: Z = wrd + 1
                            tmpW = RTBox.SelText
                            tmpW = WrdLst(X)
                            RTBox.SelText = tmpW
                        Else
                            Z = wrd + 1
                        End If
                    End If
                Loop
                wrd = 0: Z = 0
            Next
            wrd = 0
            Z = 0
        End If
    Next
    
    'If HLP Then
    '    RTBox.SetFocus
    '    RTBox.SelStart = SS
    '    RTBox.SelColor = vbBlack
    'Else
        'If Not FindForm Then RTBox.SetFocus
        'get the cursor to TOP of RTB
        RTBox.SelStart = SS
        RTBox.SelColor = vbBlack
    'End If
    RTBox.SelColor = 0
    
    DoEvents: DoEvents: DoEvents
    'SeperateSections RTBox
    
    Exit Sub
ErrorHandler:
    MsgBox "System Error Number " & Err.Number & " : " & Err.Description, vbOKOnly + vbInformation, App.Title

End Sub


