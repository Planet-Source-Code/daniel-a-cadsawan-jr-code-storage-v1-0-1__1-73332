Attribute VB_Name = "modGlobalFunction"
Option Explicit
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function SafeArrayGetDim _
                Lib "oleaut32.dll" (ByRef saArray() As Any) As Long
'file exists
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" _
    (ByVal pszPath As String) As Long

Public WordList1()  As String 'Blue
Public WordList2()  As String 'Black
Public WordList3()  As String 'Whatever
'Public CodeLocation As String

'Called by LoadWordsAndColors
Public Function LoadWordLists()
    Dim fileNum1 As Integer, filenum2 As Integer, filenum3 As Integer
    fileNum1 = FreeFile
    filenum2 = FreeFile
    filenum3 = FreeFile
    Dim tmp As String
    'MsgBox str_WLOne
    If str_WLOne <> "" Then
        If FileExists(str_WLFolder & "\" & str_WLOne) Then
            Open str_WLFolder & "\" & str_WLOne For Input As fileNum1
            tmp = Input(LOF(fileNum1), fileNum1)
            WordList1 = Split(tmp, vbCrLf)
            Close fileNum1
        End If
    End If
    'MsgBox str_WLTwo
    If str_WLTwo <> "" Then
        If FileExists(str_WLFolder & "\" & str_WLTwo) Then
            Open str_WLFolder & "\" & str_WLTwo For Input As filenum2
            tmp = Input(LOF(filenum2), filenum2)
            WordList2 = Split(tmp, vbCrLf)
            Close filenum2
        End If
    End If
    'MsgBox str_WLThree
    If str_WLThree <> "" Then
        If FileExists(str_WLFolder & "\" & str_WLThree) Then
            Open str_WLFolder & "\" & str_WLThree For Input As filenum3
            tmp = Input(LOF(filenum3), filenum3)
            WordList3 = Split(tmp, vbCrLf)
            Close filenum3
        End If
    End If
End Function

Public Function ArrayHasData(ary() As String) As Boolean
    If (SafeArrayGetDim(ary) > 0) Then
        ArrayHasData = True
    Else
        ArrayHasData = False
    End If
End Function

'get next number in records
Public Function Get_Next_Num(ByVal sTable As String, ByVal sField As String, _
ByRef sCN As ADODB.Connection) As Long
On Error GoTo Err:
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT Max(" & sTable & "." & sField & ") AS [Number] From " & sTable & " ORDER BY Max(" & sTable & "." & sField & ") DESC", sCN, adOpenStatic, adLockOptimistic
    Get_Next_Num = rs.Fields(0) + 1
    
    sTable = ""
    sField = ""
    Set rs = Nothing
Exit Function
Err:
    'Error when incounter a null value
    If Err.Number = 94 Then Get_Next_Num = 1: Resume Next
End Function

'check of textbox is empty
Public Function Is_Empty(ByRef tbox As TextBox, ByRef slbl As Label) As Boolean
If tbox.Text = "" Then
    Is_Empty = True
    MsgBox "Please check empty fields!" & vbCrLf _
    & "> : " & slbl.Caption, vbExclamation, App.Title
    tbox.SetFocus
Else
    Is_Empty = False
End If
End Function

Public Function SubstringAfter(ByVal sString1 As String, ByVal sString2 As String)
    Dim iPos As Integer
    iPos = InStr(sString1, sString2)
    If iPos <> 0 Then
        SubstringAfter = Mid$(sString1, iPos + Len(sString2))
    End If
End Function

Public Function SubstringBefore(ByVal sString1 As String, ByVal sString2 As String)
    Dim iPos As Integer
    iPos = InStr(sString1, sString2)
    If iPos <> 0 Then
        SubstringBefore = Mid$(sString1, 1, iPos - 1)
    End If
End Function

Public Function FileExists(ByVal sFileName As String) As Boolean
    If CBool(PathFileExists(sFileName)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

