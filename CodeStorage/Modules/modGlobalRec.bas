Attribute VB_Name = "modGlobalRec"
Option Explicit

Global owner_name             As String
Global owner_address          As String

'connection
Global cn                     As New ADODB.Connection
Global rs_lang                As New ADODB.Recordset
Global rs_type                As New ADODB.Recordset
Global rs_code                As New ADODB.Recordset
Global rs_qry_code            As New ADODB.Recordset
'

Public Sub Get_Connected(ByRef sConnection As ADODB.Connection, _
ByVal sDataLocation As String, ByVal sHavePassword As Boolean, _
ByVal sPassword As String)
    If sHavePassword = True Then
        sConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
        & sDataLocation & ";Persist Security Info=False;Jet OLEDB:Database Password=" _
        & sPassword
    Else
        sConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
        & sDataLocation & ";Persist Security Info=False"
    End If
End Sub

Public Sub Get_Records(ByRef sRecordset As ADODB.Recordset, _
ByRef sConnection As ADODB.Connection, ByVal sSQL As String)
With sRecordset
    .CursorLocation = adUseClient
    .Open sSQL, sConnection, adOpenKeyset, adLockOptimistic
End With
End Sub

Public Sub Delete_Record(ByRef sCONN As ADODB.Connection, ByVal sTable As String, _
ByVal sField As String, ByVal sString As String, ByVal isnumber As Boolean, _
ByVal snum As Long)
    If isnumber = True Then
        sCONN.Execute "Delete * From " & sTable & " Where " _
        & sField & " =" & snum
    Else
        sCONN.Execute "Delete * From " & sTable & " Where " _
        & sField & " ='" & sString & "'"
    End If
End Sub

