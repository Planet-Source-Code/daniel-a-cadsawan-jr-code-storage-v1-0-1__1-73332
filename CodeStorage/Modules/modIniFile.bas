Attribute VB_Name = "modIniFile"
Option Explicit

Private Declare Function GetPrivateProfileSection _
                Lib "kernel32" _
                Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, _
                                                   ByVal lpReturnedString As String, _
                                                   ByVal nSize As Long, _
                                                   ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString _
                Lib "kernel32" _
                Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                  ByVal lpKeyName As Any, _
                                                  ByVal lpDefault As String, _
                                                  ByVal lpReturnedString As String, _
                                                  ByVal nSize As Long, _
                                                  ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection _
                Lib "kernel32" _
                Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, _
                                                     ByVal lpString As String, _
                                                     ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString _
                Lib "kernel32" _
                Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                    ByVal lpKeyName As Any, _
                                                    ByVal lpString As Any, _
                                                    ByVal lpFileName As String) As Long

'reads ini string
Public Function ReadIni(FileName As String, Section As String, Key As String) As String
    Dim RetVal As String * 1000, V As Long
    V = GetPrivateProfileString(Section, Key, "", RetVal, 1000, FileName)
    ReadIni = Left(RetVal, V)
End Function
 
'reads ini section
Public Function ReadIniSection(FileName As String, Section As String) As String
    Dim RetVal As String * 1000, V As Long
    V = GetPrivateProfileSection(Section, RetVal, 1000, FileName)
    ReadIniSection = Left(RetVal, V - 1)
End Function
 
'writes ini
Public Sub WriteIni(FileName As String, Section As String, Key As String, Value As String)
    WritePrivateProfileString Section, Key, Value, FileName
End Sub
 
'writes ini section
Public Sub WriteIniSection(FileName As String, Section As String, Value As String)
    WritePrivateProfileSection Section, Value, FileName
End Sub

'sub to delete a particular key inside an ini section.
Public Sub DeleteIniKey(ByVal strSection As String, _
                        ByVal strKeyname As String, _
                        ByVal strfullpath As String)
    Call WritePrivateProfileString(strSection, strKeyname, 0&, strfullpath)
End Sub

'sub to delete an entire ini section.
Public Sub DeleteIniSection(ByVal strSection As String, ByVal strfullpath As String)
    Call WritePrivateProfileString(strSection, 0&, 0&, strfullpath)
End Sub

