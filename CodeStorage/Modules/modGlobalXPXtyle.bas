Attribute VB_Name = "modGlobalXPXtyle"
Option Explicit

Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function LoadLibrary _
               Lib "kernel32" _
               Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public m_hMod As Long

