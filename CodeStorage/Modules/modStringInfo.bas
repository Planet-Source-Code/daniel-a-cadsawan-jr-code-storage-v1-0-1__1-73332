Attribute VB_Name = "modStringInfo"
Option Explicit

Public Function Count_Spaces(Text As String) As Long
    Dim b() As Byte, i As Long
    b() = Text
    For i = 0 To UBound(b) Step 2
        ' Consider only even-numbered items.
        ' Save time and code using the function name as a local variable.
        If b(i) = 32 Then Count_Spaces = Count_Spaces + 1
    Next
End Function

Public Function Vowel_Count(ByVal strExpression As String) As Long
    Dim sVowels() As Variant
    Dim iResult   As Integer, i As Integer
    Dim lLength   As Long
    sVowels() = Array("a", "e", "i", "o", "u", "A", "E", "I", "O", "U")
    lLength = Len(strExpression)

    For lLength = 1 To lLength
        For i = LBound(sVowels) To UBound(sVowels)
            'Check whether the character is a vowel
            If Mid$(strExpression, lLength, 1) = sVowels(i) Then
                'if the character is a vowel
                iResult = iResult + 1
            End If
        Next i
    Next lLength
    Vowel_Count = iResult
End Function

Public Function Word_Count(ByVal StringToCount As String) As String
    Dim WordsToCount As Variant
    Dim FinishedText As String
    WordsToCount = Split(StringToCount, Chr(32))

    FinishedText = UBound(WordsToCount) - LBound(WordsToCount)

    Word_Count = FinishedText
End Function

Public Function Count_Lines(ByVal strTxt As String) As Integer
    Dim strLines() As String

    'Open the text file
    On Error GoTo Count_Lines_Error

    'store the lines
    strLines = Split(strTxt, vbCrLf)

    Count_Lines = UBound(strLines) - LBound(strLines)

    On Error GoTo 0
    Exit Function

Count_Lines_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure " & _
        "Count_Lines of Module Module1"
End Function

Public Sub FilterChars(Source_In As String, Digits_Out As String, _
    Letters_Out As String, OtherChars_Out As String)
  Dim lngCountDigits As Long
  Dim lngCountLetters As Long
  Dim lngCountOtherChars As Long
  Dim strChar As String
  Dim i As Long
  
  If Len(Source_In) > 0 Then
    ' Create buffer space for the filtered strings
    Digits_Out = Space(Len(Source_In))
    Letters_Out = Space(Len(Source_In))
    OtherChars_Out = Space(Len(Source_In))
    
    ' For each character in the source string, copy it into
    ' one of the filtered strings based on its character code
    For i = 1 To Len(Source_In)
      strChar = Mid$(Source_In, i, 1)
      Select Case Asc(strChar)
      Case 48 To 57 ' Digits: 0 to 9
        lngCountDigits = lngCountDigits + 1
        Mid(Digits_Out, lngCountDigits, 1) = strChar
      Case 65 To 90, 97 To 122 ' Letters: A to Z, a to z
        lngCountLetters = lngCountLetters + 1
        Mid(Letters_Out, lngCountLetters, 1) = strChar
      Case Else
        lngCountOtherChars = lngCountOtherChars + 1
        Mid(OtherChars_Out, lngCountOtherChars, 1) = strChar
      End Select
    Next i
    
    ' Trim the filtered strings to appropriate length
    Digits_Out = Left$(Digits_Out, lngCountDigits)
    Letters_Out = Left$(Letters_Out, lngCountLetters)
    OtherChars_Out = Left$(OtherChars_Out, lngCountOtherChars)
  End If
End Sub


