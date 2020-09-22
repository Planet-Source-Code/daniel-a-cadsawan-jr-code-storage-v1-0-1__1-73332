Attribute VB_Name = "modImageSize"
Option Explicit

'This function returns image height, width and type
'of JPG, GIF, BMP & PNG formats.

'Type for returning image info
Public Type ImgDimType
    height As Long
    width As Long
End Type

Public ImgSize As ImgDimType

Function getImgDim(ByVal FileName As String, ImgDim As ImgDimType) As Boolean
    'declare vars
    Dim handle       As Integer, isValidImage As Boolean
    Dim byteArr(255) As Byte, I As Integer
    
    'init vars
    isValidImage = False
    ImgDim.height = 0
    ImgDim.width = 0
    
    'open file and get 256 byte chunk
    handle = FreeFile
    On Error GoTo endFunction
    Open FileName For Binary Access Read As #handle
    Get handle, , byteArr
    Close #handle
    'check for  JPG
    If byteArr(0) = &HFF And byteArr(1) = &HD8 Then
        isValidImage = True
    Else
        GoTo checkGIF
    End If
    
    'check for SOF marker: &HFF and &HC0 TO &HCF
    For I = 0 To 255
        If byteArr(I) = &HFF And byteArr(I + 1) >= &HC0 And byteArr(I + 1) <= &HCF Then
            ImgDim.height = byteArr(I + 5) * 256 + byteArr(I + 6)
            ImgDim.width = byteArr(I + 7) * 256 + byteArr(I + 8)
            Exit For
        End If
    Next I
    GoTo endFunction
checkGIF:
    
    If byteArr(0) = &H47 And byteArr(1) = &H49 And byteArr(2) = &H46 And byteArr(3) = &H38 Then
        ImgDim.width = byteArr(7) * 256 + byteArr(6)
        ImgDim.height = byteArr(9) * 256 + byteArr(8)
        isValidImage = True
    Else
        GoTo checkBMP
    End If
    
    GoTo endFunction
checkBMP:
    'check for BMP header
    If byteArr(0) = 66 And byteArr(1) = 77 Then
        isValidImage = True
    End If
    
    'get record type info
    If byteArr(14) = 40 Then
        
        'get width and height of BMP
        ImgDim.width = byteArr(21) * 256 ^ 3 + byteArr(20) * 256 ^ 2 + byteArr(19) * 256 + byteArr(18)
        
        ImgDim.height = byteArr(25) * 256 ^ 3 + byteArr(24) * 256 ^ 2 + byteArr(23) * 256 + byteArr(22)
        
        'another kind of BMP
    ElseIf byteArr(17) = 12 Then
        
        'get width and height of BMP
        ImgDim.width = byteArr(19) * 256 + byteArr(18)
        ImgDim.height = byteArr(21) * 256 + byteArr(20)
        
    End If
    
endFunction:
    getImgDim = isValidImage
End Function


