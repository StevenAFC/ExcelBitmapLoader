Option Explicit

'BMP Offsets
Private Const WIDTH_OFFSET = 18
Private Const HEIGHT_OFFSET = 22
Private Const HEADERSIZE_OFFSET = 14
Private Const FILESIZE_OFFSET = 2

Function Main(target As Worksheet, filePath As String)
    Dim intFileNum As Integer
    Dim bytTemp As Byte
    Dim bytes() As Byte
    intFileNum = FreeFile
 
    Open filePath For Binary Access Read As intFileNum

    Dim i As Long
    
    i = 0

    Do While Not EOF(intFileNum)
        Get intFileNum, , bytTemp
        
        ReDim Preserve bytes(i + 1)
        
        bytes(i) = bytTemp
        
        i = i + 1
    Loop
    
    Close intFileNum

    Dim headerSize, width, height As Long

    width = BytesToInt(bytes(WIDTH_OFFSET + 0), bytes(WIDTH_OFFSET + 1), bytes(WIDTH_OFFSET + 2), bytes(WIDTH_OFFSET + 3))
    height = BytesToInt(bytes(HEIGHT_OFFSET + 0), bytes(HEIGHT_OFFSET + 1), bytes(HEIGHT_OFFSET + 2), bytes(HEIGHT_OFFSET + 3))
    headerSize = BytesToInt(bytes(HEADERSIZE_OFFSET + 0), bytes(HEADERSIZE_OFFSET + 1), bytes(HEADERSIZE_OFFSET + 2), bytes(HEADERSIZE_OFFSET + 3))

    Dim x, x1, y, y1, dtao, offset As Long
    Dim filler As Integer
    Dim r, g, b As Byte
    x = 0
    y = 0
    y1 = 0
    x1 = 0
    filler = width Mod 4
    
    dtao = headerSize + HEADERSIZE_OFFSET
    
    For y = height To 1 Step -1
        For x = 1 To width

            'Debug.Print "r:" & bytes(x1 * y1 * 3) & "g:" & bytes((x1 * y1 * 3) + 1) & "b:" & bytes((x1 * y1 * 3) + 2)
            offset = headerSize + HEADERSIZE_OFFSET + (y1 * filler) + (x1 * 3)
            r = bytes(offset + 2)
            g = bytes(offset + 1)
            b = bytes(offset + 0)
                        
            target.Cells(y, x).Interior.Color = RGB(r, g, b)
 
            x1 = x1 + 1
        Next x
            y1 = y1 + 1
    Next y
    
End Function

Function BytesToInt(a As Byte, b As Byte, c As Byte, d As Byte) As Double

    BytesToInt = (d * 256 ^ 3) + (c * 256 ^ 2) + (b * 256) + a

End Function
