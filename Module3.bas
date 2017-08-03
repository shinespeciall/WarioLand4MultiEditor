Attribute VB_Name = "Module3"
'////////////////////////////////////////////////////////////This part is for Wizard use as a first design/////////////////////////////////////////////////////////
'########################################## for decompress, make sure to call the function and erase these variaties after get the value in one sub or procedure
Public TextMap() As String
Public layerWidth As Integer, layerHeight As Integer
Public DataByteNumber As Long
Public BGMAPHeader() As String
'##########################################

Public Function DecompressRLE(ByVal DataOffset As String, Optional returnDataLength As Boolean) As String             'UNFINISHED
Dim strdata As String, src As Long, DecompressMap() As String, NeedRearrange As Boolean
src = 1
strdata = ReadFileHex(gbafilepath, DataOffset, Hex(Val("&H" & DataOffset) + 5120))
If Mid$(strdata, 1, 2) = "00" Then             'Althougth this is exclusive for Tile 8*8 Mode, but no problem for directly judging, 16*16-Mode-MAP will never have such small width or height
    ReDim TextMap(32, 32)
    ReDim DecompressMap(32 * 32)
    layerWidth = 32
    layerHeight = 32
ElseIf Mid$(strdata, 1, 2) = "01" Then
    ReDim TextMap(64, 32)
    ReDim DecompressMap(32 * 64)
    NeedRearrange = True
    layerWidth = 64
    layerHeight = 32
ElseIf Mid$(strdata, 1, 2) = "02" Then
    ReDim TextMap(32, 64)
    ReDim DecompressMap(32 * 64)
    layerWidth = 32
    layerHeight = 64
Else
    layerWidth = Val("&H" & Mid$(strdata, 1, 2))
    layerHeight = Val("&H" & Mid$(strdata, 3, 2))
    ReDim TextMap(layerWidth, layerHeight)
    ReDim DecompressMap(layerWidth * layerHeight)
    src = src + 2
End If
Dim ii As Integer, jj As Integer, nn As Integer, NowWord As Integer
For ii = 0 To 1
NowWord = 0
src = src + 2
If Mid$(strdata, src, 2) = "01" Then
    Do
        src = src + 2
        If Mid$(strdata, src, 2) = "00" Then
        Exit Do
        ElseIf Val("&H" & Mid$(strdata, src, 2)) < Val("&H80") Then
            nn = Val("&H" & Mid$(strdata, src, 2))
            For jj = 0 To (nn - 1)
                src = src + 2
                DecompressMap(NowWord) = Mid$(strdata, src, 2) & DecompressMap(NowWord)
                NowWord = NowWord + 1
            Next jj
        ElseIf Val("&H" & Mid$(strdata, src, 2)) >= Val("&H80") Then
            nn = Val("&H" & Mid$(strdata, src, 2)) - Val("&H80")
            If nn <> 0 Then
                src = src + 2
                For jj = 0 To (nn - 1)
                    DecompressMap(NowWord) = Mid$(strdata, src, 2) & DecompressMap(NowWord)
                    NowWord = NowWord + 1
                Next jj
            End If
        End If
    Loop
ElseIf Mid$(strdata, src, 2) = "02" Then
    Do
        src = src + 2
        If Mid$(strdata, src, 2) = "00" And Mid$(strdata, src + 2, 2) = "00" Then
            src = src + 2
            Exit Do
        ElseIf Val("&H" & Mid$(strdata, src, 2) & Mid$(strdata, src + 2, 2)) < 0 Then     'Val("&H" & Mid$(strdata, src, 2) & Mid$(strdata, src + 2, 2)) >= Abs(Val("&H8000"))
            nn = Abs(Val("&H8000")) - Abs(Val("&H" & Mid$(strdata, src, 2) & Mid$(strdata, src + 2, 2)))
            src = src + 4
            If nn > 0 Then
                For jj = 0 To (nn - 1)
                    DecompressMap(NowWord) = Mid$(strdata, src, 2) & DecompressMap(NowWord)
                    NowWord = NowWord + 1
                Next jj
            End If
        ElseIf Val("&H" & Mid$(strdata, src, 2) & Mid$(strdata, src + 2, 2)) > 0 Then   'Val("&H" & Mid$(strdata, src, 2) & Mid$(strdata, src + 2, 2)) < Abs(Abs(Val("&H8000")))
            nn = Val("&H" & Mid$(strdata, src, 2) & Mid$(strdata, src + 2, 2))
            src = src + 2
            If nn > 0 Then
                For jj = 0 To (nn - 1)
                    src = src + 2
                    DecompressMap(NowWord) = Mid$(strdata, src, 2) & DecompressMap(NowWord)
                    NowWord = NowWord + 1
                Next jj
            End If
        End If
    Loop
End If
Next ii
If returnDataLength = True Then DataByteNumber = (src - 1) / 2
For jj = 0 To layerHeight - 1
For ii = 0 To layerWidth - 1
TextMap(ii, jj) = DecompressMap(ii + layerWidth * jj)
Next ii
Next jj

If NeedRearrange = True Then
Dim rearranged1() As String, rearranged2() As String
ReDim rearranged1(32, 32)
ReDim rearranged2(32, 32)
For jj = 0 To 30 Step 2
    For ii = 0 To 31
        rearranged1(ii, jj) = TextMap(ii, jj \ 2)
        rearranged1(ii, jj + 1) = TextMap(ii + 32, jj \ 2)
    Next ii
Next jj
For jj = 1 To 31 Step 2
    For ii = 0 To 31
        rearranged2(ii, jj - 1) = TextMap(ii, jj \ 2 + 16)
        rearranged2(ii, jj) = TextMap(ii + 32, jj \ 2 + 16)
    Next ii
Next jj
For jj = 0 To 31
    For ii = 0 To 31
        TextMap(ii, jj) = rearranged1(ii, jj)
        TextMap(ii + 32, jj) = rearranged2(ii, jj)
    Next ii
Next jj
Erase rearranged1
Erase rearranged2
End If
Erase DecompressMap
DecompressRLE = Hex(layerHeight * layerWidth)
End Function
