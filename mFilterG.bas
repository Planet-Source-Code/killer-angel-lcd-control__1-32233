Attribute VB_Name = "mFilterG"
Option Explicit

'API
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Type BITMAPINFOHEADER   '40 bytes
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type

Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBQUAD
End Type

'Private Variables
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0&
Private WorkFilterG As Boolean

Private iDATA() As Byte           'holds bitmap data
Private bDATA() As Byte           'holds bitmap backup
Private PicInfo As BITMAP         'bitmap info structure
Private DIBInfo As BITMAPINFO     'Device Ind. Bitmap info structure
Private mProgress As Long         '% filter progress
Private Speed(0 To 765) As Long   'Speed up values

'-------------------------------------------FILTERS
Public Sub BurkeBW(ByVal Pic As Long, ByVal Factor As Long, ByRef pProgress As Long)
    Dim X As Long, Y As Long
    Dim R As Long, G As Long, B As Long
    Dim Erro As Long
    Dim VecErro() As Long
    Dim nCol As Long, mCol As Long
    Dim PartErr(1 To 7, -255 To 255) As Long
    Dim hdcNew As Long
    Dim oldhand As Long
    Dim ret As Long
    Dim BytesPerScanLine As Long
    Dim PadBytesPerScanLine As Long

    If WorkFilterG = True Then Exit Sub
    WorkFilterG = True
    On Error GoTo FilterError:
    'get data buffer
    Call GetObject(Pic, Len(PicInfo), PicInfo)
    hdcNew = CreateCompatibleDC(0&)
    oldhand = SelectObject(hdcNew, Pic)
    With DIBInfo.bmiHeader
        .biSize = 40
        .biWidth = PicInfo.bmWidth
        .biHeight = -PicInfo.bmHeight     'bottom up scan line is now inverted
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        BytesPerScanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
        PadBytesPerScanLine = BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
        .biSizeImage = BytesPerScanLine * Abs(.biHeight)
    End With
    'redimension  (BGR+pad,x,y)
    ReDim iDATA(1 To 4, 1 To PicInfo.bmWidth, 1 To PicInfo.bmHeight) As Byte
    ReDim bDATA(1 To 4, 1 To PicInfo.bmWidth, 1 To PicInfo.bmHeight) As Byte
    'get bytes
    ret = GetDIBits(hdcNew, Pic, 0, PicInfo.bmHeight, iDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)
    ret = GetDIBits(hdcNew, Pic, 0, PicInfo.bmHeight, bDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)
    'do it
    For X = 0 To 765
        Speed(X) = X \ 3
    Next X
    For X = -255 To 255
        PartErr(1, X) = (8 * X) \ 32
        PartErr(2, X) = (4 * X) \ 32
        PartErr(3, X) = (2 * X) \ 32
        PartErr(4, X) = (4 * X) \ 32
        PartErr(5, X) = (8 * X) \ 32
        PartErr(6, X) = (4 * X) \ 32
        PartErr(7, X) = (2 * X) \ 32
    Next X
    Erro = 0
    ReDim VecErro(1 To 2, 1 To PicInfo.bmWidth) As Long
    For X = 1 To PicInfo.bmWidth
        VecErro(1, X) = 0
        VecErro(2, X) = 0
    Next X
    pProgress = 3
    DoEvents
    For Y = 1 To PicInfo.bmHeight
        For X = 1 To PicInfo.bmWidth
            B = CLng(bDATA(1, X, Y))
            G = CLng(bDATA(2, X, Y))
            R = CLng(bDATA(3, X, Y))
            B = Speed(R + G + B)
            mCol = mCol + B
            nCol = nCol + 1
        Next X
    Next Y
    mCol = mCol \ nCol
    pProgress = 10
    DoEvents
    For Y = 1 To PicInfo.bmHeight
        For X = 1 To PicInfo.bmWidth
            B = CLng(bDATA(1, X, Y))
            G = CLng(bDATA(2, X, Y))
            R = CLng(bDATA(3, X, Y))
            B = Speed(R + G + B)
            B = B + (VecErro(1, X) * 10) \ Factor
            If B < 0 Then B = 0
            If B > 255 Then B = 255
            If B < mCol Then
                nCol = 0
                iDATA(1, X, Y) = nCol
                iDATA(2, X, Y) = nCol
                iDATA(3, X, Y) = nCol
            Else
                nCol = 255
                iDATA(1, X, Y) = 255 '238
                iDATA(2, X, Y) = 192 '204
                iDATA(3, X, Y) = 160 '170
            End If
            Erro = (B - nCol)
            If X < PicInfo.bmWidth Then VecErro(1, X + 1) = VecErro(1, X + 1) + PartErr(1, Erro)
            If X < PicInfo.bmWidth - 1 Then VecErro(1, X + 2) = VecErro(1, X + 2) + PartErr(2, Erro)
            If Y < PicInfo.bmHeight Then
                If X > 2 Then VecErro(2, X - 2) = VecErro(2, X - 2) + PartErr(3, Erro)
                If X > 1 Then VecErro(2, X - 1) = VecErro(2, X - 1) + PartErr(4, Erro)
                VecErro(2, X) = VecErro(2, X) + PartErr(5, Erro)
                If X < PicInfo.bmWidth Then VecErro(2, X + 1) = VecErro(2, X + 1) + PartErr(6, Erro)
                If X < PicInfo.bmWidth - 1 Then VecErro(2, X + 2) = VecErro(2, X + 2) + PartErr(7, Erro)
            End If
        Next X
        For X = 1 To PicInfo.bmWidth
            VecErro(1, X) = VecErro(2, X)
            VecErro(2, X) = 0
        Next X
        mProgress = 10 + (Y * 90) \ PicInfo.bmHeight
        pProgress = mProgress
        DoEvents
    Next Y
    pProgress = 100
    DoEvents
    'copy bytes to device
    ret = SetDIBits(hdcNew, Pic, 0, PicInfo.bmHeight, iDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)
    SelectObject hdcNew, oldhand
    DeleteDC hdcNew
    ReDim iDATA(1 To 4, 1 To 2, 1 To 2) As Byte
    ReDim bDATA(1 To 4, 1 To 2, 1 To 2) As Byte
    WorkFilterG = False
    Exit Sub
FilterError:
    MsgBox "Filter Error"
    WorkFilterG = False
End Sub
