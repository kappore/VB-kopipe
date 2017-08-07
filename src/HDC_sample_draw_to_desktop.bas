Attribute VB_Name = "HDC_sample_draw_to_desktop"

'
'  デスクトップにビットマップを描画するサンプル
'
'


'====================================================================='
'                               GDI32                                 '
'====================================================================='
Declare Function SetDIBitsToDevice Lib "GDI32" (ByVal hdc As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal dwWidth As Long, ByVal dwHeight As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal uStartScan As Long, ByVal cScanLines As Long, ByVal lpvBits As Long, ByVal lpbmi As Long, ByVal fuColorUse As Long) As Long
Declare Function SelectObject Lib "GDI32" (ByVal hdc As Long, ByVal hgdiobj As Long) As Long
Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
Declare Function CreateDCA Lib "GDI32" (ByVal lpszDriver As String, ByVal lpszDevice As String, ByVal lpszOutput As String, ByVal lpInitData As Long) As Long
Declare Function Rectangle Lib "GDI32" (ByVal hdc As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect) As Long
Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Declare Function GetStockObject Lib "GDI32" (ByVal fnObject As Long) As Long
Declare Function DeleteDC Lib "GDI32" (ByVal hdc As Long) As Long

Type BITMAPINFO
    biSize  As Long
    biWidth  As Long
    biHeight  As Long
    biPlanes  As Integer
    biBitCount  As Integer
    biCompression  As Long
    biSizeImage  As Long
    biXPelsPerMeter  As Long
    biYPelsPerMeter  As Long
    biClrUsed  As Long
    biClrImportant  As Long
    bmiColors As Long
End Type

Public Const WHITE_BRUSH As Long = 0


'====================================================================='
'                               GDI32                                 '
'====================================================================='
Sub test()

    Dim info As BITMAPINFO
    Dim bits(1600000) As Long
    Dim hdc As Long
    Dim i, j As Long
    
    ' ヘッダー作成
    info.biBitCount = 32
    info.biClrImportant = 0
    info.biClrUsed = 0
    info.biCompression = 0
    info.biHeight = 400
    info.biPlanes = 1
    info.biSize = 40
    info.biSizeImage = 640000
    info.biWidth = 400
    info.biXPelsPerMeter = 0
    info.biYPelsPerMeter = 0
    info.bmiColors = 0
    ' ビットマップ配列作成
    For i = 0 To 399
        For j = 0 To 399
            bits(j + i * 400) = (j Mod 255) * (2 ^ 23) + (j Mod 255) * (2 ^ 16) + (j Mod 255) * (2 ^ 8) + (j Mod 255)
        Next
    Next
    ' 描画
    hdc = CreateDCA("DISPLAY", 0, 0, 0)
    SetDIBitsToDevice hdc, 0, 100, 400, 400, 0, 0, 0, 400, VarPtr(bits(0)), VarPtr(info), 0
    DeleteDC hdc
    
End Sub
