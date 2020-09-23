Attribute VB_Name = "Declar_mod"

Option Explicit

' API Declares:

Public Const BI_RGB = 0&
Public Const DIB_RGB_COLORS = 0

Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Type BITMAPINFOHEADER '40 bytes
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
Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Public Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Public Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Public Enum MemoryAlloc
    USESYSTEMMEMORY = 0
    USEVIDEOMEMORY = 1
End Enum

' Creates a memory DC
Declare Function CreateCompatibleDC Lib "gdi32" ( _
    ByVal hDC As Long _
    ) As Long
' Places a GDI object into DC, returning the previous one:
Declare Function SelectObject Lib "gdi32" _
    (ByVal hDC As Long, ByVal hObject As Long _
    ) As Long
' Deletes a GDI object:
Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long _
    ) As Long
' Copies Bitmaps from one DC to another, can also perform
' raster operations during the transfer:
Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal dwRop As Long _
    ) As Long
' Structure used to hold bitmap information about Bitmaps
' created using GDI in memory:
Type Bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
' Get information relating to a GDI Object
Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" ( _
    ByVal hObject As Long, _
    ByVal nCount As Long, _
    lpObject As Any _
    ) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Declare Function CreateDIBSection Lib "gdi32" _
    (ByVal hDC As Long, _
    pBitmapInfo As BITMAPINFO, _
    ByVal un As Long, _
    lplpVoid As Long, _
    ByVal handle As Long, _
    ByVal dw As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long

Public ClearRec(1) As D3DRECT
Public bakm
Public RECT As RECT
Public Angle As Single
Public Background As DirectDrawSurface4  'The surface for background

Public Function REC(x, y, xa, ya) As RECT
 REC.bottom = ya
 REC.left = x
 REC.right = xa
 REC.top = y
End Function

