Attribute VB_Name = "modPicture"
Option Explicit

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Type PicBmp
    Size As Long
    Type As PictureTypeConstants
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

' Representation of 32-bit RGBA color
Type RGBQUAD
    rgbRed As Byte
    rgbGreen As Byte
    rgbBlue As Byte
    rgbReserved As Byte
End Type

Type BITMAPINFOHEADER
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

Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

' VB's array header structure
Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(1) As SAFEARRAYBOUND
End Type

' Some API calls

' Copies memory
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy&)

' Draws text to device
Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC&, ByVal lpStr$, ByVal nCount&, lpRect As RECT, ByVal wFormat&) As Long

' Creates device independent bitmap
Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC&, pBitmapInfo As BITMAPINFO, ByVal un&, lplpVoid&, ByVal handle&, ByVal dw&) As Long
' Gets info about system object
Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject&, ByVal nCount&, lpObject As Any) As Long
' Creates GDI device
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC&) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC&) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject&) As Long
' Associates device with GDI object
Declare Function SelectObject Lib "gdi32" (ByVal hDC&, ByVal hObject&) As Long
' Copies a part of bitmap to another
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC&, ByVal X&, ByVal Y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal dwRop&) As Long

' Creates OLE Picture object used in VB
Declare Function OleCreatePictureIndirect Lib "olepro32" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle&, IPic As IPicture) As Long

' Points to any variable (where VarPtr() doesn't work)
Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

' Creates Picture object of given dimensions and depth
Function CreatePicture(ByVal nWidth&, ByVal nHeight&, ByVal nBPP&) As Picture
    Dim Pic As PicBmp, IID_IDispatch As GUID, BMI As BITMAPINFO
    With BMI.bmiHeader
        .biSize = Len(BMI.bmiHeader)
        .biWidth = nWidth
        .biHeight = nHeight
        .biPlanes = 1
        .biBitCount = nBPP
    End With
    Pic.hBmp = CreateDIBSection(0, BMI, 0, 0, 0, 0)
    With IID_IDispatch
        .Data1 = &H20400: .Data4(0) = &HC0: .Data4(7) = &H46
    End With
    Pic.Size = Len(Pic)
    Pic.Type = vbPicTypeBitmap
    OleCreatePictureIndirect Pic, IID_IDispatch, 1, CreatePicture
    If CreatePicture = 0 Then Set CreatePicture = Nothing
End Function

' Copies one memory bitmap into another with edge extending
Sub CopyImage24(InArray() As Byte, OutArray() As Byte, ByVal InUBound&, Optional ByVal DisablePad As Boolean)
    Dim I&, J&
    J = InUBound + 1
    ' Copy full content of input image into output
    For I = 0 To UBound(InArray, 2)
        CopyMemory OutArray(0, I), InArray(0, I), J
    Next
    ' Fill extended pad bytes with color of edges or not?
    If DisablePad Then Exit Sub
    ' Fill left and right edges
    For J = 0 To UBound(InArray, 2)
        For I = LBound(OutArray) To -3 Step 3
            OutArray(I, J) = OutArray(0, J) ' Blue
            OutArray(I + 1, J) = OutArray(1, J) ' Green
            OutArray(I + 2, J) = OutArray(2, J) ' Red
        Next
        For I = InUBound + 1 To UBound(OutArray) - 2 Step 3
            OutArray(I, J) = OutArray(InUBound - 2, J)
            OutArray(I + 1, J) = OutArray(InUBound - 1, J)
            OutArray(I + 2, J) = OutArray(InUBound, J)
        Next
    Next
    J = UBound(OutArray) - LBound(OutArray) + 1
    InUBound = LBound(OutArray)
    ' Fill top and bottom edges
    For I = LBound(OutArray, 2) To -1
        CopyMemory OutArray(InUBound, I), OutArray(InUBound, 0), J
    Next
    For I = UBound(InArray, 2) + 1 To UBound(OutArray, 2)
        CopyMemory OutArray(InUBound, I), OutArray(InUBound, UBound(InArray, 2)), J
    Next
End Sub
