VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoCrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

 
Option Explicit
'========
'AutoCrop Version 5
'========
'
'This version not only measures for cropping, it goes ahead and does the cropping
'returning a StdPicture result.
'

Private Const WIN32_NULL As Long = 0
Private Const WIN32_TRUE = 1
Private Const S_OK = 0

Private Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDCDest As Long, _
    ByVal XDest As Long, _
    ByVal YDest As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hDCSrc As Long, _
    ByVal XSrc As Long, _
    ByVal YSrc As Long, _
    ByVal dwRop As RasterOpConstants) As Long

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function CLSIDFromString Lib "ole32" ( _
    ByVal lpsz As Long, _
    ByRef clsid As IID) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Const DIB_RGB_COLORS As Long = 0

Public Enum BiCompressionValues
    BI_RGB = 0 'We're only using this value here.
    BI_BITFIELDS = 3
End Enum

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As BiCompressionValues
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO_32
    bmiHeader As BITMAPINFOHEADER
    'GetDIBits() may return with biCompression = BI_BITFIELDS, so we need 3 DWORDS here:
    blmColors(0 To 2) As Long 'Color masks.
End Type

Private Declare Function GetDIBits Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal hBitmap As Long, _
    ByVal nStartScan As Long, _
    ByVal nNumScans As Long, _
    ByRef Bits As Long, _
    ByRef Bmi As BITMAPINFO_32, _
    ByVal wUsage As Long) As Long

Private Type PICTDESC
    cbSizeOfStruct As Long
    picType As PictureTypeConstants
    hBmp As Long
    hPal As Long
End Type

Private Declare Function OleCreatePictureIndirect Lib "oleaut32" ( _
    ByRef PICTDESC As PICTDESC, _
    ByRef riid As IID, _
    ByVal fOwn As Long, _
    ByRef IPicture As StdPicture) As Long

Private Declare Function SelectObject Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal hgdiObj As Long) As Long

Private IID_IPicture As IID

Private mLeft As Long
Private mTop As Long
Private mWidth As Long
Private mHeight As Long
Private mOrigWidth As Long
Private mOrigHeight As Long

Public Property Get Left() As Long
    Left = mLeft
End Property

Public Property Get Top() As Long
    Top = mTop
End Property

Public Property Get Width() As Long
    Width = mWidth
End Property

Public Property Get Height() As Long
    Height = mHeight
End Property

Public Property Get OrigWidth() As Long
    OrigWidth = mOrigWidth
End Property

Public Property Get OrigHeight() As Long
    OrigHeight = mOrigHeight
End Property
Private Sub LongToRGB(ByVal RGBval As Long, Optional ByRef iRed As Integer, Optional ByVal iGreen As Integer, Optional ByVal iBlue As Integer)
iRed = RGBval And 255
iGreen = (RGBval \ 256 ^ 1) And 255
iBlue = (RGBval \ 256 ^ 2) And 255
End Sub

Public Function CropByBackColor( _
    ByVal hDC As Long, _
    ByVal Picture As StdPicture, _
    ByVal BackColor As Long) As StdPicture
    'If blank, returns the original full-sized.  On error returns Nothing.
    
    Dim Bmi As BITMAPINFO_32
    Dim PixelValues() As Long
    Dim Y As Long
    Dim X As Long
    Dim Right As Long
    Dim hMemDC As Long
    Dim hMemBitmap As Long
    Dim hMemOrigBitmap As Long
    Dim LastDllError As Long
    Dim PICTDESC As PICTDESC
    Dim HRESULT As Long
    Dim vR As Integer
    Dim vG As Integer
    Dim vB As Integer

    
    'Reformat 0BGR as 0RGB:
    BackColor = (BackColor \ &H10000) _
             Or (BackColor And &HFF00&) _
             Or (BackColor And &HFF&) * &H10000
    With Bmi.bmiHeader
        'Retrieve metrics:
        .biSize = LenB(Bmi.bmiHeader)
        .biBitCount = 0 'Don't fetch color table.
        .biCompression = BI_RGB
        If GetDIBits(hDC, _
                     Picture.Handle, _
                     0, _
                     0, _
                     ByVal WIN32_NULL, _
                     Bmi, _
                     DIB_RGB_COLORS) = 0 Then
            Err.Raise &H80048000 Or (Err.LastDllError And &H7FFF&), _
                      TypeName(Me), _
                      "GetDIBits() error " & CStr(Err.LastDllError)
        End If
        mOrigWidth = .biWidth
        mOrigHeight = .biHeight
        'Retrieve pixel data:
        .biBitCount = 32
        .biCompression = BI_RGB
        .biHeight = -.biHeight
        'No padding required since we are using 32-bit (DWORD) pixels:
        ReDim PixelValues(mOrigWidth - 1, mOrigHeight - 1)
        If GetDIBits(hDC, _
                     Picture.Handle, _
                     0, _
                     -.biHeight, _
                     PixelValues(0, 0), _
                     Bmi, _
                     DIB_RGB_COLORS) = 0 Then
            Err.Raise &H80048000 Or (Err.LastDllError And &H7FFF&), _
                      TypeName(Me), _
                      "GetDIBits() error " & CStr(Err.LastDllError)
        End If
    End With
 
    For X = 0 To mOrigWidth - 1
        For Y = 0 To mOrigHeight - 1
            LongToRGB PixelValues(X, Y), vR, vG, vB
            If (vR + vG + vB) / 3 > 50 Then
              GoTo ExitLeft
            End If
'            If PixelValues(X, Y) <> BackColor Then GoTo ExitLeft
        Next Y
    Next X
ExitLeft:
    If X >= mOrigWidth Then
        mLeft = 0
        mTop = 0
        mWidth = mOrigWidth
        mHeight = mOrigHeight
        Exit Function
    Else
        If X = 0 Then mLeft = 0 Else mLeft = X
    End If

    For Y = 0 To mOrigHeight - 1
        For X = mLeft To mOrigWidth - 1
            LongToRGB PixelValues(X, Y), vR, vG, vB
            
            If (vR + vG + vB) / 3 > 80 Then
               GoTo ExitTop
            End If
            'If PixelValues(X, Y) <> BackColor Then GoTo ExitTop
        Next X
    Next Y
ExitTop:
    If Y >= mOrigHeight Or Y = 0 Then mTop = 0 Else mTop = Y

    For X = mOrigWidth - 1 To mLeft Step -1
        For Y = mTop To mOrigHeight - 1
            LongToRGB PixelValues(X, Y), vR, vG, vB
            
            If ((vR + vG + vB) / 3) > 100 Then
               
               DoEvents
               GoTo ExitRight
            End If
            'If PixelValues(X, Y) <> BackColor Then GoTo ExitRight
        Next Y
    Next X
ExitRight:
    If X < mLeft Or X = mOrigWidth - 1 Then Right = mOrigWidth Else Right = X + 1
    mWidth = Right - mLeft

    For Y = mOrigHeight - 1 To mTop Step -1
        For X = mLeft To Right - 1
             LongToRGB PixelValues(X, Y), vR, vG, vB
            If (vR + vG + vB) / 3 > 100 Then
               GoTo ExitBottom
            End If
            'If PixelValues(X, Y) <> BackColor Then GoTo ExitBottom
        Next X
    Next Y
ExitBottom:
    If Y < mTop Or Y = mOrigHeight - 1 Then mHeight = mOrigHeight Else mHeight = Y + 1
    mHeight = mHeight - mTop

    Erase PixelValues
    
    hMemDC = CreateCompatibleDC(hDC)
    hMemBitmap = CreateCompatibleBitmap(hDC, mWidth, mHeight)
    hMemOrigBitmap = SelectObject(hMemDC, hMemBitmap)
    If BitBlt(hMemDC, 0, 0, mWidth, mHeight, hDC, mLeft, mTop, vbSrcCopy) = 0 Then
        LastDllError = Err.LastDllError
        SelectObject hMemDC, hMemOrigBitmap
        DeleteObject hMemBitmap
        DeleteDC hMemDC
        Err.Raise &H80048000 Or (LastDllError And &H7FFF&), _
                  TypeName(Me), _
                  "BitBlt() error " & CStr(LastDllError)
    End If
    SelectObject hMemDC, hMemOrigBitmap
    DeleteDC hMemDC
    
    With PICTDESC
        .cbSizeOfStruct = Len(PICTDESC)
        .picType = vbPicTypeBitmap
        .hBmp = hMemBitmap
    End With
    HRESULT = OleCreatePictureIndirect(PICTDESC, IID_IPicture, WIN32_TRUE, CropByBackColor)
    If HRESULT <> S_OK Then
        DeleteObject hMemBitmap
        Err.Raise HRESULT, _
                  TypeName(Me), _
                  "OleCreatePictureIndirect error " & CStr(HRESULT)
    End If
End Function

Private Sub Class_Initialize()
    CLSIDFromString StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture
End Sub

