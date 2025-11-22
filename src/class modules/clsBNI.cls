VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBNI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const SRCCOPY = &HCC0020
Private Const BI_RGB = 0&
Private Const CBM_INIT = &H4
Private Const DIB_RGB_COLORS = 0

'Types
Private Type ScTw
    TwipWidth  As Long
    TwipHeight As Long
End Type

Private Type RGBTriple
    tRED   As Byte
    tGREEN As Byte
    tBLUE  As Byte
End Type

Private Type RGBQUAD
    rgbBlue     As Byte
    rgbGreen    As Byte
    rgbRed      As Byte
    rgbReserved As Byte
End Type

Private Type tgaheader    'declare the header
  info       As Byte      'length of image information block
  colortyp   As Byte      'DAC table or BGR format
  imagetyp   As Byte      'compressed or uncompressed
  origin     As Integer   'first entry in the DAC table
  colnumber  As Integer   'number of colors in the DAC table
  entrybits  As Byte      'entry size in the DAC table
  xvalue     As Integer   'x co-ordinate lower left corner
  yvalue     As Integer   'y co-ordinate lower left corner
  widt       As Integer   'image width
  Height     As Integer   'image height
  pixelsize  As Byte      'number of bits per pixel
  descriptor As Byte      'image descriptor
End Type

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAPINFO_24
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type


'Functions
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateDIBitmap_24 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_24, ByVal wUsage As Long) As Long

'Variables
Private Orientation As Integer
Private BitmapData() As Byte
Private Header As tgaheader         'define the header
Private i As Long
Private nWidth As Long
Private nHeight As Long
Private Scales As Integer
Private Automatic As Boolean
Private TW As ScTw
Private bm24 As BITMAPINFO_24
Private hBmp As Long

Public Function LoadTGA(ByVal PositionStart As Long, ByVal FileName As String) As StdPicture
On Error Resume Next
    Dim nFreefile As Integer
    
    nFreefile = FreeFile

    Open FileName For Binary Lock Write As #nFreefile
        Seek #nFreefile, 1
        Get #nFreefile, PositionStart, Header
    Close #nFreefile
    
    TW = PixelToTwips(CLng(Header.widt), CLng(Header.Height))
    
    Select Case Header.pixelsize
        Case 24
            Read24bit PositionStart, FileName
    End Select
End Function

Private Sub Read24bit(StartPos As Long, FileName As String)
Dim nFreefile As Integer
Dim nLineSize As Long
Dim Data() As Byte
Dim quad() As RGBQUAD
    
    nFreefile = FreeFile

    Open FileName For Binary Lock Write As #nFreefile
        Seek #nFreefile, 1
        Get #nFreefile, StartPos, Header
    
        With Header
            nWidth = .widt - .xvalue ' + 1
            nHeight = .Height - .yvalue ' + 1
            nLineSize = .widt * .pixelsize
        End With
        
        ReDim Data(LOF(nFreefile) - Len(Header))
    
        Get #nFreefile, , Data()
        Orientation = GetByte(Header.descriptor, 3)
    Close #nFreefile
    
    If Header.imagetyp = 9 Or Header.imagetyp = 10 Then
        Data = DecompressTGA(Data(), 24, nHeight, nWidth)
    End If
    
'Sort from planes into a bitmap
'==================================================
Dim x As Long, y As Long, nStartPos As Long
    
    ReDim quad(UBound(Data) / 3)
    
    For x = 0 To UBound(Data) / 3 - 1
        With quad(x)
            .rgbBlue = Data(x * 3)
            .rgbGreen = Data(x * 3 + 1)
            .rgbRed = Data(x * 3 + 2)
        End With
    Next x
        
    ReDim BitmapData(UBound(quad) * 4 + 4)
    CopyMemory BitmapData(0), quad(0), UBound(BitmapData)
    
    CreateBitmap_24 BitmapData, nWidth, nHeight, Orientation
End Sub

Private Function DecompressTGA(RLEStream() As Byte, Bits As Long, Höhe As Long, Breite As Long) As Byte()
Dim InitSize As Long, n As Long, B As Long, l As Long, fertig As Long, z As Long, Länge As Long, Byteanzahl As Long
Dim Temp() As Byte
Dim K As Boolean
    
    Byteanzahl = Bits / 8
    InitSize = CLng(Höhe * Breite * Byteanzahl)
    
    ReDim Temp(0 To InitSize)
    
    Do While fertig < InitSize
        
        If l > UBound(RLEStream) Then
            GoTo Ende
        End If
        
        z = 0
        If RLEStream(l) > 127 Then
            n = RLEStream(l) - 127
            
            For B = 0 To n - 1
                Call CopyMemory(Temp(fertig), RLEStream(l + 1), Byteanzahl)
                fertig = fertig + Byteanzahl
            Next B
        
            K = True
        Else
            n = RLEStream(l) + 1
            Länge = n * Byteanzahl
    
            CopyMemory Temp(fertig), RLEStream(l + 1), Länge
                
            K = False
            z = z + Länge
    
            fertig = fertig + z
        End If
            If K = True Then
                l = Byteanzahl + 1 + l
                n = z + (n * Byteanzahl) + 1
            Else
                l = (n * Byteanzahl + 1) + l
                n = z + n
            End If
    Loop
Ende:
    DecompressTGA = Temp
End Function

Private Sub MakeBitmap(ImageArray() As Byte, Lines As Long, BytesLine As Long)
Dim Übergabe() As Byte
Dim GrößeBMP As Long
Dim Zugabe As Integer
Dim Standort As Long
Dim nBitmapX As Long
Dim l As Long

    If (BytesLine) Mod 4 = 0 Then
        'wenn duch 4 teilbar dann ok
        nBitmapX = BytesLine - 1
    Else
        'ansonsten ändern
        nBitmapX = (BytesLine \ 4) * 4 + 3
    End If
    
    If nBitmapX + 1 <> BytesLine Then
        Zugabe = nBitmapX - BytesLine + 1
    End If
    
    GrößeBMP = Lines * (nBitmapX + 1) - 1
    
    ReDim Übergabe(UBound(ImageArray))

    CopyMemory Übergabe(0), ImageArray(0), UBound(ImageArray) + 1
    
    ReDim ImageArray(GrößeBMP)
    
    For i = 0 To BytesLine * Lines - BytesLine Step BytesLine
        CopyMemory ImageArray(Standort), Übergabe(i), BytesLine
        Standort = Standort + nBitmapX + 1
    Next i

End Sub


Private Function GetByte(Bytes As Byte, Position As Long) As Integer
    GetByte = 0
    Select Case Position
        Case 1
            If Bytes And 128 Then GetByte = 1
        Case 2
            If Bytes And 64 Then GetByte = 1
        Case 3
            If Bytes And 32 Then GetByte = 1
        Case 4
            If Bytes And 16 Then GetByte = 1
        Case 5
            If Bytes And 8 Then GetByte = 1
        Case 6
            If Bytes And 4 Then GetByte = 1
        Case 7
            If Bytes And 2 Then GetByte = 1
        Case 8
            If Bytes And 1 Then GetByte = 1
    End Select
End Function

Public Function DrawTGA(PicObj As Object)
    DrawBitmap nWidth + 4, nHeight + 4, PicObj, Automatic
End Function

Public Property Get TGAWidth() As Long
    Select Case Scales
        Case 0
            TGAWidth = nWidth + 200
        Case 1
            TGAWidth = TW.TwipWidth + 96
    End Select
End Property

Public Property Get TGAHeight() As Long
    Select Case Scales
        Case 0
            TGAHeight = nHeight
        Case 1
            TGAHeight = TW.TwipHeight
    End Select
End Property

Public Property Get Compressed() As Boolean
    Select Case Header.imagetyp
        Case 9
            Compressed = True
        Case 10
            Compressed = True
        Case Else
            Compressed = False
    End Select
End Property

Public Property Get Bpp() As Integer
    Bpp = CInt(Header.pixelsize)
End Property

Public Property Get IsTGA() As Boolean
Dim Test As Boolean
    Test = True
    Select Case Header.imagetyp
        Case 1
        Case 2
        Case 3
        Case 9
        Case 10
        Case 11
        Case 32
        Case 33
        Case Else
            Test = False
    End Select

    Select Case Header.pixelsize
        Case &H8
            '8bpp
        Case &H10
            '16bpp
        Case &H18
            '24bpp
        Case &H20
            '32bpp
        Case Else
            Test = False
    End Select
    IsTGA = Test
End Property

Public Property Get ScaleMode() As Integer
    ScaleMode = Scales
End Property

Public Property Let ScaleMode(ByVal vNewValue As Integer)
    If vNewValue > 0 Then vNewValue = 1
    If vNewValue <> Scales Then
        Scales = vNewValue
    End If
End Property

Public Property Get Autoscale() As Boolean
    Autoscale = Automatic
End Property

Public Property Let Autoscale(ByVal vNewValue As Boolean)
    Automatic = vNewValue
End Property

Private Sub Class_Initialize()
    Automatic = True
    Scales = 1
End Sub

Private Sub CreateBitmap_24(ByteArray() As Byte, PicWidth As Long, PicHeight As Long, Orientation As Integer)
Dim hdc As Long
Dim Bits() As RGBQUAD
ReDim Bits(UBound(ByteArray) / 4)
CopyMemory Bits(0), ByteArray(0), UBound(ByteArray)
    With bm24.bmiHeader
        .biSize = Len(bm24.bmiHeader)           'SizeOf Struct
        .biWidth = PicWidth                     'Bitmap Width
        If Orientation = 0 Then
            .biHeight = PicHeight               'Bitmap Height, bitmap is top down.
        Else
            .biHeight = -PicHeight
        End If
        .biPlanes = 1                           'Single plane
        .biBitCount = 32                        '32 bit alignment
        .biCompression = BI_RGB                 'No Compression
        .biSizeImage = 0                        'Default
        .biXPelsPerMeter = 0                    'Default
        .biYPelsPerMeter = 0                    'Default
        .biClrUsed = 0                          'Default
        .biClrImportant = 0                     'Default
    End With
' Get the DC.
hdc = GetDC(0)
hBmp = CreateDIBitmap_24(hdc, bm24.bmiHeader, CBM_INIT, Bits(0), bm24, DIB_RGB_COLORS)
End Sub

Public Sub DrawBitmap(PicWidth As Long, PicHeight As Long, PicObject As Object, Scalierung As Boolean)
Dim cDC As Long
Dim A As Long
Dim Übergabe As ScTw
Dim realheight As Long
Dim realwidth As Long
    PicObject.Cls
    A = PicObject.ScaleMode
    PicObject.ScaleMode = 1
    
    Select Case Scalierung
        Case True
            Übergabe = PixelToTwips(PicWidth, PicHeight)
            PicObject.Height = Übergabe.TwipHeight
            PicObject.Width = Übergabe.TwipWidth
        Case False
    End Select

    If PicObject.Height <> PicObject.ScaleHeight Then 'with Boarders
        Übergabe = Twipstopixel(PicObject.Width, PicObject.Height)
        realheight = Übergabe.TwipHeight
        realwidth = Übergabe.TwipWidth
        PicObject.Height = PicObject.Height + (PicObject.Height - PicObject.ScaleHeight)
        PicObject.Width = PicObject.Width + (PicObject.Width - PicObject.ScaleWidth)
    Else
        PicObject.ScaleMode = 3
        realheight = PicObject.ScaleHeight
        realwidth = PicObject.ScaleWidth
    End If
    
    If hBmp Then
        cDC = CreateCompatibleDC(PicObject.hdc)
        SelectObject cDC, hBmp
        Call StretchBlt(PicObject.hdc, 0, 0, realwidth, realheight, cDC, 0, 0, PicWidth, PicHeight, SRCCOPY)
        DeleteDC cDC
    End If
    
    PicObject.ScaleMode = A
End Sub

Private Function Twipstopixel(xwert As Long, ywert As Long) As ScTw
Dim ux As Long
Dim uy As Long
Dim XWert1 As Long
Dim yWert1 As Long
    ux = Screen.TwipsPerPixelX
    Twipstopixel.TwipWidth = xwert / ux
    uy = Screen.TwipsPerPixelY
    Twipstopixel.TwipHeight = ywert / uy
End Function

Private Function PixelToTwips(xwert As Long, ywert As Long) As ScTw
Dim ux As Long
Dim uy As Long
Dim XWert1 As Long
Dim yWert1 As Long
    ux = Screen.TwipsPerPixelX
    PixelToTwips.TwipWidth = xwert * ux
    uy = Screen.TwipsPerPixelY
    PixelToTwips.TwipHeight = ywert * uy
End Function

Private Function StrHex(ByVal String1 As String) As String
    Dim strTemp As String, strReturn As String, i As Long
    For i = 1 To Len(String1)
        strTemp = Hex(Asc(Mid$(String1, i, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        strReturn = strReturn & strTemp
    Next i
    StrHex = strReturn
End Function
