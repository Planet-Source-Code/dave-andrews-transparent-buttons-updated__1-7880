VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'This code researched and developed by Dave Andrews
'Check out my website:- http://www.audiokingdom.com  (independent music network)
'I know you probably don't care, but I thought I'de put a plug in for my company :)

'Feel free to use this wherever you want,
'I would just appreciate and credit / mention in your code
'-----------------------------------------------------------
Const COLOR_CAPTIONTEXT = 9
Const DT_CENTER = &H1      'centre left to right
Const DT_VCENTER = &H4     'centre top to bottom
Const DT_NOCLIP = &H100    'fast draw
Const DT_SINGLELINE = &H20 'single line only

Const DT_FLAGS = DT_SINGLELINE Or DT_CENTER Or DT_VCENTER Or DT_NOCLIP

Private parentPic As PictureBox
Private pWidth As Long
Private pHeight As Long
Private pHwnd As Long

Private UpDC As Long 'returns the address for the up image
Private UpMemPal As Long
Private UpPal As Long
Private UpMemBitmap As Long
Private UpBitmap As Long
Private UpParent As Long
Private UpWidth As Long
Private UpHeight As Long
Private UpColor As Long
Private UpMemoryFont As Long
Private UpOrginalFont As Long

Private DownDC As Long 'Returns the address for the down image
Private DownMemPal As Long
Private DownPal As Long
Private DownMemBitmap As Long
Private DownBitmap As Long
Private DownParent As Long
Private DownWidth As Long
Private DownHeight As Long
Private DownColor As Long
Private DownMemoryFont As Long
Private DownOrginalFont As Long

'----------Hook stuff----------------
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Const GWL_WNDPROC = (-4)
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208

Private gOldProc As Long

'--------HSL CONVERSION STUFF FOR BLENDING PIXELS-----------
'HSL conversion routines taken from Dan Redding's "Color Lab"
'Originally converted from the Microsoft Knowledge Base

Private Const HSLMAX As Integer = 240 '***
Const RGBMAX As Integer = 255 '***
Const UNDEFINED As Integer = (HSLMAX * 2 / 3) '***
Private Type HSLCol
    Hue As Integer
    Sat As Integer
    Lum As Integer
End Type
'-----------------------API DECLARATIONS-----------------------------------

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal color As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function GetCursor Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long


'---------------TAKEN DIRECTLY FROM TONY'S CODE FOR A VIRTUAL DC-------
Option Explicit
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const LF_FACESIZE As Long = 32
Private Const CLIP_DEFAULT_PRECIS As Long = 0
Private Const OUT_DEFAULT_PRECIS As Long = 0
Private Const DEFAULT_PITCH As Long = 0
Private Const DEFAULT_QUALITY As Long = 0
Private Const FW_NORMAL As Long = 400

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Private Type TEXTMETRIC
    tmMemoryHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY
End Type

Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

'Private Const BS_TRANSPARENT As Long = 0
'Private Const BS_OPAQUE As Long = 1

Public Enum BackStyles
    BS_TRANSPARENT = 0
    BS_OPAQUE = 1
End Enum

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

'Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long

Private Declare Function GetBkMode Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long

Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long


Private Const SRCCOPY = &HCC0020









Private Sub EmbossUp(Text As String, dX As Integer, dY As Integer)
    Dim wTextParams As DRAWTEXTPARAMS
    Dim rc As RECT
    With rc
        .Left = 2 + dX
        .Top = 2 + dY
        .Right = UpWidth
        .Bottom = UpHeight
    End With
    wTextParams.cbSize = Len(wTextParams)
    Call SetTextColor(UpDC, Brighten(GetPixel(UpDC, UpWidth / 2, UpHeight / 2), 0.4))
    Call DrawTextEx(UpDC, Text, Len(Text), rc, DT_FLAGS, wTextParams)
    With rc
        .Left = 0 + dX
        .Top = 0 + dY
        .Right = UpWidth
        .Bottom = UpHeight
    End With
    wTextParams.cbSize = Len(wTextParams)
    Call SetTextColor(UpDC, UpColor)
    Call DrawTextEx(UpDC, Text, Len(Text), rc, DT_FLAGS, wTextParams)
End Sub


Private Sub EmbossDown(Text As String, dX As Integer, dY As Integer)
    Dim wTextParams As DRAWTEXTPARAMS
    Dim rc As RECT
    With rc
        .Left = 2 + dX
        .Top = 2 + dY
        .Right = DownWidth
        .Bottom = DownHeight
    End With
    wTextParams.cbSize = Len(wTextParams)
    Call SetTextColor(DownDC, Brighten(GetPixel(DownDC, DownWidth / 2, DownHeight / 2), 0.4))
    Call DrawTextEx(DownDC, Text, Len(Text), rc, DT_FLAGS, wTextParams)
    With rc
        .Left = 0 + dX
        .Top = 0 + dY
        .Right = DownWidth
        .Bottom = DownHeight
    End With
    wTextParams.cbSize = Len(wTextParams)
    Call SetTextColor(DownDC, DownColor)
    Call DrawTextEx(DownDC, Text, Len(Text), rc, DT_FLAGS, wTextParams)
End Sub



Private Sub RaiseBevel(bPic As PictureBox, bLevel As Single, bEdge As Integer, Trans As Boolean)
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim bW As Integer
Dim bH As Integer
Dim bL As Integer
Dim bT As Integer
Dim sPix As Long
Dim fPix As Long
Dim M As Single
Dim conHDC As Long
Dim picHDC As Long
Dim BContainer As Object
Set BContainer = bPic.Container
BContainer.ScaleMode = vbPixels
bPic.ScaleMode = vbPixels
BContainer.AutoRedraw = True
bPic.AutoRedraw = True
bW = bPic.ScaleWidth - 1
bH = bPic.ScaleHeight - 1
bL = bPic.Left
bT = bPic.Top
picHDC = bPic.hdc
conHDC = BContainer.hdc
'If transparent, then copy container image as background,
If Trans = True Then
    BitBlt UpDC, 0, 0, bW, bH, conHDC, bL + 1, bT + 1, vbSrcCopy
    BitBlt picHDC, 0, 0, bW, bH, conHDC, bL + 1, bT + 1, vbSrcCopy
Else 'we copy the container content onto the holder image and the backcolor
    ClsUP bPic.BackColor
End If 'Left Edge
M = 1
For i = 0 To bEdge
    For j = i To bH - i
        sPix = Brighten(GetPixel(picHDC, i, j), M * bLevel)
        fPix = GetPixel(conHDC, i + bL, j + bT)
        SetPixel UpDC, i, j, Blend(sPix, fPix, M)
    Next j
    M = M - (1 / bEdge)
Next i
For j = bEdge To bH - bEdge
    SetPixel UpDC, bEdge, j, Brighten(GetPixel(picHDC, bEdge, j), bLevel)
Next j
'Top Edge
M = 1
For i = 0 To bEdge
    For j = i To bW - i
        sPix = Brighten(GetPixel(picHDC, j, i), M * bLevel)
        fPix = GetPixel(conHDC, j + bL, i + bT)
        SetPixel UpDC, j, i, Blend(sPix, fPix, M)
    Next j
    M = M - (1 / bEdge)
Next i
For j = bEdge To bW - bEdge
    SetPixel UpDC, j, i, Brighten(GetPixel(picHDC, j, bEdge), bLevel)
Next j
'Right Edge
M = 0
For i = bW - bEdge To bW
    For j = bW - i To bH - (bW - i)
        sPix = Darken(GetPixel(picHDC, i, j), M * bLevel)
        fPix = GetPixel(conHDC, i + bL, j + bT)
        SetPixel UpDC, i, j, Blend(sPix, fPix, M)
    Next j
    M = M + (1 / bEdge)
Next i
For j = bEdge To bH - bEdge
    SetPixel UpDC, (bW - bEdge), j, Darken(GetPixel(picHDC, (bW - bEdge), j), bLevel)
Next j
'Bottom Edge
M = 0
For i = bH - bEdge To bH
    For j = bH - i To bW - (bH - i)
        sPix = Darken(GetPixel(picHDC, j, i), M * bLevel)
        fPix = GetPixel(conHDC, j + bL, i + bT)
        SetPixel UpDC, j, i, Blend(sPix, fPix, M)
    Next j
    M = M + (1 / bEdge)
Next i
For j = bEdge To bW - bEdge
    SetPixel UpDC, j, (bH - bEdge), Darken(GetPixel(picHDC, j, (bH - bEdge)), bLevel)
Next j
Set BContainer = Nothing
End Sub

Private Sub InsetBevel(bPic As PictureBox, bLevel As Single, bEdge As Integer, Trans As Boolean)
'Creates a bevel around the perimeter of an object
'By blending the outside pixels of the objects container
'on a gradient scale.
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim bW As Integer
Dim bH As Integer
Dim bL As Integer
Dim bT As Integer
Dim sPix As Long
Dim fPix As Long
Dim M As Single
Dim picHDC As Long
Dim conHDC As Long
Dim BContainer As Object
Set BContainer = bPic.Container
BContainer.ScaleMode = vbPixels
bPic.ScaleMode = vbPixels
BContainer.AutoRedraw = True
bPic.AutoRedraw = True
bW = bPic.ScaleWidth - 1
bH = bPic.ScaleHeight - 1
bL = bPic.Left
bT = bPic.Top
picHDC = bPic.hdc
conHDC = BContainer.hdc
'If transparent, then copy container image as background,
If Trans = True Then
    BitBlt DownDC, 0, 0, bW, bH, conHDC, bL - 1, bT - 1, vbSrcCopy
    BitBlt picHDC, 0, 0, bW, bH, conHDC, bL - 1, bT - 1, vbSrcCopy
Else 'we copy the container content onto the holder image and the backcolor
    ClsDOWN bPic.BackColor
End If 'Left Edge
M = 1
For i = 0 To bEdge - 1
    For j = i To bH - i
        sPix = Darken(GetPixel(picHDC, i, j), M * bLevel)
        fPix = GetPixel(conHDC, i + bL, j + bT)
        SetPixel DownDC, i, j, Blend(sPix, fPix, M)
    Next j
    M = M - (1 / bEdge)
Next i
For j = bEdge To bH - bEdge
    SetPixel DownDC, bEdge, j, Darken(GetPixel(picHDC, bEdge, j), bLevel)
Next j
'Top Edge
M = 1
For i = 0 To bEdge - 1
    For j = i To bW - i
        sPix = Darken(GetPixel(picHDC, j, i), M * bLevel)
        fPix = GetPixel(conHDC, j + bL, i + bT)
        SetPixel DownDC, j, i, Blend(sPix, fPix, M)
    Next j
    M = M - (1 / bEdge)
Next i
For j = bEdge To bW - bEdge
    SetPixel DownDC, j, i, Darken(GetPixel(picHDC, j, bEdge), bLevel)
Next j
'Right Edge
M = 0
For i = (bW - bEdge) + 1 To bW
    For j = bW - i To bH - (bW - i)
        sPix = Brighten(GetPixel(picHDC, i, j), M * bLevel)
        fPix = GetPixel(conHDC, i + bL, j + bT)
        SetPixel DownDC, i, j, Blend(sPix, fPix, M)
    Next j
    M = M + (1 / bEdge)
Next i
For j = bEdge To bH - bEdge
    SetPixel DownDC, (bW - bEdge), j, Brighten(GetPixel(picHDC, (bW - bEdge), j), bLevel)
Next j
'Bottom Edge
M = 0
For i = (bH - bEdge) + 1 To bH
    For j = bH - i To bW - (bH - i)
        sPix = Brighten(GetPixel(picHDC, j, i), M * bLevel)
        fPix = GetPixel(conHDC, j + bL, i + bT)
        SetPixel DownDC, j, i, Blend(sPix, fPix, M)
    Next j
    M = M + (1 / bEdge)
Next i
For j = bEdge To bW - bEdge
    SetPixel DownDC, j, (bH - bEdge), Brighten(GetPixel(picHDC, j, (bH - bEdge)), bLevel)
Next j
Set BContainer = Nothing
End Sub

Private Sub RaiseRound(bPic As PictureBox, bLevel As Single, bEdge As Integer, bRad As Integer, Trans As Boolean)
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim bW As Integer
Dim bH As Integer
Dim bL As Integer
Dim bT As Integer
Dim sPix As Long
Dim fPix As Long
Dim M As Single
Dim S As Single
Dim picHDC As Long
Dim conHDC As Long
Dim BContainer As Object
Dim pi
pi = 4 * Atn(1)
Set BContainer = bPic.Container
Set BContainer = bPic.Container
BContainer.ScaleMode = vbPixels
bPic.ScaleMode = vbPixels
BContainer.AutoRedraw = True
bPic.AutoRedraw = True
bW = bPic.ScaleWidth - 1
bH = bPic.ScaleHeight - 1
bL = bPic.Left
bT = bPic.Top
picHDC = bPic.hdc
conHDC = BContainer.hdc
'If transparent, then copy container image as background,
If Trans = True Then
    BitBlt UpDC, 0, 0, bW, bH, conHDC, bL + 1, bT + 1, vbSrcCopy
    BitBlt picHDC, 0, 0, bW, bH, conHDC, bL + 1, bT + 1, vbSrcCopy
Else 'we copy the container content onto the holder image and the backcolor
    ClsUP bPic.BackColor
End If
'Left Edge
'GoTo Corners
M = 1
For i = 0 To bEdge - 1
    For j = (bRad + bEdge) To bH - (bRad + bEdge)
        sPix = Brighten(GetPixel(picHDC, i, j), M * bLevel)
        fPix = GetPixel(conHDC, i + bL, j + bT)
        SetPixel UpDC, i, j, Blend(sPix, fPix, M)
    Next j
    M = M - (1 / bEdge)
Next i
For j = (bRad + bEdge) To bH - (bRad + bEdge)
    sPix = Brighten(GetPixel(picHDC, bEdge, j), bLevel)
    SetPixel UpDC, bEdge, j, sPix
Next j
'Top Edge
M = 1
For i = 0 To bEdge - 1
    For j = bRad + bEdge To bW - (bRad + bEdge)
        sPix = Brighten(GetPixel(picHDC, j, i), M * bLevel)
        fPix = GetPixel(conHDC, j + bL, i + bT)
        SetPixel UpDC, j, i, Blend(sPix, fPix, M)
    Next j
    M = M - (1 / bEdge)
Next i
For j = bRad + bEdge To bW - (bRad + bEdge)
    sPix = Brighten(GetPixel(picHDC, j, bEdge), bLevel)
    SetPixel UpDC, j, bEdge, sPix
Next j
'Right Edge
M = 0
For i = (bW - bEdge) + 1 To bW
    For j = bRad + bEdge To bH - (bRad + bEdge)
        sPix = Darken(GetPixel(picHDC, i, j), M * bLevel)
        fPix = GetPixel(conHDC, i + bL, j + bT)
        SetPixel UpDC, i, j, Blend(sPix, fPix, M)
    Next j
    M = M + (1 / bEdge)
Next i
For j = bRad + bEdge To bH - (bRad + bEdge)
    sPix = Darken(GetPixel(picHDC, (bW - bEdge), j), bLevel)
    SetPixel UpDC, (bW - bEdge), j, sPix
Next j
'Bottom Edge
M = 0
For i = (bH - bEdge) + 1 To bH
    For j = bRad + bEdge To bW - (bRad + bEdge)
        sPix = Darken(GetPixel(picHDC, j, i), M * bLevel)
        fPix = GetPixel(conHDC, j + bL, i + bT)
        SetPixel UpDC, j, i, Blend(sPix, fPix, M)
    Next j
    M = M + (1 / bEdge)
Next i
For j = bRad + bEdge To bW - (bRad + bEdge)
    sPix = Darken(GetPixel(picHDC, j, (bH - bEdge)), bLevel)
    SetPixel UpDC, j, (bH - bEdge), sPix
Next j
'--------------CORNERS-----------------------
Corners:
Dim rX As Integer
Dim rY As Integer
Dim lX As Integer
Dim lY As Integer
Dim rStep As Integer
'Top Left Corder
M = 0
For i = bRad To bRad + bEdge
    For j = 91 To 179
        rX = (i * Cos(j * (pi / 180))) + (bRad + bEdge)
        rY = -(i * Sin(j * (pi / 180))) + (bRad + bEdge)
        If lX <> rX Or lY <> rY Then
            If i = bRad Then
                sPix = Brighten(GetPixel(picHDC, rX, rY), bLevel)
            Else
                sPix = Brighten(GetPixel(picHDC, rX, rY), M * bLevel)
            End If
            fPix = GetPixel(conHDC, rX + bL, rY + bT)
            SetPixel UpDC, rX, rY, Blend(sPix, fPix, M)
        End If
        lX = rX
        lY = rY
    Next j
    M = M + (1 / (bEdge))
Next i
'Top Right Corder
M = 0
For i = bRad To bRad + bEdge
    S = 0
    For j = 91 To 179
        rX = -(i * Cos(j * (pi / 180))) + (bW - bRad - bEdge)
        rY = -(i * Sin(j * (pi / 180))) + (bRad + bEdge)
        If lX <> rX Or lY <> rY Then
            If j < 135 Then
                If i = bRad Then
                    sPix = Brighten(GetPixel(picHDC, rX, rY), (1 - S) * bLevel)
                Else
                    sPix = Brighten(GetPixel(picHDC, rX, rY), (M - S) * bLevel)
                End If
            Else
                If i = bRad Then
                    sPix = Darken(GetPixel(picHDC, rX, rY), (1 - S) * bLevel)
                Else
                    sPix = Darken(GetPixel(picHDC, rX, rY), (M - S) * bLevel)
                End If
            End If
            fPix = GetPixel(conHDC, rX + bL, rY + bT)
            SetPixel UpDC, rX, rY, Blend(sPix, fPix, M)
        End If
        If j < 135 Then
            S = S + (1 / 45)
        Else
            S = S - (1 / 45)
        End If
        lX = rX
        lY = rY

    Next j
    M = M + (1 / (bEdge))
Next i
'Bottom Left Corder
M = 0
For i = bRad To bRad + bEdge
    S = 0
    For j = 91 To 179
        rX = (i * Cos(j * (pi / 180))) + (bRad + bEdge)
        rY = (i * Sin(j * (pi / 180))) + (bH - bRad - bEdge)
        If lX <> rX Or lY <> rY Then
            If j > 135 Then
                If i = bRad Then
                    sPix = Brighten(GetPixel(picHDC, rX, rY), (1 + S) * bLevel)
                Else
                    sPix = Brighten(GetPixel(picHDC, rX, rY), (M + S) * bLevel)
                End If
            Else
                If i = bRad Then
                    sPix = Darken(GetPixel(picHDC, rX, rY), (1 + S) * bLevel)
                Else
                    sPix = Darken(GetPixel(picHDC, rX, rY), (M + S) * bLevel)
                End If
            End If
            fPix = GetPixel(conHDC, rX + bL, rY + bT)
            SetPixel UpDC, rX, rY, Blend(sPix, fPix, M)
        End If
        lX = rX
        lY = rY
        If j > 135 Then
            S = S + (1 / 45)
        Else
            S = S - (1 / 45)
        End If
    Next j
    M = M + (1 / (bEdge))
Next i
'Bottom right Corder
M = 0
For i = bRad To bRad + bEdge
    For j = 91 To 179
        rX = -(i * Cos(j * (pi / 180))) + (bW - bRad - bEdge)
        rY = (i * Sin(j * (pi / 180))) + (bH - bRad - bEdge)
        If lX <> rX Or lY <> rY Then
            If i = bRad Then
                sPix = Darken(GetPixel(picHDC, rX, rY), bLevel)
            Else
                sPix = Darken(GetPixel(picHDC, rX, rY), M * bLevel)
            End If
            fPix = GetPixel(conHDC, rX + bL, rY + bT)
            SetPixel UpDC, rX, rY, Blend(sPix, fPix, M)
        End If
        lX = rX
        lY = rY
    Next j
    M = M + (1 / (bEdge))
Next i
Set BContainer = Nothing
End Sub

Private Sub InsetRound(bPic As PictureBox, bLevel As Single, bEdge As Integer, bRad As Integer, Trans As Boolean)
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim bW As Integer
Dim bH As Integer
Dim bL As Integer
Dim bT As Integer
Dim sPix As Long
Dim fPix As Long
Dim M As Single
Dim S As Single
Dim picHDC As Long
Dim conHDC As Long
Dim BContainer As Object
Dim pi
pi = 4 * Atn(1)
Set BContainer = bPic.Container
BContainer.ScaleMode = vbPixels
bPic.ScaleMode = vbPixels
BContainer.AutoRedraw = True
bPic.AutoRedraw = True
bW = bPic.ScaleWidth - 1
bH = bPic.ScaleHeight - 1
bL = bPic.Left
bT = bPic.Top
picHDC = bPic.hdc
conHDC = BContainer.hdc
'If transparent, then copy container image as background,
If Trans = True Then
    BitBlt DownDC, 0, 0, bW, bH, conHDC, bL - 1, bT - 1, vbSrcCopy
    BitBlt picHDC, 0, 0, bW, bH, conHDC, bL - 1, bT - 1, vbSrcCopy
Else 'we copy the container content onto the holder image and the backcolor
    ClsDOWN bPic.BackColor
End If 'Left Edge
'GoTo Corners
M = 1
For i = 0 To bEdge - 1
    For j = (bRad + bEdge) To bH - (bRad + bEdge)
        sPix = Darken(GetPixel(picHDC, i, j), M * bLevel)
        fPix = GetPixel(conHDC, i + bL, j + bT)
        SetPixel DownDC, i, j, Blend(sPix, fPix, M)
    Next j
    M = M - (1 / bEdge)
Next i
For j = (bRad + bEdge) To bH - (bRad + bEdge)
    sPix = Darken(GetPixel(picHDC, bEdge, j), bLevel)
    SetPixel DownDC, bEdge, j, sPix
Next j
'Top Edge
M = 1
For i = 0 To bEdge - 1
    For j = bRad + bEdge To bW - (bRad + bEdge)
        sPix = Darken(GetPixel(picHDC, j, i), M * bLevel)
        fPix = GetPixel(conHDC, j + bL, i + bT)
        SetPixel DownDC, j, i, Blend(sPix, fPix, M)
    Next j
    M = M - (1 / bEdge)
Next i
For j = bRad + bEdge To bW - (bRad + bEdge)
    sPix = Darken(GetPixel(picHDC, j, bEdge), bLevel)
    SetPixel DownDC, j, bEdge, sPix
Next j
'Right Edge
M = 0
For i = (bW - bEdge) + 1 To bW
    For j = bRad + bEdge To bH - (bRad + bEdge)
        sPix = Brighten(GetPixel(picHDC, i, j), M * bLevel)
        fPix = GetPixel(conHDC, i + bL, j + bT)
        SetPixel DownDC, i, j, Blend(sPix, fPix, M)
    Next j
    M = M + (1 / bEdge)
Next i
For j = bRad + bEdge To bH - (bRad + bEdge)
    sPix = Brighten(GetPixel(picHDC, (bW - bEdge), j), bLevel)
    SetPixel DownDC, (bW - bEdge), j, sPix
Next j
'Bottom Edge
M = 0
For i = (bH - bEdge) + 1 To bH
    For j = bRad + bEdge To bW - (bRad + bEdge)
        sPix = Brighten(GetPixel(picHDC, j, i), M * bLevel)
        fPix = GetPixel(conHDC, j + bL, i + bT)
        SetPixel DownDC, j, i, Blend(sPix, fPix, M)
    Next j
    M = M + (1 / bEdge)
Next i
For j = bRad + bEdge To bW - (bRad + bEdge)
    sPix = Brighten(GetPixel(picHDC, j, (bH - bEdge)), bLevel)
    SetPixel DownDC, j, (bH - bEdge), sPix
Next j
'--------------CORNERS-----------------------
Corners:
'bRad = bRad + 1 'return the value to it's original settings
Dim rX As Integer
Dim rY As Integer
Dim lX As Integer
Dim lY As Integer
Dim rStep As Integer
'Top Left Corder
M = 0
For i = bRad To bRad + bEdge
    For j = 91 To 179
        rX = (i * Cos(j * (pi / 180))) + (bRad + bEdge)
        rY = -(i * Sin(j * (pi / 180))) + (bRad + bEdge)
        If lX <> rX Or lY <> rY Then
            If i = bRad Then
                sPix = Darken(GetPixel(picHDC, rX, rY), bLevel)
            Else
                sPix = Darken(GetPixel(picHDC, rX, rY), M * bLevel)
            End If
            fPix = GetPixel(conHDC, rX + bL, rY + bT)
            SetPixel DownDC, rX, rY, Blend(sPix, fPix, M)
        End If
        lX = rX
        lY = rY
    Next j
    M = M + (1 / (bEdge))
Next i
'Top Right Corder
M = 0
For i = bRad To bRad + bEdge
    S = 0
    For j = 91 To 179
        rX = -(i * Cos(j * (pi / 180))) + (bW - bRad - bEdge)
        rY = -(i * Sin(j * (pi / 180))) + (bRad + bEdge)
        If lX <> rX Or lY <> rY Then
            If j < 135 Then
                If i = bRad Then
                    sPix = Darken(GetPixel(picHDC, rX, rY), (1 - S) * bLevel)
                Else
                    sPix = Darken(GetPixel(picHDC, rX, rY), (M - S) * bLevel)
                End If
            Else
                If i = bRad Then
                    sPix = Brighten(GetPixel(picHDC, rX, rY), (1 - S) * bLevel)
                Else
                    sPix = Brighten(GetPixel(picHDC, rX, rY), (M - S) * bLevel)
                End If
            End If
            fPix = GetPixel(conHDC, rX + bL, rY + bT)
            SetPixel DownDC, rX, rY, Blend(sPix, fPix, M)
        End If
        If j < 135 Then
            S = S + (1 / 45)
        Else
            S = S - (1 / 45)
        End If
        lX = rX
        lY = rY

    Next j
    M = M + (1 / (bEdge))
Next i
'Bottom Left Corder
M = 0
For i = bRad To bRad + bEdge
    S = 0
    For j = 91 To 179
        rX = (i * Cos(j * (pi / 180))) + (bRad + bEdge)
        rY = (i * Sin(j * (pi / 180))) + (bH - bRad - bEdge)
        If lX <> rX Or lY <> rY Then
            If j > 135 Then
                If i = bRad Then
                    sPix = Darken(GetPixel(picHDC, rX, rY), (1 + S) * bLevel)
                Else
                    sPix = Darken(GetPixel(picHDC, rX, rY), (M + S) * bLevel)
                End If
            Else
                If i = bRad Then
                    sPix = Brighten(GetPixel(picHDC, rX, rY), (1 + S) * bLevel)
                Else
                    sPix = Brighten(GetPixel(picHDC, rX, rY), (M + S) * bLevel)
                End If
            End If
            fPix = GetPixel(conHDC, rX + bL, rY + bT)
            SetPixel DownDC, rX, rY, Blend(sPix, fPix, M)
        End If
        lX = rX
        lY = rY
        If j > 135 Then
            S = S + (1 / 45)
        Else
            S = S - (1 / 45)
        End If
    Next j
    M = M + (1 / (bEdge))
Next i
'Bottom right Corder
M = 0
For i = bRad To bRad + bEdge
    For j = 91 To 179
        rX = -(i * Cos(j * (pi / 180))) + (bW - bRad - bEdge)
        rY = (i * Sin(j * (pi / 180))) + (bH - bRad - bEdge)
        If lX <> rX Or lY <> rY Then
            If i = bRad Then
                sPix = Brighten(GetPixel(picHDC, rX, rY), bLevel)
            Else
                sPix = Brighten(GetPixel(picHDC, rX, rY), M * bLevel)
            End If
            fPix = GetPixel(conHDC, rX + bL, rY + bT)
            SetPixel DownDC, rX, rY, Blend(sPix, fPix, M)
        End If
        lX = rX
        lY = rY
    Next j
    M = M + (1 / (bEdge))
Next i
Set BContainer = Nothing
End Sub




Function InitButton(pTarget As PictureBox, lText As String, bRounded As Boolean, bLevel As Single, bEdge As Integer, bRad As Integer, Clear As Boolean) As Single
Dim dT As Single
Dim bW As Long
Dim bH As Long
dT = Timer()
pTarget.ScaleMode = vbPixels
pTarget.Container.ScaleMode = vbPixels
bW = pTarget.ScaleWidth - 1  'we subtract 1 to account for the missing border
bH = pTarget.ScaleHeight - 1
pWidth = bW + 1
pHeight = bH + 1
pHwnd = pTarget.hwnd
Set parentPic = pTarget
'--------------CREATE A REGION IF IT'S ROUNDED----------------------------
If bRounded Then
    Dim NewRGN As Long
    If bW = bH Then
        NewRGN = CreateEllipticRgn(0, 0, bW + 1, bH + 1)
    Else
        NewRGN = CreateEllipticRgn(0, 0, ((bRad + bEdge) * 2) + 1, ((bRad + bEdge) * 2) + 1)
        CombineRgn NewRGN, NewRGN, CreateEllipticRgn(bW + 1, bH + 1, (bW - 1) - ((bRad + bEdge) * 2), (bH - 1) - ((bRad + bEdge) * 2)), 2
        CombineRgn NewRGN, NewRGN, CreateEllipticRgn(0, bH + 1, ((bRad + bEdge) * 2) + 1, (bH - 1) - ((bRad + bEdge) * 2)), 2
        CombineRgn NewRGN, NewRGN, CreateEllipticRgn(bW + 1, 0, (bW - 1) - ((bRad + bEdge) * 2), ((bRad + bEdge) * 2) + 1), 2
        CombineRgn NewRGN, NewRGN, CreateRectRgn(0, (bRad + bEdge) + 1, bW + 1, bH - (bRad + bEdge)), 2
        CombineRgn NewRGN, NewRGN, CreateRectRgn((bRad + bEdge) + 1, 0, bW - (bRad + bEdge), bH + 1), 2
    End If
    SetWindowRgn pTarget.hwnd, NewRGN, True
    '-----------Get Images for rounded button
    DownDC = CreateDOWN(pTarget.hdc, bW, bH)
    InsetRound pTarget, bLevel, bEdge, bRad, Clear
    EmbossDown lText, 1, 1
    UpDC = CreateUP(pTarget.hdc, bW, bH)
    RaiseRound pTarget, bLevel, bEdge, bRad, Clear
    EmbossUp lText, -1, -1
Else
    DownDC = CreateDOWN(pTarget.hdc, bW, bH)
    InsetBevel pTarget, bLevel, bEdge, Clear
    EmbossDown lText, 1, 1
    UpDC = CreateUP(pTarget.hdc, bW, bH)
    RaiseBevel pTarget, bLevel, bEdge, Clear
    EmbossUp lText, -1, -1
End If
'--------------Paint the UP state-----------
StretchBlt pTarget.hdc, 0, 0, pTarget.ScaleWidth, pTarget.ScaleHeight, UpDC, 0, 0, UpWidth, UpHeight, vbSrcCopy
pTarget.Refresh
'---------------I wish I could use the a call back within a class!----
'---------------If I could I would hook the mouse down and MOuse up events of
'---------------the target picturebox to display the states
'TriggerButton
InitButton = Timer - dT
End Function




Private Function NZ(ZInput, Optional ZDefault) As Variant
On Error GoTo SkipIt:
If IsMissing(ZDefault) Then ZDefault = ""
If (IsNull(ZInput)) Or (ZInput = Empty) Or (ZInput = "") Then
    NZ = ZDefault
Else
    NZ = ZInput
End If
Exit Function
SkipIt:
NZ = ""
End Function



Private Function Brighten(RGBColor As Long, Percent As Single)
'Brightens a color by a decimal percent
Dim HSL As HSLCol, L As Long
    If Percent <= 0 Then
        Brighten = RGBColor
        Exit Function
    End If
    
    HSL = RGBtoHSL(RGBColor)
    L = HSL.Lum + (HSLMAX * Percent)
    If L > HSLMAX Then L = HSLMAX
    HSL.Lum = L
    Brighten = HSLtoRGB(HSL)
End Function
Private Function Darken(RGBColor As Long, Percent As Single)
'Darkens a color by a percent
Dim HSL As HSLCol, L As Long
    If Percent <= 0 Then
        Darken = RGBColor
        Exit Function
    End If
    
    HSL = RGBtoHSL(RGBColor)
    L = HSL.Lum - (HSLMAX * Percent)
    If L < 0 Then L = 0
    HSL.Lum = L
    Darken = HSLtoRGB(HSL)
End Function
Private Function Blend(RGB1 As Long, RGB2 As Long, Percent As Single) As Long
'blends two colors together by a certain percent (decimal percent)
Dim R As Integer, R1 As Integer, R2 As Integer, g As Integer, G1 As Integer, G2 As Integer, b As Integer, b1 As Integer, b2 As Integer
    
    If Percent >= 1 Then
        Blend = RGB2
        Exit Function
    ElseIf Percent <= 0 Then
        Blend = RGB1
        Exit Function
    End If
    
    R1 = RGBRed(RGB1)
    R2 = RGBRed(RGB2)
    G1 = RGBGreen(RGB1)
    G2 = RGBGreen(RGB2)
    b1 = RGBBlue(RGB1)
    b2 = RGBBlue(RGB2)
    
    R = ((R2 * Percent) + (R1 * (1 - Percent)))
    g = ((G2 * Percent) + (G1 * (1 - Percent)))
    b = ((b2 * Percent) + (b1 * (1 - Percent)))
    
    Blend = RGB(R, g, b)
End Function
Private Function iMax(a As Integer, b As Integer) _
    As Integer
'Return the Larger of two values
    iMax = IIf(a > b, a, b)
End Function

Private Function iMin(a As Integer, b As Integer) _
    As Integer
'Return the smaller of two values
    iMin = IIf(a < b, a, b)
End Function

Private Function RGBRed(RGBCol As Long) As Integer
If RGBCol = -1 Then Exit Function
'Return the Red component from an RGB Color
    RGBRed = RGBCol And &HFF
End Function

Private Function RGBGreen(RGBCol As Long) As Integer
If RGBCol = -1 Then Exit Function
'Return the Green component from an RGB Color
    RGBGreen = ((RGBCol And &H100FF00) / &H100)
End Function

Private Function RGBBlue(RGBCol As Long) As Integer
If RGBCol = -1 Then Exit Function
'Return the Blue component from an RGB Color
    RGBBlue = (RGBCol And &HFF0000) / &H10000
End Function
Private Function HSLtoRGB(HueLumSat As HSLCol) As Long '***
'Converts HSL to a color value
    Dim R As Double, g As Double, b As Double
    Dim H As Double, L As Double, S As Double
    Dim Magic1 As Double, Magic2 As Double
    H = HueLumSat.Hue
    L = HueLumSat.Lum
    S = HueLumSat.Sat
    If CInt(S) = 0 Then
        R = (L * RGBMAX) / HSLMAX
        g = R
        b = R
    Else
        If L <= HSLMAX / 2 Then
            Magic2 = (L * (HSLMAX + S) + 0.5) / HSLMAX
        Else
            Magic2 = L + S - ((L * S) + 0.5) / HSLMAX
        End If
        Magic1 = 2 * L - Magic2
        R = (HuetoRGB(Magic1, Magic2, H + (HSLMAX / 3)) * RGBMAX + 0.5) / HSLMAX
        g = (HuetoRGB(Magic1, Magic2, H) * RGBMAX + 0.5) / HSLMAX
        b = (HuetoRGB(Magic1, Magic2, H - (HSLMAX / 3)) * RGBMAX + 0.5) / HSLMAX
    End If
    HSLtoRGB = RGB(CInt(R), CInt(g), CInt(b))
End Function

Private Function HuetoRGB(mag1 As Double, mag2 As Double, ByVal Hue As Double) As Double     '***
'Utility function for HSLtoRGB
    If Hue < 0 Then
        Hue = Hue + HSLMAX
    ElseIf Hue > HSLMAX Then
        Hue = Hue - HSLMAX
    End If
    Select Case Hue
        Case Is < (HSLMAX / 6)
            HuetoRGB = (mag1 + (((mag2 - mag1) * Hue + _
                (HSLMAX / 12)) / (HSLMAX / 6)))
        Case Is < (HSLMAX / 2)
            HuetoRGB = mag2
        Case Is < (HSLMAX * 2 / 3)
            HuetoRGB = (mag1 + (((mag2 - mag1) * _
                ((HSLMAX * 2 / 3) - Hue) + _
                (HSLMAX / 12)) / (HSLMAX / 6)))
        Case Else
            HuetoRGB = mag1
    End Select
End Function
Private Function RGBtoHSL(RGBCol As Long) As HSLCol '***
'Returns an HSLCol datatype containing Hue, Luminescence
'and Saturation; given an RGB Color value

Dim R As Integer, g As Integer, b As Integer
Dim cMax As Integer, cMin As Integer
Dim RDelta As Double, GDelta As Double, _
    BDelta As Double
Dim H As Double, S As Double, L As Double
Dim cMinus As Long, cPlus As Long
    
    R = RGBRed(RGBCol)
    g = RGBGreen(RGBCol)
    b = RGBBlue(RGBCol)
    
    cMax = iMax(iMax(R, g), b) 'Highest and lowest
    cMin = iMin(iMin(R, g), b) 'color values
    
    cMinus = cMax - cMin 'Used to simplify the
    cPlus = cMax + cMin  'calculations somewhat.
    
    'Calculate luminescence (lightness)
    L = ((cPlus * HSLMAX) + RGBMAX) / (2 * RGBMAX)
    
    If cMax = cMin Then 'achromatic (r=g=b, greyscale)
        S = 0 'Saturation 0 for greyscale
        H = UNDEFINED 'Hue undefined for greyscale
    Else
        'Calculate color saturation
        If L <= (HSLMAX / 2) Then
            S = ((cMinus * HSLMAX) + 0.5) / cPlus
        Else
            S = ((cMinus * HSLMAX) + 0.5) / (2 * RGBMAX - cPlus)
        End If
    
        'Calculate hue
        RDelta = (((cMax - R) * (HSLMAX / 6)) + 0.5) / cMinus
        GDelta = (((cMax - g) * (HSLMAX / 6)) + 0.5) / cMinus
        BDelta = (((cMax - b) * (HSLMAX / 6)) + 0.5) / cMinus
    
        Select Case cMax
            Case CLng(R)
                H = BDelta - GDelta
            Case CLng(g)
                H = (HSLMAX / 3) + RDelta - BDelta
            Case CLng(b)
                H = ((2 * HSLMAX) / 3) + GDelta - RDelta
        End Select
        
        If H < 0 Then H = H + HSLMAX
    End If
    
    RGBtoHSL.Hue = CInt(H)
    RGBtoHSL.Lum = CInt(L)
    RGBtoHSL.Sat = CInt(S)
End Function

Sub TriggerButton()
'Dim UpState As Boolean
'Do
'    DoEvents
'    If GetCapture = pHwnd And UpState = False Then
'        parentPic.Cls
'        StretchBlt parentPic.hdc, 0, 0, pWidth, pHeight, DownDC, 0, 0, DownWidth, DownHeight, vbSrcCopy
'        UpState = True
'    ElseIf GetCapture <> pHwnd And UpState = True Then
'       parentPic.Cls
'        StretchBlt parentPic.hdc, 0, 0, pWidth, pHeight, UpDC, 0, 0, UpWidth, UpHeight, vbSrcCopy
'        UpState = False
'    End If
'Loop

parentPic.Cls
StretchBlt parentPic.hdc, 0, 0, pWidth, pHeight, DownDC, 0, 0, DownWidth, DownHeight, vbSrcCopy
Do While pHwnd = GetCapture()
    DoEvents
Loop
parentPic.Cls
StretchBlt parentPic.hdc, 0, 0, pWidth, pHeight, UpDC, 0, 0, UpWidth, UpHeight, vbSrcCopy
End Sub

Private Sub Class_Initialize()
    UpDC = 0
    DownDC = 0
End Sub

Private Sub Class_Terminate()
    DestroyUP
    DestroyDown
    Set parentPic = Nothing
End Sub

Private Function GetColor(ByVal nColor As Long) As Long
    Const SYSCOLOR_BIT As Long = &H80000000
    If (nColor And SYSCOLOR_BIT) = SYSCOLOR_BIT Then
        nColor = nColor And (Not SYSCOLOR_BIT)
        GetColor = GetSysColor(nColor)
    Else
        GetColor = nColor
    End If
End Function

Private Function IsUpCreated() As Boolean
    IsUpCreated = (UpDC <> 0)
End Function



Private Function IsDownCreated() As Boolean
    IsDownCreated = (DownDC <> 0)
End Function

Private Function CreateUP(hParentDC As Long, Optional PixelWidth As Long = 1024, Optional PixelHeight As Long = 768) As Long
    Dim nHasPalette As Long
    Dim nPaletteSize As Long
    Dim LogPal As LOGPALETTE
    Dim tm As TEXTMETRIC
    Dim sFaceName As String * 80
    Dim fFont As StdFont

    If IsUpCreated Then DestroyUP

    UpParent = hParentDC
    UpWidth = PixelWidth
    UpHeight = PixelHeight

    ' Create a memory device context to use
    UpDC = CreateCompatibleDC(UpParent)

    ' Tell'em it's a picture (so drawings can be done on the DC)
    UpMemBitmap = CreateCompatibleBitmap(UpParent, UpWidth, UpHeight)
    UpBitmap = SelectObject(UpDC, UpMemBitmap)

    ' Get screen properties
    nHasPalette = GetDeviceCaps(UpParent, RASTERCAPS) And RC_PALETTE   ' Palette support
    nPaletteSize = GetDeviceCaps(UpParent, SIZEPALETTE)                ' Size of palette
    ' If the screen has a palette make a copy and realize it
    If nHasPalette And (nPaletteSize = 256) Then
        ' Create a copy of the system palette
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        Call GetSystemPaletteEntries(UpParent, 0&, 256, LogPal.palPalEntry(0))
        UpMemPal = CreatePalette(LogPal)
        ' Select the new palette into the memory DC and realize it
        UpPal = SelectPalette(UpDC, UpMemPal, 0&)
        Call RealizePalette(UpDC)
    End If
    Call SetBkColor(UpDC, GetBkColor(UpParent))
    UpColor = GetTextColor(UpParent)
    Call SetBkMode(UpDC, GetBkMode(UpParent))
    
    Call GetTextMetrics(UpDC, tm)
    Call GetTextFace(UpParent, 79, sFaceName)
    Set fFont = New StdFont
    With fFont
        .Bold = (tm.tmWeight > FW_NORMAL)
        .Charset = tm.tmCharSet
        .Italic = (tm.tmItalic <> 0)
        .Name = sFaceName
        .Strikethrough = (tm.tmStruckOut <> 0)
        .Underline = (tm.tmUnderlined <> 0)
        .Weight = tm.tmWeight
        .Size = (tm.tmMemoryHeight / tm.tmDigitizedAspectY) * 72 ' Size has to be calculated
    End With
    Set UPFont = fFont
    Set fFont = Nothing
    
    CreateUP = UpDC
End Function
Private Property Get UPFont() As StdFont
    If Not IsUpCreated Then Exit Property

    On Local Error Resume Next

    Dim tm As TEXTMETRIC
    Dim sFaceName As String * 80

    Call GetTextMetrics(UpDC, tm)
    Call GetTextFace(UpDC, 79, sFaceName)

    Set UPFont = New StdFont

    With UPFont
        .Bold = (tm.tmWeight > FW_NORMAL)
        .Charset = tm.tmCharSet
        .Italic = (tm.tmItalic <> 0)
        .Name = sFaceName 'StrConv(sFaceName, vbUnicode)
        .Strikethrough = (tm.tmStruckOut <> 0)
        .Underline = (tm.tmUnderlined <> 0)
        .Weight = tm.tmWeight
        .Size = (tm.tmMemoryHeight / tm.tmDigitizedAspectY) * 72 ' Size has to be calculated
    End With
End Property
Private Property Get DOWNFont() As StdFont
    If Not IsDownCreated Then Exit Property

    On Local Error Resume Next

    Dim tm As TEXTMETRIC
    Dim sFaceName As String * 80

    Call GetTextMetrics(DownDC, tm)
    Call GetTextFace(DownDC, 79, sFaceName)

    Set DOWNFont = New StdFont

    With DOWNFont
        .Bold = (tm.tmWeight > FW_NORMAL)
        .Charset = tm.tmCharSet
        .Italic = (tm.tmItalic <> 0)
        .Name = sFaceName 'StrConv(sFaceName, vbUnicode)
        .Strikethrough = (tm.tmStruckOut <> 0)
        .Underline = (tm.tmUnderlined <> 0)
        .Weight = tm.tmWeight
        .Size = (tm.tmMemoryHeight / tm.tmDigitizedAspectY) * 72 ' Size has to be calculated
    End With
End Property

Private Property Set UPFont(ByVal NewFont As StdFont)
    If Not IsUpCreated Then Exit Property

    On Local Error Resume Next

    Dim nName() As Byte, i As Byte, nSize As Byte
    Dim tFont As LOGFONT

    ' Font name is a byte array and is in ANSI (DOS) format (1 byte = 1 character)
    nName = StrConv(NewFont.Name & Chr$(0), vbFromUnicode)
    nSize = UBound(nName)
    If nSize > LF_FACESIZE Then nSize = LF_FACESIZE
    For i = 0 To nSize
        tFont.lfFaceName(i) = nName(i)
    Next

    With tFont
        .lfCharSet = NewFont.Charset
        .lfClipPrecision = CLIP_DEFAULT_PRECIS
        .lfEscapement = 0                       ' Angle to print
        .lfOrientation = .lfEscapement
        .lfWidth = 0#
        .lfItalic = IIf(NewFont.Italic, 1, 0)
        .lfOutPrecision = OUT_DEFAULT_PRECIS
        .lfPitchAndFamily = DEFAULT_PITCH
        .lfQuality = DEFAULT_QUALITY
        .lfStrikeOut = IIf(NewFont.Strikethrough, 1, 0)
        .lfUnderline = IIf(NewFont.Underline, 1, 0)
        .lfWeight = NewFont.Weight
        ' Font size (height) has to be calculated
        .lfHeight = MulDiv(NewFont.Size, GetDeviceCaps(UpDC, LOGPIXELSY), 72)
    End With

    ' Set environment (remember previous settings)
    If UpMemoryFont <> 0 Then
        ' Reset environment
        Call SelectObject(UpDC, UpOrginalFont)
        Call DeleteObject(UpMemoryFont)
    End If
    UpMemoryFont = CreateFontIndirect(tFont)
    UpOrginalFont = SelectObject(UpDC, UpMemoryFont)
End Property

Private Property Set DOWNFont(ByVal NewFont As StdFont)
    If Not IsDownCreated Then Exit Property

    On Local Error Resume Next

    Dim nName() As Byte, i As Byte, nSize As Byte
    Dim tFont As LOGFONT

    ' Font name is a byte array and is in ANSI (DOS) format (1 byte = 1 character)
    nName = StrConv(NewFont.Name & Chr$(0), vbFromUnicode)
    nSize = UBound(nName)
    If nSize > LF_FACESIZE Then nSize = LF_FACESIZE
    For i = 0 To nSize
        tFont.lfFaceName(i) = nName(i)
    Next

    With tFont
        .lfCharSet = NewFont.Charset
        .lfClipPrecision = CLIP_DEFAULT_PRECIS
        .lfEscapement = 0                       ' Angle to print
        .lfOrientation = .lfEscapement
        .lfWidth = 0#
        .lfItalic = IIf(NewFont.Italic, 1, 0)
        .lfOutPrecision = OUT_DEFAULT_PRECIS
        .lfPitchAndFamily = DEFAULT_PITCH
        .lfQuality = DEFAULT_QUALITY
        .lfStrikeOut = IIf(NewFont.Strikethrough, 1, 0)
        .lfUnderline = IIf(NewFont.Underline, 1, 0)
        .lfWeight = NewFont.Weight
        ' Font size (height) has to be calculated
        .lfHeight = MulDiv(NewFont.Size, GetDeviceCaps(DownDC, LOGPIXELSY), 72)
    End With

    ' Set environment (remember previous settings)
    If DownMemoryFont <> 0 Then
        ' Reset environment
        Call SelectObject(DownDC, DownOrginalFont)
        Call DeleteObject(DownMemoryFont)
    End If
    DownMemoryFont = CreateFontIndirect(tFont)
    DownOrginalFont = SelectObject(DownDC, DownMemoryFont)
End Property
Private Function CreateDOWN(hParentDC As Long, Optional PixelWidth As Long = 1024, Optional PixelHeight As Long = 768) As Long
    Dim nHasPalette As Long
    Dim nPaletteSize As Long
    Dim LogPal As LOGPALETTE
    Dim tm As TEXTMETRIC
    Dim sFaceName As String * 80
    Dim fFont As StdFont

    If IsDownCreated Then DestroyDown

    DownParent = hParentDC
    DownWidth = PixelWidth
    DownHeight = PixelHeight

    ' Create a memory device context to use
    DownDC = CreateCompatibleDC(DownParent)

    ' Tell'em it's a picture (so drawings can be done on the DC)
    DownMemBitmap = CreateCompatibleBitmap(DownParent, DownWidth, DownHeight)
    DownBitmap = SelectObject(DownDC, DownMemBitmap)

    ' Get screen properties
    nHasPalette = GetDeviceCaps(DownParent, RASTERCAPS) And RC_PALETTE   ' Palette sDOWNport
    nPaletteSize = GetDeviceCaps(DownParent, SIZEPALETTE)                ' Size of palette
    ' If the screen has a palette make a copy and realize it
    If nHasPalette And (nPaletteSize = 256) Then
        ' Create a copy of the system palette
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        Call GetSystemPaletteEntries(DownParent, 0&, 256, LogPal.palPalEntry(0))
        DownMemPal = CreatePalette(LogPal)
        ' Select the new palette into the memory DC and realize it
        DownPal = SelectPalette(DownDC, DownMemPal, 0&)
        Call RealizePalette(DownDC)
    End If
    Call SetBkColor(DownDC, GetBkColor(DownParent))
    DownColor = GetTextColor(DownParent)
    Call SetBkMode(DownDC, GetBkMode(DownParent))
    
    Call GetTextMetrics(DownDC, tm)
    Call GetTextFace(DownParent, 79, sFaceName)
    Set fFont = New StdFont
    With fFont
        .Bold = (tm.tmWeight > FW_NORMAL)
        .Charset = tm.tmCharSet
        .Italic = (tm.tmItalic <> 0)
        .Name = sFaceName
        .Strikethrough = (tm.tmStruckOut <> 0)
        .Underline = (tm.tmUnderlined <> 0)
        .Weight = tm.tmWeight
        .Size = (tm.tmMemoryHeight / tm.tmDigitizedAspectY) * 72 ' Size has to be calculated
    End With
    Set DOWNFont = fFont
    Set fFont = Nothing
    
    CreateDOWN = DownDC
End Function
Private Sub DestroyUP()
    If Not IsUpCreated Then Exit Sub

    '
    Call SelectObject(UpDC, UpBitmap)
    Call DeleteObject(UpMemBitmap)
    Call DeleteDC(UpDC)
    '
    UpDC = -1
End Sub

Private Sub DestroyDown()
    If Not IsDownCreated Then Exit Sub

    '
    Call SelectObject(DownDC, DownBitmap)
    Call DeleteObject(DownMemBitmap)
    Call DeleteDC(DownDC)
    '
    DownDC = -1
End Sub

Public Property Get hdcUP() As Long
    hdcUP = UpDC
End Property

Public Property Get hdcDOWN() As Long
    hdcDOWN = DownDC
End Property



Public Sub ClsUP(cColor As Long)
    Dim hBrush As Long
    Dim tRect As RECT

    hBrush = CreateSolidBrush(cColor)
    With tRect
        .Left = 0
        .Top = 0
        .Right = UpWidth
        .Bottom = UpHeight
    End With

    Call FillRect(UpDC, tRect, hBrush)
    Call DeleteObject(hBrush)
End Sub
Public Sub ClsDOWN(cColor As Long)
    Dim hBrush As Long
    Dim tRect As RECT

    hBrush = CreateSolidBrush(cColor)
    With tRect
        .Left = 0
        .Top = 0
        .Right = DownWidth
        .Bottom = DownHeight
    End With

    Call FillRect(DownDC, tRect, hBrush)
    Call DeleteObject(hBrush)
End Sub








