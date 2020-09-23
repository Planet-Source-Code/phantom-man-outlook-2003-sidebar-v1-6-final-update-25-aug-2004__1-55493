Attribute VB_Name = "mGlobalUtil"
Option Explicit
Private Type POINTAPI
    x                                                      As Long
    y                                                      As Long
End Type
Private Type RECT
    left                                                   As Long
    top                                                    As Long
    right                                                  As Long
    bottom                                                 As Long
End Type
Public m_sCurrentSystemThemename                         As String

Private Const PS_SOLID                                   As Integer = 0
'Public Const TRANSPARENT                                 As Integer = 1
Private Const DT_LEFT                                    As Long = &H0&

'Public Const CLR_INVALID                                 As Integer = -1
Private Type OSVERSIONINFO
    dwVersionInfoSize                                      As Long
    dwMajorVersion                                         As Long
    dwMinorVersion                                         As Long
    dwBuildNumber                                          As Long
    dwPlatformId                                           As Long
    szCSDVersion(0 To 127)                                 As Byte
End Type

Public Const IDC_HAND = 32649&
Public Const IDC_ARROW = 32512&

Private Const VER_PLATFORM_WIN32_NT                      As Integer = 2
Private Type TRIVERTEX
    x                                                      As Long
    y                                                      As Long
    Red                                                    As Integer
    Green                                                  As Integer
    Blue                                                   As Integer
    Alpha                                                  As Integer
End Type
Private Type GRADIENT_RECT
    UpperLeft                                              As Long
    LowerRight                                             As Long
End Type
Private Type GRADIENT_TRIANGLE
    Vertex1                                                As Long
    Vertex2                                                As Long
    Vertex3                                                As Long
End Type
Public Enum GradientFillRectType
    GRADIENT_FILL_RECT_H = 0
    GRADIENT_FILL_RECT_V = 1
End Enum
#If False Then    'Trick preserves Case of Enums when typing in IDE
    Private GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V
#End If
Private Const LF_FACESIZE                                As Integer = 32
Public Type LOGFONT
    lfHeight                                               As Long    ' The font size (see below)
    lfWidth                                                As Long    ' Normally you don't set this, just let Windows create the default
    lfEscapement                                           As Long    ' The angle, in 0.1 degrees, of the font
    lfOrientation                                          As Long    ' Leave as default
    lfWeight                                               As Long    ' Bold, Extra Bold, Normal etc
    lfItalic                                               As Byte    ' As it says
    lfUnderline                                            As Byte    ' As it says
    lfStrikeOut                                            As Byte    ' As it says
    lfCharSet                                              As Byte    ' As it says
    lfOutPrecision                                         As Byte    ' Leave for default
    lfClipPrecision                                        As Byte    ' Leave for default
    lfQuality                                              As Byte    ' Leave as default (see end of article)
    lfPitchAndFamily                                       As Byte    ' Leave as default (see end of article)
    lfFaceName(LF_FACESIZE)                                As Byte    ' The font name converted to a byte array
End Type

Public m_bHandCursor                                    As Boolean
Private m_bIsXp                                          As Boolean
Private m_bIsNt                                          As Boolean
Private m_bIs2000OrAbove                                 As Boolean
Private m_bHasGradientAndTransparency                    As Boolean

Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
        lpRect As RECT) As Long

Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, ByVal dwMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
        lpDeviceName As Any, _
        lpOutput As Any, _
        lpInitData As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
        ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
        ByVal nCount As Long, _
        lpObject As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
        ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
        ByVal nWidth As Long, _
        ByVal crColor As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
        ByVal crColor As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, _
        ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, _
        ByVal nBkMode As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, _
        ByVal lpStr As String, _
        ByVal nCount As Long, _
        lpRect As RECT, _
        ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, _
        ByVal lpStr As Long, _
        ByVal nCount As Long, _
        lpRect As RECT, _
        ByVal wFormat As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, _
        ByVal X1 As Long, _
        ByVal y1 As Long, _
        ByVal x2 As Long, _
        ByVal y2 As Long) As Long

'Public Declare Function InflateRect Lib "user32" (lpRect As RECT, _
 ByVal x As Long, _
 ByVal y As Long) As Long
'Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
 ByVal x As Long, _
 ByVal y As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, _
        ByVal ptX As Long, _
        ByVal ptY As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
        lpRect As RECT, _
        ByVal hBrush As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, _
        lpvSource As Any, _
        ByVal cbCopy As Long)
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
        ByVal HPALETTE As Long, _
        pccolorref As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long
Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, _
        pVertex As TRIVERTEX, _
        ByVal dwNumVertex As Long, _
        pMesh As GRADIENT_RECT, _
        ByVal dwNumMesh As Long, _
        ByVal dwMode As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
        ByVal x As Long, _
        ByVal y As Long) As Long
'Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
        lpPoint As POINTAPI) As Long
'Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, _
 lpPoint As POINTAPI) As Long
Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long

Public Function AppThemed() As Boolean

    On Error Resume Next
    AppThemed = IsAppThemed()
    On Error GoTo 0

End Function

Public Property Get dBlendColor(ByVal oColorFrom As OLE_COLOR, _
                               ByVal oColorTo As OLE_COLOR, _
                               Optional ByVal Alpha As Long = 128) As Long

    Dim lSrcR  As Long

    Dim lSrcG  As Long
    Dim lSrcB  As Long
    Dim lDstR  As Long
    Dim lDstG  As Long
    Dim lDstB  As Long
    Dim lCFrom As Long
    Dim lCTo   As Long
    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
    dBlendColor = RGB(((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))

End Property

Private Sub GradientFillRect(ByVal lHDC As Long, _
                             tR As RECT, _
                             ByVal oStartColor As OLE_COLOR, _
                             ByVal oEndColor As OLE_COLOR, _
                             ByVal eDir As GradientFillRectType)

    Dim tTV(0 To 1) As TRIVERTEX
    Dim tGR         As GRADIENT_RECT
    Dim hBrush      As Long
    Dim lStartColor As Long
    Dim lEndColor   As Long

    'Dim lR As Long
    ' Use GradientFill:
    If (HasGradientAndTransparency) Then
        lStartColor = TranslateColor(oStartColor)
        lEndColor = TranslateColor(oEndColor)
        setTriVertexColor tTV(0), lStartColor
        tTV(0).x = tR.left
        tTV(0).y = tR.top
        setTriVertexColor tTV(1), lEndColor
        tTV(1).x = tR.right
        tTV(1).y = tR.bottom
        tGR.UpperLeft = 0
        tGR.LowerRight = 1
        GradientFill lHDC, tTV(0), 2, tGR, 1, eDir
    Else
        ' Fill with solid brush:
        hBrush = CreateSolidBrush(TranslateColor(oEndColor))
        FillRect lHDC, tR, hBrush
        DeleteObject hBrush
    End If

End Sub

Public Property Get HasGradientAndTransparency()

    HasGradientAndTransparency = m_bHasGradientAndTransparency

End Property

Public Property Get Is2000OrAbove() As Boolean

    Is2000OrAbove = m_bIs2000OrAbove

End Property

Public Property Get IsNt() As Boolean

    IsNt = m_bIsNt

End Property

Public Property Get IsXp() As Boolean

    IsXp = m_bIsXp

End Property

Public Function lTranslateColor(ByVal oClr As OLE_COLOR, _
                               Optional hPal As Long = 0) As Long

' Convert Automation color to Windows color

    If OleTranslateColor(oClr, hPal, lTranslateColor) Then
        lTranslateColor = CLR_INVALID
    End If

End Function

Private Sub setTriVertexColor(tTV As TRIVERTEX, _
                              ByVal lColor As Long)


    Dim lRed   As Long
    Dim lGreen As Long
    Dim lBlue  As Long

    lRed = (lColor And &HFF&) * &H100&
    lGreen = (lColor And &HFF00&)
    lBlue = (lColor And &HFF0000) \ &H100&
    With tTV
        setTriVertexColorComponent .Red, lRed
        setTriVertexColorComponent .Green, lGreen
        setTriVertexColorComponent .Blue, lBlue
    End With    'tTV

End Sub

Private Sub setTriVertexColorComponent(ByRef iColor As Integer, _
                                       ByVal lComponent As Long)

    If (lComponent And &H8000&) = &H8000& Then
        iColor = (lComponent And &H7F00&)
        iColor = iColor Or &H8000
    Else
        iColor = lComponent
    End If

End Sub

Public Sub UtilDrawBackground(ByVal lngHdc As Long, _
                              ByVal colorStart As Long, _
                              ByVal colorEnd As Long, _
                              ByVal lngLeft As Long, _
                              ByVal lngTop As Long, _
                              ByVal lngWidth As Long, _
                              ByVal lngHeight As Long, _
                              Optional ByVal horizontal As Boolean = False)


    Dim tR As RECT

    With tR
        .left = lngLeft
        .top = lngTop
        .right = lngLeft + lngWidth
        .bottom = lngTop + lngHeight
        ' gradient fill vertical:
    End With    'tR
    GradientFillRect lngHdc, tR, colorStart, colorEnd, IIf(horizontal, GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V)

End Sub

Public Sub UtilDrawBorderRectangle(ByVal lngHdc As Long, _
                                   ByVal lColor As Long, _
                                   ByVal lngLeft As Long, _
                                   ByVal lngTop As Long, _
                                   ByVal lngWidth As Long, _
                                   ByVal lngHeight As Long, _
                                   ByVal bInset As Boolean)


    Dim tJ      As POINTAPI
    Dim hPen    As Long
    Dim hPenOld As Long

    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(lngHdc, hPen)
    MoveToEx lngHdc, lngLeft, lngTop + lngHeight - 1, tJ
    LineTo lngHdc, lngLeft, lngTop
    LineTo lngHdc, lngLeft + lngWidth - 1, lngTop
    LineTo lngHdc, lngLeft + lngWidth - 1, lngTop + lngHeight - 1
    LineTo lngHdc, lngLeft, lngTop + lngHeight - 1
    SelectObject lngHdc, hPenOld
    DeleteObject hPen

End Sub

Public Sub UtilDrawText(ByVal lngHdc As Long, _
                        ByVal sCaption As String, _
                        ByVal lTextX As Long, _
                        ByVal lTextY As Long, _
                        ByVal lTextX1 As Long, _
                        ByVal lTextY1 As Long, _
                        ByVal bEnabled As Boolean, _
                        ByVal color As Long, _
                        ByVal bCentreHorizontal As Boolean, _
                        Optional RightAlign As Boolean = False)


    Dim rcText As RECT

    SetTextColor lngHdc, TranslateColor(color)
    'Dim lFlags As Long
    If Not bEnabled Then
        SetTextColor lngHdc, GetSysColor(vbGrayText And &H1F&)
    End If
    
    With rcText
        .left = lTextX
        .top = lTextY
        .right = lTextX1
        .bottom = lTextY1
    End With
    If m_bIsNt Then
        DrawTextW lngHdc, StrPtr(sCaption), -1, rcText, IIf(RightAlign, DT_RIGHT, DT_LEFT) Or DT_END_ELLIPSIS
    Else
        DrawTextA lngHdc, sCaption, -1, rcText, IIf(RightAlign, DT_RIGHT, DT_LEFT) Or DT_END_ELLIPSIS
    End If
    If Not bEnabled Then
        SetTextColor lngHdc, GetSysColor(color)
    End If

End Sub

Public Sub VerInitialise()

    Dim tOSV As OSVERSIONINFO

    tOSV.dwVersionInfoSize = Len(tOSV)
    GetVersionEx tOSV
    m_bIsNt = ((tOSV.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
    If (tOSV.dwMajorVersion > 5) Then
        m_bHasGradientAndTransparency = True
        m_bIsXp = True
        m_bIs2000OrAbove = True
    ElseIf (tOSV.dwMajorVersion = 5) Then
        m_bHasGradientAndTransparency = True
        m_bIs2000OrAbove = True
        If (tOSV.dwMinorVersion >= 1) Then
            m_bIsXp = True
        End If
    ElseIf (tOSV.dwMajorVersion = 4) Then    ' NT4 or 9x/ME/SE
        If (tOSV.dwMinorVersion >= 10) Then
            m_bHasGradientAndTransparency = True
        End If
    Else    ' Too old
    End If

End Sub

Public Sub yDrawText(ByVal lHDC As Long, _
                    ByVal sText As String, _
                    ByVal lLength As Long, _
                    tR As RECT, _
                    ByVal lFlags As Long)


    Dim lPtr As Long

    If (m_bIsNt) Then
        lPtr = StrPtr(sText)
        If Not (lPtr = 0) Then    ' NT4 crashes with ptr = 0
            DrawTextW lHDC, lPtr, -1, tR, lFlags
        End If
    Else
        DrawTextA lHDC, sText, -1, tR, lFlags
    End If

End Sub

Public Sub DrawChev(hwnd As Long, lHDC As Long, x As Long, y As Long, X1 As Long, y1 As Long, bEnabled As Boolean, Optional lColour As Long = vbBlack)
    Dim tWR As RECT
    Dim tTR As RECT
    Dim rc As RECT
    Dim tR As RECT
    Dim hPen As Long
    Dim hPenOld As Long
    Dim tJunk As POINTAPI

    tR.top = y    ' - 1
    tR.bottom = y1
    tR.left = x
    tR.right = X1

    '
    LSet tWR = tR
    tWR.top = tWR.bottom - 20
    tWR.right = tWR.right - 2
    tWR.bottom = tWR.bottom - 1


    ' draw the chevron:
    If Not bEnabled Then
        hPen = CreatePen(PS_SOLID, 1, TranslateColor(vbGrayText))
        Else
        hPen = CreatePen(PS_SOLID, 1, TranslateColor(lColour))
    End If
    hPenOld = SelectObject(lHDC, hPen)
    LSet tTR = tWR
    tTR.left = ((tTR.right - tTR.left) \ 2) - 3 + tTR.left
    tTR.top = tTR.top + 2

    MoveToEx lHDC, tTR.left, tTR.top, tJunk
    LineTo lHDC, tTR.left + 3, tTR.top + 3
    MoveToEx lHDC, tTR.left, tTR.top + 1, tJunk
    LineTo lHDC, tTR.left + 3, tTR.top + 3 + 1

    MoveToEx lHDC, tTR.left, tTR.top + 4, tJunk
    LineTo lHDC, tTR.left + 3, tTR.top + 3 + 4
    MoveToEx lHDC, tTR.left, tTR.top + 1 + 4, tJunk
    LineTo lHDC, tTR.left + 3, tTR.top + 3 + 1 + 4

    MoveToEx lHDC, tTR.left + 4, tTR.top, tJunk
    LineTo lHDC, tTR.left + 4 - 3, tTR.top + 3
    MoveToEx lHDC, tTR.left + 4, tTR.top + 1, tJunk
    LineTo lHDC, tTR.left + 4 - 3, tTR.top + 3 + 1

    MoveToEx lHDC, tTR.left + 4, tTR.top + 4, tJunk
    LineTo lHDC, tTR.left + 4 - 3, tTR.top + 3 + 4
    MoveToEx lHDC, tTR.left + 4, tTR.top + 1 + 4, tJunk
    LineTo lHDC, tTR.left + 4 - 3, tTR.top + 3 + 1 + 4


    SelectObject lHDC, hPenOld
    DeleteObject hPen

End Sub

Public Sub DrawDimple(hdc As Long, ByVal x As Long, ByVal y As Long, Optional ByVal Raised As Boolean = False)
    Dim RT As RECT
    Raised = True
    SetRect RT, x + 1, y + 1, x + 3, y + 3
    DrawSolidRect hdc, RT, vb3DHighlight
    SetRect RT, x, y, x + 2, y + 2
    DrawSolidRect hdc, RT, dBlendColor(vbActiveTitleBar, vbBlack, 100)
End Sub

Private Sub DrawSolidRect(hdc As Long, RT As RECT, Optional FillColor As Long = 0, Optional BorderColor As Long = -1, Optional ByVal BorderWidth As Long = 1, Optional ByVal RoundingW As Long = 0, Optional ByVal RoundingH As Long = 0)
    Dim lBrush    As Long, _
            lPen      As Long, _
            oldBrush  As Long, _
            oldPen    As Long

    
    If BorderColor = -1 Then BorderColor = FillColor
    lBrush = CreateSolidBrush(TranslateColor(FillColor))
    lPen = CreatePen(PS_SOLID, BorderWidth, TranslateColor(BorderColor))
    oldPen = SelectObject(hdc, lPen)
    oldBrush = SelectObject(hdc, lBrush)
    If (RoundingW <> 0) Or (RoundingH <> 0) Then
        RoundRect hdc, RT.left, RT.top, RT.right, RT.bottom, RoundingW, RoundingH
    Else
        Rectangle hdc, RT.left, RT.top, RT.right, RT.bottom
    End If
    lBrush = SelectObject(hdc, oldBrush)
    lPen = SelectObject(hdc, oldPen)
    DeleteObject lBrush: DeleteObject lPen
End Sub

' Desc: Get the "Real" Hand Cursor
Public Sub UtilSetCursor(bHand As Boolean)
    If bHand = True Then
        SetCursor LoadCursor(0, IDC_HAND)
        m_bHandCursor = True
    Else

        SetCursor LoadCursor(0, IDC_ARROW)
        m_bHandCursor = False

    End If
End Sub

Public Sub GetThemeName(hwnd As Long)
    'Gett the current Theme name, ans Scheme Color
    Dim hTheme As Long
    Dim sShellStyle As String
    Dim sThemeFile As String
    Dim lPtrThemeFile As Long, lPtrColorName As Long, hres As Long
    Dim iPos As Long
    On Error Resume Next
    hTheme = OpenThemeData(hwnd, StrPtr("ExplorerBar"))
   
   If Not hTheme = 0 Then
      ReDim bThemeFile(0 To 260 * 2) As Byte
      lPtrThemeFile = VarPtr(bThemeFile(0))
      ReDim bColorName(0 To 260 * 2) As Byte
      lPtrColorName = VarPtr(bColorName(0))
      hres = GetCurrentThemeName(lPtrThemeFile, 260, lPtrColorName, 260, 0, 0)
      
      sThemeFile = bThemeFile
      iPos = InStr(sThemeFile, vbNullChar)
      If (iPos > 1) Then sThemeFile = left(sThemeFile, iPos - 1)
      m_sCurrentSystemThemename = bColorName
      iPos = InStr(m_sCurrentSystemThemename, vbNullChar)
      If (iPos > 1) Then m_sCurrentSystemThemename = left(m_sCurrentSystemThemename, iPos - 1)
      
      sShellStyle = sThemeFile
      For iPos = Len(sThemeFile) To 1 Step -1
         If (Mid(sThemeFile, iPos, 1) = "\") Then
            sShellStyle = left(sThemeFile, iPos)
            Exit For
         End If
      Next iPos
      sShellStyle = sShellStyle & "Shell\" & m_sCurrentSystemThemename & "\ShellStyle.dll"
      CloseThemeData hTheme
    Else
        m_sCurrentSystemThemename = "Classic"
    End If
Debug.Print m_sCurrentSystemThemename

End Sub

