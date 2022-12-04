Attribute VB_Name = "mDeclares"
Option Explicit

Public Const TOOLWINDOWPARENTWINDOWHWND = "EDC:ToolWindow:ParenthWnd"
Public Const EDCCHEVRONMENUCONST = &H56291024

Public Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

' =======================================================================
' MENU Declares:
' =======================================================================
' Menu information:
Public Type tMenuItem
    sHelptext As String
    sInputCaption As String
    sCaption As String
    sAccelerator As String
    sShortCutDisplay As String
    iShortCutShiftMask As Integer
    iShortCutShiftKey As Integer
    lID As Long
    lActualID As Long    ' The ID gets modified if we add a sub-menu to the hMenu of the popup
    lItemData As Long
    lIndex As Long
    lParentId As Long
    lIconIndex As Long
    bChecked As Boolean
    bRadioCheck As Boolean
    bEnabled As Boolean
    hMenu As Long
    lHeight As Long
    lWidth As Long
    bCreated As Boolean
    bIsAVBMenu As Boolean
    lShortCutStartPos As Long
    bMarkToDestroy As Boolean
    sKey As String
    lParentIndex As Long
    bTitle As Boolean
    bDefault As Boolean
    bOwnerDraw As Boolean
    bMenuBarBreak As Boolean
    bMenuBreak As Boolean
    bVisible As Boolean
    bDragOff As Boolean
    bInfrequent As Boolean
    bTextBox As Boolean
    bComboBox As Boolean
    bChevronAppearance As Boolean
    bChevronBehaviour As Boolean
    bShowCheckAndIcon As Boolean
End Type

' Menu flag constants:
Public Const MF_APPEND = &H100&
Public Const MF_BITMAP = &H4&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const MF_CALLBACKS = &H8000000
Public Const MF_CHANGE = &H80&
Public Const MF_CHECKED = &H8&
Public Const MF_CONV = &H40000000
Public Const MF_DELETE = &H200&
Public Const MF_DISABLED = &H2&
Public Const MF_ENABLED = &H0&
Public Const MF_END = &H80
Public Const MF_ERRORS = &H10000000
Public Const MF_GRAYED = &H1&
Public Const MF_HELP = &H4000&
Public Const MF_HILITE = &H80&
Public Const MF_HSZ_INFO = &H1000000
Public Const MF_INSERT = &H0&
Public Const MF_LINKS = &H20000000
Public Const MF_MASK = &HFF000000
Public Const MF_MENUBARBREAK = &H20&
Public Const MF_MENUBREAK = &H40&
Public Const MF_MOUSESELECT = &H8000&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_POPUP = &H10&
Public Const MF_POSTMSGS = &H4000000
Public Const MF_REMOVE = &H1000&
Public Const MF_SENDMSGS = &H2000000
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const MF_SYSMENU = &H2000&
Public Const MF_UNCHECKED = &H0&
Public Const MF_UNHILITE = &H0&
Public Const MF_USECHECKBITMAPS = &H200&
Public Const MF_DEFAULT = &H1000&

Public Const MFT_STRING = MF_STRING
Public Const MFT_BITMAP = MF_BITMAP
Public Const MFT_MENUBARBREAK = MF_MENUBARBREAK
Public Const MFT_MENUBREAK = MF_MENUBREAK
Public Const MFT_OWNERDRAW = MF_OWNERDRAW
Public Const MFT_RADIOCHECK = &H200&
Public Const MFT_SEPARATOR = MF_SEPARATOR
Public Const MFT_RIGHTORDER = &H2000&
'Public Const MFT_RIGHTJUSTIFY = MF_RIGHTJUSTIFY

' New versions of the names...
Public Const MFS_GRAYED = &H3&
Public Const MFS_DISABLED = MFS_GRAYED
Public Const MFS_CHECKED = MF_CHECKED
Public Const MFS_HILITE = MF_HILITE
Public Const MFS_ENABLED = MF_ENABLED
Public Const MFS_UNCHECKED = MF_UNCHECKED
Public Const MFS_UNHILITE = MF_UNHILITE
Public Const MFS_DEFAULT = MF_DEFAULT

Public Const MIIM_STATE = &H1&
Public Const MIIM_ID = &H2&
Public Const MIIM_SUBMENU = &H4&
Public Const MIIM_CHECKMARKS = &H8&
Public Const MIIM_TYPE = &H10&
Public Const MIIM_DATA = &H20&

' Track popup menu constants:
Public Const TPM_CENTERALIGN = &H4&
Public Const TPM_LEFTALIGN = &H0&
Public Const TPM_LEFTBUTTON = &H0&
Public Const TPM_RIGHTALIGN = &H8&
Public Const TPM_RIGHTBUTTON = &H2&

Public Const TPM_NONOTIFY = &H80&    '/* Don't send any notification msgs */
Public Const TPM_RETURNCMD = &H100
Public Const TPM_HORIZONTAL = &H0    '/* Horz alignment matters more */
Public Const TPM_VERTICAL = &H40    '/* Vert alignment matters more */

Public Const TPM_RECURSE = &H1
Public Const TPM_HORPOSANIMATION = &H400&
Public Const TPM_HORNEGANIMATION = &H800&
Public Const TPM_VERPOSANIMATION = &H1000&
Public Const TPM_VERNEGANIMATION = &H2000&
Public Const TPM_NOANIMATION = &H4000&

' Owner draw information:
Public Const ODS_CHECKED = &H8
Public Const ODS_DISABLED = &H4
Public Const ODS_FOCUS = &H10
Public Const ODS_GRAYED = &H2
Public Const ODS_SELECTED = &H1
Public Const ODT_BUTTON = 4
Public Const ODT_COMBOBOX = 3
Public Const ODT_LISTBOX = 2
Public Const ODT_MENU = 1

Public Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemWidth As Long
    itemHeight As Long
    ItemData As Long
End Type

Public Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECT
    ItemData As Long
End Type

Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
End Type

Public Type MENUITEMTEMPLATE
    mtOption As Integer
    mtID As Integer
    mtString As Byte
End Type
Public Type MENUITEMTEMPLATEHEADER
    versionNumber As Integer
    offset As Integer
End Type

Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long

Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
Public Declare Function GetMenuContextHelpId Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function GetMenuItemRect Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
Public Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long

Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long

Public Declare Function AppendMenuBylong Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Public Declare Function AppendMenuByString Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Declare Function ModifyMenuByLong Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function InsertMenuByLong Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Public Declare Function InsertMenuByString Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByVal lpcMenuItemInfo As MENUITEMINFO) As Long

Public Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Public Declare Function CheckMenuRadioItem Lib "user32" (ByVal hMenu As Long, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Long
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Public Declare Function HiliteMenuItem Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal wIDHiliteItem As Long, ByVal wHilite As Long) As Long

Public Declare Function MenuItemFromPoint Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal ptScreen As POINTAPI) As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long
Public Declare Function TrackPopupMenuByLong Lib "user32" Alias "TrackPopupMenu" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Type TPMPARAMS
    cbSize As Long
    rcExclude As RECT
End Type
Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal un As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal hwnd As Long, lpTPMParams As TPMPARAMS) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long

' =======================================================================
' General Window Declares
' =======================================================================
'
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageAsAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function getActiveWindow Lib "user32" Alias "GetActiveWindow" () As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_ABSOLUTE = &H8000    '  absolute move
Public Const MOUSEEVENTF_LEFTDOWN = &H2    '  left button down
Public Const MOUSEEVENTF_LEFTUP = &H4    '  left button up
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20    '  middle button down
Public Const MOUSEEVENTF_MIDDLEUP = &H40    '  middle button up
Public Const MOUSEEVENTF_MOVE = &H1    '  mouse move
Public Const MOUSEEVENTF_RIGHTDOWN = &H8    '  right button down
Public Const MOUSEEVENTF_RIGHTUP = &H10    '  right button up


' =======================================================================
' GDI Declares:
' =======================================================================
' GDI object functions:
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
        lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Const BITSPIXEL = 12
Public Const LOGPIXELSX = 88    '  Logical pixels/inch in X
Public Const LOGPIXELSY = 90    '  Logical pixels/inch in Y
' System metrics:
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXICON = 11
Public Const SM_CYICON = 12
Public Const SM_CXFRAME = 32
Public Const SM_CYCAPTION = 4
Public Const SM_CYFRAME = 33
Public Const SM_CYBORDER = 6
Public Const SM_CXBORDER = 5
Public Const SM_CYMENU = 15

' Region paint and fill functions:
Public Declare Function PaintRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Public Const FLOODFILLBORDER = 0
Public Const FLOODFILLSURFACE = 1

' Pen functions:
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Const PS_DASH = 1
Public Const PS_DASHDOT = 3
Public Const PS_DASHDOTDOT = 4
Public Const PS_DOT = 2
Public Const PS_SOLID = 0
Public Const PS_NULL = 5

' Brush functions:
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

' Line functions:
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function DrawEdgeAPI Lib "user32" Alias "DrawEdge" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Const BF_LEFT = &H1
Public Const BF_BOTTOM = &H8
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BDR_INNER = &HC
Public Const BDR_OUTER = &H3
Public Const BDR_RAISED = &H5
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKEN = &HA
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

' Colour functions:
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Const OPAQUE = 2
Public Const TRANSPARENT = 1
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_ADJ_MAX = 100
Public Const COLOR_ADJ_MIN = -100
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_MENU = 4
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_WINDOWTEXT = 8
Public Const COLORONCOLOR = 3

' Shell Extract icon functions:
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

' GDI icon functions:
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

' Blitting functions
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020
Public Const SRCINVERT = &H660046
Public Const BLACKNESS = &H42
Public Const WHITENESS = &HFF0062
Public Const SRCAND = &H8800C6
Public Const SRCERASE = &H440328
Public Const SRCPAINT = &HEE0086

Public Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function LoadBitmapBynum Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As Long) As Long
Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Long
    bmBitsPixel As Integer
    bmBits As Long
End Type
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function LoadImageByNum Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADTRANSPARENT = &H20
Public Const IMAGE_BITMAP = 0

' Text functions:
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptX As Long, ByVal ptY As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Const DT_BOTTOM = &H8
Public Const DT_CENTER = &H1
Public Const DT_LEFT = &H0
Public Const DT_CALCRECT = &H400
Public Const DT_WORDBREAK = &H10
Public Const DT_VCENTER = &H4
Public Const DT_TOP = &H0
Public Const DT_TABSTOP = &H80
Public Const DT_SINGLELINE = &H20
Public Const DT_RIGHT = &H2
Public Const DT_NOCLIP = &H100
Public Const DT_INTERNAL = &H1000
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_EXPANDTABS = &H40
Public Const DT_CHARSTREAM = 4
Public Const DT_EDITCONTROL = &H2000&
Public Const DT_PATH_ELLIPSIS = &H4000&
Public Const DT_END_ELLIPSIS = &H8000&
Public Const DT_MODIFYSTRING = &H10000
Public Const DT_RTLREADING = &H20000
Public Const DT_WORD_ELLIPSIS = &H40000

Public Declare Function GrayString Lib "user32" Alias "GrayStringA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpOutputFunc As Long, ByVal lpData As Long, ByVal nCount As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
Public Const DI_MASK = 1
Public Const DI_IMAGE = 2
Public Const DI_NORMAL = 3
Public Const DI_COMPAT = 4
Public Const DI_DEFAULTSIZE = 8

Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SW_SHOWNOACTIVATE = 4

' Scrolling and region functions:
Public Declare Function ScrollDC Lib "user32" (ByVal hdc As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Public Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Public Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long)
Public Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal hSavedDC As Long) As Long

Public Const LF_FACESIZE = 32
Public Type LOGFONT
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
Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700
Public Const FF_DONTCARE = 0
Public Const DEFAULT_QUALITY = 0
Public Const DEFAULT_PITCH = 0
Public Const DEFAULT_CHARSET = 1
Public Declare Function CreateFontIndirect& Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT)
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Public Declare Function DrawState Lib "user32" Alias "DrawStateA" _
        (ByVal hdc As Long, _
        ByVal hBrush As Long, _
        ByVal lpDrawStateProc As Long, _
        ByVal lParam As Long, _
        ByVal wParam As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal cX As Long, _
        ByVal cY As Long, _
        ByVal fuFlags As Long) As Long
'/* Image type */
Public Const DST_COMPLEX = &H0
Public Const DST_TEXT = &H1
Public Const DST_PREFIXTEXT = &H2
Public Const DST_ICON = &H3
Public Const DST_BITMAP = &H4

' /* State type */
Public Const DSS_NORMAL = &H0
Public Const DSS_UNION = &H10    ' /* Gray string appearance */
Public Const DSS_DISABLED = &H20
Public Const DSS_MONO = &H80
Public Const DSS_RIGHT = &H8000

'/* flags for DrawFrameControl */
Public Enum DFCFlags
    DFC_CAPTION = 1
    DFC_MENU = 2
    DFC_SCROLL = 3
    DFC_BUTTON = 4
    'Win98/2000 only
    DFC_POPUPMENU = 5
End Enum

Public Enum DFCCaptionTypeFlags
' Caption types:
    DFCS_CAPTIONCLOSE = &H0&
    DFCS_CAPTIONMIN = &H1&
    DFCS_CAPTIONMAX = &H2&
    DFCS_CAPTIONRESTORE = &H3&
    DFCS_CAPTIONHELP = &H4&
End Enum
Public Enum DFCMenuTypeFlags
' Menu types:
    DFCS_MENUARROW = &H0&
    DFCS_MENUCHECK = &H1&
    DFCS_MENUBULLET = &H2&
    DFCS_MENUARROWRIGHT = &H4&
End Enum
Public Enum DFCScrollTypeFlags
' Scroll types:
    DFCS_SCROLLUP = &H0&
    DFCS_SCROLLDOWN = &H1&
    DFCS_SCROLLLEFT = &H2&
    DFCS_SCROLLRIGHT = &H3&
    DFCS_SCROLLCOMBOBOX = &H5&
    DFCS_SCROLLSIZEGRIP = &H8&
    DFCS_SCROLLSIZEGRIPRIGHT = &H10&
End Enum
Public Enum DFCButtonTypeFlags
' Button types:
    DFCS_BUTTONCHECK = &H0&
    DFCS_BUTTONRADIOIMAGE = &H1&
    DFCS_BUTTONRADIOMASK = &H2&
    DFCS_BUTTONRADIO = &H4&
    DFCS_BUTTON3STATE = &H8&
    DFCS_BUTTONPUSH = &H10&
End Enum
Public Enum DFCStateTypeFlags
' Styles:
    DFCS_INACTIVE = &H100&
    DFCS_PUSHED = &H200&
    DFCS_CHECKED = &H400&
    ' Win98/2000 only
    DFCS_TRANSPARENT = &H800&
    DFCS_HOT = &H1000&
    'End Win98/2000 only
    DFCS_ADJUSTRECT = &H2000&
    DFCS_FLAT = &H4000&
    DFCS_MONO = &H8000&
End Enum

Public Declare Function DrawFrameControl Lib "user32" (ByVal lHDC As Long, tR As RECT, ByVal eFlag As DFCFlags, ByVal eStyle As Long) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Const CLR_INVALID = -1

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type
Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpFn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Const WH_KEYBOARD As Long = 2
Private Const MSGF_MENU = 2
Private Const HC_ACTION = 0

' =======================================================================
' Image list Declares:
' =======================================================================
' Create/Destroy functions:
Declare Function ImageList_Create Lib "COMCTL32.DLL" ( _
        ByVal cX As Long, _
        ByVal cY As Long, _
        ByVal fMask As Long, _
        ByVal cInitial As Long, _
        ByVal cGrow As Long _
        ) As Long
Public Const ILC_MASK = 1&
Public Const ILC_COLOR = 0&
Public Const ILC_COLORDDB = &HFE&
Public Const ILC_COLOR4 = &H4&
Public Const ILC_COLOR8 = &H8&
Public Const ILC_COLOR16 = &H10&
Public Const ILC_COLOR24 = &H18&
Public Const ILC_COLOR32 = &H20&
Public Const ILC_PALETTE = &H800&

Declare Function ImageList_Destroy Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long _
        ) As Long

' Add functions:
Declare Function ImageList_Add Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal hBmp As Long, _
        ByVal hBmpMask As Long _
        ) As Long
Declare Function ImageList_AddMasked Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal hBmp As Long, _
        ByVal crMask As Long _
        ) As Long
Declare Function ImageList_AddIcon Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal hIcon As Long _
        ) As Long
Declare Function ImageList_LoadImage Lib "COMCTL32.DLL" ( _
        ByVal hInst As Long, _
        ByVal lpBmp As String, _
        ByVal cX As Long, _
        ByVal cGrow As Long, _
        ByVal crMask As Long, _
        ByVal uType As Long, _
        ByVal uFlags As Long _
        ) As Long

' Modification/deletion functions:
Declare Function ImageList_Remove Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long _
        ) As Long
Declare Function ImageList_Replace Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal hBmpImage As Long, _
        ByVal hBmpMask As Long _
        ) As Long
Declare Function ImageList_ReplaceIcon Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal hIcon As Long _
        ) As Long

' Image information functions:
Declare Function ImageList_GetImageCount Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long _
        ) As Long
Declare Function ImageList_GetImageRect Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        prcImage As RECT _
        ) As Long
Declare Function ImageList_GetIconSize Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal cX As Long, _
        ByVal cY As Long _
        ) As Long
Type IMAGEINFO
    hBitmapImage As Long
    hBitmapMask As Long
    cPlanes As Long
    cBitsPerPixel As Long
    rcImage As RECT
End Type
Declare Function ImageList_GetImageInfo Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        pImageInfo As IMAGEINFO _
        )

' Create a new icon based on an image list icon:
Declare Function ImageList_GetIcon Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal diIgnore As Long _
        ) As Long

' Merge and move functions:
Declare Function ImageList_Merge Lib "COMCTL32.DLL" ( _
        ByVal hIml1 As Long, _
        ByVal i As Long, _
        ByVal hIml2 As Long, _
        ByVal i2 As Long, _
        ByVal dx As Long, _
        ByVal dy As Long _
        ) As Long
Declare Sub ImageList_CopyDitherImage Lib "COMCTL32.DLL" ( _
        ByVal hImlDst As Long, _
        ByVal iDst As Integer, _
        ByVal xDst As Long, _
        ByVal yDst As Long, _
        ByVal hImlSrc As Long, _
        ByVal iSrc As Long _
        )
Declare Function ImageList_AddFromImageList Lib "COMCTL32.DLL" ( _
        ByVal hImlDest As Long, _
        ByVal hImlSrc As Long, _
        ByVal iSrc As Long _
        ) As Long

' Get/Set Background Colour:
Declare Function ImageList_SetBkColor Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal clrBk As Long _
        ) As Long
Public Const CLR_NONE = -1
Public Const CLR_DEFAULT = -16777216
Public Const CLR_HILIGHT = -16777216
Declare Function ImageList_GetBkColor Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long _
        ) As Long

' Draw:
Declare Function ImageList_Draw Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal hdcDst As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal fStyle As Long _
        ) As Long
Type IMAGELISTDRAWPARAMS
    cbSize As Long
    hIml As Long
    i As Long
    hdcDst As Long
    x As Long
    y As Long
    cX As Long
    cY As Long
    xBitmap As Long    '        // x offest from the upperleft of bitmap
    yBitmap As Long    '        // y offset from the upperleft of bitmap
    rgbBk As Long
    rgbFg As Long
    fStyle As Long
    dwRop As Long
End Type
Declare Function ImageList_DrawIndirect Lib "COMCTL32.DLL" (pimldp As IMAGELISTDRAWPARAMS) As Long

Declare Function ImageList_SetOverlayImage Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal iImage As Long, _
        ByVal iOverlay As Long _
        ) As Long
Public Const ILD_NORMAL = 0
Public Const ILD_TRANSPARENT = 1
Public Const ILD_BLEND25 = 2
Public Const ILD_SELECTED = 4
Public Const ILD_FOCUS = 4
Public Const ILD_MASK = &H10&
Public Const ILD_IMAGE = &H20&
Public Const ILD_ROP = &H40&
Public Const ILD_OVERLAYMASK = 3840

Declare Function ImageList_BeginDrag Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal dxHotSpot As Long, _
        ByVal dyHotSpot As Long _
        ) As Long
Declare Function ImageList_DragMove Lib "COMCTL32.DLL" ( _
        ByVal x As Long, _
        ByVal y As Long _
        ) As Long
Declare Function ImageList_DragShow Lib "COMCTL32.DLL" ( _
        ByVal fShow As Long _
        ) As Long
Declare Function ImageList_EndDrag Lib "COMCTL32.DLL" () As Long
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

' Work DC
Private m_hdcMono As Long
Private m_hbmpMono As Long
Private m_hBmpOld As Long

' Keyboard hook (for accelerators):
Private m_hKeyHook As Long
Private m_lKeyHookPtr() As Long
Private m_iKeyHookCount As Long
Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long



Public Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Public Declare Function GetGDIObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Public Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type



Public Sub AttachKeyboardHook(cThis As IAPP_PopupMenu)
    Dim lpFn As Long
    Dim lPtr As Long
    Dim i As Long

    If m_iKeyHookCount = 0 Then
        lpFn = HookAddress(AddressOf KeyboardFilter)
        m_hKeyHook = SetWindowsHookEx(WH_KEYBOARD, lpFn, 0&, GetCurrentThreadId())
        Debug.Assert (m_hKeyHook <> 0)
    End If
    lPtr = ObjPtr(cThis)
    For i = 1 To m_iKeyHookCount
        If lPtr = m_lKeyHookPtr(i) Then
            ' we already have it:
            Debug.Assert False
            Exit Sub
        End If
    Next i
    ReDim Preserve m_lKeyHookPtr(1 To m_iKeyHookCount + 1) As Long
    m_iKeyHookCount = m_iKeyHookCount + 1
    m_lKeyHookPtr(m_iKeyHookCount) = lPtr

End Sub

Public Function BitmapToPicture(ByVal hBmp As Long) As IPicture

    If (hBmp = 0) Then Exit Function

    Dim oNewPic As Picture, tPicConv As PictDesc, IGuid As Guid

    ' Fill PictDesc structure with necessary parts:
    With tPicConv
        .cbSizeofStruct = Len(tPicConv)
        .picType = vbPicTypeBitmap
        .hImage = hBmp
    End With

    ' Fill in IDispatch Interface ID
    With IGuid
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    ' Create a picture object:
    OleCreatePictureIndirect tPicConv, IGuid, True, oNewPic

    ' Return it:
    Set BitmapToPicture = oNewPic


End Function

Public Property Get BlendColor(ByVal oColorFrom As OLE_COLOR, ByVal oColorTo As OLE_COLOR) As Long
    Dim lCFrom As Long
    Dim lCTo As Long
    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)
    Dim lCRetR As Long
    Dim lCRetG As Long
    Dim lCRetB As Long
    lCRetR = (lCFrom And &HFF) + ((lCTo And &HFF) - (lCFrom And &HFF)) \ 2
    If (lCRetR > 255) Then lCRetR = 255 Else If (lCRetR < 0) Then lCRetR = 0
    lCRetG = ((lCFrom \ &H100) And &HFF&) + (((lCTo \ &H100) And &HFF&) - ((lCFrom \ &H100) And &HFF&)) \ 2
    If (lCRetG > 255) Then lCRetG = 255 Else If (lCRetG < 0) Then lCRetG = 0
    lCRetB = ((lCFrom \ &H10000) And &HFF&) + (((lCTo \ &H10000) And &HFF&) - ((lCFrom \ &H10000) And &HFF&)) \ 2
    If (lCRetB > 255) Then lCRetB = 255 Else If (lCRetB < 0) Then lCRetB = 0
    BlendColor = RGB(lCRetR, lCRetG, lCRetB)
End Property

Public Sub ClearUpWorkDC()
    If m_hBmpOld <> 0 Then
        SelectObject m_hdcMono, m_hBmpOld
        m_hBmpOld = 0
    End If
    If m_hbmpMono <> 0 Then
        DeleteObject m_hbmpMono
        m_hbmpMono = 0
    End If
    If m_hdcMono <> 0 Then
        DeleteDC m_hdcMono
        m_hdcMono = 0
    End If
End Sub

Public Sub DetachKeyboardHook(cThis As IAPP_PopupMenu)
    Dim i As Long
    Dim lPtr As Long
    Dim iThis As Long

    lPtr = ObjPtr(cThis)
    For i = 1 To m_iKeyHookCount
        If m_lKeyHookPtr(i) = lPtr Then
            iThis = i
            Exit For
        End If
    Next i
    If iThis <> 0 Then
        If m_iKeyHookCount > 1 Then
            For i = iThis To m_iKeyHookCount - 1
                m_lKeyHookPtr(i) = m_lKeyHookPtr(i + 1)
            Next i
        End If
        m_iKeyHookCount = m_iKeyHookCount - 1
        If m_iKeyHookCount >= 1 Then
            ReDim Preserve m_lKeyHookPtr(1 To m_iKeyHookCount) As Long
        Else
            Erase m_lKeyHookPtr
        End If
    Else
        ' Trying to detach a toolbar which was never attached...
        ' This will happen at design time
    End If

    If m_iKeyHookCount <= 0 Then
        If (m_hKeyHook <> 0) Then
            UnhookWindowsHookEx m_hKeyHook
            m_hKeyHook = 0
        End If
    End If

End Sub

Public Function DrawEdge( _
      ByVal hdc As Long, _
      qrc As RECT, _
      ByVal edge As Long, _
      ByVal grfFlags As Long, _
      ByVal bOfficeXpStyle As Boolean _
   ) As Long
    If (bOfficeXpStyle) Then
        Dim junk As POINTAPI
        Dim hPenOld As Long
        Dim hPen As Long
        If (qrc.bottom > qrc.top) Then
            hPen = CreatePen(PS_SOLID, 1, TranslateColor(vbHighlight))
        Else
            hPen = CreatePen(PS_SOLID, 1, TranslateColor(vb3DShadow))
        End If
        hPenOld = SelectObject(hdc, hPen)
        MoveToEx hdc, qrc.left, qrc.top, junk
        LineTo hdc, qrc.right - 1, qrc.top
        If (qrc.bottom > qrc.top) Then
            LineTo hdc, qrc.right - 1, qrc.bottom - 1
            LineTo hdc, qrc.left, qrc.bottom - 1
            LineTo hdc, qrc.left, qrc.top
        End If
        SelectObject hdc, hPenOld
        DeleteObject hPen
    Else
        DrawEdgeAPI hdc, qrc, edge, grfFlags
    End If
End Function

Public Sub DrawGradient( _
      ByVal hdc As Long, _
      ByRef rct As RECT, _
      ByVal lEndColour As Long, _
      ByVal lStartColour As Long, _
      ByVal bVertical As Boolean _
   )
    Dim lStep As Long
    Dim lPos As Long, lSize As Long
    Dim bRGB(1 To 3) As Integer
    Dim bRGBStart(1 To 3) As Integer
    Dim dR(1 To 3) As Double
    Dim dPos As Double, d As Double
    Dim hBr As Long
    Dim tR As RECT

    LSet tR = rct
    If bVertical Then
        lSize = (tR.bottom - tR.top)
    Else
        lSize = (tR.right - tR.left)
    End If
    lStep = lSize \ 255
    If (lStep < 3) Then
        lStep = 3
    End If

    bRGB(1) = lStartColour And &HFF&
    bRGB(2) = (lStartColour And &HFF00&) \ &H100&
    bRGB(3) = (lStartColour And &HFF0000) \ &H10000
    bRGBStart(1) = bRGB(1): bRGBStart(2) = bRGB(2): bRGBStart(3) = bRGB(3)
    dR(1) = (lEndColour And &HFF&) - bRGB(1)
    dR(2) = ((lEndColour And &HFF00&) \ &H100&) - bRGB(2)
    dR(3) = ((lEndColour And &HFF0000) \ &H10000) - bRGB(3)

    For lPos = lSize To 0 Step -lStep
        ' Draw bar:
        If bVertical Then
            tR.top = tR.bottom - lStep
        Else
            tR.left = tR.right - lStep
        End If
        If tR.top < rct.top Then
            tR.top = rct.top
        End If
        If tR.left < rct.left Then
            tR.left = rct.left
        End If

        'Debug.Print tR.Right, tR.left, (bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1))
        hBr = CreateSolidBrush((bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1)))
        FillRect hdc, tR, hBr
        DeleteObject hBr

        ' Adjust colour:
        dPos = ((lSize - lPos) / lSize)
        If bVertical Then
            tR.bottom = tR.top
            bRGB(1) = bRGBStart(1) + dR(1) * dPos
            bRGB(2) = bRGBStart(2) + dR(2) * dPos
            bRGB(3) = bRGBStart(3) + dR(3) * dPos
        Else
            tR.right = tR.left
            bRGB(1) = bRGBStart(1) + dR(1) * dPos
            bRGB(2) = bRGBStart(2) + dR(2) * dPos
            bRGB(3) = bRGBStart(3) + dR(3) * dPos
        End If

    Next lPos

End Sub

Public Sub DrawMaskedFrameControl( _
    ByVal hdcDest As Long, _
    ByRef trWhere As RECT, _
    ByVal kind As DFCFlags, _
    ByVal Style As Long _
   )
    Dim hbrMenu As Long, hbrStockWhite As Long
    Dim saveBkMode As Long, saveBkColor As Long, saveBrush As Long
    Dim tRWhereOnTmp As RECT
    Dim bgcolor As Long
    Static s_lLastRight As Long, s_lLastBottom As Long

    With tRWhereOnTmp
        .right = trWhere.right - trWhere.left
        .bottom = trWhere.bottom - trWhere.top
        If .right > s_lLastRight Or .bottom > s_lLastBottom Or (m_hdcMono = 0) Or (m_hbmpMono = 0) Or (m_hBmpOld = 0) Then
            ClearUpWorkDC
            ' Create memory device context for our temporary mask
            m_hdcMono = CreateCompatibleDC(0)
            If m_hdcMono <> 0 Then
                ' Create monochrome bitmap and select it into DC
                m_hbmpMono = CreateCompatibleBitmap(m_hdcMono, .right, .bottom)
                If m_hbmpMono <> 0 Then
                    m_hBmpOld = SelectObject(m_hdcMono, m_hbmpMono)
                    SetBkColor m_hdcMono, &HFFFFFF
                End If
            End If
            If m_hBmpOld = 0 Then
                ' Failed...
                ClearUpWorkDC
            End If
        End If
        s_lLastRight = .right
        s_lLastBottom = .bottom
    End With


    DrawFrameControl m_hdcMono, tRWhereOnTmp, kind, Style
    ' We have black where tick & white elsewhere
    SetBkColor hdcDest, &HFFFFFF
    BitBlt hdcDest, trWhere.left, trWhere.top, trWhere.right, trWhere.bottom, m_hdcMono, 0, 0, vbSrcAnd

    ' Clean up everything.
    If saveBrush <> 0 Then
        SelectObject hdcDest, saveBrush
    End If
    If hbrMenu <> 0 Then
        DeleteObject hbrMenu
    End If
    If saveBkMode <> 0 Then
        SetBkMode hdcDest, saveBkMode
    End If
    If saveBkColor <> 0 Then
        SetBkColor hdcDest, saveBkColor
    End If

End Sub

Public Sub HLSToRGB( _
     ByVal h As Single, ByVal s As Single, ByVal l As Single, _
     r As Long, g As Long, b As Long _
     )
    Dim rR As Single, rG As Single, rB As Single
    Dim Min As Single, Max As Single

    If s = 0 Then
        ' Achromatic case:
        rR = l: rG = l: rB = l
    Else
        ' Chromatic case:
        ' delta = Max-Min
        If l <= 0.5 Then
            's = (Max - Min) / (Max + Min)
            ' Get Min value:
            Min = l * (1 - s)
        Else
            's = (Max - Min) / (2 - Max - Min)
            ' Get Min value:
            Min = l - s * (1 - l)
        End If
        ' Get the Max value:
        Max = 2 * l - Min

        ' Now depending on sector we can evaluate the h,l,s:
        If (h < 1) Then
            rR = Max
            If (h < 0) Then
                rG = Min
                rB = rG - h * (Max - Min)
            Else
                rB = Min
                rG = h * (Max - Min) + rB
            End If
        ElseIf (h < 3) Then
            rG = Max
            If (h < 2) Then
                rB = Min
                rR = rB - (h - 2) * (Max - Min)
            Else
                rR = Min
                rB = (h - 2) * (Max - Min) + rR
            End If
        Else
            rB = Max
            If (h < 4) Then
                rR = Min
                rG = rR - (h - 4) * (Max - Min)
            Else
                rG = Min
                rR = (h - 4) * (Max - Min) + rG
            End If

        End If

    End If
    r = rR * 255: g = rG * 255: b = rB * 255
End Sub

Private Function HookAddress(ByVal lPtr As Long) As Long
    HookAddress = lPtr
End Function

Public Sub ImageListDrawIcon( _
        ByVal ptrVb6ImageList As Long, _
        ByVal hdc As Long, _
        ByVal hIml As Long, _
        ByVal iIconIndex As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        Optional ByVal bSelected As Boolean = False, _
        Optional ByVal bBlend25 As Boolean = False, Optional IsCheckAndIcon As Boolean = False _
    )
    Dim playImage As Long
    Dim lFlags As Long
    Dim lR As Long

    lFlags = ILD_TRANSPARENT
    If (bSelected) Then
        lFlags = lFlags Or ILD_SELECTED
    End If
    If (bBlend25) Then
        lFlags = lFlags Or ILD_BLEND25
    End If
    If (ptrVb6ImageList <> 0) Then
        Dim o As Object
        On Error Resume Next
        Set o = ObjectFromPtr(ptrVb6ImageList)
        If Not (o Is Nothing) Then
'            o.ListImages(iIconIndex + 1).Draw hdc, lX * Screen.TwipsPerPixelX, lY * Screen.TwipsPerPixelY, lFlags
            If IsCheckAndIcon Then
                'UtilDrawBorderRectangle hdc, TranslateColor(vbActiveTitleBar), lX - 3, lY - 2, 20, 20, False
            End If
            DrawIconEx hdc, lX, 3, ResizeIconImage(o.ListImages(iIconIndex + 1).ExtractIcon()), 16, 16, 0, 0, &H3 Or ILD_BLEND25
        
        End If
        On Error GoTo 0
    Else
        lR = ImageList_Draw( _
                hIml, _
                iIconIndex, _
                hdc, _
                lX, _
                lY, _
                lFlags)
        If (lR = 0) Then
            Debug.Print "Failed to draw Image: " & iIconIndex & " onto hDC " & hdc, "ImageListDrawIcon"
        End If
    End If
End Sub

Public Sub ImageListDrawIconDisabled( _
        ByVal ptrVb6ImageList As Long, _
        ByVal hdc As Long, _
        ByVal hIml As Long, _
        ByVal iIconIndex As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        ByVal lSize As Long, _
        Optional ByVal asShadow As Boolean _
    )
    Dim lR As Long
    Dim hIcon As Long
Exit Sub
    hIcon = 0
    If (ptrVb6ImageList <> 0) Then
        Dim o As Object
        On Error Resume Next
        Set o = ObjectFromPtr(ptrVb6ImageList)
        If Not (o Is Nothing) Then

            Dim lhDCDisp As Long
            Dim lHDC As Long
            Dim lhBmp As Long
            Dim lhBmpOld As Long
            Dim lhIml As Long

            lhDCDisp = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
            lHDC = CreateCompatibleDC(lhDCDisp)
            lhBmp = CreateCompatibleBitmap(lhDCDisp, o.ImageWidth, o.ImageHeight)
            DeleteDC lhDCDisp
            lhBmpOld = SelectObject(lHDC, lhBmp)
            o.ListImages.Item(iIconIndex + 1).Draw lHDC, 0, 0, 0
            SelectObject lHDC, lhBmpOld
            DeleteDC lHDC
            lhIml = ImageList_Create(o.ImageWidth, o.ImageHeight, ILC_MASK Or ILC_COLOR32, 1, 1)
            ImageList_AddMasked lhIml, lhBmp, TranslateColor(o.BackColor)
            DeleteObject lhBmp
            hIcon = ImageList_GetIcon(lhIml, 0, 0)
            ImageList_Destroy lhIml

        End If
        On Error GoTo 0
    Else
        hIcon = ImageList_GetIcon(hIml, iIconIndex, 0)
    End If
    If (hIcon <> 0) Then
        If (asShadow) Then
            Dim hBr As Long
            hBr = GetSysColorBrush(vb3DShadow And &H1F)
            lR = DrawState(hdc, hBr, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_MONO)
            DeleteObject hBr
        Else
            lR = DrawState(hdc, 0, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_DISABLED)
        End If
        DestroyIcon hIcon
    End If

End Sub

Private Function KeyboardFilter(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim bKeyUp As Boolean
    Dim bAlt As Boolean, bCtrl As Boolean, bShift As Boolean
    Dim bFKey As Boolean, bEscape As Boolean, bDelete As Boolean
    Dim wMask As KeyCodeConstants
    Dim cT As IAPP_PopupMenu
    Dim i As Long

    On Error GoTo ErrorHandler

    If nCode = HC_ACTION And m_iKeyHookCount > 0 Then
        ' Key up or down:
        bKeyUp = ((lParam And &H80000000) = &H80000000)
        If Not bKeyUp Then
            bShift = (GetAsyncKeyState(vbKeyShift) <> 0)
            bAlt = ((lParam And &H20000000) = &H20000000)
            bCtrl = (GetAsyncKeyState(vbKeyControl) <> 0)
            bFKey = ((wParam >= vbKeyF1) And (wParam <= vbKeyF12))
            bEscape = (wParam = vbKeyEscape)
            bDelete = (wParam = vbKeyDelete)
            If bAlt Or bCtrl Or bFKey Or bEscape Or bDelete Then
                wMask = Abs(bShift * vbShiftMask) Or Abs(bCtrl * vbCtrlMask) Or Abs(bAlt * vbAltMask)
                For i = m_iKeyHookCount To 1 Step -1
                    If m_lKeyHookPtr(i) <> 0 Then
                        ' Alt- or Ctrl- key combination pressed:
                        Set cT = PopupMenuFromPtr(m_lKeyHookPtr(i))
                        If Not cT Is Nothing Then
                            If cT.AcceleratorPress(wParam, wMask) Then
                                KeyboardFilter = 1
                                Exit Function
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    End If
    KeyboardFilter = CallNextHookEx(m_hKeyHook, nCode, wParam, lParam)

    Exit Function

ErrorHandler:
    Debug.Print "Keyboard Hook Error!"
    Exit Function

End Function

Public Property Get LighterColour(ByVal oColor As OLE_COLOR) As Long
    Dim lC As Long
    Dim h As Single, s As Single, l As Single
    Dim lR As Long, lG As Long, lB As Long
    Static s_lColLast As Long
    Static s_lLightColLast As Long

    lC = TranslateColor(oColor)
    If (lC <> s_lColLast) Then
        s_lColLast = lC
        RGBToHLS lC And &HFF&, (lC \ &H100) And &HFF&, (lC \ &H10000) And &HFF&, h, s, l
        If (l > 0.99) Then
            l = l * 0.8
        Else
            l = l * 1.1
            If (l > 1) Then
                l = 1
            End If
        End If
        HLSToRGB h, s, l, lR, lG, lB
        s_lLightColLast = RGB(lR, lG, lB)
    End If
    LighterColour = s_lLightColLast
End Property

 Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
    If (rR > rG) Then
        If (rR > rB) Then
            Maximum = rR
        Else
            Maximum = rB
        End If
    Else
        If (rB > rG) Then
            Maximum = rB
        Else
            Maximum = rG
        End If
    End If
End Function

 Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
    If (rR < rG) Then
        If (rR < rB) Then
            Minimum = rR
        Else
            Minimum = rB
        End If
    Else
        If (rB < rG) Then
            Minimum = rB
        Else
            Minimum = rG
        End If
    End If
End Function

Public Property Get NoPalette(Optional ByVal bForce As Boolean = False) As Boolean
    Static bOnce As Boolean
    Static bNoPalette As Boolean
    Dim lHDC As Long
    Dim lBits As Long
    If (bForce) Then
        bOnce = False
    End If
    If Not (bOnce) Then
        lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        If (lHDC <> 0) Then
            lBits = GetDeviceCaps(lHDC, BITSPIXEL)
            If (lBits <> 0) Then
                bOnce = True
            End If
            bNoPalette = (lBits > 8)
            DeleteDC lHDC
        End If
    End If
    NoPalette = bNoPalette
End Property

Public Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
    Dim oTemp As Object
    ' Turn the pointer into an illegal, uncounted interface
    CopyMemory oTemp, lPtr, 4
    ' Do NOT hit the End button here! You will crash!
    ' Assign to legal reference
    Set ObjectFromPtr = oTemp
    ' Still do NOT hit the End button here! You will still crash!
    ' Destroy the illegal reference
    CopyMemory oTemp, 0&, 4
    ' OK, hit the End button if you must--you'll probably still crash,
    ' but it will be because of the subclass, not the uncounted reference
End Property

Private Property Get PopupMenuFromPtr(ByVal lPtr As Long) As IAPP_PopupMenu
    Dim oTemp As Object
    If lPtr <> 0 Then
        ' Turn the pointer into an illegal, uncounted interface
        CopyMemory oTemp, lPtr, 4
        ' Do NOT hit the End button here! You will crash!
        ' Assign to legal reference
        Set PopupMenuFromPtr = oTemp
        ' Still do NOT hit the End button here! You will still crash!
        ' Destroy the illegal reference
        CopyMemory oTemp, 0&, 4
        ' OK, hit the End button if you must--you'll probably still crash,
        ' but it will be because of the subclass, not the uncounted reference
    End If
End Property

Public Sub RGBToHLS( _
     ByVal r As Long, ByVal g As Long, ByVal b As Long, _
     h As Single, s As Single, l As Single _
     )
    Dim Max As Single
    Dim Min As Single
    Dim delta As Single
    Dim rR As Single, rG As Single, rB As Single

    rR = r / 255: rG = g / 255: rB = b / 255

    '{Given: rgb each in [0,1].
    ' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
    Max = Maximum(rR, rG, rB)
    Min = Minimum(rR, rG, rB)
    l = (Max + Min) / 2    '{This is the lightness}
    '{Next calculate saturation}
    If Max = Min Then
        'begin {Acrhomatic case}
        s = 0
        h = 0
        'end {Acrhomatic case}
    Else
        'begin {Chromatic case}
        '{First calculate the saturation.}
        If l <= 0.5 Then
            s = (Max - Min) / (Max + Min)
        Else
            s = (Max - Min) / (2 - Max - Min)
        End If
        '{Next calculate the hue.}
        delta = Max - Min
        If rR = Max Then
            h = (rG - rB) / delta    '{Resulting color is between yellow and magenta}
        ElseIf rG = Max Then
            h = 2 + (rB - rR) / delta    '{Resulting color is between cyan and yellow}
        ElseIf rB = Max Then
            h = 4 + (rR - rG) / delta    '{Resulting color is between magenta and cyan}
        End If
        'end {Chromatic Case}
    End If
End Sub

Public Property Get SlightlyLighterColour(ByVal oColor As OLE_COLOR) As Long
    Dim lC As Long
    Dim h As Single, s As Single, l As Single
    Dim lR As Long, lG As Long, lB As Long
    Static s_lColLast As Long
    Static s_lLightColLast As Long

    lC = TranslateColor(oColor)
    If (lC <> s_lColLast) Then
        s_lColLast = lC
        RGBToHLS lC And &HFF&, (lC \ &H100) And &HFF&, (lC \ &H10000) And &HFF&, h, s, l
        If (l > 0.99) Then
            l = l * 0.95
        Else
            l = l * 1.05
            If (l > 1) Then
                l = 1
            End If
        End If
        HLSToRGB h, s, l, lR, lG, lB
        s_lLightColLast = RGB(lR, lG, lB)
    End If
    SlightlyLighterColour = s_lLightColLast
End Property

Public Sub TileArea( _
        ByVal hdcTo As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal Width As Long, _
        ByVal Height As Long, _
        ByVal hDcSrc As Long, _
        ByVal SrcWidth As Long, _
        ByVal SrcHeight As Long, _
        ByVal lOffsetY As Long _
    )
    Dim lSrcX As Long
    Dim lSrcY As Long
    Dim lSrcStartX As Long
    Dim lSrcStartY As Long
    Dim lSrcStartWidth As Long
    Dim lSrcStartHeight As Long
    Dim lDstX As Long
    Dim lDstY As Long
    Dim lDstWidth As Long
    Dim lDstHeight As Long

    lSrcStartX = (x Mod SrcWidth)
    lSrcStartY = ((y + lOffsetY) Mod SrcHeight)
    lSrcStartWidth = (SrcWidth - lSrcStartX)
    lSrcStartHeight = (SrcHeight - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY

    lDstY = y
    lDstHeight = lSrcStartHeight

    Do While lDstY < (y + Height)
        If (lDstY + lDstHeight) > (y + Height) Then
            lDstHeight = y + Height - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = x
        lSrcX = lSrcStartX
        Do While lDstX < (x + Width)
            If (lDstX + lDstWidth) > (x + Width) Then
                lDstWidth = x + Width - lDstX
                If (lDstWidth = 0) Then
                    lDstWidth = 4
                End If
            End If
            'If (lDstWidth > Width) Then lDstWidth = Width
            'If (lDstHeight > Height) Then lDstHeight = Height
            BitBlt hdcTo, lDstX, lDstY, lDstWidth, lDstHeight, hDcSrc, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = SrcWidth
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = SrcHeight
    Loop
End Sub

Public Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Public Function WinAPIError(ByVal lLastDLLError As Long) As String
    Dim sBuff As String
    Dim lCount As Long

    ' Return the error message associated with LastDLLError:
    sBuff = String$(256, 0)
    lCount = FormatMessage( _
            FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
            0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
    If lCount Then
        WinAPIError = left$(sBuff, lCount)
    End If

End Function

Public Function ResizeIconImage(hImg As Long) As Long
            Dim icoInfo As ICONINFO
            Dim newICOinfo As ICONINFO
            Dim icoBMPinfo As BITMAP
            Call GetIconInfo(hImg, icoInfo)
            If ResizeIconImage Then DestroyIcon ResizeIconImage

            ' start a new icon structure
            CopyMemory newICOinfo, icoInfo, Len(icoInfo)

            ' get the icon dimensions from the bitmap portion of the icon
            GetGDIObject icoInfo.hbmColor, Len(icoBMPinfo), icoBMPinfo
            

            ResizeIconImage = CreateIconIndirect(newICOinfo)
                        
            DeleteObject newICOinfo.hbmMask
            DeleteObject newICOinfo.hbmColor
End Function
