Attribute VB_Name = "Module1"
'------------------------------------------------------------------------------
'-- Module Name.....: mDialogCallBack
'-- Description.....: Module that provides the callback function for Dialoge Boxes
'--
'-- Notes...........: The callback procedure receives messages or notifications intended for
'--                   the default dialog box procedure of dialog boxes, so you can subclass
'--                   the standard controls of the common dialog box. In this implementation
'--                   we can center the standard dialog on the screen or we can set another
'--                   title. If you want to perform other changes modify the following function.
'--
'--
'-- Author, date....: Gary Noble (TDLcom) , 16 March 2002
'--
'--
'-- Property             Data Type     Description
'-- ------------------   ---------     --------------------------------------
'--
'-- Method(Public)       Description
'-- ------------------   --------------------------------------
'-- FontDialogCallBack   Callback (global) routine for ICDLG_FontDialogHandler. It is used to center the
'--                      dialog and to set the caption text.
'-- FileOpenSaveDialogCallbackEx - Old-style callback (global) routine for ICDLG_FileOpenSaveHandler. It is used
'--                                to center the dialog box. Must be used with eFileOpenSaveFlag_Explorer.
'-- FileOpenSaveDialogCallback     Old-style callback (global) routine for ICDLG_FileOpenSaveHandler. It is used
'--                                to center the dialog box. Does not cover eFileOpenSaveFlag_Explorer.
'-- ColourDialogCallBack Callback (global) routine for the ICDLG_ColorDialogHandler class. It is used to
'--                      center the dialog box and to set the caption text.
'--
'-- BrowseForFolderCallBack - Callback (global) routine for the ICDLG_BrowseForFolderHandler class. It is used to
'--                           center the dialog and to set the caption text.
'--
'-- Method(Private)      Description
'-- ------------------   --------------------------------------
'------------------------------------------------------------------------------

Option Explicit

Private m_bShowPreview As Boolean
Private m_oPreview As Object
Private PreviewObjectParentHWND As Long

'-- Private constants - BrowseForFolder
Private Const MAX_PATH = 512
Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELECTIONCHANGED = 2
Private Const BFFM_SETSTATUSTEXTA = (WM_USER + 100)
Private Const BFFM_SETSTATUSTEXT = BFFM_SETSTATUSTEXTA
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTION = BFFM_SETSELECTIONA
Private Const WM_NOTIFY         As Long = &H4E

'------------------------------------------------------------------------------
' Constants specific to the common dialog - ListView HWND and two notification messages.
'------------------------------------------------------------------------------
Private Const ID_LIST           As Long = &H460
' Notification codes:
Private Const H_MAX As Long = &HFFFF + 1
Private Const CDN_FIRST = (H_MAX - 601)
Private Const CDN_LAST = (H_MAX - 699)


'// Notifications when Open or Save dialog status changes
Private Const CDN_INITDONE = (CDN_FIRST - &H0)
Private Const CDN_SELCHANGE = (CDN_FIRST - &H1)
Private Const CDN_FOLDERCHANGE = (CDN_FIRST - &H2)
Private Const CDN_SHAREVIOLATION = (CDN_FIRST - &H3)
Private Const CDN_HELP = (CDN_FIRST - &H4)
Private Const CDN_FILEOK = (CDN_FIRST - &H5)
Private Const CDN_TYPECHANGE = (CDN_FIRST - &H6)
Private Const CDN_INCLUDEITEM = (CDN_FIRST - &H7)


Private Const CDM_FIRST = WM_USER + 100
Private Const CDM_GETFILEPATH = CDM_FIRST + &H1

Private Type RECT
    Left     As Long
    Top      As Long
    Right    As Long
    Bottom   As Long
End Type
Private Type POINTAPI
    x As Long
    y As Long
End Type

'------------------------------------------------------------------------------
' Used with the SetWindowPlacement API function
'------------------------------------------------------------------------------
Private Type WINDOWPLACEMENT
    length              As Long
    Flags               As Long
    showCmd             As Long
    ptMinPosition       As POINTAPI
    ptMaxPosition       As POINTAPI
    rcNormalPosition    As RECT
End Type
 
Private Type NMHDR
    hwndFrom            As Long    ' Window handle of control sending message
    idFrom              As Long    ' Identifier of control sending message
    code                As Long    ' Specifies the notification code
End Type

'------------------------------------------------------------------------------
'-- Private class constants
'------------------------------------------------------------------------------
Private Const WM_INITDIALOG = &H110


'------------------------------------------------------------------------------
' OPENFILENAME structure.
'------------------------------------------------------------------------------
Private Type OPENFILENAME
    lStructSize         As Long
    HWndOwner           As Long
    hInstance           As Long
    lpstrFilter         As String
    lpstrCustomFilter   As String
    nMaxCustFilter      As Long
    nFilterIndex        As Long
    lpstrFile           As String
    nMaxFile            As Long
    lpstrFileTitle      As String
    nMaxFileTitle       As Long
    lpstrInitialDir     As String
    lpstrTitle          As String
    Flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As String
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As String
End Type

Private Const SW_HIDE = 0
Private Const SW_NORMAL = 1

'------------------------------------------------------------------------------
'-- Private class API function declarations
'------------------------------------------------------------------------------
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

'------------------------------------------------------------------------------
'-- Private API function declarations - BrowseForFolder
'------------------------------------------------------------------------------
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

'------------------------------------------------------------------------------
'-- Public variables used for communication between a ICDLG_FontDialogHandler object and
'-- the callback routine implemented in this module
'------------------------------------------------------------------------------
Public g_bCenterFontDialog   As Boolean
Public g_sFontDialogTitle    As String

'------------------------------------------------------------------------------
'-- Public variable used for communication between a ICDLG_ColorDialogHandler object and the callback routine
'-- implemented in this module
'------------------------------------------------------------------------------
Public g_bCenterColourDialog  As Boolean
Public g_sColourDialogTitle   As String
Public g_sCenterOpenDialog    As Boolean


'------------------------------------------------------------------------------
'-- Function    : FontDialogCallback
'-- Notes       : Callback (global) routine for ICDLG_FontDialogHandler. It is used to center the
'--               dialog and to set the caption text.
'------------------------------------------------------------------------------
Public Function FontDialogCallback(ByVal hwnd As Long, ByVal uMsg As Long, lParam As Long, ByVal lpData As Long) As Long
    On Error Resume Next
    
    Dim rcHeight     As Long
    Dim rcWidth      As Long
    Dim RC           As RECT
    Dim rcDesk       As RECT
    
    Select Case uMsg
        
        Case WM_INITDIALOG
            '-- Set the new title
            If Len(Trim$(g_sFontDialogTitle)) > 0 Then SetWindowText hwnd, g_sFontDialogTitle
            
            '-- Center the window
            If g_bCenterFontDialog Then
                Call GetWindowRect(GetDesktopWindow, rcDesk)
                Call GetWindowRect(hwnd, RC)
            
                rcHeight = RC.Bottom - RC.Top
                rcWidth = RC.Right - RC.Left
                RC.Left = Abs(((rcDesk.Right - rcDesk.Left) - rcWidth) / 2)
                RC.Top = Abs(((rcDesk.Bottom - rcDesk.Top) - rcHeight) / 2)
            
                MoveWindow hwnd, RC.Left, RC.Top, rcWidth, rcHeight, 1
            End If
        Case Else
            '
    
    End Select
    
    FontDialogCallback = 0&
End Function
