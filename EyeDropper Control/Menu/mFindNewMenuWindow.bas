Attribute VB_Name = "mFindNewMenuWindow"
Option Explicit


Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Const WM_COMMAND = &H111

Private m_iCount As Long
Private m_hWnd() As Long

Private Function ClassName(ByVal lhWnd As Long) As String
    Dim lLen As Long
    Dim sBuf As String
    lLen = 260
    sBuf = String$(lLen, 0)
    lLen = GetClassName(lhWnd, sBuf, lLen)
    If (lLen <> 0) Then
        ClassName = left$(sBuf, lLen)
    End If
End Function

Public Function EnumerateWindows() As Long
    m_iCount = 0
    Erase m_hWnd
    EnumWindows AddressOf EnumWindowsProc, 0
End Function

Public Property Get EnumerateWindowsCount() As Long
    EnumerateWindowsCount = m_iCount
End Property

Public Property Get EnumerateWindowshWnd(ByVal iIndex As Long) As Long
    EnumerateWindowshWnd = m_hWnd(iIndex)
End Property

Private Function EnumWindowsProc( _
        ByVal hwnd As Long, _
        ByVal lParam As Long _
    ) As Long
    Dim sClass As String
    sClass = ClassName(hwnd)
    If sClass = "#32768" Then    ' Menu Window Class Name
        If IsWindowVisible(hwnd) Then
            m_iCount = m_iCount + 1
            ReDim Preserve m_hWnd(1 To m_iCount) As Long
            m_hWnd(m_iCount) = hwnd
            ' Debug.Print "Menu:", hWnd
        End If
    End If
End Function

Private Function WindowTitle(ByVal lhWnd As Long) As String
    Dim lLen As Long
    Dim sBuf As String

    ' Get the Window Title:
    lLen = GetWindowTextLength(lhWnd)
    If (lLen > 0) Then
        sBuf = String$(lLen + 1, 0)
        lLen = GetWindowText(lhWnd, sBuf, lLen + 1)
        WindowTitle = left$(sBuf, lLen)
    End If

End Function
