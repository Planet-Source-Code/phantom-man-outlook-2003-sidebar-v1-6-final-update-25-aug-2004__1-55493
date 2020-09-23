VERSION 5.00
Begin VB.UserControl EyeDropper 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MouseIcon       =   "EyeDropper.ctx":0000
   ScaleHeight     =   3735
   ScaleWidth      =   5880
   ToolboxBitmap   =   "EyeDropper.ctx":0152
   Begin VB.PictureBox picSplitter 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   100
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   241
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.PictureBox picExtraItems 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   5880
      TabIndex        =   0
      Top             =   3285
      Width           =   5880
   End
   Begin VB.Menu mnuItemMain 
      Caption         =   ""
      Begin VB.Menu mnuItems 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMore 
         Caption         =   "Show More Items"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuLess 
         Caption         =   "Show Less Items"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddRem 
         Caption         =   "Add Or Remove"
         Begin VB.Menu mnuAddRemA 
            Caption         =   ""
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "EyeDropper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 0

' --------------------------------------------------------------------------------------
' Name:     EyeDropper.cls
' Author:   Gary Noble (gwnoble@msn.com)
' Date:     19/04/2004
'
' Requires: Nothing
'
' Copyright © 2004 Gary Noble
' --------------------------------------------------------------------------------------
'
' A Control Implementing The Outlook 2003 Side Menu Interface.
'
' ImageLists Can Be A Standard VB Image List Or A VBAccelerator Image List
'
' --------------------------------------------------------------------------------------
'
'   This Control Also Uses Code From Other Authors, All The Original Copyrights
'   Credits Can Be Found Where They Put Them.
'
'   Give Credit Where Credit Is Due.
'
'
' --------------------------------------------------------------------------------------
' History:
'           19/04/2004 - Initial Implementation (Gary Noble)
'           21/04/2004 - Added Add Remove item Functionality (Gary Noble)
'           22/04/2004 - Added Visible Property (Gary Noble)
'                      - This Allows You To Hide The Item But Not Delete It
'                      - It Is Then Placed In The Add Remove Menu Items
'           05/05/2004 - Fixed The Virtual Memory Low Error in WinXP (Gary Noble)
'                      - Added Right To Left Support (Gary Noble)
'           06/05/2004 - Added The Ability To Set Mouse Pointer To Hand (Gary Noble)
'                      - Added DrawToolbarItems RightToLeft Just To Make (Gary Noble)
'                        More Like The Original
'                      - Cleaned Up Code and Test Until I Was Blue In The Face!!! (Can't Say My Name Anymore)
'           11/05/2004 - Added Custom Properties Function
'                      - Added Header Colour
'                      - Update The Default Colours To Conincide With The Original Control
'           14/05/2004 - Added Visible Item SaveState when The User First Resize
'                        What This Does Is Return The Number of Visible Items To Its Original
'                        State Before Sizing.
'                      - Updated Caption Drawing - It Now Has It's Own Sub pDrawCaption
'                      - Updated The Paint Order So The Caption Does'nt Paint Under The Splitter
'                        And Over The Items. (Much More Pro!)
'                      - Updated The Drawing Of The Line to use Api Instead Of .line (x,y)-(x1,y1)
'                      - Removed memDC Drawing As It Wasn't Having Any Effect On The control
'
' --------------------------------------------------------------------------------------
'
'  *********  If You Use This Control Please Give Credit  *********
'
' --------------------------------------------------------------------------------------
' Copyright (c) 2004 Gary Noble
' ---------------------------------------------------------------------
'
' Redistribution and use in source and binary forms, with or
' without modification, are permitted provided that the following
' conditions are met:
'
' 1. Redistributions of source code must retain the above copyright
'    notice, this list of conditions and the following disclaimer.
'
' 2. Redistributions in binary form must reproduce the above copyright
'    notice, this list of conditions and the following disclaimer in
'    the documentation and/or other materials provided with the distribution.
'
' 3. The end-user documentation included with the redistribution, if any,
'    must include the following acknowledgment:
'
'  "This product includes software developed by Gary Noble"
'
' Alternately, this acknowledgment may appear in the software itself, if
' and wherever such third-party acknowledgments normally appear.
'
' THIS SOFTWARE IS PROVIDED "AS IS" AND ANY EXPRESSED OR IMPLIED WARRANTIES,
' INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY
' AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL
' GARY NOBLE OR ANY CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
' INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
' BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF
' USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
' THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
' THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
' ---------------------------------------------------------------------


Implements ISubclass

Private Const m_const_lToolbarHeight As Long = 450
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Dim xMenu                                       As Long
Dim yMenu                                       As Long
Dim sObject                                     As Object
Dim bBeginSize                                  As Boolean

Private srcPicObj                               As StdPicture
Private playImage                               As Long
Private sourceWidth                             As Long
Private sourceHeight                            As Long

Private Const WM_SYSCOLORCHANGE                As Long = &H15
Private Const WM_ENTERSIZEMOVE                 As Long = &H231&
Private Const WM_EXITSIZEMOVE                  As Long = &H232&

Private Type POINTAPI
   x                                            As Long
   y                                            As Long
End Type
Private Type RECT
   left                                         As Long
   top                                          As Long
   right                                        As Long
   bottom                                       As Long
End Type
' Built in ImageList drawing methods:
Private Const ILD_TRANSPARENT                  As Integer = 1&
Private Const ILD_BLEND25                      As Integer = 2&
Private Const ILD_SELECTED                     As Integer = 4&
Private Const TRANSPARENT                      As Integer = 1
Private Const DST_ICON                         As Long = &H3
Private Const DSS_DISABLED                     As Long = &H20
Private Const DSS_MONO                         As Long = &H80
Private Const CLR_INVALID                      As Integer = -1
Private m_fntOrig                              As StdFont
Private m_lPanelTopOffset As Long



'-- Item Data
Private Type ItemInfo
   sCaption                                     As String
   sKey                                         As String
   oRect                                        As RECT
   sToolTipText                                 As String
   lItemData                                    As Long
   sTag                                         As String
   lIconIndex                                   As Long
   lObjPtrPanel                                 As Long
   lID                                          As Long
   tItemR                                        As RECT
   bVisible                                     As Boolean
   bEnabled                                     As Boolean
   bToolbar                                     As Boolean
End Type

Private WithEvents MTimer                      As IAPP_Timer
Attribute MTimer.VB_VarHelpID = -1



Private mobjPCurrentPanel                      As Object

Private m_lCustColorOneNormal                  As OLE_COLOR
Private m_lCustColorTwoNormal                  As OLE_COLOR
Private m_lCustColorOneSelected                As OLE_COLOR
Private m_lCustColorTwoSelected                As OLE_COLOR
Private m_lCustColorHeaderColorOne             As OLE_COLOR
Private m_lCustColorHeaderColorTwo             As OLE_COLOR
Private m_lCustColorHeaderForeColor            As OLE_COLOR
Private m_bCustUseGradient                     As Boolean

Private m_lColorOneSelectedNormal          As OLE_COLOR
Private m_lColorTwoSelectedNormal          As OLE_COLOR
Private m_lColorOneNormal                  As OLE_COLOR
Private m_lColorTwoNormal                  As OLE_COLOR
Private m_lColorOneSelected                As OLE_COLOR
Private m_lColorTwoSelected                As OLE_COLOR
Private m_lColorHeaderColorOne             As OLE_COLOR
Private m_lColorHeaderColorTwo             As OLE_COLOR
Private m_lColorHeaderForeColor            As OLE_COLOR
Private m_lColorHotOne                     As OLE_COLOR
Private m_lColorHotTwo                     As OLE_COLOR
Private m_lColorBorder                     As OLE_COLOR


Private m_lDefTop                              As Long
Private m_lLastY                               As Long
Private m_lPanelTop                            As Long
Private m_lBtnDown                             As Long
Private m_lItemHover                           As Long
Private m_lIdGenerator                         As Long
Private m_hIml                                 As Long
Private m_lptrVb6ImageList                     As Long
Private m_lIconWidth                           As Long
Private m_lIconHeight                          As Long
Private m_iItemCount                           As Long
Private m_lSelItem                             As Long
Private m_lVisibleItems                        As Long
Private m_lDefaultItemHeight                   As Long
Private m_tItem()                              As ItemInfo
Private m_bMenuShown                           As Boolean
Private m_bSplitDown                           As Boolean
Private m_bDown                                As Boolean
Private m_bDesignMode                          As Boolean
Private m_SelectedItemFontBold                 As Boolean
Private m_bDrawTopCaptionIcon                  As Boolean
Private m_bRedraw                              As Boolean
Private m_hWnd                                 As Long
Private m_lVisibleItemsMax                     As Long
Private m_objLastitem                          As Object
Private Const m_def_VisibleItems               As Integer = 1
Private Const m_def_DefaultItemHeight          As Integer = 50
Private Const mDef_DefaultPanelHeight          As Long = 150
Private Const m_def_SelectedItemFontBold       As Boolean = True
Private Const m_def_DrawTopCaptionIcon         As Boolean = True
Private Const m_def_Redraw                     As Boolean = True
Private m_vImageList                           As Variant
Private m_bMouseOut                            As Boolean
Private mVisItemsMove                          As Long

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, _
      ByVal yPoint As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
      lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
      lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, _
      ByVal x As Long, _
      ByVal y As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
      lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, _
      lpPoint As POINTAPI) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
      ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, _
      ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, _
      ByVal nBkMode As Long) As Long
Private Declare Function ImageList_Draw Lib "COMCTL32.DLL" (ByVal hIml As Long, _
      ByVal i As Long, _
      ByVal hdcDst As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal fStyle As Long) As Long
Private Declare Function ImageList_GetIcon Lib "COMCTL32.DLL" (ByVal hIml As Long, _
      ByVal i As Long, _
      ByVal diIgnore As Long) As Long
Private Declare Function ImageList_GetImageCount Lib "COMCTL32.DLL" (ByVal hIml As Long) As Long
Private Declare Function ImageList_GetImageRect Lib "COMCTL32.DLL" (ByVal hIml As Long, _
      ByVal i As Long, _
      prcImage As RECT) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, _
      ByVal hBrush As Long, _
      ByVal lpDrawStateProc As Long, _
      ByVal lParam As Long, _
      ByVal wParam As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal cX As Long, _
      ByVal cY As Long, _
      ByVal fuFlags As Long) As Long

Private Declare Function InflateRect Lib "user32" (lpRect As RECT, _
      ByVal x As Long, _
      ByVal y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
      ByVal x As Long, _
      ByVal y As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, _
      ByVal X1 As Long, _
      ByVal y1 As Long, _
      ByVal x2 As Long, _
      ByVal y2 As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
      ByVal HPALETTE As Long, _
      pccolorref As Long) As Long


Const m_def_DisplayAddRemoveItemMenu = True
Const m_def_DisplayIconsInMenu = False
Dim m_DisplayAddRemoveItemMenu As Boolean
Dim m_DisplayIconsInMenu As Boolean

'-- Public Events
Event ItemRightClick(edItem As cEDItem)
Event ItemHidden(edItem As cEDItem)
Event ItemSelected(edItem As cEDItem)
Event ItemAdded(edItem As cEDItem)
Event ItemRemoved(edItem As cEDItem)
Event HoverItem(edItem As cEDItem)
Event ItemVisibleFromMenu(edItem As cEDItem)
Event ItemHiddenFromMenu(edItem As cEDItem)
Event HoverItemLeave()
Event BeginSizing()
Event EndSizing()
Event MenuShown()
Event MenuDestroyed()
Event MenuItemHover(Caption As String, IsAddRemoveItems As Boolean)

Private WithEvents m_Menus As IAPP_PopupMenu
Attribute m_Menus.VB_VarHelpID = -1

Const m_def_HideInfrequentlyUsedMenuItems = True
Dim m_HideInfrequentlyUsedMenuItems As Boolean
Dim m_PrarentHwnd As Long
Dim m_DisplayBannersInMenu As Boolean
Dim m_UseHandCursor As Boolean
Dim m_RightToLeft As Boolean
Dim m_DrawToolbarItemsRightToLeft As Boolean

Const m_def_DisplayBannersInMenu = True
Const m_def_UseHandCursor = True
Const m_def_RightToLeft = False
Const m_def_DrawToolbarItemsRightToLeft = True

Private m_lCustomHeaderColourOne    As OLE_COLOR
Private m_lCustomHeaderColourTwo    As OLE_COLOR
Private m_lCustomItemColourOne    As OLE_COLOR
Private m_lCustomItemColourTwo    As OLE_COLOR
Private m_lCustomItemSelectedColourOne    As OLE_COLOR
Private m_lCustomItemSelectedColourTwo    As OLE_COLOR
Private m_lCustomItemHoverColourOne    As OLE_COLOR
Private m_lCustomItemHoverColourTwo    As OLE_COLOR
Private m_lCustomBorderColour    As OLE_COLOR
Private m_lCustomItemSelectedDownColourOne As OLE_COLOR
Private m_lCustomItemSelectedDownColourTwo As OLE_COLOR

'Default Property Values:
'Const m_def_CaptionFontBold = True
Const m_def_SelectedItemForeColor = vbBlack
Const m_def_DisplayMenuChevron = True
Const m_def_DisplayHeader = True
Const m_def_HeaderTextColor = vbWhite
Const m_def_ButtonPressedOffset = 3
Const m_def_UseCustomColors = False
'Property Variables:
Dim m_CaptionFont As Font
'Dim m_CaptionFontBold As Boolean
Dim m_SelectedItemForeColor As OLE_COLOR
Dim m_DisplayMenuChevron As Boolean
Dim m_DisplayHeader As Boolean
Dim m_HeaderTextColor As OLE_COLOR
Dim m_ButtonPressedOffset As Integer
Dim m_UseCustomColors As Boolean

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
   BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   UserControl.BackColor() = New_BackColor
   PropertyChanged "BackColor"
End Property

Public Property Get ButtonPressedOffset() As Integer
   ButtonPressedOffset = m_ButtonPressedOffset
End Property

Public Property Let ButtonPressedOffset(ByVal New_ButtonPressedOffset As Integer)
   If New_ButtonPressedOffset > 3 Then
      New_ButtonPressedOffset = 3
   ElseIf New_ButtonPressedOffset < 0 Then
      New_ButtonPressedOffset = 0
   End If
   m_ButtonPressedOffset = New_ButtonPressedOffset
   PropertyChanged "ButtonPressedOffset"
End Property

Public Property Get Controls() As Object

   Set Controls = UserControl.Controls

End Property

Private Property Get DefaultItemHeight() As Long

   DefaultItemHeight = m_lDefaultItemHeight

End Property

Private Property Let DefaultItemHeight(ByVal New_DefaultItemHeight As Long)

   m_lDefaultItemHeight = New_DefaultItemHeight
   PropertyChanged "DefaultItemHeight"

End Property

Public Property Get DisplayAddRemoveItemMenu() As Boolean
   DisplayAddRemoveItemMenu = m_DisplayAddRemoveItemMenu
End Property

Public Property Let DisplayAddRemoveItemMenu(ByVal New_DisplayAddRemoveItemMenu As Boolean)
   m_DisplayAddRemoveItemMenu = New_DisplayAddRemoveItemMenu
   PropertyChanged "DisplayAddRemoveItemMenu"
End Property

Public Property Get DisplayBannersInMenu() As Boolean
   DisplayBannersInMenu = m_DisplayBannersInMenu
End Property

Public Property Let DisplayBannersInMenu(ByVal New_DisplayBannersInMenu As Boolean)
   m_DisplayBannersInMenu = New_DisplayBannersInMenu
   PropertyChanged "DisplayBannersInMenu"
End Property

Public Property Get DisplayIconsInMenu() As Boolean
   DisplayIconsInMenu = m_DisplayIconsInMenu
End Property

Public Property Let DisplayIconsInMenu(ByVal New_DisplayIconsInMenu As Boolean)
   m_DisplayIconsInMenu = New_DisplayIconsInMenu
   PropertyChanged "DisplayIconsInMenu"
End Property

'---------------------------------------------------------------------------------------
' Procedure : DrawControl
' DateTime  : 27/04/2004 16:12
' Author    : G_Noble
' Purpose   : Draws The Actual Control Items
'---------------------------------------------------------------------------------------
'
Private Sub DrawControl()

   Dim oColour    As OLE_COLOR
   Dim i          As Long
   Dim iVisCount  As Long
   Dim lHDC       As Long
   Dim lastleft   As Long
   Dim lBot       As Long
   Dim pDC        As Long
   Dim tR         As RECT
   Dim tRItem     As RECT

   Dim rcText     As RECT
   Dim RCCap      As RECT
   Dim tRItemX    As RECT
   Dim oGradOne   As OLE_COLOR
   Dim oGradTwo   As OLE_COLOR
   Dim bDrawChev As Boolean
   Dim hTheme As Long
   Dim hres As Long
   Dim va As Long
   Dim lHDCTo As Long
   Dim bChange As Boolean


   Dim bAppThemed As Boolean
   On Error Resume Next
   If Not Redraw Then
      Exit Sub
   End If


   '-- Setup Our Defaults
   bDrawChev = Me.DisplayMenuChevron

   bAppThemed = AppThemed
   DoEvents
   Cls
   picExtraItems.Cls

   lHDC = UserControl.hdc
   lHDCTo = lHDC
   pDC = picExtraItems.hdc
   lHDCTo = lHDC

   '-- Used For Drawing Onto picExtraItems
   lastleft = IIf(DrawToolbarItemsRightToLeft, ((picExtraItems.Width - (IIf(bDrawChev, 400, 220))) \ Screen.TwipsPerPixelX) - 10, 7)   '(m_lIconWidth / IIf(m_lIconWidth > 16, 1.2, 2)), 7)
   '-- Get The userControl Rect
   GetItemWindowRect tR
   tR.top = tR.top - 100
   '-- Set The PicExtraItems (Simulated Toolbar) Rect
   GetWindowRect picExtraItems.hwnd, tRItemX
   '-- Keep The Original UserColour
   oColour = UserControl.ForeColor
   '-- If We Have An ImageList Then We Offset The
   '-- Default Height
   If m_hIml > 0 Then
      DefaultItemHeight = m_lIconHeight + 3   '/ 1.3
   Else
      DefaultItemHeight = 3.5 + (TextHeight("A") \ Screen.TwipsPerPixelY)
   End If
   '-- UserControl Background
   UtilDrawBackground lHDC, Me.BackColor, Me.BackColor, 0, 0, tR.right - tR.left, tR.bottom - tR.top, False
   '-- Gradient Our Simulated Toolbar
   '-- And Draw A Line Round It
   If m_UseCustomColors Then
      UtilDrawBackground pDC, m_lCustomItemColourOne, m_lCustomItemColourTwo, 0, 0, picExtraItems.Width, (picExtraItems.ScaleHeight / Screen.TwipsPerPixelY)
      UtilDrawBorderRectangle pDC, m_lCustomBorderColour, 0, 0, (picExtraItems.ScaleWidth \ Screen.TwipsPerPixelX), (picExtraItems.ScaleHeight \ Screen.TwipsPerPixelY), False
   Else
      UtilDrawBackground pDC, m_lColorOneNormal, m_lColorTwoNormal, 0, 0, picExtraItems.Width, (picExtraItems.ScaleHeight / Screen.TwipsPerPixelY)
      UtilDrawBorderRectangle pDC, m_lColorBorder, 0, 0, (picExtraItems.ScaleWidth \ Screen.TwipsPerPixelX), (picExtraItems.ScaleHeight \ Screen.TwipsPerPixelY), False
   End If
   '-- Draw The Control Border
   If m_UseCustomColors Then
      UtilDrawBorderRectangle lHDC, m_lCustomBorderColour, tR.left, tR.top, tR.right - tR.left, tR.bottom - tR.top, False
   Else
      UtilDrawBorderRectangle lHDC, IIf(bAppThemed, GetSysColor(&H80000002 And &H1F&), GetSysColor(vbApplicationWorkspace And &H1F&)), tR.left, tR.top, tR.right - tR.left, tR.bottom - tR.top, False
   End If
   picSplitter.Visible = m_lSelItem > 0

   '-- Start To Draw Our Items
   If m_iItemCount > 0 Then
      '-- Set The BackMode
      SetBkMode lHDC, TRANSPARENT
      tR.top = picExtraItems.top / Screen.TwipsPerPixelY - ((((TextHeight("W") / Screen.TwipsPerPixelY)) + DefaultItemHeight) * (VisibleItems) - 1)
      '-- Make Sure If We Are Going To Selected
      '-- Anything The Mouse Pointer Is Within CoOrdinates
      m_lDefTop = tR.top
      '-- Draw Our Splitter
      '-- Gradient and Move The Splitter To the Correct CoOrdinates
      With picSplitter
         GetWindowRect .hwnd, rcText
         If m_UseCustomColors Then
            UtilDrawBackground .hdc, m_lCustomHeaderColourOne, m_lCustomHeaderColourTwo, 0, 0, .Width, (.Height \ Screen.TwipsPerPixelY) + 1.5
            picSplitter.Refresh
         Else
            UtilDrawBackground .hdc, m_lColorHeaderColorOne, m_lColorHeaderColorTwo, 0, 0, .Width, (.Height \ Screen.TwipsPerPixelY) + 1
            'UtilDrawBorderRectangle .hdc, m_lColorBorder, -5, 0, (picSplitter.ScaleWidth + 10), (picSplitter.ScaleHeight + 5), False
            picSplitter.Refresh
         End If

         If .top <> IIf(m_iItemCount <= 0, 0, 8 + ((tR.top) * Screen.TwipsPerPixelY) - .Height) Then
            .Move 15, IIf(m_iItemCount <= 0, 0, 8 + ((tR.top) * Screen.TwipsPerPixelY) - .Height), ScaleWidth - 25, .Height
         End If
         If m_iItemCount > 0 Then
            '-- Draw The Grab Handles
            Dim xs As Long
            xs = (Width \ Screen.TwipsPerPixelX) \ 2
            DrawDimple .hdc, xs, 1, False
            For i = 1 To 4
               DrawDimple .hdc, xs + (5 * i), 1, False
            Next
            For i = 4 To 1 Step -1
               DrawDimple .hdc, xs - (5 * i), 1, False
            Next
            picSplitter.Refresh
         End If

         '-- Keep A Note Of The Panels Default Height
      End With   'picSplitter
      m_lPanelTop = tR.top
      LSet tRItem = tR
      '-- Set The CoOrdinates For The First Visible Item
      With tRItem
         .left = .left
         .right = .right - 1
         .top = .top   '+ 1
         .bottom = .top + ((TextHeight("W") / Screen.TwipsPerPixelY)) + DefaultItemHeight
         If m_iItemCount > 0 Then
            picExtraItems.Height = m_const_lToolbarHeight
            If Not picExtraItems.Visible Then picExtraItems.Visible = True
         Else
            picExtraItems.Visible = False
         End If
      End With

      pDrawCaption

      If Me.Redraw Then
         '-- Start Drawing our Items
         For i = 1 To m_iItemCount
            With m_tItem(i)
               .bToolbar = False
               '-- Empty The Rect
               .oRect.top = 0
               .oRect.bottom = 0
               .oRect.left = 0
               .oRect.right = 0
            End With   'm_tItem(i)
            If iVisCount < VisibleItems Then
               If m_tItem(i).bVisible Then
                  iVisCount = iVisCount + 1
                  '-- Set The item Rect
                  LSet m_tItem(i).oRect = tRItem
                  m_tItem(i).bToolbar = False
                  '-- Draw State
                  If m_lItemHover = i Then

                     If i = m_lSelItem Then

                        If m_UseCustomColors Then
                           UtilDrawBackground lHDC, m_lCustomItemSelectedDownColourOne, m_lCustomItemSelectedDownColourTwo, tRItem.left, tRItem.top, tRItem.right, tRItem.bottom - tRItem.top
                        Else
                           UtilDrawBackground lHDC, m_lColorOneSelected, m_lColorTwoSelected, tRItem.left, tRItem.top, tRItem.right, tRItem.bottom - tRItem.top
                        End If
                        If m_tItem(i).bEnabled Then
                           ImageListDrawIcon m_lptrVb6ImageList, lHDC, m_hIml, m_tItem(i).lIconIndex, IIf(RightToLeft, (ScaleWidth \ Screen.TwipsPerPixelX) - m_lIconWidth - 10, m_lIconWidth / 2), tRItem.top + ((tRItem.bottom - tRItem.top) - m_lIconHeight) \ 2
                        Else
                           ImageListDrawIconDisabled m_lptrVb6ImageList, lHDC, m_hIml, m_tItem(i).lIconIndex, IIf(RightToLeft, (ScaleWidth \ Screen.TwipsPerPixelX) - m_lIconWidth - 10, m_lIconWidth / 2) + 2, 2 + tRItem.top + ((tRItem.bottom - tRItem.top) - m_lIconHeight) \ 2, m_lIconHeight, True
                        End If
                     ElseIf m_lBtnDown = i Then
                        If m_UseCustomColors Then
                           UtilDrawBackground lHDC, m_lCustomItemSelectedDownColourOne, m_lCustomItemSelectedDownColourTwo, tRItem.left, tRItem.top, tRItem.right, tRItem.bottom - tRItem.top
                        Else
                           UtilDrawBackground lHDC, m_lColorOneSelected, m_lColorTwoSelected, tRItem.left, tRItem.top, tRItem.right, tRItem.bottom - tRItem.top
                        End If

                        If m_tItem(i).bEnabled Then
                           ImageListDrawIcon m_lptrVb6ImageList, lHDC, m_hIml, m_tItem(i).lIconIndex, IIf(RightToLeft, (ScaleWidth \ Screen.TwipsPerPixelX) - m_lIconWidth - 10, m_lIconWidth / 2) + ButtonPressedOffset + 1, tRItem.top + ButtonPressedOffset + 1 + ((tRItem.bottom - tRItem.top) - m_lIconHeight) \ 2
                        Else
                           ImageListDrawIconDisabled m_lptrVb6ImageList, lHDC, m_hIml, m_tItem(i).lIconIndex, IIf(RightToLeft, (ScaleWidth \ Screen.TwipsPerPixelX) - m_lIconWidth - 10, m_lIconWidth / 2) + ButtonPressedOffset, ButtonPressedOffset + tRItem.top + ((tRItem.bottom - tRItem.top) - m_lIconHeight) \ 2, m_lIconHeight, True
                        End If
                     Else
                        If m_UseCustomColors Then
                           UtilDrawBackground lHDC, m_lCustomItemHoverColourOne, m_lCustomItemHoverColourTwo, tRItem.left, tRItem.top, tRItem.right, tRItem.bottom - tRItem.top
                        Else

                           UtilDrawBackground lHDC, m_lColorHotOne, m_lColorHotTwo, tRItem.left, tRItem.top, tRItem.right, tRItem.bottom - tRItem.top

                        End If
                        If m_tItem(i).bEnabled Then
                           'ImageListDrawIconDisabled m_lptrVb6ImageList, lHDC, m_hIml, m_tItem(i).lIconIndex, IIf(RightToLeft, (ScaleWidth \ Screen.TwipsPerPixelX) - m_lIconWidth - 10, m_lIconWidth / 2) + 2, 2 + tRItem.top + ((tRItem.bottom - tRItem.top) - m_lIconHeight) \ 2, m_lIconHeight, True
                           ImageListDrawIcon m_lptrVb6ImageList, lHDC, m_hIml, m_tItem(i).lIconIndex, IIf(RightToLeft, (ScaleWidth \ Screen.TwipsPerPixelX) - m_lIconWidth - 10, m_lIconWidth / 2), tRItem.top + ((tRItem.bottom - tRItem.top) - m_lIconHeight) \ 2
                        Else
                           ImageListDrawIconDisabled m_lptrVb6ImageList, lHDC, m_hIml, m_tItem(i).lIconIndex, IIf(RightToLeft, (ScaleWidth \ Screen.TwipsPerPixelX) - m_lIconWidth - 10, m_lIconWidth / 2) + 2, 2 + tRItem.top + ((tRItem.bottom - tRItem.top) - m_lIconHeight) \ 2, m_lIconHeight, True
                        End If
                     End If
                  Else
                     If i = m_lSelItem Then
                        If m_UseCustomColors Then
                           UtilDrawBackground lHDC, m_lCustomItemSelectedColourOne, m_lCustomItemSelectedColourTwo, tRItem.left, tRItem.top, tRItem.right, tRItem.bottom - tRItem.top
                        Else
                           UtilDrawBackground lHDC, m_lColorOneSelectedNormal, m_lColorTwoSelectedNormal, tRItem.left, tRItem.top, tRItem.right, tRItem.bottom - tRItem.top
                        End If
                     Else
                        If m_UseCustomColors Then
                           UtilDrawBackground lHDC, m_lCustomItemColourOne, m_lCustomItemColourTwo, tRItem.left, tRItem.top, tRItem.right, tRItem.bottom - tRItem.top
                        Else
                           UtilDrawBackground lHDC, m_lColorOneNormal, m_lColorTwoNormal, tRItem.left, tRItem.top, tRItem.right, tRItem.bottom - tRItem.top
                        End If
                     End If
                     If m_tItem(i).bEnabled Then
                        ImageListDrawIcon m_lptrVb6ImageList, lHDC, m_hIml, m_tItem(i).lIconIndex, IIf(RightToLeft, (ScaleWidth \ Screen.TwipsPerPixelX) - m_lIconWidth - 10, m_lIconWidth / 2), tRItem.top + ((tRItem.bottom - tRItem.top) - m_lIconHeight) \ 2
                     Else
                        ImageListDrawIconDisabled m_lptrVb6ImageList, lHDC, m_hIml, m_tItem(i).lIconIndex, IIf(RightToLeft, (ScaleWidth \ Screen.TwipsPerPixelX) - m_lIconWidth - 10, m_lIconWidth / 2) + 2, 2 + tRItem.top + ((tRItem.bottom - tRItem.top) - m_lIconHeight) \ 2, m_lIconHeight, True
                     End If
                  End If
                  tRItem.left = 5
                  LSet rcText = tRItem
                  If Not RightToLeft Then
                     rcText.left = (m_lIconWidth / 2) + m_lIconWidth + 10
                  Else
                     rcText.left = 0
                     rcText.right = (ScaleWidth \ Screen.TwipsPerPixelX) - ((m_lIconWidth / 2) + m_lIconWidth + 10)
                  End If

                  rcText.top = tRItem.top + ((tRItem.bottom - tRItem.top) - ((TextHeight("A") / Screen.TwipsPerPixelY))) \ 2   ' rcText.Top + (m_lIconHeight \ 4) '((rcText.Bottom - rcText.Top) - m_lIconHeight) \ 2

                  If (i = m_lBtnDown And i = m_lItemHover) And i <> m_lSelItem Then
                     OffsetRect rcText, ButtonPressedOffset, ButtonPressedOffset
                  End If
                  bChange = False

                  If SelectedItemFontBold Then
                     If i = m_lSelItem Then
                        If Not Me.Font.Bold Then Font.Bold = True: bChange = True
                     End If
                  End If


                  If i = m_lSelItem Then
                     UtilDrawText lHDC, m_tItem(i).sCaption, rcText.left, rcText.top, rcText.right, TextHeight(m_tItem(i).sCaption) * Screen.TwipsPerPixelY, IIf(UserControl.Enabled, IIf(m_tItem(i).bEnabled, True, False), False), SelectedItemForeColor, False, IIf(RightToLeft, True, False)
                  Else
                     UtilDrawText lHDC, m_tItem(i).sCaption, rcText.left, rcText.top, rcText.right, TextHeight(m_tItem(i).sCaption) * Screen.TwipsPerPixelY, IIf(UserControl.Enabled, IIf(m_tItem(i).bEnabled, True, False), False), oColour, False, IIf(RightToLeft, True, False)
                  End If

                  If bChange Then
                     If i = m_lSelItem Then
                        Font.Bold = Not Font.Bold
                     End If
                  End If


                  '-- Draw The Rectangle Round The Item
                  With tRItem
                     .left = 0
                     .right = .right + 2
                     '-- Draw The Border Rectangle
                     If m_UseCustomColors Then
                        UtilDrawBorderRectangle lHDC, m_lCustomBorderColour, .left, .top, (.right - .left) + 1, 1 + (.bottom - .top), False
                     Else
                        UtilDrawBorderRectangle lHDC, m_lColorBorder, .left, .top, (.right - .left) + 1, 1 + (.bottom - .top), False
                     End If
                     '-- Set The Next Item Rect
                     .top = .bottom
                     .bottom = .top + ((TextHeight("W") / Screen.TwipsPerPixelY)) + DefaultItemHeight
                     .right = .right - 2
                  End With
               End If
            Else
               If m_tItem(i).bVisible Then
                  '-- Draw Extra Items
                  '-- This Is Where We Simulate A Toolbar
                  m_tItem(i).bToolbar = True

                  With tRItemX
                     .top = 3
                     .left = (lastleft - 5)
                     .right = (.left + 26)
                     .bottom = (picExtraItems.ScaleHeight - 20) \ Screen.TwipsPerPixelY
                  End With
                  '-- Leave Space For The Chevrons
                  '-- Paint Only What We Can See
                  If IIf(Not DrawToolbarItemsRightToLeft, tRItemX.right < (picExtraItems.ScaleWidth / Screen.TwipsPerPixelX) - IIf(Me.DisplayMenuChevron, 5, 0), tRItemX.left > 0) Then

                     '-- Set The item Rect
                     LSet m_tItem(i).oRect = tRItemX
                     tRItemX.left = tRItemX.left + 2
                     tRItemX.top = tRItemX.top - 2
                     If m_lItemHover = i Then
                        If i = m_lSelItem Then
                           If m_UseCustomColors Then
                              UtilDrawBackground pDC, m_lCustomItemSelectedDownColourOne, m_lCustomItemSelectedDownColourTwo, tRItemX.left, tRItemX.top, tRItemX.right - tRItemX.left, tRItemX.bottom - tRItemX.top
                           Else
                              UtilDrawBackground pDC, m_lColorOneSelected, m_lColorTwoSelected, tRItemX.left, tRItemX.top, tRItemX.right - tRItemX.left, tRItemX.bottom - tRItemX.top
                           End If
                           If m_tItem(i).bEnabled Then
                              ImageListDrawIcon m_lptrVb6ImageList, pDC, m_hIml, m_tItem(i).lIconIndex, lastleft, (tRItemX.top + ((tRItemX.bottom - tRItemX.top) - m_lIconHeight) \ 2) - 2
                           Else
                              ImageListDrawIconDisabled m_lptrVb6ImageList, pDC, m_hIml, m_tItem(i).lIconIndex, lastleft + 2, 2 + tRItemX.top + ((tRItemX.bottom - tRItemX.top) - m_lIconHeight) \ 2, m_lIconHeight, True
                           End If
                        ElseIf m_lBtnDown = i Then
                           If m_UseCustomColors Then
                              UtilDrawBackground pDC, m_lCustomItemSelectedDownColourOne, m_lCustomItemSelectedDownColourTwo, tRItemX.left, tRItemX.top, tRItemX.right - tRItemX.left, tRItemX.bottom - tRItemX.top
                           Else
                              UtilDrawBackground pDC, m_lColorOneSelected, m_lColorTwoSelected, tRItemX.left, tRItemX.top, tRItemX.right - tRItemX.left, tRItemX.bottom - tRItemX.top
                           End If
                           If m_tItem(i).bEnabled Then
                              ImageListDrawIcon m_lptrVb6ImageList, pDC, m_hIml, m_tItem(i).lIconIndex, lastleft + 1, tRItemX.top + 1 + (((tRItemX.bottom - tRItemX.top) - m_lIconHeight) \ 2)
                           Else
                              ImageListDrawIconDisabled m_lptrVb6ImageList, pDC, m_hIml, m_tItem(i).lIconIndex, lastleft + 2, 2 + tRItemX.top + ((tRItemX.bottom - tRItemX.top) - m_lIconHeight) \ 2, m_lIconHeight, True
                           End If
                        Else   'NOT m_lBtnDown...
                           If m_UseCustomColors Then
                              UtilDrawBackground pDC, m_lCustomItemHoverColourOne, m_lCustomItemHoverColourTwo, tRItemX.left, tRItemX.top, tRItemX.right - tRItemX.left, tRItemX.bottom - tRItemX.top
                           Else
                              UtilDrawBackground pDC, m_lColorHotOne, m_lColorHotTwo, tRItemX.left, tRItemX.top, tRItemX.right - tRItemX.left, tRItemX.bottom - tRItemX.top
                           End If
                           If m_tItem(i).bEnabled Then
                              'ImageListDrawIconDisabled m_lptrVb6ImageList, pDC, m_hIml, m_tItem(i).lIconIndex, lastleft + 2, 2 + tRItemX.top + ((tRItemX.bottom - tRItemX.top) - 20) \ 2, 20, True
                              ImageListDrawIcon m_lptrVb6ImageList, pDC, m_hIml, m_tItem(i).lIconIndex, lastleft, tRItemX.top + ((tRItemX.bottom - tRItemX.top) - 20) \ 2
                           Else
                              ImageListDrawIconDisabled m_lptrVb6ImageList, pDC, m_hIml, m_tItem(i).lIconIndex, lastleft + 2, 2 + tRItemX.top + ((tRItemX.bottom - tRItemX.top) - 20) \ 2, 20, True
                           End If
                        End If
                     Else
                        If i = m_lSelItem Then
                           If m_UseCustomColors Then
                              UtilDrawBackground pDC, m_lCustomItemSelectedColourOne, m_lCustomItemSelectedColourTwo, tRItemX.left, tRItemX.top, tRItemX.right - tRItemX.left, tRItemX.bottom - tRItemX.top
                           Else
                              UtilDrawBackground pDC, m_lColorOneSelectedNormal, m_lColorTwoSelectedNormal, tRItemX.left, tRItemX.top, tRItemX.right - tRItemX.left, tRItemX.bottom - tRItemX.top
                           End If
                        End If
                        If m_tItem(i).bEnabled Then
                           ImageListDrawIcon m_lptrVb6ImageList, pDC, m_hIml, m_tItem(i).lIconIndex, lastleft, tRItemX.top + ((tRItemX.bottom - tRItemX.top) - 20) \ 2
                        Else
                           ImageListDrawIconDisabled m_lptrVb6ImageList, pDC, m_hIml, m_tItem(i).lIconIndex, lastleft + 2, 2 + tRItemX.top + ((tRItemX.bottom - tRItemX.top) - 20) \ 2, 20, True
                        End If

                     End If
                     LSet m_tItem(i).oRect = tRItemX
                     If i = m_lSelItem Or i = m_lItemHover Then
                        If m_UseCustomColors Then
                           'UtilDrawBorderRectangle pDC, m_lCustomBorderColour, (lastleft - 5), 2, ((m_lIconWidth + 10)), (tRItemX.bottom - 1), False
                        Else
                           '   UtilDrawBorderRectangle pDC, IIf(bAppThemed, dBlendColor((oGradTwo), vbBlack, 200), dBlendColor((oGradTwo), vbBlack, 200)), (lastleft - 5), 2, ((m_lIconWidth + 10)), (tRItemX.bottom - 1), False
                        End If
                     End If

                     '-- Set The Corodinates of The Next Item
                     If Me.DrawToolbarItemsRightToLeft Then
                        lastleft = lastleft - (16 + 8)
                        lastleft = lastleft - 1
                     Else
                        lastleft = lastleft + (16 + 8)
                        lastleft = lastleft + 1
                     End If
                  Else
                     Exit For
                  End If
               End If
            End If
         Next i
      End If
   End If

   If Me.DisplayMenuChevron Then

      '-- Chevron Button
      If m_iItemCount > 0 Then
         Dim xRC As RECT
         GetWindowRect picExtraItems.hwnd, xRC
         Dim xI As Integer
         Dim iNum As Integer
         iNum = picExtraItems.Height - 10

         If m_bMenuShown Or m_lItemHover = 99999 Then
            If m_UseCustomColors Then
                If m_lBtnDown = -2 Then
                    UtilDrawBackground pDC, m_lCustomHeaderColourTwo, m_lCustomHeaderColourOne, (picExtraItems.Width - (200)) \ Screen.TwipsPerPixelX, 1, 16, (xRC.bottom - xRC.top) - 2, False
                Else
                    UtilDrawBackground pDC, m_lCustomHeaderColourOne, m_lCustomHeaderColourTwo, (picExtraItems.Width - (200)) \ Screen.TwipsPerPixelX, 1, 16, (xRC.bottom - xRC.top) - 2, False
                End If
            Else
               If m_lBtnDown = -2 Then
                  UtilDrawBackground pDC, IIf(bAppThemed, m_lColorOneSelected, GetSysColor(vbApplicationWorkspace And &H1F&)), IIf(bAppThemed, m_lColorTwoSelected, dBlendColor(GetSysColor(vbApplicationWorkspace And &H1F&), vbWhite, 150)), (picExtraItems.Width - (200)) \ Screen.TwipsPerPixelX, 1, 16, (xRC.bottom - xRC.top) - 2, False
               Else
                  UtilDrawBackground pDC, IIf(bAppThemed, m_lColorTwoSelected, dBlendColor(GetSysColor(vbApplicationWorkspace And &H1F&), vbWhite, 150)), IIf(bAppThemed, m_lColorOneSelected, GetSysColor(vbApplicationWorkspace And &H1F&)), (picExtraItems.Width - (200)) \ Screen.TwipsPerPixelX, 1, 16, (xRC.bottom - xRC.top) - 2, False
               End If
            End If
            DrawChev picExtraItems.hwnd, picExtraItems.hdc, (picExtraItems.Width - (130)) \ Screen.TwipsPerPixelX, xRC.top, (xRC.right - xRC.left) - 2, xRC.bottom - xRC.top, Me.Enabled, dBlendColor(vbBlack, vbBlack, 150)
            '     If m_UseCustomColors Then UtilDrawBorderRectangle pDC, m_lCustomBorderColour, (picExtraItems.Width - (200)) \ Screen.TwipsPerPixelX, 0, ((xRC.right - xRC.left) + 20) \ Screen.TwipsPerPixelX, (xRC.bottom - xRC.top), False
         Else
            DrawChev picExtraItems.hwnd, picExtraItems.hdc, (picExtraItems.Width - (130)) \ Screen.TwipsPerPixelX, xRC.top, (xRC.right - xRC.left) - 2, xRC.bottom - xRC.top, Me.Enabled, vbBlack
         End If

      End If
   End If

   '-- Display The Panel if Any
   pPanelSize


   If m_UseCustomColors Then
      UtilDrawBorderRectangle pDC, m_lCustomBorderColour, 0, 0, (picExtraItems.ScaleWidth \ Screen.TwipsPerPixelX), (picExtraItems.ScaleHeight \ Screen.TwipsPerPixelY), False
      UtilDrawBorderRectangle lHDC, m_lCustomBorderColour, 0, 0, (ScaleWidth \ Screen.TwipsPerPixelX), (ScaleHeight \ Screen.TwipsPerPixelY), False
   Else
      UtilDrawBorderRectangle pDC, m_lColorBorder, 0, 0, (picExtraItems.ScaleWidth \ Screen.TwipsPerPixelX), (picExtraItems.ScaleHeight \ Screen.TwipsPerPixelY), False
      UtilDrawBorderRectangle lHDC, m_lColorBorder, 0, 0, (ScaleWidth \ Screen.TwipsPerPixelX), (ScaleHeight \ Screen.TwipsPerPixelY), False
   End If

   On Error GoTo 0
End Sub

Public Property Get DrawToolbarItemsRightToLeft() As Boolean
   DrawToolbarItemsRightToLeft = m_DrawToolbarItemsRightToLeft
End Property

Public Property Let DrawToolbarItemsRightToLeft(ByVal New_DrawToolbarItemsRightToLeft As Boolean)
   m_DrawToolbarItemsRightToLeft = New_DrawToolbarItemsRightToLeft
   PropertyChanged "DrawToolbarItemsRightToLeft"
   If m_bRedraw Then DrawControl
End Property

Public Property Get DrawTopCaptionIcon() As Boolean

   DrawTopCaptionIcon = m_bDrawTopCaptionIcon

End Property

Public Property Let DrawTopCaptionIcon(ByVal New_DrawTopCaptionIcon As Boolean)

   m_bDrawTopCaptionIcon = New_DrawTopCaptionIcon
   PropertyChanged "DrawTopCaptionIcon"
   DrawControl

End Property

Public Property Get Enabled() As Boolean

   Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

   UserControl.Enabled() = New_Enabled
   PropertyChanged "Enabled"
   DrawControl

End Property

Public Property Get EyeDropperItems() As IAPP_EDITemCollection

   Dim cT As New IAPP_EDITemCollection

   cT.Init ObjPtr(Me), Me.hwnd
   Set EyeDropperItems = cT

End Property

'---------------------------------------------------------------------------------------
' Procedure : fAdd
' DateTime  : 27/04/2004 16:15
' Author    : G_Noble
' Purpose   : Internal Add Calling Function
'---------------------------------------------------------------------------------------
Friend Function fAdd(Optional Key As Variant, _
                     Optional KeyBefore As Variant, _
                     Optional ByVal strCaption As String, _
                     Optional IconIndex As Long = -1) As cEDItem

   Dim i            As Long
   Dim iIndexBefore As Long
   Dim iItemIndex    As Long
   Dim cT           As New cEDItem
   Dim sKey         As String

   If Not IsMissing(Key) Then
      ' validate key.
      If IsNumeric(Key) Then
         ' invalid key
         Err.Raise 13, App.EXEName & ".EDITemControl"
         Exit Function
      End If
      On Error Resume Next
      sKey = Key
      If (Err.Number <> 0) Then
         ' invalid key
         On Error GoTo 0
         Exit Function
      End If
      On Error GoTo 0
      For i = 1 To m_iItemCount
         If (m_tItem(i).sKey = sKey) Then
            ' duplicate key
            Err.Raise 457, App.EXEName & ".EDITemControl"
            Exit Function
         End If
      Next i
   End If
   ' Check KeyBefore:
   iIndexBefore = 0
   If Not IsMissing(KeyBefore) Then
      On Error Resume Next
      iIndexBefore = ItemForKey(KeyBefore)
      If (Err.Number <> 0) Then
         Err.Raise Err.Number, App.EXEName & ".EDITemControl", Err.Description
         On Error GoTo 0
         Exit Function
      End If
      On Error GoTo 0
   End If
   ' Ok all checks passed. We can add the item.
   ' Check if this is an insert:
   m_iItemCount = m_iItemCount + 1
   If (m_iItemCount = 1) Then
      m_lSelItem = 1
   End If
   ReDim Preserve m_tItem(1 To m_iItemCount) As ItemInfo
   If (iIndexBefore > 0) Then
      ' Fix: should step backwards!
      For i = m_iItemCount - 1 To iIndexBefore Step -1
         LSet m_tItem(i + 1) = m_tItem(i)
      Next i
      iItemIndex = iIndexBefore
   Else
      iItemIndex = m_iItemCount
   End If
   ' set the info:
   With m_tItem(iItemIndex)
      .sCaption = strCaption
      .lIconIndex = IconIndex
      .bEnabled = True
      .bVisible = True
      .lID = nextId()
   End With   'm_tItem(iItemIndex)
   If LenB(sKey) = 0 Then
      m_tItem(iItemIndex).sKey = "I" & m_tItem(iItemIndex).lID
   Else
      m_tItem(iItemIndex).sKey = sKey
   End If
   If m_lVisibleItems = 0 Then
      m_lVisibleItems = 1
   End If
   cT.fInit ObjPtr(Me), m_hWnd, m_tItem(iItemIndex).lID
   RaiseEvent ItemAdded(Me.fItem(iItemIndex))
   Set fAdd = cT
   If m_iItemCount > 0 Then
      If VisibleItems = 1 Then
         VisibleItems = 0
      End If
   End If
   picSplitter.Visible = m_iItemCount > 0
   If Redraw Then
      DrawControl
   End If

End Function

Friend Property Get fEyeDropperItemselected(ByVal lID As Long) As Boolean

   Dim lIndex As Long

   If (getItemForId(lID, lIndex)) Then
      fEyeDropperItemselected = (lIndex = m_lSelItem)
   End If

End Property

Friend Property Let fEyeDropperItemselected(ByVal lID As Long, _
                                            ByVal bSelected As Boolean)


   Dim lIndex As Long

   If (getItemForId(lID, lIndex)) Then
      If Not (lIndex = m_lSelItem) Then
         m_lSelItem = lIndex
         DrawControl
         'pPanelSize
      End If
   End If

End Property

Friend Function fItem(Key As Variant)


   Dim cT     As New cEDItem
   Dim iIndex As Long

   On Error Resume Next
   iIndex = ItemForKey(Key)
   If (Err.Number <> 0) Then
      Err.Raise Err.Number, App.EXEName & ".EDITemControl", Err.Description
      On Error GoTo 0
      Exit Function
   End If
   On Error GoTo 0
   cT.fInit ObjPtr(Me), m_hWnd, m_tItem(iIndex).lID
   Set fItem = cT

End Function

Friend Property Get fItemCaption(ByVal lID As Long) As String

   Dim lIndex As Long

   If (getItemForId(lID, lIndex)) Then
      fItemCaption = m_tItem(lIndex).sCaption
   End If

End Property

Friend Property Let fItemCaption(ByVal lID As Long, _
                                 ByVal sCaption As String)

   Dim lIndex As Long

   If (getItemForId(lID, lIndex)) Then
      m_tItem(lIndex).sCaption = sCaption
      DrawControl
   End If

End Property

Friend Property Get fItemCount() As Long

   fItemCount = m_iItemCount

End Property

Friend Property Get fItemEnabled(ByVal lID As Long) As Boolean

   Dim lIndex As Long

   If (getItemForId(lID, lIndex)) Then
      fItemEnabled = m_tItem(lIndex).bEnabled
   End If

End Property

Friend Property Let fItemEnabled(ByVal lID As Long, _
                                 ByVal bEnabled As Boolean)

   Dim lIndex As Long

   If (getItemForId(lID, lIndex)) Then
      m_tItem(lIndex).bEnabled = bEnabled
      DrawControl
   End If

End Property

Friend Property Get fItemIconIndex(ByVal lID As Long) As Long

   Dim lIndex As Long

   If (getItemForId(lID, lIndex)) Then
      fItemIconIndex = m_tItem(lIndex).lIconIndex
   End If

End Property

Friend Property Let fItemIconIndex(ByVal lID As Long, _
                                   ByVal lIconIndex As Long)

   Dim o      As Object
   Dim lIndex As Long

   If (getItemForId(lID, lIndex)) Then
      m_tItem(lIndex).lIconIndex = lIconIndex
      If lIconIndex > -1 Then
         If m_hIml > 0 Then
            Set o = ObjectFromPtr(m_lptrVb6ImageList)
            '   ctxHookMenu1.SetBitmap mnuItems(lId + 1), o.ListImages(lIconIndex).Picture
         End If
      End If
      DrawControl
   End If

End Property

Friend Property Get fItemIndex(ByVal lID As Long) As Long

   Dim lIndex As Long

   If (getItemForId(lID, lIndex)) Then
      fItemIndex = lIndex
   End If

End Property

Friend Property Let fItemIndex(ByVal lID As Long, _
                               ByVal lIndex As Long)

   Dim lCurrentIndex As Long

   If (getItemForId(lID, lCurrentIndex)) Then
      If Not (lIndex = lCurrentIndex) Then
         If (lIndex > 0) And (lIndex <= m_iItemCount) Then
            ' replaceWithCandidate lCurrentIndex, lIndex
         Else
            ' New index out of range
            Err.Raise 9, App.EXEName & ".EDITemControl"
         End If
      End If
   End If

End Property

Friend Property Get fItemItemData(ByVal lID As Long) As Long

   Dim lIndex As Long

   If (getItemForId(lID, lIndex)) Then
      fItemItemData = m_tItem(lIndex).lItemData
   End If

End Property

Friend Property Let fItemItemData(ByVal lID As Long, _
                                  ByVal lItemData As Long)

   Dim lIndex As Long

   If (getItemForId(lID, lIndex)) Then
      m_tItem(lIndex).lItemData = lItemData
   End If

End Property

Friend Property Get fItemKey(ByVal lID As Long) As String

   Dim lIndex As Long

   If (getItemForId(lID, lIndex)) Then
      fItemKey = m_tItem(lIndex).sKey
   End If

End Property

Friend Property Get fItemPanel(ByVal lID As Long) As Object

   Dim ctlThis As Object
   Dim lIndex  As Long

   If (getItemForId(lID, lIndex)) Then
      ' Fix thanks to Matt Funnell: use lIndex not lId to find panel:
      If pbGetItemPanel(lIndex, ctlThis) Then
         Set fItemPanel = ctlThis
      End If
   End If

End Property

Friend Property Let fItemPanel(ByVal lID As Long, _
                               ByVal ctlThis As Object)

   Dim ctlPanel As Object
   Dim lIndex   As Long

   If (getItemForId(lID, lIndex)) Then
      If pbGetItemPanel(lIndex, ctlPanel) Then
         pbPanelVisible ctlPanel, False
      End If
      Set ctlThis.Container = UserControl.Extender
      m_tItem(lIndex).lObjPtrPanel = ObjPtr(ctlThis)
      If (lIndex = m_lSelItem) Then
         pPanelSize
      Else
         pbPanelVisible ctlThis, False
      End If
   End If

End Property

Friend Property Get fItemTag(ByVal lID As Long) As String

   Dim lIndex As Long

   If (getItemForId(lID, lIndex)) Then
      fItemTag = m_tItem(lIndex).sTag
   End If

End Property

Friend Property Let fItemTag(ByVal lID As Long, _
                             ByVal sTag As String)

   Dim lIndex As Long

   If (getItemForId(lID, lIndex)) Then
      m_tItem(lIndex).sTag = sTag
   End If

End Property

Friend Property Get fItemToolTipText(ByVal lID As Long) As String

   Dim lIndex As Long

   If (getItemForId(lID, lIndex)) Then
      fItemToolTipText = m_tItem(lIndex).sToolTipText
   End If

End Property

Friend Property Let fItemToolTipText(ByVal lID As Long, _
                                     ByVal sToolTipText As String)

   Dim lIndex As Long

   If (getItemForId(lID, lIndex)) Then
      m_tItem(lIndex).sToolTipText = sToolTipText
   End If

End Property

Public Property Get Font() As Font

   Set Font = UserControl.Font
   'MappingInfo=UserControl,UserControl,-1,Font

End Property

Public Property Set Font(ByVal New_Font As Font)

   Set UserControl.Font = New_Font
 '  Set m_fntOrig = UserControl.Font

   PropertyChanged "Font"
  
End Property

Public Property Get ForeColor() As OLE_COLOR

   ForeColor = UserControl.ForeColor
   'MappingInfo=UserControl,UserControl,-1,ForeColor

End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)

   UserControl.ForeColor() = New_ForeColor
   PropertyChanged "ForeColor"

End Property

Friend Sub fRemove(Key As Variant)

' Get Item to remove:

   Dim CTL       As Control
   Dim i         As Long
   Dim iToRemove As Long

   On Error Resume Next
   iToRemove = ItemForKey(Key)
   If (Err.Number <> 0) Then
      On Error GoTo 0
      Err.Raise Err.Number, App.EXEName & ".EDITemControl", Err.Description
      Exit Sub
   End If
   On Error GoTo 0
   RaiseEvent ItemRemoved(Me.EyeDropperItems.Item(iToRemove))
   ' its valid.
   If (pbGetItemPanel(iToRemove, CTL)) Then
      pbPanelVisible CTL, False
   End If
   If (m_iItemCount = 1) Then
      m_iItemCount = 0
      m_lSelItem = 0
      Erase m_tItem
   Else
      If (m_lSelItem = iToRemove) Then
         If (m_lSelItem = m_iItemCount) Then
            m_lSelItem = m_iItemCount - 1
         End If
      Else   'NOT (M_ISELITEM...
         If (m_lSelItem = m_iItemCount) Then
            m_lSelItem = m_iItemCount - 1
         Else   'NOT (M_ISELITEM...
            m_lSelItem = m_lSelItem - 1
         End If
      End If
      For i = iToRemove + 1 To m_iItemCount
         LSet m_tItem(i - 1) = m_tItem(i)
      Next i
      m_iItemCount = m_iItemCount - 1
      ReDim Preserve m_tItem(1 To m_iItemCount) As ItemInfo
   End If
   VisibleItems = VisibleItems - 1
   If m_iItemCount < m_lVisibleItems Then
      m_lVisibleItems = m_lVisibleItems - 1
   End If
   If VisibleItems <= 0 Then
      If m_iItemCount > 1 Then
         m_lVisibleItems = 1
      End If
   End If
   On Error Resume Next
   If m_iItemCount > 0 Then
      If Not m_tItem(m_lSelItem).bVisible Then
         m_lSelItem = 0
         For i = 1 To m_iItemCount
            If m_tItem(i).bVisible Then
               m_lSelItem = i
               Exit For
            Else
               m_lSelItem = 0
            End If
         Next i
      End If
   End If
   On Error GoTo 0
   VisibleItems = VisibleItems
   'picSplitter.Visible = m_iItemCount > 0
   'picSplitter.Visible = m_lSelItem > 0
   DrawControl

End Sub

Friend Property Get fVisible(ByVal lID As Long) As Boolean

   Dim lIndex As Long

   If (getItemForId(lID, lIndex)) Then
      fVisible = m_tItem(lIndex).bVisible
   End If

End Property

Friend Property Let fVisible(ByVal lID As Long, _
                             ByVal bVisible As Boolean)

   Dim i      As Long
   Dim lIndex As Long

   If (getItemForId(lID, lIndex)) Then

      If Not bVisible Then RaiseEvent ItemHidden(Me.EyeDropperItems.Item(lIndex))

      With m_tItem(lIndex)
         .bVisible = bVisible
         If Not .bToolbar Then
            EyeDropperItems.Item(.sKey).LastStateShown = True
            VisibleItems = VisibleItems
         Else
            EyeDropperItems.Item(.sKey).LastStateShown = False
            VisibleItems = VisibleItems
         End If
      End With   'm_tItem(lIndex)
   End If
   On Error Resume Next
   '-- Make Sure We Can Select
   If m_iItemCount > 0 Then
      If Not m_tItem(m_lSelItem).bVisible Then
         For i = 1 To m_iItemCount
            If m_tItem(i).bVisible Then
               m_lSelItem = i
               Exit For
            Else
               m_lSelItem = 0
            End If
         Next i
      End If
   End If
   VisibleItems = VisibleItems
   On Error GoTo 0

End Property

Private Function getItemForId(ByVal lID As Long, _
                              ByRef lIndex As Long) As Boolean

   Dim i As Long

   On Error Resume Next
   For i = 1 To m_iItemCount
      If (m_tItem(i).lID = lID) Then
         lIndex = i
         getItemForId = True
         Exit Function
      End If
   Next i
   'Err.Raise 9, App.EXEName & ".EDITemControl"
   On Error GoTo 0

End Function

Private Sub GetItemWindowRect(tR As RECT)

   GetClientRect m_hWnd, tR
   tR.bottom = picExtraItems.top
   tR.top = mDef_DefaultPanelHeight - 100

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbwhite
Public Property Get HeaderTextColor() As OLE_COLOR
   HeaderTextColor = m_HeaderTextColor
End Property

Public Property Let HeaderTextColor(ByVal New_HeaderTextColor As OLE_COLOR)
   m_HeaderTextColor = New_HeaderTextColor
   PropertyChanged "HeaderTextColor"
End Property

Public Property Get HideInfrequentlyUsedMenuItems() As Boolean
   HideInfrequentlyUsedMenuItems = m_HideInfrequentlyUsedMenuItems
End Property

Public Property Let HideInfrequentlyUsedMenuItems(ByVal New_HideInfrequentlyUsedMenuItems As Boolean)
   m_HideInfrequentlyUsedMenuItems = New_HideInfrequentlyUsedMenuItems
   PropertyChanged "HideInfrequentlyUsedMenuItems"
End Property

Private Function HitTest(x As Long, _
                         y As Long) As Object
   On Error GoTo errHitTest

   Dim i  As Long
   Dim rc As RECT
   Dim PT As POINTAPI

   GetCursorPos PT

   If Me.hwnd = WindowFromPoint(PT.x, PT.y) Then
      If y < m_lDefTop * Screen.TwipsPerPixelY Then
         Set m_objLastitem = Nothing
         Set HitTest = Nothing
         'RaiseEvent HoverItemLeave
         Exit Function
      End If
   End If
   If picExtraItems.hwnd = WindowFromPoint(PT.x, PT.y) Then
      If x >= picExtraItems.ScaleWidth - IIf(Me.DisplayMenuChevron, 200, 10) Then
         If y > 50 Then
            Set HitTest = picExtraItems
            If TypeName(sObject) = "cEDItem" Then
               RaiseEvent HoverItemLeave
            End If
            Exit Function
         End If
      End If
   End If
   For i = 1 To m_iItemCount
      LSet rc = m_tItem(i).oRect
      If PtInRect(rc, x \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY) Then
         Set HitTest = Me.fItem(i)
         If sObject Is Nothing Then RaiseEvent HoverItem(Me.fItem(i))
         If Not sObject Is Nothing Then
            If TypeName(HitTest) = "cEDItem" Then
               If HitTest.Caption <> sObject.Caption Then
                  RaiseEvent HoverItem(Me.fItem(i))
                  Debug.Print "Draw"
               End If
            Else
               RaiseEvent HoverItem(Me.fItem(i))
            End If
         End If

         Exit For
      End If
   Next i
CleanExit:
   On Error GoTo 0

   Exit Function
errHitTest:
   Set HitTest = Nothing
   Resume CleanExit

End Function

Public Property Get hwnd() As Long

   hwnd = UserControl.hwnd

End Property

Public Property Let ImageList(ByRef vImageList As Variant)
   On Error Resume Next

   Dim o  As Object

   Dim rc As RECT
   m_hIml = 0
   m_lptrVb6ImageList = 0
   m_lIconWidth = 0
   m_lIconHeight = 0
   m_hIml = 0
   m_lptrVb6ImageList = 0
   Set m_vImageList = vImageList

   If TypeName(vImageList) <> "ImageList" Then
      MsgBox "Imagelist:" & vbCrLf & "Image List Has To Be A Standard VB One" & vbCrLf & _
           "No Images Will Be Shown!", vbInformation
      Me.DrawTopCaptionIcon = False
      Exit Property '

   End If

   If (VarType(vImageList) = vbLong) Then
      ' Assume a handle to an image list:
      m_hIml = vImageList
   ElseIf (VarType(vImageList) = vbObject) Then
      ' Assume a VB image list:
      On Error Resume Next
      ' Get the image list initialised..
      vImageList.ListImages(1).Draw 0, 0, 0, 1
      m_hIml = vImageList.hImageList
      If (Err.Number = 0) Then
         ' Check for VB6 image list:
         If (TypeName(vImageList) = "ImageList") Then
'            If (vImageList.ListImages.Count = ImageList_GetImageCount(m_hIml)) Then
               Set o = vImageList
               m_lptrVb6ImageList = ObjPtr(o)
 '           End If
         End If
      Else
         Debug.Print "Failed to Get Image list Handle", "EDCTRL.ImageList"
      End If
      On Error GoTo 0
   End If
   If (m_hIml <> 0) Then
      If (m_lptrVb6ImageList <> 0) Then
         m_lIconWidth = vImageList.ImageWidth
         m_lIconHeight = vImageList.ImageHeight
         If (UserControl.Extender.Align = vbAlignLeft) Or (UserControl.Extender.Align = vbAlignRight) Then
            UserControl_Resize
         End If
      Else   'NOT (m_lptrVb6ImageList...
         ImageList_GetImageRect m_hIml, 0, rc
         m_lIconWidth = rc.right - rc.left
         m_lIconHeight = rc.bottom - rc.top
         If (UserControl.Extender.Align = vbAlignLeft) Or (UserControl.Extender.Align = vbAlignRight) Then
            UserControl_Resize
         End If
      End If
   End If


   On Error GoTo 0
   VisibleItems = VisibleItems
   DrawControl

End Property

Private Sub ImageListDrawIcon(ByVal ptrVb6ImageList As Long, _
                              ByVal lngHdc As Long, _
                              ByVal hIml As Long, _
                              ByVal iIconIndex As Long, _
                              ByVal lX As Long, _
                              ByVal lY As Long, _
                              Optional ByVal bSelected As Boolean = False, _
                              Optional ByVal bBlend25 As Boolean = False, Optional IsHeaderIcon As Boolean = False)


   Dim o      As Object
   Dim lFlags As Long
   Dim lR     As Long

   If Not Me.Enabled Then
      ImageListDrawIconDisabled ptrVb6ImageList, lngHdc, hIml, iIconIndex, lX, lY, m_lIconHeight, True
      Exit Sub
   End If
   lFlags = ILD_TRANSPARENT
   If (bSelected) Then
      lFlags = lFlags Or ILD_SELECTED
   End If
   If (ptrVb6ImageList <> 0) Then
      On Error Resume Next
      Set o = ObjectFromPtr(ptrVb6ImageList)
      If Not (o Is Nothing) Then
         If ((lFlags And ILD_SELECTED) = ILD_SELECTED) Then
            lFlags = 2   ' best we can do in VB6
         End If

         '            If (bBlend25) Then
         lFlags = lFlags Or ILD_BLEND25
         '           End If

         Dim icoInfo As ICONINFO
         Dim newICOinfo As ICONINFO
         Dim icoBMPinfo As BITMAP
         Call GetIconInfo(o.ListImages(iIconIndex + 1).ExtractIcon(), icoInfo)
         If playImage Then DestroyIcon playImage

         ' start a new icon structure
         CopyMemory newICOinfo, icoInfo, Len(icoInfo)

         ' get the icon dimensions from the bitmap portion of the icon
         GetGDIObject icoInfo.hbmColor, Len(icoBMPinfo), icoBMPinfo
         sourceWidth = IIf(m_lIconWidth > 16, 18, m_lIconWidth)
         sourceHeight = IIf(m_lIconWidth > 16, 18, m_lIconWidth)

         playImage = CreateIconIndirect(newICOinfo)

         If lngHdc = hdc Then
            If Not IsHeaderIcon Then
               DrawIconEx lngHdc, lX, lY, playImage, m_lIconWidth, m_lIconHeight, 0, 0, IIf(InIDE, &H3 Or lFlags, lFlags)
            Else
               DrawIconEx lngHdc, lX, lY, playImage, sourceWidth, sourceHeight, 0, 0, IIf(InIDE, &H3 Or lFlags, lFlags)
            End If

         Else
            DrawIconEx lngHdc, lX, 6, playImage, sourceWidth, sourceHeight, 0, 0, &H3 Or ILD_BLEND25
         End If
         DeleteObject newICOinfo.hbmMask
         DeleteObject newICOinfo.hbmColor

      End If
      On Error GoTo 0
   Else
      If (bBlend25) Then
         lFlags = lFlags Or ILD_BLEND25
      End If

      lR = ImageList_Draw(hIml, iIconIndex, lngHdc, lX, lY, lFlags)
      If (lR = 0) Then
         'Debug.Print "Failed to draw Image: " & iIconIndex & " onto hDC " & hdc, "ImageListDrawIcon"
      End If
   End If


End Sub

Private Sub ImageListDrawIconDisabled(ByVal ptrVb6ImageList As Long, _
                                      ByVal lngHdc As Long, _
                                      ByVal hIml As Long, _
                                      ByVal iIconIndex As Long, _
                                      ByVal lX As Long, _
                                      ByVal lY As Long, _
                                      ByVal lSize As Long, _
                                      Optional ByVal asShadow As Boolean)


   Dim o     As Object
   Dim hBr   As Long


   'Dim lR    As Long
   Dim hIcon As Long
   hIcon = 0


   If (ptrVb6ImageList <> 0) Then
      On Error Resume Next
      Set o = ObjectFromPtr(ptrVb6ImageList)
      If Not (o Is Nothing) Then
         hIcon = o.ListImages(iIconIndex + 1).ExtractIcon()
      End If
      On Error GoTo 0
   Else
      hIcon = ImageList_GetIcon(hIml, iIconIndex, 0)
   End If

   If (hIcon <> 0) Then
      If (asShadow) Then
         hBr = GetSysColorBrush(vb3DShadow And &H1F)
         If lngHdc = hdc Then
            Call DrawState(lngHdc, hBr, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_MONO)
         Else
            Call DrawState(lngHdc, hBr, 0, hIcon, 0, lX, lY + 4, 16, 16, DST_ICON Or DSS_MONO)
         End If
         DeleteObject hBr
      Else
         Call DrawState(lngHdc, 0, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_DISABLED)
      End If
      DestroyIcon hIcon
   End If

End Sub

Private Function InIDE() As Boolean

   On Error Resume Next
   'This function determines whether or not
   '     you're in development mode.
   On Error Resume Next
   Debug.Print (1 / 0)
   InIDE = (Err.Number <> 0)
   On Error GoTo 0

End Function

Public Function Initialise()

   If Not MTimer Is Nothing Then Set MTimer = Nothing
   Set MTimer = New IAPP_Timer
   MTimer.Interval = -1
   pCreateSubClass

End Function

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
' Do Nothing

End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
' Do Nothing
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

   If iMsg = WM_SYSCOLORCHANGE Then GetThemeName hwnd: GetGradientColors: DrawControl
   '-- Set The Visible Item State
   '-- This Is Used When The User Resizes The Parent Form
   If iMsg = WM_ENTERSIZEMOVE Then bBeginSize = True: mVisItemsMove = VisibleItems

End Function

Private Function ItemForKey(Key As Variant) As Long

   Dim lCheckIndex As Long
   Dim i           As Long

   If IsNumeric(Key) Then
      lCheckIndex = Key
      If (lCheckIndex < 0) Or (lCheckIndex > m_iItemCount) Then
         Err.Raise 9, App.EXEName & ".EDITemControl"
      Else
         ItemForKey = lCheckIndex
      End If
   Else
      For i = 1 To m_iItemCount
         If (m_tItem(i).sKey = Key) Then
            ItemForKey = i
            Exit Function
         End If
      Next i
      Err.Raise 9, App.EXEName & ".EDITemControl"
   End If

End Function

Private Sub m_Menus_Click(ItemNumber As Long)

   Dim IDX As Long
   Dim sKey As String
   sKey = m_Menus.ItemKey(ItemNumber)

   Select Case left(sKey, 1)
   Case "S"
      m_lSelItem = ItemForKey(right(sKey, Len(sKey) - 1))
   Case "V"
      Me.EyeDropperItems.Item(ItemForKey(right(sKey, Len(sKey) - 1))).Visible = Not Me.EyeDropperItems.Item(ItemForKey(right(sKey, Len(sKey) - 1))).Visible
      m_Menus.Checked(ItemNumber) = Not m_Menus.Checked(ItemNumber)
   Case "K"
      Select Case sKey
      Case "KM"
         VisibleItems = VisibleItems + 1
      Case "KL"
         VisibleItems = VisibleItems - 1
      End Select
   End Select

End Sub

Private Sub m_Menus_ItemHighlight(ItemNumber As Long, bEnabled As Boolean, bSeparator As Boolean)

   If Not bSeparator Then
      If left(m_Menus.ItemKey(ItemNumber), 1) = "V" Then
         RaiseEvent MenuItemHover(m_Menus.Caption(ItemNumber), True)
      Else
         RaiseEvent MenuItemHover(m_Menus.Caption(ItemNumber), False)
      End If
   End If



End Sub

Private Sub m_Menus_PopupMenuTerminated()

   m_lBtnDown = -1
   m_lItemHover = -1
   m_bMenuShown = False
   DrawControl
   RaiseEvent MenuDestroyed

End Sub

Private Sub mnuAddRemA_Click(Index As Integer)

   EyeDropperItems.Item(m_tItem(mnuAddRemA(Index).Tag).sKey).Visible = Not mnuAddRemA(Index).Checked

End Sub

Private Sub mnuItems_Click(Index As Integer)

   m_lSelItem = mnuItems(Index).Tag
   DrawControl

End Sub

Private Sub mnuLess_Click()

   Me.VisibleItems = Me.VisibleItems - 1

End Sub

Private Sub mnuMore_Click()

   Me.VisibleItems = Me.VisibleItems + 1

End Sub

Private Sub mTimer_ThatTime()

   Dim PT As POINTAPI

   GetCursorPos PT
   If ((hwnd) <> WindowFromPoint(PT.x, PT.y)) Then
      If (picExtraItems.hwnd) <> WindowFromPoint(PT.x, PT.y) Then

         If m_lBtnDown > 0 Then
            m_lItemHover = m_lItemHover   ' -1
            m_lBtnDown = m_lBtnDown
         Else
            m_lItemHover = -1
            m_lBtnDown = -1
         End If

         DrawControl
         'RaiseEvent HoverItemLeave
         UserControl.MousePointer = vbDefault
         m_bMouseOut = True
         MTimer.Interval = -1
         m_bMenuShown = False
      End If
   End If

End Sub

Private Function nextId() As Long

   m_lIdGenerator = m_lIdGenerator + 1
   nextId = m_lIdGenerator

End Function

Private Function pbGetItemPanel(ByVal lIndex As Long, _
                               ByRef ctlThis As Object) As Boolean

   Dim CTL  As Control
   Dim lPtr As Long

   For Each CTL In UserControl.ContainedControls
      lPtr = ObjPtr(CTL)
      If lPtr = m_tItem(lIndex).lObjPtrPanel Then
         Set ctlThis = CTL
         pbGetItemPanel = True
      End If
   Next   '  CTL CTL

End Function

Private Sub pbPanelVisible(ByRef ctlThis As Object, _
                           ByVal bState As Boolean)


   ctlThis.Visible = bState
   picExtraItems.ZOrder 0

End Sub

Private Sub pCreateSubClass()
   On Error Resume Next

   AttachMessage Me, hwnd, WM_SYSCOLORCHANGE
   AttachMessage Me, Parent.hwnd, WM_ENTERSIZEMOVE

   On Error GoTo 0
End Sub

Private Sub pDrawCaption()
   On Error Resume Next
   '-- Caption
   Dim RCCap As RECT
   Dim oGradOne As OLE_COLOR
   Dim oGradTwo As OLE_COLOR
   Dim oColor As OLE_COLOR
   Dim bChange As Boolean
   Dim bFnt As StdFont
   Set bFnt = Font

   Set Font = Me.CaptionFont

   '-- Store The Control Fore Color
   oColor = UserControl.ForeColor

   '-- Set The Caption Rectangle
   RCCap.top = 0
   RCCap.bottom = (TextHeight("W") \ Screen.TwipsPerPixelY) + 8  '+ DefaultItemHeight  '(picExtraItems.Height \ Screen.TwipsPerPixelY) - 5
   m_lPanelTopOffset = 0
   RCCap.right = ScaleWidth
   RCCap.left = 0
   If Not Me.DisplayHeader Then Exit Sub
   m_lPanelTopOffset = (TextHeight("W") \ Screen.TwipsPerPixelY) + 8   '+ DefaultItemHeight

   '-- Gradient The Caption Rectangle
   If m_lSelItem > 0 Then
      If m_UseCustomColors Then
         UtilDrawBackground hdc, m_lCustomHeaderColourOne, m_lCustomHeaderColourTwo, RCCap.left, RCCap.top, RCCap.right - RCCap.left, RCCap.bottom - RCCap.top
      Else
         UtilDrawBackground hdc, m_lColorHeaderColorOne, m_lColorHeaderColorTwo, RCCap.left, RCCap.top, RCCap.right - RCCap.left, RCCap.bottom - RCCap.top
      End If
      '-- Draw The Caption Text
      If Not RightToLeft Then
         RCCap.left = IIf(Me.DrawTopCaptionIcon, 30, 5)
         RCCap.right = ScaleWidth \ Screen.TwipsPerPixelX
      Else
         RCCap.left = 0
         RCCap.right = (ScaleWidth \ Screen.TwipsPerPixelX) - IIf(Me.DrawTopCaptionIcon, 40, 10)
      End If

      If Me.DrawTopCaptionIcon Then
         ImageListDrawIcon m_lptrVb6ImageList, hdc, m_hIml, m_tItem(m_lSelItem).lIconIndex, IIf(RightToLeft, (ScaleWidth \ Screen.TwipsPerPixelX) - 30, 8), RCCap.top + ((RCCap.bottom - RCCap.top) - IIf(m_lIconHeight > 16, 18, m_lIconHeight)) \ 2, , , True
      End If


      If m_lSelItem > 0 Then

         RCCap.top = RCCap.top + ((RCCap.bottom - RCCap.top) - (TextHeight("W") / Screen.TwipsPerPixelY)) \ 2
         RCCap.top = RCCap.top - 4
         RCCap.bottom = RCCap.bottom
         UserControl.ForeColor = HeaderTextColor
         UserControl.Font.Size = UserControl.Font.Size + 4
         UtilDrawText hdc, m_tItem(m_lSelItem).sCaption, RCCap.left, RCCap.top, (RCCap.right - RCCap.left), RCCap.bottom - RCCap.top, Me.Enabled, HeaderTextColor, False, RightToLeft
         UserControl.Font.Size = UserControl.Font.Size - 4

      End If
      '-- Draw The Border Round The Caption
      If m_UseCustomColors Then
         UtilDrawBorderRectangle hdc, m_lCustomBorderColour, 0, 0, ScaleWidth, (picExtraItems.Height \ Screen.TwipsPerPixelY) - 5, False
      Else
         UtilDrawBorderRectangle hdc, IIf(AppThemed, dBlendColor((oGradTwo), vbBlack, 200), GetSysColor(vbApplicationWorkspace And &H1F&)), 0, 0, ScaleWidth, m_lPanelTopOffset, False
      End If
      '-- ReDraw The Border Round The Control
   End If

   UserControl.ForeColor = oColor

   Set Font = bFnt

End Sub

Private Function pGetCurrentThemeColour(oColor As Long, oColor2 As Long, Optional IsRect As Boolean = False) As String


   oColor2 = dBlendColor(SlightlyLighterColour(GetSysColor(-2147483621 And &HFF&)), dBlendColor(GetSysColor(-2147483621 And &HFF&), vbWhite, 177), 250)
   oColor = dBlendColor(GetSysColor(-214748362 And &HFF&), vbWhite, 150)

   If AppThemed Then
      If IsRect Then oColor2 = vbActiveTitleBar: oColor = vbInactiveTitleBar
   Else
      If IsRect Then oColor2 = vbButtonShadow: oColor = dBlendColor(vbButtonShadow, vbWhite, 150)
   End If

End Function

Private Sub picExtraItems_MouseDown(Button As Integer, _
                                    Shift As Integer, _
                                    x As Single, _
                                    y As Single)

   UserControl_MouseDown Button, Shift, x, y

End Sub

Private Sub picExtraItems_MouseMove(Button As Integer, _
                                    Shift As Integer, _
                                    x As Single, _
                                    y As Single)

   UserControl_MouseMove Button, Shift, x, y
   xMenu = x
   yMenu = y

End Sub

Private Sub picExtraItems_MouseUp(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)

   UserControl_MouseUp Button, Shift, x, y

End Sub

Private Sub picSplitter_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)

   Dim PT As POINTAPI

   m_bSplitDown = True
   UserControl.MousePointer = 7
   GetCursorPos PT
   ScreenToClient Me.hwnd, PT
   m_lLastY = PT.y
   RaiseEvent BeginSizing


End Sub

Private Sub picSplitter_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)

   Dim PT As POINTAPI

   If m_bSplitDown Then
      GetCursorPos PT
      ScreenToClient Me.hwnd, PT
      If y < m_lLastY - DefaultItemHeight Then
         If y <= (m_lLastY) - ((TextHeight("W") / Screen.TwipsPerPixelY)) + DefaultItemHeight Then
            If (PT.y \ 2) > (picExtraItems.Height \ Screen.TwipsPerPixelY) Then
               VisibleItems = VisibleItems + 1
               m_lLastY = DefaultItemHeight / 2 - (y \ Screen.TwipsPerPixelY)
            End If
         End If
      Else
         If y + DefaultItemHeight >= (m_lLastY) + ((TextHeight("W") / Screen.TwipsPerPixelY)) + DefaultItemHeight Then
            VisibleItems = VisibleItems - 1
            m_lLastY = DefaultItemHeight / 2 + y \ Screen.TwipsPerPixelY
         End If
      End If
   End If

End Sub

Private Sub picSplitter_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)

   UserControl.MousePointer = vbDefault
   m_bSplitDown = False
   RaiseEvent EndSizing

End Sub

Private Sub pPanelSize()


   Dim ctlPanel As Control
   Dim CTL      As Control
   Dim lItem     As Long

   If m_iItemCount > 0 Then
      lItem = m_lSelItem
      If lItem > 0 Then
         If pbGetItemPanel(lItem, ctlPanel) Then
            On Error Resume Next
            If picSplitter.top \ Screen.TwipsPerPixelY > IIf(Me.DisplayHeader, (picExtraItems.Height \ Screen.TwipsPerPixelY), 10) Then
               If Me.DisplayHeader Then
                  ctlPanel.Move 20, IIf(Me.DisplayHeader, m_lPanelTopOffset * Screen.TwipsPerPixelY, 20), ScaleWidth - 30, 80 + ((m_lPanelTop * Screen.TwipsPerPixelY) - (m_lPanelTopOffset * Screen.TwipsPerPixelY) - picSplitter.Height - 70)
               Else
                  ctlPanel.Move 20, IIf(Me.DisplayHeader, m_lPanelTopOffset * Screen.TwipsPerPixelY, 20), ScaleWidth - 30, 80 + ((m_lPanelTop * Screen.TwipsPerPixelY) - (190))
               End If
            Else
               ctlPanel.Move 20, -50, ScaleWidth - 30, 0
            End If
            On Error GoTo 0
            pbPanelVisible ctlPanel, True
            
         End If
      End If
   End If
   For Each CTL In UserControl.ContainedControls
      If CTL Is ctlPanel Then
      Else
         pbPanelVisible CTL, False
      End If
   Next

End Sub

Private Sub pShowMenu(Optional bShow As Boolean = False)

   Dim o      As ImageList
   Dim IDX    As Long
   Dim idxX   As Long
   Dim i      As Integer
   Dim j As Long
   Dim k As Long
   Dim iIndex As Long
   Dim iCount As Long
   Dim bNoImages As Boolean

   On Error Resume Next


   If Me.EyeDropperItems.Count <= 0 Then Exit Sub


   With m_Menus   'NewMenu

      bNoImages = m_lptrVb6ImageList = 0

      .hWndOwner = picExtraItems.hwnd

      '-- Set The Style
      .OfficeXpStyle = IIf(IsXp, IIf(AppThemed, True, False), False)

      '-- Clear The Existing Items If Any
      If Not .CurrentlyRestoredKey = "Customise" Then
         '.Restore "Customise"
         .Clear
      End If

      'If Not bHide Then
      .HideInfrequentlyUsed = HideInfrequentlyUsedMenuItems

      '-- Add The Header Form Main items
      If DisplayBannersInMenu Then
         k = .AddItem("Selection Options")
         m_Menus.Header(k) = True
         m_Menus.HeaderStyle = ecnmHeaderCaptionBar
      End If

      '-- Add The Visible Items
      For i = 1 To Me.EyeDropperItems.Count
         If Me.EyeDropperItems.Item(i).Visible Then
            iCount = iCount + 1
            If Me.DisplayIconsInMenu Then
               If bNoImages Then
                  k = .AddItem(Me.EyeDropperItems.Item(i).Caption, , , , , IIf(i = m_lSelItem, True, False), , "S" & Me.EyeDropperItems.Item(i).Key)
               Else
                  k = .AddItem(Me.EyeDropperItems.Item(i).Caption, , , , Me.EyeDropperItems.Item(i).IconIndex, IIf(i = m_lSelItem, True, False), , "S" & Me.EyeDropperItems.Item(i).Key)
               End If
            Else
               k = .AddItem(Me.EyeDropperItems.Item(i).Caption, , , , , IIf(i = m_lSelItem, True, False), , "S" & Me.EyeDropperItems.Item(i).Key)
            End If

            If i Mod 2 Then
               If HideInfrequentlyUsedMenuItems Then
                  .ItemInfrequentlyUsed(k) = IIf(i <> m_lSelItem, True, False)
               End If
            End If

         End If
      Next

      '-- Seperator
      k = .AddItem("-", , , j)
      .ItemInfrequentlyUsed(k) = HideInfrequentlyUsedMenuItems

      '-- Add Header For Options
      If DisplayBannersInMenu Then
         k = .AddItem("Display Options")
         .Header(k) = True
         .HeaderStyle = ecnmHeaderCaptionBar
         .ItemInfrequentlyUsed(k) = HideInfrequentlyUsedMenuItems
      End If
      k = .AddItem("Show More..." & vbTab & "Ctrl+M", , , j, , , IIf((VisibleItems < m_lVisibleItemsMax), True, False), "KM")
      .ItemInfrequentlyUsed(k) = HideInfrequentlyUsedMenuItems
      k = .AddItem("Show Less..." & vbTab & "Ctrl+L", , , j, , , (VisibleItems > 1 And m_iItemCount > 1), "KL")
      .ItemInfrequentlyUsed(k) = HideInfrequentlyUsedMenuItems

      If Me.EyeDropperItems.Count > 0 Then
         If Me.DisplayAddRemoveItemMenu Then
            k = .AddItem("-", , , j)
            .ItemInfrequentlyUsed(k) = HideInfrequentlyUsedMenuItems
            j = .AddItem("&Add or Remove Buttons")

            If DisplayBannersInMenu Then
               idxX = .AddItem("Add Remove Buttons", , , j)
               m_Menus.Header(idxX) = True
               m_Menus.HeaderStyle = ecnmHeaderCaptionBar
            End If

            For i = 1 To Me.EyeDropperItems.Count
               If Me.DisplayIconsInMenu Then
                  If bNoImages Then
                     k = .AddItem(Me.EyeDropperItems.Item(i).Caption, , , j, , Me.EyeDropperItems.Item(i).Visible, , "V" & Me.EyeDropperItems.Item(i).Key)
                     .ShowCheckAndIcon(k) = False
                  Else
                     k = .AddItem(Me.EyeDropperItems.Item(i).Caption, , , j, Me.EyeDropperItems.Item(i).IconIndex, Me.EyeDropperItems.Item(i).Visible, , "V" & Me.EyeDropperItems.Item(i).Key)
                     .ShowCheckAndIcon(k) = Me.DisplayIconsInMenu
                  End If
               Else
                  k = .AddItem(Me.EyeDropperItems.Item(i).Caption, , , j, , Me.EyeDropperItems.Item(i).Visible, , "V" & Me.EyeDropperItems.Item(i).Key)
                  .ShowCheckAndIcon(k) = False
               End If
            Next
         End If
      End If


      If Not bNoImages Then
         .ImageList = m_vImageList
      End If

      If Me.DisplayIconsInMenu Then
         If bNoImages Then
            .ResetImageHeightAndWidth = 16
         Else
            .ResetImageHeightAndWidth = 16
         End If

      Else
         .ResetImageHeightAndWidth = 16
      End If

      m_bMenuShown = True

      Dim PT As POINTAPI
      ClientToScreen picExtraItems.hwnd, PT
      GetCursorPos PT
      'ClientToScreen picExtraItems.hwnd, PT
      PT.x = picExtraItems.ScaleWidth - 150
      PT.y = picExtraItems.ScaleTop - (picExtraItems.Height)
      ClientToScreen picExtraItems.hwnd, PT
      RaiseEvent MenuShown

      .ShowPopupMenu xMenu, yMenu

      If (iIndex > 0) Then
         m_Menus_Click iIndex
         .Store "Customise"
      End If
   End With

   On Error GoTo 0

End Sub

Public Property Get Redraw() As Boolean

   Redraw = m_bRedraw

End Property

Public Property Let Redraw(ByVal New_Redraw As Boolean)

   m_bRedraw = New_Redraw
   PropertyChanged "Redraw"
   DoEvents
   If New_Redraw Then
      VisibleItems = VisibleItems
      DrawControl
   End If

End Property

Public Property Get RightToLeft() As Boolean
   RightToLeft = m_RightToLeft
End Property

Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)
   m_RightToLeft = New_RightToLeft
   PropertyChanged "RightToLeft"
End Property

Public Property Get ScaleMode() As Integer

   ScaleMode = UserControl.ScaleMode

End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)

   UserControl.ScaleMode() = New_ScaleMode
   PropertyChanged "ScaleMode"

End Property

Public Property Get SelectedItem() As cEDItem

   Dim cT As New cEDItem

   If (m_lSelItem > 0) Then
      If (m_iItemCount > 0) Then
         cT.fInit ObjPtr(Me), m_hWnd, m_tItem(m_lSelItem).lID
         Set SelectedItem = cT
      End If
   End If

End Property

Public Property Get SelectedItemFontBold() As Boolean

   SelectedItemFontBold = m_SelectedItemFontBold
   'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
   'MemberInfo=0,0,0,True

End Property

Public Property Let SelectedItemFontBold(ByVal New_SelectedItemFontBold As Boolean)

   m_SelectedItemFontBold = New_SelectedItemFontBold
   PropertyChanged "SelectedItemFontBold"

End Property

Public Function SetCustomProperties(HeaderColorOne As OLE_COLOR, _
                HeaderColorTwo As OLE_COLOR, ItemColorOne As OLE_COLOR, _
                ItemColorTwo As OLE_COLOR, ItemSelectedColorOne As OLE_COLOR, _
                ItemSelectedColorTwo As OLE_COLOR, ItemHoverColorOne As OLE_COLOR, _
                ItemHoverColorTwo As OLE_COLOR, BorderColor As OLE_COLOR, ItemSelectedDownColourOne As OLE_COLOR, _
                ItemSelectedDownColourTwo As OLE_COLOR)


   m_lCustomHeaderColourOne = HeaderColorOne
   m_lCustomHeaderColourTwo = HeaderColorTwo
   m_lCustomItemColourOne = ItemColorOne
   m_lCustomItemColourTwo = ItemColorTwo
   m_lCustomItemSelectedColourOne = ItemSelectedColorOne
   m_lCustomItemSelectedColourTwo = ItemSelectedColorTwo
   m_lCustomItemHoverColourOne = ItemHoverColorOne
   m_lCustomItemHoverColourTwo = ItemHoverColorTwo
   m_lCustomBorderColour = BorderColor
   m_lCustomItemSelectedDownColourOne = ItemSelectedDownColourOne
   m_lCustomItemSelectedDownColourTwo = ItemSelectedDownColourTwo
   GetGradientColors
   DrawControl

End Function

Public Function TextHeight(ByVal strSTR As String) As Single


   TextHeight = UserControl.TextHeight(strSTR)

End Function

Public Function TextWidth(ByVal strSTR As String) As Single


   TextWidth = UserControl.TextWidth(strSTR)

End Function

Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                                Optional hPal As Long = 0) As Long

' Convert Automation color to Windows color
'--------- Drawing

   If OleTranslateColor(oClr, hPal, TranslateColor) Then
      TranslateColor = CLR_INVALID
   End If

End Function

Public Property Get UseCustomColors() As Boolean
   UseCustomColors = m_UseCustomColors
End Property

Public Property Let UseCustomColors(ByVal New_UseCustomColors As Boolean)
   m_UseCustomColors = New_UseCustomColors
   DrawControl
   PropertyChanged "UseCustomColors"
End Property

Public Property Get UseHandCursor() As Boolean
   UseHandCursor = m_UseHandCursor
End Property

Public Property Let UseHandCursor(ByVal New_UseHandCursor As Boolean)
   m_UseHandCursor = New_UseHandCursor
   PropertyChanged "UseHandCursor"
End Property

Private Sub UserControl_Initialize()

   Call VerInitialise

   m_hWnd = UserControl.hwnd

   Set MTimer = New IAPP_Timer
   Set m_Menus = New IAPP_PopupMenu

   AppThemed



End Sub

Private Sub UserControl_InitProperties()

   m_lDefaultItemHeight = m_def_DefaultItemHeight
   Set UserControl.Font = Ambient.Font
   m_lVisibleItems = m_def_VisibleItems
   m_bRedraw = m_def_Redraw
   m_bDrawTopCaptionIcon = m_def_DrawTopCaptionIcon
   m_SelectedItemFontBold = m_def_SelectedItemFontBold

   m_DisplayAddRemoveItemMenu = m_def_DisplayAddRemoveItemMenu
   m_DisplayIconsInMenu = m_def_DisplayIconsInMenu
   m_HideInfrequentlyUsedMenuItems = m_def_HideInfrequentlyUsedMenuItems

   m_DrawToolbarItemsRightToLeft = m_def_DrawToolbarItemsRightToLeft
   m_RightToLeft = m_def_RightToLeft
   m_UseHandCursor = m_def_UseHandCursor
   m_DisplayBannersInMenu = m_def_DisplayBannersInMenu
   m_UseCustomColors = m_def_UseCustomColors
   m_ButtonPressedOffset = m_def_ButtonPressedOffset
   m_HeaderTextColor = m_def_HeaderTextColor
   m_DisplayHeader = m_def_DisplayHeader
   m_DisplayMenuChevron = m_def_DisplayMenuChevron
   m_SelectedItemForeColor = m_def_SelectedItemForeColor
   '    m_CaptionFontBold = m_def_CaptionFontBold
   Set m_CaptionFont = Ambient.Font
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
   MsgBox KeyCode

End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)

   Dim oItem As Object
   Dim PT    As POINTAPI

   GetCursorPos PT
   If Me.hwnd = WindowFromPoint(PT.x, PT.y) Then
      If y <= m_lDefTop Then
         Exit Sub
      End If
   End If
   If Button = vbLeftButton Then
      Set oItem = HitTest(CLng(x), CLng(y))
      If Not oItem Is Nothing Then
         If TypeName(oItem) = "cEDItem" Then
            m_lBtnDown = oItem.Index
         ElseIf TypeName(oItem) = "PictureBox" Then
            m_lBtnDown = -2
            m_bDown = True
            Debug.Print "IOIOIOIO"
            DoEvents
            DrawControl
            'pShowMenu True
            
            
         Else   'NOT NOT...
            m_lBtnDown = -1
         End If
      End If
      m_bDown = True
   ElseIf Button = vbRightButton Then
      Set oItem = HitTest(CLng(x), CLng(y))
      If Not oItem Is Nothing Then
         If TypeName(oItem) = "cEDItem" Then
            RaiseEvent ItemRightClick(Me.EyeDropperItems.Item(oItem.Index))
         End If
      End If

   End If
   DrawControl

End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)

   If m_iItemCount = 0 Then Exit Sub

   Dim PT As POINTAPI


   Dim oItem      As Object

   MTimer.Interval = 100
   m_lItemHover = -1

   If m_bMouseOut Then
      Set sObject = Nothing
   End If

   m_bMouseOut = False

   Set oItem = HitTest(CLng(x), CLng(y))

   If TypeName(oItem) = "cEDItem" Then

      m_lItemHover = oItem.Index
      picExtraItems.ToolTipText = oItem.Caption
      If Not sObject Is Nothing Then
         If TypeName(sObject) = "cEDItem" Then
            If oItem.Caption <> sObject.Caption Then
               DrawControl
            End If
         End If
      Else
         DrawControl
      End If

      If UserControl.MousePointer <> vbCustom Then If UseHandCursor Then UserControl.MousePointer = vbCustom

   ElseIf TypeName(oItem) = "PictureBox" Then

      If Me.DisplayMenuChevron Then
         m_lItemHover = 99999
         DrawControl
         If UserControl.MousePointer <> vbCustom Then If UseHandCursor Then UserControl.MousePointer = vbCustom
         Set sObject = Nothing
      End If
   Else

      picExtraItems.ToolTipText = ""
      m_lItemHover = -1
      UserControl.MousePointer = vbDefault
      If Not sObject Is Nothing Then RaiseEvent HoverItemLeave
      DrawControl
   End If

   Set sObject = oItem



End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)

   Dim oItem As Object   ' cEDItem

   m_bDown = False
   Set oItem = HitTest(CLng(x), CLng(y))
   If Not oItem Is Nothing Then
      If TypeName(oItem) = "cEDItem" Then
         If Button = vbLeftButton Then
            RaiseEvent ItemSelected(Me.EyeDropperItems.Item(oItem.Index))
         Else
            RaiseEvent ItemRightClick(Me.EyeDropperItems.Item(oItem.Index))
         End If
         With oItem
            If m_lBtnDown = .Index Then
               m_lSelItem = .Index
               m_lItemHover = .Index
               pPanelSize
            End If
         End With   'oItem
         m_lBtnDown = -1
      ElseIf TypeName(oItem) = "PictureBox" Then
        m_lBtnDown = -2
            DrawControl
            pShowMenu True
            m_lBtnDown = -1
         'm_lItemHover = -
      Else   'NOT NOT...
         m_lItemHover = -1
         m_lSelItem = m_lSelItem
      End If
   End If
   DrawControl

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   m_bDesignMode = (UserControl.Ambient.UserMode)
   GetThemeName Me.hwnd
   GetGradientColors
   With PropBag
      m_lDefaultItemHeight = .ReadProperty("DefaultItemHeight", m_def_DefaultItemHeight)
      Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
      UserControl.Enabled = .ReadProperty("Enabled", True)
      UserControl.ForeColor = .ReadProperty("ForeColor", &H80000008)
      m_lVisibleItems = .ReadProperty("VisibleItems", m_def_VisibleItems)
   End With   'PropBag
   With PropBag
      UserControl.ScaleMode = .ReadProperty("ScaleMode", 1)
      m_bRedraw = .ReadProperty("Redraw", m_def_Redraw)
      m_bDrawTopCaptionIcon = .ReadProperty("DrawTopCaptionIcon", m_def_DrawTopCaptionIcon)
      m_SelectedItemFontBold = .ReadProperty("SelectedItemFontBold", m_def_SelectedItemFontBold)
   End With   'PropBag

   m_DisplayAddRemoveItemMenu = PropBag.ReadProperty("DisplayAddRemoveItemMenu", m_def_DisplayAddRemoveItemMenu)
   m_DisplayIconsInMenu = PropBag.ReadProperty("DisplayIconsInMenu", m_def_DisplayIconsInMenu)
   m_HideInfrequentlyUsedMenuItems = PropBag.ReadProperty("HideInfrequentlyUsedMenuItems", m_def_HideInfrequentlyUsedMenuItems)
   m_DrawToolbarItemsRightToLeft = PropBag.ReadProperty("DrawToolbarItemsRightToLeft", m_def_DrawToolbarItemsRightToLeft)
   m_RightToLeft = PropBag.ReadProperty("RightToLeft", m_def_RightToLeft)
   m_UseHandCursor = PropBag.ReadProperty("UseHandCursor", m_def_UseHandCursor)
   m_DisplayBannersInMenu = PropBag.ReadProperty("DisplayBannersInMenu", m_def_DisplayBannersInMenu)
   UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
   m_UseCustomColors = PropBag.ReadProperty("UseCustomColors", m_def_UseCustomColors)
   m_ButtonPressedOffset = PropBag.ReadProperty("ButtonPressedOffset", m_def_ButtonPressedOffset)
   m_HeaderTextColor = PropBag.ReadProperty("HeaderTextColor", m_def_HeaderTextColor)
   m_DisplayHeader = PropBag.ReadProperty("DisplayHeader", m_def_DisplayHeader)
   m_DisplayMenuChevron = PropBag.ReadProperty("DisplayMenuChevron", m_def_DisplayMenuChevron)
   m_SelectedItemForeColor = PropBag.ReadProperty("SelectedItemForeColor", m_def_SelectedItemForeColor)
   '    m_CaptionFontBold = PropBag.ReadProperty("CaptionFontBold", m_def_CaptionFontBold)
   Set m_CaptionFont = PropBag.ReadProperty("CaptionFont", Ambient.Font)
End Sub

Private Sub UserControl_Resize()

   On Error Resume Next

   '-- Reurn The visible Items To The Original Save State
   VisibleItems = mVisItemsMove

   '-- Sets The Visible Items To The Max Or Min Available
   If Me.DisplayHeader Then
      m_lVisibleItemsMax = ((ScaleHeight \ Screen.TwipsPerPixelY) / (m_lIconHeight * 1.5)) - 3
   Else
      m_lVisibleItemsMax = ((ScaleHeight \ Screen.TwipsPerPixelY) / (picExtraItems.Height \ Screen.TwipsPerPixelY)) - 3
   End If
   If m_lVisibleItems > m_lVisibleItemsMax Then
      VisibleItems = m_lVisibleItemsMax
   End If

   '-- Move The Splitter
   picSplitter.ZOrder 0
   If m_lVisibleItems > 0 Then
      picSplitter.Move 15, picSplitter.top, ScaleWidth - 30, picSplitter.Height
      DrawControl
   End If
   On Error GoTo 0

End Sub

Private Sub UserControl_Show()

   m_hWnd = UserControl.hwnd
   DrawControl

End Sub

Private Sub UserControl_Terminate()

   On Error Resume Next
   Erase m_tItem
   DetachMessage Me, hwnd, WM_SYSCOLORCHANGE
   DetachMessage Me, Parent.hwnd, WM_ENTERSIZEMOVE
   Set MTimer = Nothing
   Set m_Menus = Nothing

   On Error GoTo 0

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      'Write property values to storage
      Call .WriteProperty("DefaultItemHeight", m_lDefaultItemHeight, m_def_DefaultItemHeight)
      Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
      Call .WriteProperty("Enabled", UserControl.Enabled, True)
      Call .WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
      Call .WriteProperty("VisibleItems", m_lVisibleItems, m_def_VisibleItems)
   End With   'PropBag
   With PropBag
      Call .WriteProperty("ScaleMode", UserControl.ScaleMode, 1)
      Call .WriteProperty("Redraw", m_bRedraw, m_def_Redraw)
      Call .WriteProperty("DrawTopCaptionIcon", m_bDrawTopCaptionIcon, m_def_DrawTopCaptionIcon)
      Call .WriteProperty("SelectedItemFontBold", m_SelectedItemFontBold, m_def_SelectedItemFontBold)
   End With   'PropBag

   Call PropBag.WriteProperty("DisplayAddRemoveItemMenu", m_DisplayAddRemoveItemMenu, m_def_DisplayAddRemoveItemMenu)
   Call PropBag.WriteProperty("DisplayIconsInMenu", m_DisplayIconsInMenu, m_def_DisplayIconsInMenu)
   Call PropBag.WriteProperty("HideInfrequentlyUsedMenuItems", m_HideInfrequentlyUsedMenuItems, m_def_HideInfrequentlyUsedMenuItems)
   Call PropBag.WriteProperty("DrawToolbarItemsRightToLeft", m_DrawToolbarItemsRightToLeft, m_def_DrawToolbarItemsRightToLeft)
   Call PropBag.WriteProperty("RightToLeft", m_RightToLeft, m_def_RightToLeft)
   Call PropBag.WriteProperty("UseHandCursor", m_UseHandCursor, m_def_UseHandCursor)
   Call PropBag.WriteProperty("DisplayBannersInMenu", m_DisplayBannersInMenu, m_def_DisplayBannersInMenu)
   Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000005)
   Call PropBag.WriteProperty("UseCustomColors", m_UseCustomColors, m_def_UseCustomColors)
   Call PropBag.WriteProperty("ButtonPressedOffset", m_ButtonPressedOffset, m_def_ButtonPressedOffset)
   Call PropBag.WriteProperty("HeaderTextColor", m_HeaderTextColor, m_def_HeaderTextColor)
   Call PropBag.WriteProperty("DisplayHeader", m_DisplayHeader, m_def_DisplayHeader)
   Call PropBag.WriteProperty("DisplayMenuChevron", m_DisplayMenuChevron, m_def_DisplayMenuChevron)
   Call PropBag.WriteProperty("SelectedItemForeColor", m_SelectedItemForeColor, m_def_SelectedItemForeColor)
   '    Call PropBag.WriteProperty("CaptionFontBold", m_CaptionFontBold, m_def_CaptionFontBold)
   Call PropBag.WriteProperty("CaptionFont", m_CaptionFont, Ambient.Font)
End Sub

Public Property Get VisibleItems() As Long

   VisibleItems = m_lVisibleItems

End Property

Public Property Let VisibleItems(ByVal New_VisibleItems As Long)

'-- Set The Max Visible Items
   If Not bBeginSize Then mVisItemsMove = New_VisibleItems
   Dim i     As Long
   Dim iNum  As Long
   Static bFirst As Boolean

   '-- Stop The Flicker When We First Start Splitting
   If m_bSplitDown Then
      If Not bFirst Then bFirst = True: Exit Property
      bFirst = False
   End If

   If m_iItemCount > 0 Then
      For i = UBound(m_tItem()) To LBound(m_tItem()) Step -1
         If m_tItem(i).bVisible Then
            iNum = iNum + 1
         End If
      Next i
   End If
   If Me.DisplayHeader Then
      m_lVisibleItemsMax = Round(((ScaleHeight \ Screen.TwipsPerPixelY) / IIf(m_lIconHeight > 16, (m_lIconHeight * 1.5), (picExtraItems.Height \ Screen.TwipsPerPixelY))) - 3)
   Else
      m_lVisibleItemsMax = Round(((ScaleHeight \ Screen.TwipsPerPixelY) / (picExtraItems.Height \ Screen.TwipsPerPixelY)) - 4)
   End If
   m_lVisibleItems = New_VisibleItems
   If m_lVisibleItems > iNum Then
      m_lVisibleItems = iNum
   End If
   PropertyChanged "VisibleItems"
   If m_lVisibleItemsMax > iNum Then
      m_lVisibleItemsMax = iNum
   End If
   If m_lVisibleItems > m_lVisibleItemsMax Then
      m_lVisibleItems = m_lVisibleItemsMax
   End If
   If m_lVisibleItems <= 0 Then
      If m_iItemCount >= 1 Then
         m_lVisibleItems = 0
      End If
   End If
   If m_lVisibleItemsMax <= 0 Then
      m_lVisibleItems = 0
   End If
   If m_lVisibleItems = Me.EyeDropperItems.Count Then
      m_lVisibleItems = iNum
   End If
   If Redraw Then
      DrawControl
   End If

End Property


Sub GetGradientColors()

   m_lColorOneSelected = 1
   m_lColorTwoSelected = 1
   m_lColorHeaderColorOne = 1
   m_lColorHeaderColorTwo = 1
   m_lColorHeaderForeColor = 1
   m_lColorHotOne = 1
   m_lColorHotTwo = 1

   If AppThemed Then

      Select Case m_sCurrentSystemThemename
      Case "HomeStead"
         m_lColorOneNormal = RGB(228, 235, 200)
         m_lColorTwoNormal = RGB(175, 194, 142)
         m_lColorBorder = RGB(100, 144, 88)
         m_lColorHeaderColorOne = RGB(165, 182, 121)
         m_lColorHeaderColorTwo = dBlendColor(RGB(99, 122, 68), vbBlack, 200)
      Case "NormalColor"
         m_lColorOneNormal = RGB(197, 221, 250)
         m_lColorTwoNormal = RGB(128, 167, 225)
         m_lColorBorder = RGB(0, 45, 150)
         m_lColorHeaderColorOne = RGB(81, 128, 208)
         m_lColorHeaderColorTwo = dBlendColor(RGB(11, 63, 153), vbBlack, 230)
      Case "Metallic"
         m_lColorOneNormal = RGB(219, 220, 232)
         m_lColorTwoNormal = RGB(149, 147, 177)
         m_lColorBorder = RGB(119, 118, 151)
         m_lColorHeaderColorOne = RGB(163, 162, 187)
         m_lColorHeaderColorTwo = dBlendColor(RGB(112, 111, 145), vbBlack, 200)
      Case Else

         m_lColorOneNormal = dBlendColor(vbButtonFace, vbWhite, 120)
         m_lColorTwoNormal = vbButtonFace
         m_lColorBorder = dBlendColor(vbButtonFace, vbBlack, 200)
         m_lColorHeaderColorOne = vbButtonFace
         m_lColorHeaderColorTwo = dBlendColor(vbInactiveTitleBar, vbBlack, 200)
         m_lColorBorder = TranslateColor(vbInactiveTitleBar)

      End Select
      m_lColorOneSelectedNormal = RGB(248, 216, 126)
      m_lColorTwoSelectedNormal = RGB(240, 160, 38)

      m_lColorHotOne = dBlendColor(vbWindowBackground, vbButtonFace, 220)
      m_lColorHotTwo = RGB(248, 216, 126)

      m_lColorOneSelected = RGB(240, 160, 38)
      m_lColorTwoSelected = RGB(248, 216, 126)

   Else
      m_lColorOneNormal = dBlendColor(vbButtonFace, vbWhite, 120)
      m_lColorTwoNormal = vbButtonFace
      m_lColorBorder = dBlendColor(vbButtonFace, vbBlack, 200)
      m_lColorHeaderColorOne = vbButtonFace
      m_lColorHeaderColorTwo = dBlendColor(vbInactiveTitleBar, dBlendColor(vbBlack, vbButtonFace, 10), 200)
      m_lColorBorder = TranslateColor(vbInactiveTitleBar)
      m_lColorHotTwo = dBlendColor(vbInactiveTitleBar, dBlendColor(vbButtonFace, vbWhite, 50), 10)
      m_lColorHotOne = m_lColorHotTwo
      m_lColorOneSelected = dBlendColor(vbInactiveTitleBar, dBlendColor(vbButtonFace, vbWhite, 150), 100)
      m_lColorTwoSelected = m_lColorOneSelected
      m_lColorOneSelectedNormal = dBlendColor(vbInactiveTitleBar, dBlendColor(vbButtonFace, vbWhite, 150), 130)
      m_lColorTwoSelectedNormal = m_lColorOneSelectedNormal
   End If


End Sub

Public Property Get DisplayHeader() As Boolean
   DisplayHeader = m_DisplayHeader
End Property

Public Property Let DisplayHeader(ByVal New_DisplayHeader As Boolean)
   m_DisplayHeader = New_DisplayHeader
   PropertyChanged "DisplayHeader"
   DrawControl
End Property

Public Property Get DisplayMenuChevron() As Boolean
   DisplayMenuChevron = m_DisplayMenuChevron
End Property

Public Property Let DisplayMenuChevron(ByVal New_DisplayMenuChevron As Boolean)
   m_DisplayMenuChevron = New_DisplayMenuChevron
   PropertyChanged "DisplayMenuChevron"
   DrawControl
End Property

Public Sub ShowAboutBox()
Attribute ShowAboutBox.VB_UserMemId = -552
   frmAbout.Show vbModal
End Sub

Public Property Get SelectedItemForeColor() As OLE_COLOR
   SelectedItemForeColor = m_SelectedItemForeColor
End Property

Public Property Let SelectedItemForeColor(ByVal New_SelectedItemForeColor As OLE_COLOR)
   m_SelectedItemForeColor = New_SelectedItemForeColor
   PropertyChanged "SelectedItemForeColor"
   DrawControl
End Property
Public Property Get CaptionFont() As Font
   Set CaptionFont = m_CaptionFont
End Property

Public Property Set CaptionFont(ByVal New_CaptionFont As Font)

   Set m_CaptionFont = New_CaptionFont
   PropertyChanged "CaptionFont"
   DrawControl
   
    
   
End Property

