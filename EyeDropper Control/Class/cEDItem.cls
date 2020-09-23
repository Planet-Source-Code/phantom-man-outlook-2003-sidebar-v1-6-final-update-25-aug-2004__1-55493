VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEDItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ======================================================================================
' Name:     cEDItem.cls
' Author:   Gary Noble
' Date:     19/04/2004
'
' Requires: -
'
' Copyright © 2004 Gary Noble
' --------------------------------------------------------------------------------------
'
' Public access to EyeDropper properties for the EyeDropper control.
'
' ======================================================================================
Private m_lOwner              As Long
Private m_hWnd                As Long
Private m_lId                 As Long
Private m_bLastStateSown      As Boolean
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
        pSrc As Any, _
        ByVal ByteLen As Long)

Public Property Get Caption() As String

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        Caption = oT.fItemCaption(m_lId)
    End If

End Property

Public Property Let Caption(ByVal sCaption As String)

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        oT.fItemCaption(m_lId) = sCaption
    End If

End Property

Private Property Get EyeDropperObject() As EyeDropper

    Dim oT As Object

    If Not (m_lOwner = 0) Then
        CopyMemory oT, m_lOwner, 4
        Set EyeDropperObject = oT
        CopyMemory oT, 0&, 4
        Exit Property
    End If
    Err.Raise 9, App.EXEName & ".EyeDropper"

End Property

Friend Sub fInit(ByVal lPtr As Long, _
                 ByVal lhWnd As Long, _
                 ByVal lID As Long)

    m_lOwner = lPtr
    m_hWnd = lhWnd
    m_lId = lID

End Sub

Public Property Get IconIndex() As Long

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        IconIndex = oT.fItemIconIndex(m_lId)
    End If

End Property

Public Property Let IconIndex(ByVal lIconIndex As Long)

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        oT.fItemIconIndex(m_lId) = lIconIndex
    End If

End Property

Public Property Get Index() As Long

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        Index = oT.fItemIndex(m_lId)
    End If

End Property

Public Property Let Index(ByVal lIndex As Long)

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        oT.fItemIndex(m_lId) = lIndex
    End If

End Property

Public Property Get ItemData() As Long

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        ItemData = oT.fItemItemData(m_lId)
    End If

End Property

Public Property Let ItemData(ByVal lItemData As Long)

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        oT.fItemItemData(m_lId) = lItemData
    End If

End Property

Public Property Get Key() As String

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        Key = oT.fItemKey(m_lId)
    End If

End Property

Friend Property Let LastStateShown(ByVal bState As Boolean)


    m_bLastStateSown = bState

End Property

Public Property Get Panel() As Object

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        Set Panel = oT.fItemPanel(m_lId)
    End If

End Property

Public Property Let Panel(ByRef ctlThis As Object)

    pSetPanel ctlThis

End Property

Public Property Set Panel(ByRef ctlThis As Object)

    pSetPanel ctlThis

End Property

Private Sub pSetPanel(ByRef ctlThis As Object)

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        oT.fItemPanel(m_lId) = ctlThis
    End If

End Sub

Public Property Get Selected() As Boolean

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        Selected = oT.fEyeDropperItemselected(m_lId)
    End If

End Property

Public Property Let Selected(ByVal bState As Boolean)

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        oT.fEyeDropperItemselected(m_lId) = bState
    End If

End Property

Public Property Get Tag() As String

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        Tag = oT.fItemTag(m_lId)
    End If

End Property

Public Property Let Tag(ByVal sTag As String)

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        oT.fItemTag(m_lId) = sTag
    End If

End Property



Public Property Get Enabled() As Boolean

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        Enabled = oT.fItemEnabled(m_lId)
    End If

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        oT.fItemEnabled(m_lId) = New_Enabled
    End If

End Property




Public Property Get ToolTipText() As String

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        ToolTipText = oT.fItemToolTipText(m_lId)
    End If

End Property

Public Property Let ToolTipText(ByVal sToolTipText As String)

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        oT.fItemToolTipText(m_lId) = sToolTipText
    End If

End Property

Public Property Get Visible() As Boolean

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        Visible = oT.fVisible(m_lId)
    End If

End Property

Public Property Let Visible(ByVal bState As Boolean)

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        oT.fVisible(m_lId) = bState
        If m_bLastStateSown Then
            oT.VisibleItems = oT.VisibleItems + 1
        Else
            oT.VisibleItems = oT.VisibleItems
        End If
    End If

End Property
