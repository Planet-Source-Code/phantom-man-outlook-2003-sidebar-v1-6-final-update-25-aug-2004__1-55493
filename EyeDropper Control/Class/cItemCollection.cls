VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAPP_EDITemCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ======================================================================================
' Name:     IAPP_EDITemCollection.cls
' Author:   Gary Noble
' Date:     19/04/2004
'
' Requires: -
'
' Copyright © 2004 Gary Noble
' --------------------------------------------------------------------------------------
'
' Strongly-typed collection of tab objects for EyeDropper control
'
' ======================================================================================
Private m_lOwner     As Long
Private m_hWnd       As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
        pSrc As Any, _
        ByVal ByteLen As Long)

Public Function Add(Optional Key As Variant, _
                    Optional KeyBefore As Variant, _
                    Optional strCaption As String, _
                    Optional IconIndex As Long) As cEDItem


    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        Set Add = oT.fAdd(Key, KeyBefore, strCaption, IconIndex)
    End If

End Function

Public Sub Clear()

    Dim i  As Long
    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        For i = oT.fItemCount To 1 Step -1
            oT.fRemove i
        Next i
    End If

End Sub

Public Property Get Count() As Long

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        Count = oT.fItemCount
    End If

End Property

Private Property Get EyeDropperObject() As EyeDropper

    Dim oT As Object

    If (IsWindow(m_hWnd)) Then
        If Not (m_lOwner = 0) Then
            CopyMemory oT, m_lOwner, 4
            Set EyeDropperObject = oT
            CopyMemory oT, 0&, 4
        End If
    End If

End Property

Friend Sub Init(ByVal lPtr As Long, _
                ByVal lhWnd As Long)

    m_lOwner = lPtr
    m_hWnd = lhWnd

End Sub

Public Property Get Item(Key As Variant) As cEDItem

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        Set Item = oT.fItem(Key)
    End If

End Property

Public Sub Remove(Key As Variant)

    Dim oT As EyeDropper

    Set oT = EyeDropperObject()
    If Not oT Is Nothing Then
        oT.fRemove Key
    End If

End Sub

