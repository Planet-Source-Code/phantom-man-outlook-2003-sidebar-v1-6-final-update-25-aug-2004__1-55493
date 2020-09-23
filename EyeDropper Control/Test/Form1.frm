VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\EyeDropper.vbp"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Sidebar Demo Application"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command19 
      Caption         =   "Header Font"
      Height          =   375
      Left            =   4920
      TabIndex        =   37
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Item Font"
      Height          =   375
      Left            =   3360
      TabIndex        =   38
      Top             =   5400
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6960
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Show About"
      Height          =   375
      Left            =   7800
      TabIndex        =   36
      Top             =   120
      Width           =   2175
   End
   Begin EyeDropperTab.EyeDropper EyeDropper1 
      Align           =   3  'Align Left
      Height          =   7470
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   13176
      DefaultItemHeight=   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      VisibleItems    =   0
      BackColor       =   -2147483633
      SelectedItemForeColor=   192
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1gg 
      Caption         =   "Header Yes/No"
      Height          =   375
      Left            =   3360
      TabIndex        =   33
      Top             =   5040
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties And Events"
      Height          =   6615
      Left            =   3120
      TabIndex        =   0
      Top             =   720
      Width           =   7095
      Begin VB.CommandButton Command17 
         Caption         =   "Display Menu Chevron"
         Height          =   375
         Left            =   4200
         TabIndex        =   34
         Top             =   4680
         Width           =   2775
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Custom Colours"
         Height          =   375
         Left            =   4200
         TabIndex        =   31
         Top             =   4320
         Width           =   2775
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Hide Banners In Menu"
         Height          =   375
         Left            =   4200
         TabIndex        =   30
         Top             =   3960
         Width           =   2775
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Draw Items Right To Left"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   3960
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Reload Items"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   2400
         Width           =   3135
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Don't Draw Icon In Caption"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   3600
         Width           =   3135
      End
      Begin VB.CommandButton cmdCommand4 
         Caption         =   "View Less Items"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   2040
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "View More Items"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   3135
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Hide Selected Item"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   3240
         Width           =   3135
      End
      Begin VB.CommandButton cmdCommand3 
         Caption         =   "Disable"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   2880
         Width           =   3135
      End
      Begin VB.CommandButton cmdCommand2 
         Caption         =   "Remove Selected Item"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   3135
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Use Normal Cursor"
         Height          =   375
         Left            =   4200
         TabIndex        =   2
         Top             =   3600
         Width           =   2775
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Draw Toolbar Items Left To Right"
         Height          =   375
         Left            =   4200
         TabIndex        =   4
         Top             =   3240
         Width           =   2775
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Show Infrequently Used Menu Items"
         Height          =   375
         Left            =   4200
         TabIndex        =   5
         Top             =   2880
         Width           =   2775
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Hide Add/Remove Item Menu"
         Height          =   375
         Left            =   4200
         TabIndex        =   7
         Top             =   2520
         Width           =   2775
      End
      Begin VB.CommandButton Command14 
         Caption         =   "None"
         Height          =   375
         Left            =   4200
         TabIndex        =   1
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Large Images"
         Height          =   375
         Left            =   4200
         TabIndex        =   9
         Top             =   960
         Width           =   2775
      End
      Begin VB.PictureBox picTure4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   4080
         ScaleHeight     =   255
         ScaleWidth      =   1575
         TabIndex        =   23
         Top             =   6720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox picTure2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   4080
         ScaleHeight     =   255
         ScaleWidth      =   1575
         TabIndex        =   22
         Top             =   6480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox picTure3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   4080
         ScaleHeight     =   255
         ScaleWidth      =   1575
         TabIndex        =   21
         Top             =   6840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox picTure1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   4080
         ScaleHeight     =   255
         ScaleWidth      =   1575
         TabIndex        =   20
         Top             =   7200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdCommand1 
         Caption         =   "Add Item"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   3135
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Small Images"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4200
         TabIndex        =   10
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Display icons In Menu"
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   8
         Top             =   2160
         Width           =   2775
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1095
         Left            =   240
         TabIndex        =   6
         Top             =   5400
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   1931
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   16711680
         BackColor       =   16777215
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Event"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   4800
         Top             =   5760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   15
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":031A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":076C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":6A06
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":6D20
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":7172
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":75C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":CDB6
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":13050
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1336A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":13684
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1991E
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1A3B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1A80A
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1AC5C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4200
         Top             =   5760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1AF76
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1B3C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1B81A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1BC6C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1C0BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1C3D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1C6F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1C84C
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1C9A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1DEA8
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1EBB2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   495
         Left            =   5520
         TabIndex        =   24
         Top             =   5760
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Appearance      =   0
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Events"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   5160
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Items"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Imagelist"
         Height          =   255
         Left            =   4200
         TabIndex        =   26
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Misc"
         Height          =   255
         Left            =   4200
         TabIndex        =   25
         Top             =   1920
         Width           =   2055
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sidebar Control V1.6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   32
      Top             =   120
      Width           =   2595
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Demo Application"
      Height          =   375
      Left            =   5160
      TabIndex        =   29
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cDLG As New ICDLG_FontDialogHandlerHandler
Private Sub cmdCommand1_Click()
Randomize

  Dim oItem As cEDItem

    With Me.EyeDropper1
        .Redraw = False
        
        Set oItem = .eyedropperitems.Add("New Item: " & Me.EyeDropper1.eyedropperitems.Count + 1, , "New Item: " & Me.EyeDropper1.eyedropperitems.Count + 1)
        oItem.IconIndex = Int((8 * Rnd) + 1)
        oItem.Selected = True
        '.visibleitems = .visibleitems + 1
        .Redraw = True
    End With
    
End Sub

Private Sub cmdCommand2_Click()

    On Error Resume Next
    Me.EyeDropper1.eyedropperitems.Remove (Me.EyeDropper1.SelectedItem.Index)
    On Error GoTo 0

End Sub

Private Sub cmdCommand3_Click()

    Me.EyeDropper1.Enabled = Not Me.EyeDropper1.Enabled
    If Me.EyeDropper1.Enabled Then
        cmdCommand3.Caption = "Disable"
    Else
        cmdCommand3.Caption = "Enable"
    End If

End Sub

Private Sub cmdCommand4_Click()

    Me.EyeDropper1.visibleitems = Me.EyeDropper1.visibleitems - 1

End Sub






Private Sub Command1_Click(Index As Integer)
    EyeDropper1.visibleitems = EyeDropper1.visibleitems + 1
End Sub

Private Sub Command10_Click()
Me.EyeDropper1.HideInfrequentlyUsedMenuItems = Not Me.EyeDropper1.HideInfrequentlyUsedMenuItems
If Me.EyeDropper1.HideInfrequentlyUsedMenuItems Then
    Command10.Caption = "Show Infrequently Used Menu Items"
Else
    Command10.Caption = "Hide Infrequently Used Menu Items"
End If
End Sub

Private Sub Command11_Click()
    Me.EyeDropper1.DrawToolbarItemsRightToLeft = Not Me.EyeDropper1.DrawToolbarItemsRightToLeft
    If Me.EyeDropper1.DrawToolbarItemsRightToLeft Then
        Command11.Caption = "Draw Toolbar Items Left To Right"
    Else
        Command11.Caption = "Draw Toolbar Items Right To Left"
    End If
End Sub

Private Sub Command12_Click()
Me.EyeDropper1.RightToLeft = Not Me.EyeDropper1.RightToLeft
Me.EyeDropper1.Redraw = True
If Not Me.EyeDropper1.RightToLeft Then
    Command12.Caption = "Draw Items Right To Left"
Else
    Command12.Caption = "Draw Items Left To Right"
End If
End Sub

Private Sub Command13_Click()
    
    Me.EyeDropper1.usehandcursor = Not Me.EyeDropper1.usehandcursor
    If Me.EyeDropper1.usehandcursor Then
        Command13.Caption = "Use Normal Cursor"
    Else
        Command13.Caption = "Use Hand Cursor"
    End If
End Sub

Private Sub Command14_Click()
 EyeDropper1.ImageList = vbNullString
 Command6.Enabled = True
 Command7.Enabled = True
End Sub

Private Sub Command15_Click()
    Me.EyeDropper1.DisplayBannersInMenu = Not Me.EyeDropper1.DisplayBannersInMenu
    
    If Me.EyeDropper1.DisplayBannersInMenu Then
        Command15.Caption = "Hide Banners In Menu"
    Else
        Command15.Caption = "Display Banners In Menu"
    End If
End Sub

Private Sub Command16_Click()
    Me.EyeDropper1.UseCustomColors = Not Me.EyeDropper1.UseCustomColors
End Sub

Private Sub Command17_Click()
Me.EyeDropper1.displaymenuchevron = Not Me.EyeDropper1.displaymenuchevron
End Sub

Private Sub Command18_Click()
    Me.EyeDropper1.ShowAboutBox
End Sub

Private Sub Command19_Click()
    With cDLG
        .Init EyeDropper1.CaptionFont, "Selct Header Font", hwnd
        .Show
        Set EyeDropper1.CaptionFont = .Font
    End With
    
End Sub

Private Sub Command1gg_Click()
    With Me
        .EyeDropper1.displayheader = Not .EyeDropper1.displayheader
        Me.EyeDropper1.eyedropperitems.Item(Me.EyeDropper1.SelectedItem.Index).Index = 1
         Me.EyeDropper1.Redraw = True
         
        
    End With
End Sub

Private Sub Command2_Click()

    Call LoadItems

End Sub

Private Sub Command20_Click()
    With cDLG
        .Init EyeDropper1.Font, "Selct Header Font", hwnd, eFontFlag_InitToLogFontStruct
        .Show
        Set EyeDropper1.Font = .Font
        EyeDropper1.Redraw = True
    End With
    
End Sub

Private Sub Command3_Click()


    With Me
        .EyeDropper1.Redraw = False
        .EyeDropper1.eyedropperitems.Clear
        .EyeDropper1.Redraw = True
    End With 'Me

End Sub

Private Sub Command4_Click()

    Me.EyeDropper1.DrawTopCaptionIcon = Not Me.EyeDropper1.DrawTopCaptionIcon
If Me.EyeDropper1.DrawTopCaptionIcon Then
    Command4.Caption = "Don't Draw Icon In Caption"
Else
    Command4.Caption = "Draw Icon In Caption"
End If
End Sub

Private Sub Command5_Click()

    On Error Resume Next
    Me.EyeDropper1.eyedropperitems.Item(Me.EyeDropper1.SelectedItem.Index).Visible = Not Me.EyeDropper1.eyedropperitems.Item(Me.EyeDropper1.SelectedItem.Index).Visible
    On Error GoTo 0

End Sub

Private Sub Command6_Click()
    Me.EyeDropper1.ImageList = Me.ImageList1
    Command6.Enabled = False
    Command7.Enabled = True
End Sub

Private Sub Command7_Click()
Me.EyeDropper1.ImageList = Me.ImageList2
Command7.Enabled = False
    Command6.Enabled = True
End Sub


Private Sub Command8_Click(Index As Integer)
Me.EyeDropper1.displayiconsinmenu = Not Me.EyeDropper1.displayiconsinmenu
If Me.EyeDropper1.displayiconsinmenu Then
    Command8(0).Caption = "Hide icons In Menu"
Else
    Command8(0).Caption = "Display icons In Menu"
End If

End Sub

Private Sub Command9_Click()
Me.EyeDropper1.DisplayAddRemoveItemMenu = Not Me.EyeDropper1.DisplayAddRemoveItemMenu

If Me.EyeDropper1.DisplayAddRemoveItemMenu Then
    Command9.Caption = "Hide Add/Remove Item Menu"
    Else
    Command9.Caption = "Display Add/Remove Item Menu"
End If

End Sub

Private Sub EyeDropper1_BeginSizing()
    AddEventMessage "Begin Sizing", ""
End Sub

Private Sub EyeDropper1_EndSizing()
    AddEventMessage "End Sizing", ""
End Sub

Private Sub EyeDropper1_HoverItem(edItem As EyeDropperTab.cEDItem)
        AddEventMessage "Hovering Item: ", edItem.Caption
        Caption = edItem.Caption
End Sub

Private Sub EyeDropper1_HoverItemLeave()
    AddEventMessage "Hovering Leave: ", ""
End Sub

Private Sub EyeDropper1_ItemAdded(edItem As EyeDropperTab.cEDItem)
        AddEventMessage "Item Added: ", edItem.Caption
End Sub

Private Sub EyeDropper1_ItemHidden(edItem As EyeDropperTab.cEDItem)
        AddEventMessage "Item Hidden: ", edItem.Caption
End Sub

Private Sub EyeDropper1_ItemHiddenFromMenu(edItem As EyeDropperTab.cEDItem)
    AddEventMessage "Item Made InVisible From Menu: ", edItem.Caption
End Sub

Private Sub EyeDropper1_ItemRemoved(edItem As EyeDropperTab.cEDItem)
        AddEventMessage "Item Removed: ", edItem.Caption
End Sub

Private Sub EyeDropper1_ItemRightClick(edItem As EyeDropperTab.cEDItem)
        AddEventMessage "Item right Click: ", edItem.Caption
End Sub

Private Sub EyeDropper1_ItemSelected(edItem As EyeDropperTab.cEDItem)
        AddEventMessage "Item Selected: ", edItem.Caption
End Sub

Private Sub EyeDropper1_ItemVisibleFromMenu(edItem As EyeDropperTab.cEDItem)
    AddEventMessage "Item Made Visible From Menu: ", edItem.Caption
End Sub

Private Sub EyeDropper1_MenuDestroyed()
    AddEventMessage "Menu Destroyed", ""
End Sub



Private Sub EyeDropper1_MenuItemHover(Caption As String, IsAddRemoveItems As Boolean)
    If IsAddRemoveItems Then
        AddEventMessage "Hover Menu Item: " & Caption, "From Add Remove Items Menu"
    Else
        AddEventMessage "Hover Menu Item: " & Caption, "Normal"
    End If
End Sub

Private Sub EyeDropper1_MenuShown()
    AddEventMessage "Menu Initiated", ""
End Sub

Private Sub fdg_Click()

End Sub


Private Sub Form_Load()

    '- Only do This once
    Me.EyeDropper1.Initialise
        
    '-- Set The Custom Colour Properties
    '-- I've Opted For A Purple Look
    '-- You Can Set The Colours To What ever You Want
    Me.EyeDropper1.SetCustomProperties &HFF80FF, &H800080, vbWhite, _
                &H800080, &HFF80FF, &HC000C0, vbWhite, _
                &HFFC0FF, &H800080, &H800080, &HFFC0FF
    
    '-- Load Sample Items
    Call LoadItems
    
End Sub

Private Sub LoadItems()
  Dim i     As Long
  Dim iItems As Long
  
  Dim oItem As cEDItem
  Dim oNode As Node
 
    With Me.EyeDropper1
        'Me.vbalImageList1
        .ImageList = Me.ImageList1
        .Redraw = False
        .eyedropperitems.Clear
        .Redraw = False
        '-- Add Some Sample nodes
        TreeView1.Nodes.Clear
        
        For i = 1 To 10
            Set oNode = TreeView1.Nodes.Add(, , "Main" & i, "TopLevel Node " & i)
            For iItems = 1 To 10
                Set oNode = TreeView1.Nodes.Add("Main" & i, tvwChild, "MainSub" & iItems & "Main" & i, "Sub Level Node " & iItems)
            Next
        Next
        
        '-- Add Some Sample Panels
        For i = 1 To 5
        
            Select Case i
                Case Is = 1
                Set oItem = .eyedropperitems.Add("Item: " & i, , "Mail ")
                Case Is = 2
                Set oItem = .eyedropperitems.Add("Item: " & i, , "Calender " & i)
                Case Is = 3
                Set oItem = .eyedropperitems.Add("Item: " & i, , "Contacts " & i)
                Case Is = 4
                Set oItem = .eyedropperitems.Add("Item: " & i, , "Tasks: " & i)
                Case Is = 5
                Set oItem = .eyedropperitems.Add("Item: " & i, , "Notes " & i)
            End Select
            If i = 1 Then
                oItem.Panel = Me.TreeView1
            End If
            If i = 2 Then
                oItem.Panel = Me.picTure2
            End If
            If i = 3 Then
                oItem.Panel = Me.picTure3
            End If
            
            oItem.IconIndex = i
            If i = 1 Then
                oItem.Selected = True
            End If
            If i = 4 Or i = 9 Then
                oItem.Visible = False
             Else
                oItem.Visible = True
            End If
        Next i
        .visibleitems = 5
        .Redraw = True
        
    End With

End Sub
Private Sub AddEventMessage(strMessage As String, ItemName As String)
Dim xItem As ListItem

    Set xItem = Me.ListView1.ListItems.Add(, , strMessage)
        xItem.SubItems(1) = ItemName
        xItem.Selected = True
        xItem.EnsureVisible
        
        
End Sub
