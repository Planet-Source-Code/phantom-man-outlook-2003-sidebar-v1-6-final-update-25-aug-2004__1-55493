VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About MyApp"
   ClientHeight    =   3000
   ClientLeft      =   4920
   ClientTop       =   3330
   ClientWidth     =   5505
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   2070.653
   ScaleMode       =   0  'User
   ScaleWidth      =   5169.479
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      X1              =   676.117
      X2              =   5070.878
      Y1              =   331.305
      Y2              =   331.305
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   480
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   4725
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   2265
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   4695
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
    lblTitle.Caption = App.Title & " Control " & App.Major & "." & App.Minor
    lblDisclaimer.Caption = "A Freeware/OpenSource Outlook 2003 Style Sidebar Control." & _
             String(2, vbCrLf) & "All I Ask Is Three Things:" & vbCrLf & "1) If You Enhance This Control Send It To gwnoble@msn.com" & _
             String(1, vbCrLf) & "2) Keep The Copyright Notice." & _
             String(1, vbCrLf) & "3) Keep It Free. Don't Palm This Of As Your Own Work." & _
             String(2, vbCrLf) & "This Control Uses A Modified Version Of The PopMenu By VBAccelerator (www.VBAccelerator.com)." & _
             String(2, vbCrLf) & "Copyright © 2004 Gary Noble" & vbCrLf

            
End Sub

Private Sub lblDisclaimer_Click()
Unload Me
End Sub

Private Sub lblTitle_Click()
Unload Me
End Sub
