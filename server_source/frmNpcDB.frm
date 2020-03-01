VERSION 5.00
Begin VB.Form frmNpcDB 
   Caption         =   "Npc Database"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3975
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNpcDB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   397
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   3735
   End
   Begin VB.ListBox lstNPC 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5340
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmNpcDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************
'  Streetwars Online 2 Version 1.00
'  Copyright 2000 - B.Smith aka (Wuzzbent)
'  All Rights Reserved
'  wuzzbent@swbell.net
'
'  By using this source code, you agree to the following
'  terms and conditions.
'
'  You may use this source code for your own personal
'  pleasure and use.  You may freely distribute it along with
'  any modification(s) made to it.  You may NOT remove, modify,
'  or adjust this copyright information.  You may NOT attempt
'  to charge for the use of this software under any conditions.
'
'  Support Free Software....
'
'******************************************************

Option Explicit

Private Sub cmdExit_Click()

Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub Form_Load()
'Disable X (Close Button) on main form
Dim hMenu As Long
Dim menuItemCount As Long
hMenu = GetSystemMenu(Me.hwnd, 0)
If hMenu Then
menuItemCount = GetMenuItemCount(hMenu)
Call RemoveMenu(hMenu, menuItemCount - 1, _
MF_REMOVE Or MF_BYPOSITION)
Call RemoveMenu(hMenu, menuItemCount - 2, _
MF_REMOVE Or MF_BYPOSITION)
Call DrawMenuBar(Me.hwnd)
End If

End Sub


Private Sub Form_Resize()

If WindowState <> vbMinimized Then
   lstNPC.Move 0, 0, ScaleWidth, (ScaleHeight - cmdExit.Height)
   cmdExit.Move 0, (ScaleHeight - cmdExit.Height), ScaleWidth, cmdExit.Height
End If

End Sub

