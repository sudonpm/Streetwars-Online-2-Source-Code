VERSION 5.00
Begin VB.Form frmTravel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Airport Checkout"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3660
   Icon            =   "frmTravel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   278
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   244
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdForgetIt 
      Caption         =   "Forget It"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   3450
   End
   Begin VB.CommandButton cmdNewJersey 
      Caption         =   "New Jersey"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   3120
      Width           =   1650
   End
   Begin VB.CommandButton cmdChicago 
      Caption         =   "Chicago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   1650
   End
   Begin VB.CommandButton cmdMiami 
      Caption         =   "Miami"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   2280
      Width           =   1650
   End
   Begin VB.CommandButton cmdHouston 
      Caption         =   "Houston"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1650
   End
   Begin VB.CommandButton cmdLosAngeles 
      Caption         =   "Los Angeles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   1650
   End
   Begin VB.CommandButton cmdNewYork 
      Caption         =   "New York"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   1650
   End
   Begin VB.Label lblNewJersey 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   2760
      Width           =   1650
   End
   Begin VB.Label lblChicago 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   1650
   End
   Begin VB.Label lblMiami 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   1920
      Width           =   1650
   End
   Begin VB.Label lblHouston 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1650
   End
   Begin VB.Label lblLosAngeles 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   1080
      Width           =   1650
   End
   Begin VB.Label lblNewYork 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1650
   End
   Begin VB.Line lneMain 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   4
      X1              =   10
      X2              =   235
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   810
      Left            =   120
      Picture         =   "frmTravel.frx":0442
      Top             =   120
      Width           =   3435
   End
End
Attribute VB_Name = "frmTravel"
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

Private Sub cmdChicago_Click()

'Fly to Chicago
frmMain.wsk.SendData Chr$(255) & Chr$(7) & "chicago" & Chr$(0)
DoEvents
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdForgetIt_Click()

Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdHouston_Click()

'Fly to Houston
frmMain.wsk.SendData Chr$(255) & Chr$(7) & "houston" & Chr$(0)
DoEvents
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdLosAngeles_Click()

'Fly to Los Angeles
frmMain.wsk.SendData Chr$(255) & Chr$(7) & "los angeles" & Chr$(0)
DoEvents
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdMiami_Click()

'Fly to Miami
frmMain.wsk.SendData Chr$(255) & Chr$(7) & "miami" & Chr$(0)
DoEvents
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdNewJersey_Click()

'Fly to New Jersey
frmMain.wsk.SendData Chr$(255) & Chr$(7) & "new jersey" & Chr$(0)
DoEvents
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdNewYork_Click()

'Fly to New York
frmMain.wsk.SendData Chr$(255) & Chr$(7) & "new york" & Chr$(0)
DoEvents
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub
Private Sub Form_Load()

Dim hMenu As Long
Dim menuItemCount As Long
'Obtain the handle to the form's system menu
hMenu = GetSystemMenu(Me.hWnd, 0)
If hMenu Then
'Obtain the number of items in the menu
menuItemCount = GetMenuItemCount(hMenu)
'Remove the system menu Close menu item.
'The menu item is 0-based, so the last
'item on the menu is menuItemCount - 1
Call RemoveMenu(hMenu, menuItemCount - 1, _
MF_REMOVE Or MF_BYPOSITION)
'Remove the system menu separator line
Call RemoveMenu(hMenu, menuItemCount - 2, _
MF_REMOVE Or MF_BYPOSITION)
'Force a redraw of the menu. This
'refreshes the titlebar, dimming the X
Call DrawMenuBar(Me.hWnd)
End If

End Sub


