VERSION 5.00
Begin VB.Form frmPawnShop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Pawn Shop"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   Icon            =   "frmPawnShop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   295
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   517
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
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
      Left            =   2400
      TabIndex        =   7
      Top             =   3960
      Width           =   3015
   End
   Begin VB.CommandButton cmdSell 
      Caption         =   "<<< Sell <<<"
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
      Left            =   2400
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   ">>> Buy >>>"
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
      Left            =   3960
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.ListBox lstInv 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2790
      Left            =   5520
      TabIndex        =   1
      Top             =   1500
      Width           =   2055
   End
   Begin VB.ListBox lstShop 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2790
      Left            =   240
      TabIndex        =   0
      Top             =   1500
      Width           =   2055
   End
   Begin VB.Label lblItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Item"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3600
      TabIndex        =   17
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblItemDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   16
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblRank 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblRankDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rank:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblCash 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblCashDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cash:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   160
      X2              =   360
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Label lblCanBuy 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Item"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblCanBuyDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Can Buy:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Item"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblPriceDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Price:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Welcome to the Pawn Shop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label lblInventory 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Your Inventory"
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
      Left            =   5520
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblShop 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Shop Inventory"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Left            =   120
      Picture         =   "frmPawnShop.frx":0442
      Top             =   120
      Width           =   7560
   End
End
Attribute VB_Name = "frmPawnShop"
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

Private Sub cmdBuy_Click()

cmdBuy.Enabled = False
cmdSell.Enabled = False
cmdExit.SetFocus
frmMain.wsk.SendData Chr$(254) & Chr$(5) & lstShop.ListIndex & Chr$(0)
DoEvents

End Sub
Private Sub cmdExit_Click()

Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub


Private Sub cmdSell_Click()

cmdBuy.Enabled = False
cmdSell.Enabled = False
cmdExit.SetFocus
frmMain.wsk.SendData Chr$(254) & Chr$(6) & lstInv.ListIndex & Chr$(0)
DoEvents

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


Private Sub lstInv_Click()

cmdBuy.Enabled = False
cmdSell.Enabled = False

End Sub

Private Sub lstInv_DblClick()
   
If lstInv.Text <> "<Empty>" Then
   cmdSell.Enabled = True
   cmdBuy.Enabled = False
   frmMain.wsk.SendData Chr$(254) & Chr$(3) & lstInv.ListIndex & Chr$(0)
   DoEvents
ElseIf lstInv.Text = "<Empty>" Then
   cmdSell.Enabled = False
   cmdBuy.Enabled = False
End If

End Sub

Private Sub lstShop_Click()

cmdBuy.Enabled = False
cmdSell.Enabled = False

End Sub

Private Sub lstShop_DblClick()
   
   cmdBuy.Enabled = True
   cmdSell.Enabled = False
   frmMain.wsk.SendData Chr$(254) & Chr$(4) & lstShop.ListIndex & Chr$(0)
   DoEvents

End Sub


