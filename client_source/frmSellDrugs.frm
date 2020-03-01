VERSION 5.00
Begin VB.Form frmSellDrugs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sell Drugs"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "frmSellDrugs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   258
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   383
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
      Left            =   120
      TabIndex        =   2
      Top             =   3420
      Width           =   1455
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
      Left            =   1800
      TabIndex        =   1
      Top             =   3420
      Width           =   1455
   End
   Begin VB.ListBox lstInventory 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2595
      Left            =   3420
      TabIndex        =   0
      Top             =   1200
      Width           =   2205
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Alright,  let me see what you got."
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label lblCash 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Item"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblCashDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   16
      X2              =   208
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Item"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblDrug 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Item"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblPriceDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Price"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblDrugDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Drug"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Shape shpMain 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   2055
      Left            =   120
      Top             =   1230
      Width           =   3135
   End
   Begin VB.Image imgMain 
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Left            =   120
      Picture         =   "frmSellDrugs.frx":0442
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmSellDrugs"
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


Private Sub cmdSell_Click()

If lstInventory.ListIndex < 0 Or _
   lstInventory.ListIndex > 19 Then
   Exit Sub
End If

cmdSell.Enabled = False
frmMain.wsk.SendData Chr$(253) & Chr$(4) & lstInventory.ListIndex & Chr$(0)
DoEvents
cmdExit.SetFocus

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


Private Sub lstInventory_Click()

cmdSell.Enabled = False

End Sub


Private Sub lstInventory_DblClick()

If lstInventory.ListIndex < 0 Or _
   lstInventory.ListIndex > 19 Then
   Exit Sub
End If

If lstInventory.Text = "<Empty>" Then
   cmdSell.Enabled = False
ElseIf lstInventory.Text <> "<Empty>" Then
   cmdSell.Enabled = True
   frmMain.wsk.SendData Chr$(253) & Chr$(3) & lstInventory.ListIndex & Chr$(0)
   DoEvents
End If

End Sub


