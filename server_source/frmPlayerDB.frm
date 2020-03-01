VERSION 5.00
Begin VB.Form frmPlayerDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Player Database"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2370
   Icon            =   "frmPlayerDB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   352
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   158
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4860
      Width           =   2175
   End
   Begin VB.ListBox lstPlayerDB 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4350
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblNOU 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label lblNumberOfUsers 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Num Of Users:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuMainInfo 
         Caption         =   "User Information"
      End
      Begin VB.Menu mnuMainBarOne 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMainDelete 
         Caption         =   "Delete User"
      End
      Begin VB.Menu mnuMainBarTwo 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMainBan 
         Caption         =   "Ban User IP"
      End
   End
End
Attribute VB_Name = "frmPlayerDB"
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


Private Sub lstPlayerDB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
   PopupMenu mnuMain
End If

End Sub

Private Sub mnuMainDelete_Click()
Dim a As Integer
Dim b As Integer
Dim c As Integer

a = MsgBox("Delete User " & lstPlayerDB.Text, vbYesNo, "Confirm Delete")

If a = vbNo Then
   Exit Sub
End If

For a = 0 To UBound(UserDB)
   If Trim$(LCase$(lstPlayerDB.Text)) = Trim$(LCase$(UserDB(a).UName)) Then
      Exit For
   ElseIf a = UBound(UserDB) Then
      Exit Sub
   End If
Next a

For b = 1 To MaxUsers
   If UserDB(a).UName = User(b).UName And _
      UserDB(a).UserGUID = User(b).UserGUID And _
      User(b).Status = "Playing" Then
      frmMain.wsk(b).SendData Chr$(2) & "Your account has been deleted from the server database..." & vbCrLf & vbCrLf & "Have a Wonderfull Day!!!" & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      frmMain.wsk(b).Close
      frmMain.lstUsers.List(b - 1) = "<Waiting>"
      Call ResetIndex(b)
   End If
Next b
      
For b = 0 To UBound(Item)
   If Item(b).ItemGUID = UserDB(a).UserGUID Then
      Call ResetItem(b)
   End If
Next b

For b = 0 To UBound(City)
   If City(b).OwnerGUID = UserDB(a).UserGUID Then
      City(b).OwnerGUID = ""
      City(b).CDesc = "An old run down Mafia Boss House."
         For c = 0 To UBound(City(b).Storage)
            If City(b).Storage(c) <> -1 Then
               Call ResetItem(City(b).Storage(c))
            End If
            City(b).Storage(c) = -1
         Next c
   End If
Next b
      
Call ResetUserDB(a)

frmPlayerDB.lstPlayerDB.Clear
For a = 0 To UBound(UserDB)
   If UserDB(a).UName <> "" And _
      UserDB(a).UserGUID <> "" Then
   frmPlayerDB.lstPlayerDB.AddItem UserDB(a).UName
   End If
Next a

frmPlayerDB.lblNOU.Caption = UBound(UserDB) + 1

End Sub
