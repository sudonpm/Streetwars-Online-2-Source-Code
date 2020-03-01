VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Street Wars Online II (Alpha)"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10875
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   483
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   725
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Left            =   480
      Top             =   0
   End
   Begin MSWinsockLib.Winsock wsk 
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstInventory 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2010
      Left            =   8520
      TabIndex        =   9
      Top             =   3420
      Width           =   2295
   End
   Begin VB.ListBox lstUsers 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2790
      Left            =   8520
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   300
      Width           =   2295
   End
   Begin VB.CommandButton cmdTravel 
      Caption         =   "Travel"
      Height          =   315
      Left            =   4920
      TabIndex        =   5
      Top             =   6840
      Width           =   1140
   End
   Begin VB.CommandButton cmdPawnShop 
      Caption         =   "Pawn Shop"
      Height          =   315
      Left            =   3720
      TabIndex        =   4
      Top             =   6840
      Width           =   1140
   End
   Begin VB.CommandButton cmdMap 
      Caption         =   "Map"
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   6840
      Width           =   1140
   End
   Begin VB.TextBox txtInput 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      MaxLength       =   200
      TabIndex        =   0
      Top             =   6480
      Width           =   8310
   End
   Begin VB.TextBox txtOutput 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3180
      Width           =   8310
   End
   Begin VB.TextBox txtNews 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2370
      Width           =   8310
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   408
      X2              =   560
      Y1              =   464
      Y2              =   464
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   160
      X2              =   8
      Y1              =   464
      Y2              =   464
   End
   Begin VB.Label lblLastSell 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8640
      TabIndex        =   36
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Label lblLastSellDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Sell Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   8640
      TabIndex        =   35
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label lblLastBuy 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8640
      TabIndex        =   34
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label lblLastBuyDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Buy Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   8640
      TabIndex        =   33
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Shape shpNavigation 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1620
      Left            =   8520
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label lblAmmo 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6840
      TabIndex        =   32
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6840
      TabIndex        =   31
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblAmmoDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ammo:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   30
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblArmorDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Armor:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   29
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6840
      TabIndex        =   28
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblWeaponDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Weapon:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   27
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblEquip 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Equipment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   26
      Top             =   1125
      Width           =   2535
   End
   Begin VB.Label lblKills 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   25
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblRank 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   24
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblLocation 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblHomeTown 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   22
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblKillsDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Kills:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   21
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblRankDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Rank:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   20
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblLocationDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Location:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblHomeTownDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Home Town:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblBank 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   17
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblCash 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   16
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblHealth 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   15
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblName 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblBankDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblCashDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblHealthDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Health:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblNameDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblInventory 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inventory"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8520
      TabIndex        =   8
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label lblDealers 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dealers Online"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8520
      TabIndex        =   6
      Top             =   0
      Width           =   2295
   End
   Begin VB.Shape shpMain 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   1215
      Left            =   135
      Top             =   1065
      Width           =   8280
   End
   Begin VB.Image imgMain 
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Left            =   120
      Picture         =   "frmMain.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8310
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuFileDisconnect 
         Caption         =   "&Disconnect"
      End
      Begin VB.Menu mnuFileBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuQuickKeys 
      Caption         =   "&Q-Keys"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpGuide 
         Caption         =   "Street Wars Online II Help Guide"
      End
      Begin VB.Menu mnuHelpVisitSite 
         Caption         =   "Visit Street Wars Online II Website"
      End
   End
   Begin VB.Menu mnuInventory 
      Caption         =   "Inventory"
      Visible         =   0   'False
      Begin VB.Menu mnuInventoryEquip 
         Caption         =   "Equip"
      End
      Begin VB.Menu mnuInventoryUnequip 
         Caption         =   "Un-Equip"
      End
      Begin VB.Menu mnuInventoryExamine 
         Caption         =   "Examine"
      End
      Begin VB.Menu mnuInventoryUse 
         Caption         =   "Use"
      End
      Begin VB.Menu mnuInventoryDrop 
         Caption         =   "Drop"
      End
   End
End
Attribute VB_Name = "frmMain"
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

Private Sub cmdMap_Click()
On Error Resume Next

frmMain.wsk.SendData Chr$(253) & Chr$(5) & Chr$(0)
DoEvents

End Sub


Private Sub cmdPawnShop_Click()
On Error Resume Next

frmMain.wsk.SendData Chr$(254) & Chr$(2) & Chr$(0)
DoEvents

End Sub
Private Sub cmdSkills_Click()

End Sub

Private Sub cmdTravel_Click()
On Error Resume Next

frmMain.wsk.SendData Chr$(255) & Chr$(6) & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If MoveDelay = True Then
   Exit Sub
End If

If KeyUsed = False Then
If KeyCode = vbKeyUp Then
   frmMain.wsk.SendData "n" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyRight Then
   frmMain.wsk.SendData "e" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyDown Then
   frmMain.wsk.SendData "s" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyLeft Then
   frmMain.wsk.SendData "w" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyF1 Then
   frmMain.wsk.SendData "punch" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyF2 Then
   frmMain.wsk.SendData "strike" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyF3 Then
   frmMain.wsk.SendData "fire" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyF4 Then
   frmMain.wsk.SendData "look" & Chr$(0)
   DoEvents
   KeyUsed = True
End If
End If

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyUp Then
   KeyUsed = False
ElseIf KeyCode = vbKeyRight Then
   KeyUsed = False
ElseIf KeyCode = vbKeyDown Then
   KeyUsed = False
ElseIf KeyCode = vbKeyLeft Then
   KeyUsed = False
ElseIf KeyCode = vbKeyF1 Then
   KeyUsed = False
ElseIf KeyCode = vbKeyF2 Then
   KeyUsed = False
ElseIf KeyCode = vbKeyF3 Then
   KeyUsed = False
ElseIf KeyCode = vbKeyF4 Then
   KeyUsed = False
End If

End Sub
Private Sub Form_Load()
Dim a As Integer 'Counter

'Setup initial inventory slots
For a = 0 To 19
  lstInventory.AddItem "<Empty>"
Next a

txtNews.BackColor = vbBlack
txtNews.ForeColor = vbWhite

End Sub
Private Sub imgEast_Click()

If frmMain.wsk.State <> sckClosed Then
   frmMain.wsk.SendData "e" & Chr$(0)
   DoEvents
End If

End Sub

Private Sub imgNorth_Click()

If frmMain.wsk.State <> sckClosed Then
   frmMain.wsk.SendData "n" & Chr$(0)
   DoEvents
End If


End Sub

Private Sub imgSouth_Click()

If frmMain.wsk.State <> sckClosed Then
   frmMain.wsk.SendData "s" & Chr$(0)
   DoEvents
End If

End Sub

Private Sub imgWest_Click()

If frmMain.wsk.State <> sckClosed Then
   frmMain.wsk.SendData "w" & Chr$(0)
   DoEvents
End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

wsk.Close

End Sub
Private Sub lstInventory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
   PopupMenu mnuInventory
End If

End Sub


Private Sub mnuFileConnect_Click()
On Error Resume Next
Dim iServ As String

iServ = InputBox("Enter the server IP or Host name you wish to connect to.", "Connect To Server", "127.0.0.1")

If iServ = "" Then Exit Sub

'Disable menus
frmMain.mnuFileConnect.Enabled = False
frmMain.mnuFileExit.Enabled = False

'Connect to the server
With wsk
  .Close
  .Protocol = sckTCPProtocol
  .RemotePort = ServerPort
  .RemoteHost = iServ
  .Connect
End With

Call ShowText("Connecting to the Street Wars Online II central server, please stand by..." & vbCrLf & vbCrLf)

End Sub
Private Sub mnuFileDisconnect_Click()
'Disconnect and enable menus

wsk.Close
frmMain.mnuFileConnect.Enabled = True
frmMain.mnuFileExit.Enabled = True

End Sub

Private Sub mnuFileExit_Click()

   'Close winsock and shut down the game
   wsk.Close
   Unload Me
   End

End Sub

Private Sub mnuHelpGuide_Click()

Call OpenLocation("http://streetwars.8m.com/street_wars_online_ii_online_hel.htm", SW_SHOWNORMAL)

End Sub

Private Sub mnuHelpVisitSite_Click()

Call OpenLocation("http://streetwars.8m.com", SW_SHOWNORMAL)

End Sub
Private Sub mnuInventoryDrop_Click()

frmMain.wsk.SendData Chr$(7) & lstInventory.ListIndex & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub

Private Sub mnuInventoryEquip_Click()

frmMain.wsk.SendData Chr$(255) & Chr$(3) & frmMain.lstInventory.ListIndex & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub

Private Sub mnuInventoryExamine_Click()

'Examine the item
frmMain.wsk.SendData Chr$(255) & Chr$(2) & frmMain.lstInventory.ListIndex & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub

Private Sub mnuInventoryUnequip_Click()

frmMain.wsk.SendData Chr$(255) & Chr$(4) & frmMain.lstInventory.ListIndex & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub

Private Sub mnuInventoryUse_Click()

frmMain.wsk.SendData Chr$(255) & Chr$(5) & frmMain.lstInventory.ListIndex & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
On Error GoTo Failed 'Error Handler

'Send textbox text to the server
If (KeyAscii = 13) And (txtInput.Text <> "") Then
  KeyAscii = 0
  If wsk.State <> sckClosed Then
     If InputDelay = True Then
        Exit Sub
     End If
    wsk.SendData Trim$(txtInput.Text) & Chr$(0)
    DoEvents
    txtInput.Text = ""
  End If
End If
Exit Sub

'If an error occurs,  close the socket and reset
'everything
Failed:
wsk.Close
With txtOutput
  .Text = .Text & "An error has occured while sending data to the server, your connection has been reset." & vbCrLf & vbCrLf
  .SelStart = Len(.Text)
End With
txtInput.Text = ""
tmrMain.Enabled = False
mnuFileConnect.Enabled = True
mnuFileExit.Enabled = True

End Sub
Private Sub txtNews_GotFocus()
  'Don't allow textbox to have focus
  txtInput.SetFocus
End Sub


Private Sub txtOutput_GotFocus()
  'Don't allow textbox to get focus
  txtInput.SetFocus
End Sub


Private Sub wsk_Connect()

frmMain.wsk.SendData ClientVer & Chr$(0)
DoEvents

End Sub

Private Sub wsk_DataArrival(ByVal bytesTotal As Long)
Dim a As Integer 'Counter
Dim Msg As String 'String to hold data off the wire
Dim SplitMsg() As String 'String array to parse data

'Pull data off the wire
wsk.GetData Msg, vbString

'Split the string array
SplitMsg = Split(Msg, Chr$(0))

'Loop through data and process accordingly
For a = 0 To UBound(SplitMsg) - 1
   
   Select Case Left$(SplitMsg(a), 2)
      Case Chr$(255) & Chr$(2)
         Call TravelMenu(Mid$(SplitMsg(a), 3))
      Case Chr$(255) & Chr$(3)
         Call PawnShopMenu(Mid$(SplitMsg(a), 3))
      Case Chr$(255) & Chr$(4)
         Call UpdateCashRank(Mid$(SplitMsg(a), 3))
      Case Chr$(255) & Chr$(5)
         Call PawnShopItemInfo(Mid$(SplitMsg(a), 3))
      Case Chr$(255) & Chr$(6)
         Call PawnShopPlayerInventoryUpdate(Mid$(SplitMsg(a), 3))
      Case Chr$(255) & Chr$(7)
         Call UpdateGeneralInfo(Mid$(SplitMsg(a), 3))
      Case Chr$(254) & Chr$(2)
         Call UpdateGearInfo(Mid$(SplitMsg(a), 3))
      Case Chr$(254) & Chr$(3)
         Call UpdatePlayerList(Mid$(SplitMsg(a), 3))
      Case Chr$(254) & Chr$(4)
         Call BuyDrugMenu(Mid$(SplitMsg(a), 3))
      Case Chr$(254) & Chr$(5)
         Call CloseDrugDealMenu
      Case Chr$(254) & Chr$(6)
         Call DrugDealItemInfo(Mid$(SplitMsg(a), 3))
      Case Chr$(254) & Chr$(7)
         Call DrugDealMessage(Mid$(SplitMsg(a), 3))
      Case Chr$(253) & Chr$(2)
         Call UpdateDealerInventory(Mid$(SplitMsg(a), 3))
      Case Chr$(253) & Chr$(3)
         Call SellDrugMenu(Mid$(SplitMsg(a), 3))
      Case Chr$(253) & Chr$(4)
         Call CloseDruggieMenu
      Case Chr$(253) & Chr$(5)
         Call DruggieMenuMessage(Mid$(SplitMsg(a), 3))
      Case Chr$(253) & Chr$(6)
         Call DruggieMenuItemInfo(Mid$(SplitMsg(a), 3))
      Case Chr$(253) & Chr$(7)
         Call ReUpdateDruggieInventory(Mid$(SplitMsg(a), 3))
      Case Chr$(252) & Chr$(2)
         Call ShowMap(Mid$(SplitMsg(a), 3))
      Case Chr$(252) & Chr$(3)
         Call UpdateNews(Mid$(SplitMsg(a), 3))
      Case Chr$(252) & Chr$(4)
         frmMain.lblLastBuy.Caption = Mid$(SplitMsg(a), 3)
      Case Chr$(252) & Chr$(5)
         frmMain.lblLastSell.Caption = Mid$(SplitMsg(a), 3)
   End Select
   
   Select Case Left$(SplitMsg(a), 1)
      Case Chr$(2)
         Call ShowText(Mid$(SplitMsg(a), 2))
      Case Chr$(3)
         Call NewAccount
      Case Chr$(4)
         Call DupeName
      Case Chr$(5)
         Call AccountCreated
      Case Chr$(6)
         Call UpdateFullInventory(Mid$(SplitMsg(a), 2))
      Case Chr$(7)
         Call UpdateSingleItem(Mid$(SplitMsg(a), 2))
   End Select
Next a

End Sub
