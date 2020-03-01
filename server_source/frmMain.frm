VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Street Wars Online II Server"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7560
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   304
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   504
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   0
   End
   Begin MSWinsockLib.Winsock wsk 
      Index           =   0
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtInput 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   7335
   End
   Begin VB.TextBox txtOutput 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4155
      Left            =   2280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   5175
   End
   Begin VB.ListBox lstUsers 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4155
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileStart 
         Caption         =   "&Start Server"
      End
      Begin VB.Menu mnuFileStop 
         Caption         =   "&Stop Server"
      End
      Begin VB.Menu mnuFileBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsUserDataBase 
         Caption         =   "User Database"
      End
      Begin VB.Menu mnuOptionsNpcDB 
         Caption         =   "Npc Database"
      End
      Begin VB.Menu mnuOptionsSpawn 
         Caption         =   "Spawn NPCs"
      End
      Begin VB.Menu mnuOptionsRestock 
         Caption         =   "Restock"
      End
      Begin VB.Menu mnuOptionsBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsPurge 
         Caption         =   "NPC Purge"
      End
   End
   Begin VB.Menu mnuUser 
      Caption         =   "User"
      Visible         =   0   'False
      Begin VB.Menu mnuUserKillConnection 
         Caption         =   "Kill Connection"
      End
      Begin VB.Menu mnuUserTempIpBan 
         Caption         =   "Temp IP Ban"
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

Private Sub Form_Load()
Dim a As Integer 'Counter
Dim b As Integer 'Counter

mnuOptions.Enabled = False

'Disable Menu
mnuFileStop.Enabled = False

'Disable Timers
tmrMain.Enabled = False

'Reset All User Data Type Structures To Game Default
For a = 0 To MaxUsers
   Call ResetIndex(a)
Next a

'Redim All City Arrays
For a = 0 To UBound(City)
   ReDim City(a).CItem(0)
'   ReDim City(a).CNpc(0)
   City(a).CItem(0) = -1
'   City(a).CNpc(0) = -1
   For b = 0 To 49
      City(a).Storage(b) = -1
   Next b
Next a

For a = 0 To UBound(City)
   For b = 0 To 9
      City(a).CNpc(b) = -1
   Next b
Next a

For a = 0 To MaxUsers
   RunCode(a) = False
Next a

'Dim All Arrays
ReDim UserDB(0)
ReDim Item(0)
ReDim Npc(0)
ReDim IPBan(0)
ResetUserDB (0)
ResetItem (0)
ResetNPC (0)


'Test Code
Call LoadStaticItems
Call LoadStaticNPCs

'Item Slot Index Tracker
Dim z As Integer
z = 0
For a = 0 To UBound(ItemDB)
   If ItemDB(a).ForSale = True Then
      ReDim Preserve SlotID(z)
      SlotID(z) = a
      z = z + 1
   End If
Next a

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


Private Sub lstUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
   PopupMenu mnuUser
End If

End Sub

Private Sub mnuFileExit_Click()
   Unload Me
   End
End Sub

Private Sub mnuFileStart_Click()
Dim a As Integer 'Counter
Dim b As Integer 'Counter

SaveNew = GetTickCount()
SaveOld = GetTickCount()

With frmMain.txtOutput
   .Text = .Text & "Loading Server Data..." & vbCrLf
   .SelStart = Len(.Text)
End With
DoEvents

mnuFileStart.Enabled = False 'Disable Menus
mnuFileExit.Enabled = False
mnuFileStop.Enabled = True

'Load World
Call LoadCitys
Call LoadPlayerData
Call LoadNPCs
Call LoadItems
Call LoadMap

'Find Map Airports
Call AirportLocations

'Link Game Objects
Call LinkItems

'Recalculate weapons/armor/ammo and relink
Call SetGearValues

'Load Main Listening Socket
With wsk(0)
  .Protocol = sckTCPProtocol
  .LocalPort = ServerPort
  .Listen
End With

'Load Player Sockets
If RunOnce = False Then
  For a = 1 To MaxUsers
    Load wsk(a)
    wsk(a).Protocol = sckTCPProtocol
    wsk(a).LocalPort = ServerPort
  Next a
RunOnce = True
End If

'Set User List Box Text
For a = 0 To MaxUsers - 1
  lstUsers.List(a) = "<Waiting>"
Next a

'For a = 0 To 1000
'Call AddNpc(N_Dealer, Int(5399 - 1) * Rnd + 1)
'Next a

tmrMain.Enabled = True


With frmMain.txtOutput
   .Text = .Text & "Server Started..." & vbCrLf & _
   UBound(UserDB) + 1 & " Players Loaded" & vbCrLf & _
   UBound(Item) + 1 & " Items Loaded" & vbCrLf & _
   UBound(Npc) & " NPC's Loaded" & vbCrLf
   .SelStart = Len(.Text)
End With

mnuOptions.Enabled = True

End Sub
Private Sub mnuFileStop_Click()
Dim a As Integer 'Counter

mnuOptions.Enabled = False

With frmMain.txtOutput
   .Text = .Text & "Shutting Down Server, Saving Data..." & vbCrLf
   .SelStart = Len(.Text)
End With
DoEvents

'Close all sockets
If RunOnce = True Then
   For a = 0 To MaxUsers
     wsk(a).Close
   Next a
End If

mnuFileStart.Enabled = True 'Enable Menus
mnuFileExit.Enabled = True
mnuFileStop.Enabled = False

'Clear User List Box
lstUsers.Clear

'Save All Game Data
Call SaveCitys
Call SavePlayerData
Call SaveItems
Call SaveNPCs

'Disable Main Timer
tmrMain.Enabled = False

With frmMain.txtOutput
   .Text = .Text & "Server Successfully Shut Down..." & vbCrLf
   .SelStart = Len(.Text)
End With

End Sub

Private Sub mnuOptionsNpcDB_Click()
Dim a As Integer
Dim Msg As String

frmMain.Enabled = False
frmNpcDB.Show
DoEvents

For a = 0 To UBound(Npc)
If Npc(a).NName <> "" And _
   Npc(a).NpcGUID <> "" Then
   Msg = Npc(a).NName & " the " & Npc(a).NameTag
   If Npc(a).NGear(0) <> -1 Then
   Msg = Msg & "  |  " & Item(Npc(a).NGear(0)).IName
   End If
   If Npc(a).NGear(1) <> -1 Then
   Msg = Msg & "  |  " & Item(Npc(a).NGear(1)).IName
   End If
   If Npc(a).NGear(2) <> -1 Then
   Msg = Msg & "  |  " & Item(Npc(a).NGear(2)).IName
   End If
   Msg = Msg & "  |  " & City(Npc(a).NLocation).CName
   Msg = Msg & "  " & City(Npc(a).NLocation).Compass
   frmNpcDB.lstNPC.AddItem Msg
End If
Next a

End Sub

Private Sub mnuOptionsPurge_Click()
Dim a As Integer
Dim b As Integer

a = MsgBox("Are You Sure?", vbYesNo, "Confirm")

If a = vbNo Then
   Exit Sub
End If

frmMain.tmrMain.Enabled = False

For a = 0 To UBound(Npc)
   For b = 0 To UBound(Item)
      If Item(b).ItemGUID = Npc(a).NpcGUID Then
         Call ResetItem(b)
      End If
   Next b
   ResetNPC (a)
Next a

For a = 0 To UBound(UserDB)
   For b = 0 To UBound(Item)
      If Item(b).ItemGUID = UserDB(a).UserGUID And _
         Item(b).OnPlayer = True Then
         Exit For
      ElseIf b = UBound(Item) Then
         Call ResetItem(b)
      End If
   Next b
Next a

End Sub

Private Sub mnuOptionsRestock_Click()

Call RestockDrugs

End Sub
Private Sub mnuOptionsSpawn_Click()

'Spawn Drug Dealers/Druggies every Hour
   Call SpawnNPC(N_Dealer, 4) 'Spawn 4 Dealers in Each City
   Call SpawnNPC(N_Druggie, 4) 'Spawn 4 Druggies in Each city
   Call GenNpcInventory
   Call SpawnNPC(N_Cop, 2) 'Spawn 2 Police Officers
   Call SpawnNPC(N_Bum, 2) 'Spawn 2 Street Bums
   Call SpawnNPC(N_Tweaker, 2) 'Spawn 2 Tweakers
   Call GenNpcInventory

End Sub

Private Sub mnuOptionsUserDataBase_Click()
Dim a As Integer

frmMain.Enabled = False
frmPlayerDB.Show
DoEvents

For a = 0 To UBound(UserDB)
   If UserDB(a).UName <> "" And _
      UserDB(a).UserGUID <> "" Then
   frmPlayerDB.lstPlayerDB.AddItem UserDB(a).UName
   End If
Next a

frmPlayerDB.lblNOU.Caption = UBound(UserDB) + 1

End Sub

Private Sub mnuUserTempIpBan_Click()
Dim a As Integer

For a = 0 To UBound(IPBan)
   If IPBan(a) = "" Then
      IPBan(a) = wsk(lstUsers.ListIndex + 1).RemoteHostIP
      frmMain.wsk(lstUsers.ListIndex + 1).SendData Chr$(2) & "You have been kicked and banned from the server.  If you feel this action was unjust, email the server administrator at x-net@swbell.net to resolve the issue." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      wsk(lstUsers.ListIndex + 1).Close
      lstUsers.List(lstUsers.ListIndex) = "<Waiting>"
      If User(lstUsers.ListIndex + 1).Status = "Playing" Then
         UserDB(User(lstUsers.ListIndex + 1).DataBaseID) = User(lstUsers.ListIndex + 1)
      End If
      Call ResetIndex(lstUsers.ListIndex + 1)
      Call UpdatePlayerList
      Exit Sub
   ElseIf a = UBound(IPBan) Then
      ReDim Preserve IPBan(UBound(IPBan) + 1)
      IPBan(UBound(IPBan)) = wsk(lstUsers.ListIndex + 1).RemoteHostIP
      frmMain.wsk(lstUsers.ListIndex + 1).SendData Chr$(2) & "You have been kicked and banned from the server.  If you feel this action was unjust, email the server administrator at x-net@swbell.net to resolve the issue." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      wsk(lstUsers.ListIndex + 1).Close
      lstUsers.List(lstUsers.ListIndex) = "<Waiting>"
      If User(lstUsers.ListIndex + 1).Status = "Playing" Then
         UserDB(User(lstUsers.ListIndex + 1).DataBaseID) = User(lstUsers.ListIndex + 1)
      End If
      Call ResetIndex(lstUsers.ListIndex + 1)
      Call UpdatePlayerList
      Exit Sub
   End If
Next a

End Sub
Private Sub tmrMain_Timer()
On Error GoTo TimerFail
Dim a As Integer
Static StateChange As Boolean

'Save data every 15 minutes
SaveNew = GetTickCount()
If SaveNew - SaveOld > SaveTick Then
   Call SaveMessage
   Call SaveCitys
   Call SavePlayerData
   Call SaveItems
   Call SaveNPCs
   SaveOld = GetTickCount()
End If

'Spawn Drug Dealers/Druggies every Hour
DealerNew = GetTickCount()
If DealerNew - DealerOld > DealerSpawn Then
   Call SpawnNPC(N_Dealer, 4) 'Spawn 4 Dealers in Each City
   Call SpawnNPC(N_Druggie, 4) 'Spawn 4 Druggies in Each city
   Call GenNpcInventory
   DealerOld = GetTickCount()
End If

'Spawn All other NPCs
SpawnNew = GetTickCount()
If SpawnNew - SpawnOld > SpawnTime Then
   Call SpawnNPC(N_Cop, 2) 'Spawn 2 Police Officers
   Call SpawnNPC(N_Bum, 2) 'Spawn 2 Street Bums
   Call SpawnNPC(N_Tweaker, 2) 'Spawn 2 Tweakers
   Call GenNpcInventory
   SpawnOld = GetTickCount()
End If

'Stock NPC's
StockNew = GetTickCount()
If StockNew - StockOld > StockTime Then
   Call RestockDrugs
   StockOld = GetTickCount()
End If

'Run Npc Combat Code
NpcCombatNew = GetTickCount()
If NpcCombatNew - NpcCombatOld > NpcCombatTick Then
   Call NpcCombat
   NpcCombatOld = GetTickCount()
End If

'NPC Walk Code
WalkNew = GetTickCount()
If WalkNew - WalkOld > WalkTime Then
   Call NpcMove
   WalkOld = GetTickCount()
End If

'Check for Decayed Items and clean them up
DecayNew = GetTickCount()
If DecayNew - DecayOld > DecayTime Then
   Call RemoveDecay
   DecayOld = GetTickCount()
End If

'Heal Player 1 point every 15 seconds
HealNew = GetTickCount()
If HealNew - HealOld > HealTime Then
   Call PlayerHealth
   HealOld = GetTickCount()
End If

Exit Sub

TimerFail:
'Error Handler
Dim ff As Integer 'Free File
ff = FreeFile
Open App.Path & "\error.log" For Append As ff
Print #ff, "[BOE]"
Print #ff, "Timer Error - Combat Error"
Print #ff, "[EOE]"
Close ff

End Sub
Private Sub txtOutput_GotFocus()
   'Change Focus
   txtInput.SetFocus
End Sub


Private Sub wsk_Close(Index As Integer)

If User(Index).UName <> "" And _
   User(Index).Status = "Playing" Then
   UserDB(User(Index).DataBaseID) = User(Index)
End If

wsk(Index).Close
Call ResetIndex(Index)
Call UpdatePlayerList

frmMain.lstUsers.List(Index - 1) = "<Waiting>"

End Sub
Private Sub wsk_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim a As Integer 'Counter

If Index = 0 Then
   For a = 1 To MaxUsers
      With wsk(a)
         If .State = sckClosed Then
            .Accept requestID
            User(Index).Status = ""
            lstUsers.List(a - 1) = "<Connecting>"
            Exit For
         End If
      End With
   Next a
End If

End Sub

Private Sub wsk_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo Failed
Dim a As Integer 'Counter
Dim Msg As String
Dim SplitMsg() As String 'Message Array

wsk(Index).GetData Msg, vbString
SplitMsg = Split(Msg, Chr$(0))

'Anti-Hammer Check
'If RunCode(Index) = True Then
'   frmMain.wsk(Index).SendData Chr$(2) & "Quit hammering the server!" & vbCrLf & vbCrLf & Chr$(0)
'   DoEvents
'   Exit Sub
'Else
'   RunCode(Index) = True
'End If

For a = 0 To UBound(SplitMsg) - 1
   Select Case User(Index).Status
      Case ""
         Call SendWelcome(Index, Trim$(Left$(SplitMsg(a), Len(SplitMsg(a)))))
      Case "GetName"
         Call VerifyName(Index, Left$(SplitMsg(a), Len(SplitMsg(a))))
      Case "YesNo"
         Call YesNo(Index, Left$(SplitMsg(a), Len(SplitMsg(a))))
      Case "NewAccount"
         Call NewAccount(Index, Left$(SplitMsg(a), Len(SplitMsg(a))))
      Case "GetPass"
         Call GetPassword(Index, Left$(SplitMsg(a), Len(SplitMsg(a))))
      Case "Playing"
         Call DoCommand(Index, Left$(SplitMsg(a), Len(SplitMsg(a))))
   End Select
Next a

'RunCode(Index) = False

Exit Sub

'Error Handler
Failed:
Dim ff As Integer 'Free File
ff = FreeFile
Open App.Path & "\error.log" For Append As ff
Print #ff, "[BOE]"
Print #ff, "Tracking Error - UserIP/User Name/Command"
Print #ff, wsk(Index).RemoteHostIP & " | " & User(Index).UName & " | " & SplitMsg(a)
Print #ff, "[EOE]"
Close ff

End Sub

Private Sub wsk_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

CloseSocket (Index)
Err.Clear

End Sub
