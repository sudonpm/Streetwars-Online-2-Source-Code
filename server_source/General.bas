Attribute VB_Name = "General"
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


'Maximum concurrent connections
Public Const MaxUsers = 20

'Anti-Hammer Boolean
Public RunCode(MaxUsers) As Boolean

'Global Heal Price
Public Const HealPrice = 110

'Global Save Times
Public Const SaveTick = 900000
Public SaveNew As Long
Public SaveOld As Long

Public IPBan() As String

'Server Port
Public Const ServerPort = 5002

'Client Version Check
Public Const SVersion = 7000

'Flag to avoid loading sockets that are already loaded
Public RunOnce As Boolean

'Decay Settings
Public Const DecayTime = 600000
Public DecayOld As Long
Public DecayNew As Long

'Player Heal
Public Const HealTime = 20000
Public HealNew As Long
Public HealOld As Long

'The Skill Delay
Public Const SkillDelayTick = 3000

'AirPort Locations
Public NY_Location As Integer
Public HO_Location As Integer
Public MI_Location As Integer
Public CH_Location As Integer
Public LA_Location As Integer
Public NJ_Location As Integer

'Travel Prices
Public Const NY_Price = 339
Public Const HO_Price = 421
Public Const MI_Price = 283
Public Const CH_Price = 199
Public Const LA_Price = 213
Public Const NJ_Price = 241

'Map String Used When Server is Loaded
Public NYMap As String
Public NJMap As String
Public MIMap As String
Public HOMap As String
Public LAMap As String
Public CHMap As String

'Disable X (Close Button) on main form
Public Const MF_BYPOSITION = &H400
Public Const MF_REMOVE = &H1000
Public Declare Function DrawMenuBar Lib "user32" _
(ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" _
(ByVal hMenu As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" _
(ByVal hwnd As Long, _
ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32" _
(ByVal hMenu As Long, _
ByVal nPosition As Long, _
ByVal wFlags As Long) As Long

'Timing Counter
Public Declare Function GetTickCount Lib "kernel32" _
() As Long

'Global Unique Identifier
Private Type GUID
    PartOne As Long
    PartTwo As Integer
    PartThree As Integer
    PartFour(7) As Byte
End Type
      
'Globl Unique Identifier
Private Declare Function CoCreateGuid Lib "OLE32.DLL" _
(ptrGuid As GUID) As Long


Public Function GUID() As String
    Dim lRetVal As Long
    Dim udtGuid As GUID
    
    Dim sPartOne As String
    Dim sPartTwo As String
    Dim sPartThree As String
    Dim sPartFour As String
    Dim iDataLen As Integer
    Dim iStrLen As Integer
    Dim iCtr As Integer
    Dim sAns As String
   
    On Error GoTo errorhandler
    sAns = ""
    
    lRetVal = CoCreateGuid(udtGuid)
    
    If lRetVal = 0 Then
    
       'First 8 chars
        sPartOne = Hex$(udtGuid.PartOne)
        iStrLen = Len(sPartOne)
        iDataLen = Len(udtGuid.PartOne)
        sPartOne = String((iDataLen * 2) - iStrLen, "0") _
        & Trim$(sPartOne)
        
        'Next 4 Chars
        sPartTwo = Hex$(udtGuid.PartTwo)
        iStrLen = Len(sPartTwo)
        iDataLen = Len(udtGuid.PartTwo)
        sPartTwo = String((iDataLen * 2) - iStrLen, "0") _
        & Trim$(sPartTwo)
           
        'Next 4 Chars
        sPartThree = Hex$(udtGuid.PartThree)
        iStrLen = Len(sPartThree)
        iDataLen = Len(udtGuid.PartThree)
        sPartThree = String((iDataLen * 2) - iStrLen, "0") _
        & Trim$(sPartThree)   'Next 2 bytes (4 hex digits)
           
        'Final 16 chars
        For iCtr = 0 To 7
            sPartFour = sPartFour & _
            Format$(Hex$(udtGuid.PartFour(iCtr)), "00")
        Next
 
     'To create GUID with "-", change line below to:
     'sAns = sPartOne & "-" & sPartTwo & "-" & sPartThree _
     '& "-" & sPartFour
       
       sAns = sPartOne & sPartTwo & sPartThree & sPartFour
            
        End If
        
        GUID = sAns
Exit Function

errorhandler:
'return a blank string if there's an error
Exit Function

End Function

Public Sub ResetIndex(Index As Integer)
Dim a As Integer

'Reset Index Type Data For Fresh Login
   User(Index).UName = ""
   User(Index).UPass = ""
   User(Index).UserGUID = ""
   User(Index).Purge = Date
   User(Index).Status = ""
   User(Index).DataBaseID = -1
   User(Index).Location = 0
   User(Index).HomeTown = ""
   User(Index).CurrTown = ""
   User(Index).HomeAbv = ""
   User(Index).Reputation = 0
   User(Index).Rank = ""
   User(Index).Kills = 0
   User(Index).Cash = 0
   User(Index).Bank = 0
   User(Index).TargetNum = -1
   User(Index).TargetGUID = ""
   User(Index).NpcTrade = -1
   User(Index).Accuracy = 0#
   User(Index).Hiding = 0#
   User(Index).Search = 0#
   User(Index).Tracking = 0#
   User(Index).Chemistry = 0#
   User(Index).Pimping = 0#
   User(Index).Snooping = 0#
   User(Index).Stealing = 0#
   User(Index).IsHiding = False
   User(Index).Health = 0
   User(Index).Mute = False
   User(Index).AccessLevel = 0
   
   For a = 0 To 19
      User(Index).Item(a) = -1
   Next a

End Sub

Public Sub ResetUserDB(Index As Integer)
Dim a As Integer

'Reset Index Type Data For Fresh Login
   UserDB(Index).UName = ""
   UserDB(Index).UPass = ""
   UserDB(Index).UserGUID = ""
   UserDB(Index).Purge = Date
   UserDB(Index).Status = ""
   UserDB(Index).DataBaseID = -1
   UserDB(Index).Location = 0
   UserDB(Index).HomeTown = ""
   UserDB(Index).CurrTown = ""
   UserDB(Index).HomeAbv = ""
   UserDB(Index).Reputation = 0
   UserDB(Index).Rank = ""
   UserDB(Index).Kills = 0
   UserDB(Index).Cash = 0
   UserDB(Index).Bank = 0
   UserDB(Index).TargetNum = -1
   UserDB(Index).TargetGUID = ""
   UserDB(Index).NpcTrade = -1
   UserDB(Index).Accuracy = 0#
   UserDB(Index).Hiding = 0#
   UserDB(Index).Search = 0#
   UserDB(Index).Tracking = 0#
   UserDB(Index).Chemistry = 0#
   UserDB(Index).Pimping = 0#
   UserDB(Index).Snooping = 0#
   UserDB(Index).Stealing = 0#
   UserDB(Index).IsHiding = False
   UserDB(Index).Health = 0
   UserDB(Index).Mute = False
   UserDB(Index).AccessLevel = 0
   

   For a = 0 To UBound(UserDB(Index).Item)
      UserDB(Index).Item(a) = -1
   Next a

End Sub
Public Sub ResetItem(Index As Integer)
   
'Reset The Item To Allow an Overwrite
   Item(Index).IName = ""
   Item(Index).IDesc = ""
   Item(Index).ItemGUID = ""
   Item(Index).OnPlayer = False
   Item(Index).Equip = False
   Item(Index).Amount = -1
   Item(Index).Damage = -1
   Item(Index).Armor = -1
   Item(Index).Condition = -1
   Item(Index).IType = -1
   Item(Index).Price = -1
   Item(Index).Multiple = False
   Item(Index).Movable = False
   Item(Index).CanBuy = 0
   Item(Index).Decay = -1
   Item(Index).ILocation = -1

End Sub

Public Function IsFile(InFile As String) As Boolean

If Len(Dir$(InFile)) > 0 Then
   IsFile = True
Else
   IsFile = False
End If

End Function


Public Sub LinkItems()
Dim a As Integer, b As Integer, c As Integer 'Counters

'Link all items to players/npcs
For a = 0 To UBound(Item)
   For b = 0 To UBound(UserDB)
      If Item(a).ItemGUID = UserDB(b).UserGUID And _
         Item(a).OnPlayer = True Then
         For c = 0 To 19
            If UserDB(b).Item(c) = -1 Then
               UserDB(b).Item(c) = a
               Exit For
            End If
         Next c
      End If
   Next b
Next a
         
'Link items to npc
For a = 0 To UBound(Npc)
   For b = 0 To UBound(Item)
      If Item(b).ItemGUID = Npc(a).NpcGUID And _
         Item(b).OnPlayer = True Then
            If Item(b).IType = C_Gun And _
               Item(b).Equip = True Then
               Npc(a).NGear(0) = b
            ElseIf Item(b).IType = C_Melee And _
                   Item(b).Equip = True Then
                   Npc(a).NGear(0) = b
            ElseIf Item(b).IType = C_Armor And _
                   Item(b).Equip = True Then
               Npc(a).NGear(1) = b
            ElseIf Item(b).IType = C_Ammo And _
                   Item(b).Equip = True Then
               Npc(a).NGear(2) = b
            Else
               For c = 0 To 19
                  If Npc(a).NItem(c) = -1 Then
                     Npc(a).NItem(c) = b
                     Exit For
                  End If
               Next c
            End If
      End If
   Next b
Next a

End Sub
Public Sub RemoveDecay()
Dim a As Integer 'Counter
Dim b As Integer 'Counter

'Remove decayed items from the game
For a = 0 To UBound(Item)
   If Item(a).Decay <> -1 Then
      If DecayNew - Item(a).Decay > DecayTime Then
         For b = 0 To UBound(City(Item(a).ILocation).CItem)
            If City(Item(a).ILocation).CItem(b) = a Then
               City(Item(a).ILocation).CItem(b) = -1
            End If
         Next b
         Call ResetItem(a)
      End If
   End If
Next a
      
End Sub

Public Sub AirportLocations()
Dim a As Integer 'Counter

'Find all the airports and set locations
For a = 0 To UBound(City)
   If City(a).CName = "Los Angeles" And _
      City(a).AirPort = True Then
      LA_Location = a
   ElseIf City(a).CName = "Miami" And _
      City(a).AirPort = True Then
      MI_Location = a
   ElseIf City(a).CName = "Houston" And _
      City(a).AirPort = True Then
      HO_Location = a
   ElseIf City(a).CName = "New York" And _
      City(a).AirPort = True Then
      NY_Location = a
   ElseIf City(a).CName = "Chicago" And _
      City(a).AirPort = True Then
      CH_Location = a
   ElseIf City(a).CName = "New Jersey" And _
      City(a).AirPort = True Then
      NJ_Location = a
   End If
Next a
   
End Sub

Public Function MaleNames() As String
Dim a As Integer

Randomize
a = Int(50 - 1) * Rnd + 1

Select Case a
   Case 1
      MaleNames = "Bob"
   Case 2
      MaleNames = "Fred"
   Case 3
      MaleNames = "Ted"
   Case 4
      MaleNames = "Mike"
   Case 5
      MaleNames = "Brian"
   Case 6
      MaleNames = "Seth"
   Case 7
      MaleNames = "George"
   Case 8
      MaleNames = "Craig"
   Case 9
      MaleNames = "Chris"
   Case 10
      MaleNames = "Hoagan"
   Case 11
      MaleNames = "Ricky"
   Case 12
      MaleNames = "Michael"
   Case 13
      MaleNames = "Stanley"
   Case 14
      MaleNames = "Greg"
   Case 15
      MaleNames = "Brandon"
   Case 16
      MaleNames = "Harold"
   Case 17
      MaleNames = "Matt"
   Case 18
      MaleNames = "Daniel"
   Case 19
      MaleNames = "Danny"
   Case 20
      MaleNames = "Aaron"
   Case 21
      MaleNames = "Spencer"
   Case 22
      MaleNames = "Kyle"
   Case 23
      MaleNames = "Mark"
   Case 24
      MaleNames = "Richard"
   Case 25
      MaleNames = "Jonathan"
   Case 26
      MaleNames = "Eric"
   Case 27
      MaleNames = "David"
   Case 28
      MaleNames = "Ryan"
   Case 29
      MaleNames = "Patrick"
   Case 30
      MaleNames = "Jonny"
   Case 31
      MaleNames = "John"
   Case 32
      MaleNames = "Neil"
   Case 33
      MaleNames = "Justin"
   Case 34
      MaleNames = "Chad"
   Case 35
      MaleNames = "Tom"
   Case 36
      MaleNames = "Thomas"
   Case 37
      MaleNames = "Jason"
   Case 38
      MaleNames = "Chase"
   Case 39
      MaleNames = "Shawn"
   Case 40
      MaleNames = "Sean"
   Case 41
      MaleNames = "Garrett"
   Case 42
      MaleNames = "Adam"
   Case 43
      MaleNames = "Sylvester"
   Case 44
      MaleNames = "Bruce"
   Case 45
      MaleNames = "Arnold"
   Case 46
      MaleNames = "Paul"
   Case 47
      MaleNames = "Billy"
   Case 48
      MaleNames = "Steve"
   Case 49
      MaleNames = "Al"
   Case 50
      MaleNames = "Jake"
End Select

End Function


Public Sub ResetNPC(NpcIndex As Integer)
Dim a As Integer

'Remove the NPC from the game
With Npc(NpcIndex)
   Npc(NpcIndex).NName = ""
   Npc(NpcIndex).NameTag = ""
   Npc(NpcIndex).NLocation = -1
   Npc(NpcIndex).NpcGUID = ""
   Npc(NpcIndex).NHealth = 0
   Npc(NpcIndex).NCash = 0
   Npc(NpcIndex).GearGun = 0
   Npc(NpcIndex).GearArmor = 0
   Npc(NpcIndex).GearAmmo = 0
   Npc(NpcIndex).NpcType = -1
   Npc(NpcIndex).NPCOwner = ""
   Npc(NpcIndex).NTargetID = -1
   Npc(NpcIndex).NTargetGUID = ""
   Npc(NpcIndex).CanMove = -1
   Npc(NpcIndex).NCity = ""
   Npc(NpcIndex).NMovable = True
   Npc(NpcIndex).Accuracy = 0#
End With

For a = 0 To 19
   Npc(NpcIndex).NItem(a) = -1
Next a

For a = 0 To 2
   Npc(NpcIndex).NGear(a) = -1
Next a

End Sub

Public Function FemaleNames() As String
Dim a As Integer

Randomize
a = Int(50 - 1) * Rnd + 1

Select Case a
   Case 1
      FemaleNames = "Susan"
End Select

End Function

Public Sub SaveMessage()
On Error Resume Next
Dim a As Integer

For a = 1 To MaxUsers
   If User(a).Status = "Playing" Then
      frmMain.wsk(a).SendData Chr$(2) & ">>>-- Saving Server Data --<<<" & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
   End If
Next a

End Sub

Public Sub ChatLog(Index As Integer, Msg As String)
Const ChatLog = 5000
Dim ff As Integer, a As Integer
Dim SplitLog() As String

With frmMain.txtOutput
  If Len(.Text) > ChatLog Then
  SplitLog = Split(.Text, vbCrLf)
    ff = FreeFile
    Open App.Path & "\server.log" For Append As ff
    For a = 0 To UBound(SplitLog) - 1
      Print #ff, SplitLog(a)
    Next a
    Close ff
    .Text = "Saved Server Log - Clearing Buffer" & vbCrLf
  End If
End With

With frmMain.txtOutput
  .Text = .Text & "[" & User(Index).UName & "] - " & Msg & vbCrLf
  .SelStart = Len(.Text)
End With

End Sub


Public Sub CloseSocket(Index As Integer)

If User(Index).UName <> "" And _
   User(Index).Status = "Playing" Then
   UserDB(User(Index).DataBaseID) = User(Index)
End If

frmMain.wsk(Index).Close
Call ResetIndex(Index)
Call UpdatePlayerList

frmMain.lstUsers.List(Index - 1) = "<Waiting>"

End Sub
