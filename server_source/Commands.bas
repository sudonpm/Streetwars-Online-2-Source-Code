Attribute VB_Name = "Commands"
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




Public Sub DoCommand(Index As Integer, Msg As String)

If LCase$(Msg) = "look" Then
   Call ShowCity(Index)
   Exit Sub
ElseIf LCase$(Msg) = "n" Then
   Call North(Index)
   Exit Sub
ElseIf LCase$(Msg) = "e" Then
   Call East(Index)
   Exit Sub
ElseIf LCase$(Msg) = "s" Then
   Call South(Index)
   Exit Sub
ElseIf LCase$(Msg) = "w" Then
   Call West(Index)
   Exit Sub
ElseIf Left$(Msg, 1) = ";" Then
   Call SendChat(Index, Trim$(Mid$(Msg, 2)))
   Exit Sub
ElseIf LCase$(Msg) = "mute" Then
   Call Mute(Index)
   Exit Sub
ElseIf LCase$(Left$(Msg, 8)) = "/additem" Then
   Call AddItemGM(Index, Trim$(Mid$(Msg, 9)))
   Exit Sub
ElseIf LCase$(Msg) = "/listitems" Then
   Call ListItemGM(Index)
   Exit Sub
ElseIf LCase$(Left$(Msg, 3)) = "get" Then
   Call GetItem(Index, Trim$(Mid$(Msg, 4)))
   Exit Sub
ElseIf Left$(Msg, 1) = Chr$(7) Then
   Call DropItem(Index, Mid$(Msg, 2))
   Exit Sub
ElseIf Left$(Msg, 2) = Chr$(255) & Chr$(2) Then
   Call ExamineItem(Index, Mid$(Msg, 3))
   Exit Sub
ElseIf Left$(Msg, 2) = Chr$(255) & Chr$(3) Then
   Call EquipItem(Index, Mid$(Msg, 3))
   Exit Sub
ElseIf Left$(Msg, 2) = Chr$(255) & Chr$(4) Then
   Call UnEquip(Index, Mid$(Msg, 3))
   Exit Sub
ElseIf Left$(Msg, 2) = Chr$(255) & Chr$(5) Then
   Call UseItem(Index, Mid$(Msg, 3))
   Exit Sub
ElseIf Msg = Chr$(255) & Chr$(6) Then
   Call TravelMenu(Index)
   Exit Sub
ElseIf Left$(Msg, 2) = Chr$(255) & Chr$(7) Then
   Call Travel(Index, Mid$(Msg, 3))
   Exit Sub
ElseIf Left$(Msg, 2) = Chr$(254) & Chr$(2) Then
   Call PawnShopMenu(Index)
   Exit Sub
ElseIf Left(Msg, 2) = Chr$(254) & Chr$(3) Then
   Call PlayerItemInfo(Index, Mid$(Msg, 3))
   Exit Sub
ElseIf Left(Msg, 2) = Chr$(254) & Chr$(4) Then
   Call ShopItemInfo(Index, Mid$(Msg, 3))
   Exit Sub
ElseIf Left(Msg, 2) = Chr$(254) & Chr$(5) Then
   Call BuyItem(Index, Mid$(Msg, 3))
   Exit Sub
ElseIf Left(Msg, 2) = Chr$(254) & Chr$(6) Then
   Call SellItem(Index, Mid$(Msg, 3))
   Exit Sub
ElseIf LCase$(Left$(Msg, 7)) = "/addmob" Then
   Call AddNpcGM(Index, Trim$(Mid$(Msg, 8)))
   Exit Sub
ElseIf LCase$(Left$(Msg, 3)) = "buy" Then
   Call BuyDrugMenu(Index, Trim$(Mid$(Msg, 4)))
   Exit Sub
ElseIf Left$(Msg, 2) = Chr$(254) & Chr$(7) Then
   Call DrugDealItemInfo(Index, Trim$(Mid$(Msg, 3)))
   Exit Sub
ElseIf Left$(Msg, 2) = Chr$(253) & Chr$(2) Then
   Call BuyNpcDrug(Index, Mid$(Msg, 3))
   Exit Sub
ElseIf LCase(Left$(Msg, 4)) = "sell" Then
   Call DruggieMenu(Index, Trim$(Mid$(Msg, 5)))
   Exit Sub
ElseIf Left$(Msg, 2) = Chr$(253) & Chr$(3) Then
   Call DruggieItemInfo(Index, Trim$(Mid$(Msg, 3)))
   Exit Sub
ElseIf LCase$(Msg) = "/listnpcs" Then
   Call ListNPCs(Index)
   Exit Sub
ElseIf Left$(Msg, 2) = Chr$(253) & Chr$(4) Then
   Call SellDruggieItem(Index, Trim$(Mid$(Msg, 3)))
   Exit Sub
ElseIf Left$(Msg, 2) = Chr$(253) & Chr$(5) Then
   Call SendMap(Index)
   Exit Sub
ElseIf LCase$(Left$(Msg, 3)) = "aim" Then
   Call Aim(Index, Trim$(Mid$(Msg, 4)))
   Exit Sub
ElseIf LCase$(Left$(Msg, 5)) = "/goto" Then
   Call GotoPlayer(Index, Trim$(Mid$(Msg, 6)))
   Exit Sub
ElseIf LCase$(Msg) = "punch" Then
   Call Punch(Index)
   Exit Sub
ElseIf LCase$(Msg) = "fire" Then
   Call Fire(Index)
   Exit Sub
ElseIf LCase$(Msg) = "strike" Then
   Call Strike(Index)
   Exit Sub
ElseIf LCase$(Msg) = "hide" Then
   Call Hide(Index)
   Exit Sub
ElseIf LCase$(Left$(Msg, 4)) = "flee" Then
   Call Flee(Index, Trim$(Mid$(Msg, 5)))
   Exit Sub
ElseIf LCase$(Msg) = "healinfo" Then
   Call HealInfo(Index)
   Exit Sub
ElseIf LCase$(Left$(Msg, 6)) = "healme" Then
   Call HealMe(Index, Trim$(Mid$(Msg, 7)))
   Exit Sub
ElseIf LCase$(Msg) = "skills" Then
   Call ShowSkills(Index)
   Exit Sub
ElseIf LCase$(Left$(Msg, 7)) = "deposit" Then
   Call Deposit(Index, Trim$(Mid$(Msg, 8)))
   Exit Sub
ElseIf LCase$(Left$(Msg, 8)) = "withdraw" Then
   Call Withdraw(Index, Trim$(Mid$(Msg, 9)))
   Exit Sub
ElseIf LCase$(Left$(Msg, 5)) = "track" Then
   Call TrackPlayer(Index, Trim$(Mid$(Msg, 6)))
   Exit Sub
Else
   frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

End Sub
Public Sub North(Index As Integer)

'Check to see if player is hiding
Call NoHiding(Index)

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

'Move the player north
If City(User(Index).Location).North <> -1 Then
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " the " & User(Index).Rank & " stroll off to the north." & vbCrLf & vbCrLf & Chr$(0))
   User(Index).Location = City(User(Index).Location).North
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " the " & User(Index).Rank & " stroll in from south." & vbCrLf & vbCrLf & Chr$(0))
   Call ShowCity(Index)
ElseIf City(User(Index).Location).North = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You find no exit leading to the North." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
End If

End Sub
Public Sub East(Index As Integer)

'Check to see if player is hiding
Call NoHiding(Index)

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

'Move the player east
If City(User(Index).Location).East <> -1 Then
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " the " & User(Index).Rank & " stroll off to the east." & vbCrLf & vbCrLf & Chr$(0))
   User(Index).Location = City(User(Index).Location).East
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " the " & User(Index).Rank & " stroll in from west." & vbCrLf & vbCrLf & Chr$(0))
   Call ShowCity(Index)
ElseIf City(User(Index).Location).East = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You find no exit leading to the East." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
End If

End Sub
Public Sub South(Index As Integer)

'Check to see if player is hiding
Call NoHiding(Index)

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

'Move the player south
If City(User(Index).Location).South <> -1 Then
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " the " & User(Index).Rank & " stroll off to the south." & vbCrLf & vbCrLf & Chr$(0))
   User(Index).Location = City(User(Index).Location).South
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " the " & User(Index).Rank & " stroll in from north." & vbCrLf & vbCrLf & Chr$(0))
   Call ShowCity(Index)
ElseIf City(User(Index).Location).South = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You find no exit leading to the South." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
End If

End Sub

Public Sub West(Index As Integer)

'Check to see if player is hiding
Call NoHiding(Index)

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

'Move the player west
If City(User(Index).Location).West <> -1 Then
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " the " & User(Index).Rank & " stroll off to the west." & vbCrLf & vbCrLf & Chr$(0))
   User(Index).Location = City(User(Index).Location).West
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " the " & User(Index).Rank & " stroll in from east." & vbCrLf & vbCrLf & Chr$(0))
   Call ShowCity(Index)
ElseIf City(User(Index).Location).West = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You find no exit leading to the West." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
End If

End Sub
Public Sub ShowWatchers(Index As Integer, Msg As String)
On Error Resume Next
Dim a As Integer 'Counter

'Show message to all players exept Index Player
For a = 1 To MaxUsers
   If User(a).Status = "Playing" And _
      User(a).Location = User(Index).Location And _
      Index <> a Then
         frmMain.wsk(a).SendData Msg
         DoEvents
   End If
Next a

End Sub

Public Sub SendChat(Index As Integer, Msg As String)
On Error Resume Next
Dim a As Integer 'Counter

Call ChatLog(Index, Msg)

'Dont allow player to send a global msg with mute on
If User(Index).Mute = True Then
   frmMain.wsk(Index).SendData Chr$(2) & "You cannot transmit a global message while you're in mute mode." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

'Display global yell message
For a = 1 To MaxUsers
   If User(a).Status = "Playing" And _
      User(a).Mute = False Then
         frmMain.wsk(a).SendData Chr$(2) & "[" & User(Index).UName & " yells] - " & Msg & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
   End If
Next a

End Sub
Public Sub Mute(Index As Integer)

If User(Index).Mute = False Then
   User(Index).Mute = True
   frmMain.wsk(Index).SendData Chr$(2) & "You have selected to mute all global messages.  You can still transmit and recieve private messages while in mute mode." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
ElseIf User(Index).Mute = True Then
   User(Index).Mute = False
   frmMain.wsk(Index).SendData Chr$(2) & "You have selected to recieve all global messages." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
End If

End Sub

Public Sub NoHiding(Index As Integer)

'Check to see if the player is hiding
If User(Index).IsHiding = True Then
   User(Index).IsHiding = False
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " the " & User(Index).Rank & " emerge from the shadows." & vbCrLf & vbCrLf & Chr$(0))
   frmMain.wsk(Index).SendData Chr$(2) & "You are no longer in hiding." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
End If

End Sub

Public Sub AddItemGM(Index As Integer, Msg As String)
Dim a As Integer, b As Integer 'Counters

'Check to make sure player access level is high enough
If User(Index).AccessLevel <> 5 Then
   frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

For a = 0 To UBound(ItemDB)
   If LCase$(Msg) = LCase$(Left$(ItemDB(a).IName, Len(Msg))) Then
      For b = 0 To UBound(City(User(Index).Location).CItem)
         If City(User(Index).Location).CItem(b) = -1 Then
            ReDim Preserve Item(UBound(Item) + 1)
            Item(UBound(Item)) = ItemDB(a)
            Item(UBound(Item)).ItemGUID = "" 'city(user(index).Location).CityGUID
            Item(UBound(Item)).ILocation = User(Index).Location
            Item(UBound(Item)).Decay = GetTickCount()
            City(User(Index).Location).CItem(b) = UBound(Item)
            frmMain.wsk(Index).SendData Chr$(2) & "You toss a " & Item(UBound(Item)).IName & " on the ground." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " toss a " & Item(UBound(Item)).IName & " on the ground." & vbCrLf & vbCrLf & Chr$(0))
            Exit Sub
         ElseIf b = UBound(City(User(Index).Location).CItem) Then
            With City(User(Index).Location)
            ReDim Preserve .CItem(UBound(.CItem) + 1)
            ReDim Preserve Item(UBound(Item) + 1)
            Item(UBound(Item)) = ItemDB(a)
            Item(UBound(Item)).ItemGUID = "" 'City(User(Index).Location).CityGUID
            Item(UBound(Item)).ILocation = User(Index).Location
            Item(UBound(Item)).Decay = GetTickCount()
            .CItem(UBound(.CItem)) = UBound(Item)
            End With
            frmMain.wsk(Index).SendData Chr$(2) & "You toss a " & Item(UBound(Item)).IName & " on the ground." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " toss a " & Item(UBound(Item)).IName & " on the ground." & vbCrLf & vbCrLf & Chr$(0))
            Exit Sub
         End If
      Next b
   End If
Next a

End Sub

Public Sub ListItemGM(Index As Integer)
Dim a As Integer 'Counter
Dim Msg As String 'String
Msg = Chr$(2)

'Check to make sure player access level is high enough
If User(Index).AccessLevel <> 5 Then
   frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

'List Items to GM's Only
For a = 0 To UBound(ItemDB)
   Msg = Msg & ItemDB(a).IName & "     "
Next a

Msg = Msg & vbCrLf & vbCrLf & Chr$(0)
frmMain.wsk(Index).SendData Msg
DoEvents

End Sub

Public Function InventoryFull(Index As Integer) As Boolean
Dim a As Integer 'Counter

'Check to see if a players inventory is full
For a = 0 To 19
   If User(Index).Item(a) = -1 Then
      InventoryFull = False
      Exit Function
   ElseIf a = 19 Then
      InventoryFull = True
      frmMain.wsk(Index).SendData Chr$(2) & "You have no room in your inventory, try selling something." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Exit Function
   End If
Next a

End Function

Public Sub GetItem(Index As Integer, Msg As String)
Dim a As Integer 'Counter
Dim b As Integer 'Counter

'Check for full inventory
If InventoryFull(Index) = True Then
   Exit Sub
End If

'Pick the item up off the ground
For a = 0 To UBound(City(User(Index).Location).CItem)
   If City(User(Index).Location).CItem(a) <> -1 Then
   If LCase$(Msg) = LCase$(Left$(Item(City(User(Index).Location).CItem(a)).IName, Len(Msg))) Then
      For b = 0 To 19
         If User(Index).Item(b) = -1 Then
            User(Index).Item(b) = City(User(Index).Location).CItem(a)
            City(User(Index).Location).CItem(a) = -1
            Item(User(Index).Item(b)).ItemGUID = User(Index).UserGUID
            Item(User(Index).Item(b)).OnPlayer = True
            Item(User(Index).Item(b)).Decay = -1
            Item(User(Index).Item(b)).ILocation = -1
            frmMain.wsk(Index).SendData Chr$(2) & "You pick up a " & Item(User(Index).Item(b)).IName & " and put it in your pack." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " pick up a " & Item(User(Index).Item(b)).IName & "." & vbCrLf & vbCrLf & Chr$(0))
            Call UpdateSingleItem(Index, b)
            Exit Sub
         End If
      Next b
   End If
   End If
Next a
           
'Runs if no item in room matches get message
frmMain.wsk(Index).SendData Chr$(2) & "You can't pick up what isn't there." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

           
           
End Sub

Public Sub DropItem(Index As Integer, Msg As String)
Dim a As Integer 'Counter
Dim b As Integer 'Counter

If IsNumeric(Msg) = True Then
   b = Msg
Else
   Exit Sub
End If

If b > 19 Or b < 0 Then
   Exit Sub
End If

If User(Index).Item(b) = -1 Then
   Exit Sub
End If

'Drop the item the players chooses
For a = 0 To UBound(City(User(Index).Location).CItem)
   If City(User(Index).Location).CItem(a) = -1 Then
      
      If User(Index).Item(b) = User(Index).Weapon Then
         User(Index).Weapon = -1
         Call UpdateGearInfo(Index)
      ElseIf User(Index).Item(b) = User(Index).Armor Then
         User(Index).Armor = -1
         Call UpdateGearInfo(Index)
      ElseIf User(Index).Item(b) = User(Index).Ammo Then
         User(Index).Ammo = -1
         Call UpdateGearInfo(Index)
      End If
      
      City(User(Index).Location).CItem(a) = User(Index).Item(b)
      Item(User(Index).Item(b)).OnPlayer = False
      Item(User(Index).Item(b)).Equip = False
      Item(User(Index).Item(b)).Decay = GetTickCount()
      Item(User(Index).Item(b)).ItemGUID = ""
      Item(User(Index).Item(b)).ILocation = User(Index).Location
      User(Index).Item(b) = -1
      frmMain.wsk(Index).SendData Chr$(2) & "You toss a " & Item(City(User(Index).Location).CItem(a)).IName & " on the ground." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " toss a " & Item(City(User(Index).Location).CItem(a)).IName & " on the ground." & vbCrLf & vbCrLf & Chr$(0))
      Call UpdateSingleItem(Index, b)
      Exit Sub
   ElseIf a = UBound(City(User(Index).Location).CItem) Then
      With City(User(Index).Location)
      ReDim Preserve .CItem(UBound(.CItem) + 1)
      
      If User(Index).Item(b) = User(Index).Weapon Then
         User(Index).Weapon = -1
         Call UpdateGearInfo(Index)
      ElseIf User(Index).Item(b) = User(Index).Armor Then
         User(Index).Armor = -1
         Call UpdateGearInfo(Index)
      ElseIf User(Index).Item(b) = User(Index).Ammo Then
         User(Index).Ammo = -1
         Call UpdateGearInfo(Index)
      End If

      .CItem(UBound(.CItem)) = User(Index).Item(b)
      Item(User(Index).Item(b)).OnPlayer = False
      Item(User(Index).Item(b)).Equip = False
      Item(User(Index).Item(b)).Decay = GetTickCount()
      Item(User(Index).Item(b)).ItemGUID = ""
      Item(User(Index).Item(b)).ILocation = User(Index).Location
      User(Index).Item(b) = -1
      frmMain.wsk(Index).SendData Chr$(2) & "You toss a " & Item(City(User(Index).Location).CItem(UBound(.CItem))).IName & " on the ground." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " toss a " & Item(City(User(Index).Location).CItem(UBound(.CItem))).IName & " on the ground." & vbCrLf & vbCrLf & Chr$(0))
      Call UpdateSingleItem(Index, b)
      End With
      Exit Sub
   End If
Next a

End Sub
Public Sub ExamineItem(Index As Integer, Msg As String)
Dim a As Integer

If IsNumeric(Msg) = True Then
   a = Msg
Else
   Exit Sub
End If

If a < 0 Or a > 19 Then
   Exit Sub
End If

If User(Index).Item(a) = -1 Then
   Exit Sub
End If

frmMain.wsk(Index).SendData Chr$(2) & Item(User(Index).Item(a)).IDesc & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub
Public Sub EquipItem(Index As Integer, Msg As String)
Dim a As Integer 'Counter
Dim b As Integer 'Counter

'Make sure the index is correct
If IsNumeric(Msg) = True Then
   b = Msg
Else
   Exit Sub
End If

If Msg < 0 Or Msg > 19 Then
   Exit Sub
End If

If User(Index).Item(b) = -1 Then
   Exit Sub
End If

If User(Index).Item(b) <> -1 Then
   If Item(User(Index).Item(b)).IType = C_Gun Or _
      Item(User(Index).Item(b)).IType = C_Armor Or _
      Item(User(Index).Item(b)).IType = C_Ammo Or _
      Item(User(Index).Item(b)).IType = C_Melee Then
      For a = 0 To 19
      If User(Index).Item(a) <> -1 Then
         If Item(User(Index).Item(a)).IType = _
            Item(User(Index).Item(b)).IType Then
            Item(User(Index).Item(a)).Equip = False
         End If
         'Make sure guns/melee are unquipted if opposite
         'No dual weapon weilding
            If Item(User(Index).Item(b)).IType = C_Melee And _
               Item(User(Index).Item(a)).IType = C_Gun And _
               Item(User(Index).Item(a)).Equip = True Then
                  Item(User(Index).Item(a)).Equip = False
            ElseIf Item(User(Index).Item(b)).IType = C_Gun And _
               Item(User(Index).Item(a)).IType = C_Melee And _
               Item(User(Index).Item(a)).Equip = True Then
                  Item(User(Index).Item(a)).Equip = False
            End If
      End If
      Next a
         If Item(User(Index).Item(b)).IType = C_Gun Then
            User(Index).Weapon = User(Index).Item(b)
         ElseIf Item(User(Index).Item(b)).IType = C_Armor Then
            User(Index).Armor = User(Index).Item(b)
         ElseIf Item(User(Index).Item(b)).IType = C_Ammo Then
            User(Index).Ammo = User(Index).Item(b)
         ElseIf Item(User(Index).Item(b)).IType = C_Melee Then
            User(Index).Weapon = User(Index).Item(b)
         End If
   Item(User(Index).Item(b)).Equip = True
   Call FullInventoryUpdate(Index)
   Call UpdateGearInfo(Index)
   End If
End If

End Sub
Public Sub UpdateSingleItem(Index As Integer, ItemNo As Integer)
Dim Msg As String
Msg = Chr$(7) & ItemNo & Chr$(1)

   If User(Index).Item(ItemNo) = -1 Then
      Msg = Msg & "<Empty>"
   ElseIf User(Index).Item(ItemNo) <> -1 Then
            
      'Check to see if item is multiple
      If Item(User(Index).Item(ItemNo)).IType = C_Ammo And _
         Item(User(Index).Item(ItemNo)).Amount > 0 And _
         Item(User(Index).Item(ItemNo)).Multiple = True Then
            Msg = Msg & "(" & Item(User(Index).Item(ItemNo)).Amount & ") "
      End If
            
      'Check to see if the item is equipted
      If Item(User(Index).Item(ItemNo)).Equip = True Then
            Msg = Msg & "<E> "
      End If
      
      'Add Item Name
      Msg = Msg & "<" & Item(User(Index).Item(ItemNo)).IName & ">"
   End If

Msg = Msg & Chr$(0)
frmMain.wsk(Index).SendData Msg
DoEvents

End Sub

Public Sub UnEquip(Index As Integer, Msg As String)
Dim a As Integer

If IsNumeric(Msg) = True Then
   a = Msg
Else
   Exit Sub
End If

If User(Index).Item(a) = -1 Then
   Exit Sub
End If

If Item(User(Index).Item(a)).Equip = False Then
   Exit Sub
End If

'Remove Gear Index
If User(Index).Item(a) <> -1 Then
   If Item(User(Index).Item(a)).IType = C_Melee Or _
      Item(User(Index).Item(a)).IType = C_Gun Then
         User(Index).Weapon = -1
   ElseIf Item(User(Index).Item(a)).IType = C_Armor Then
         User(Index).Armor = -1
   ElseIf Item(User(Index).Item(a)).IType = C_Ammo Then
         User(Index).Ammo = -1
   End If
Call UpdateGearInfo(Index)
End If

'UnEquip users item
If User(Index).Item(a) <> -1 Then
   Item(User(Index).Item(a)).Equip = False
   Call UpdateSingleItem(Index, a)
End If

End Sub

Public Sub UseItem(Index As Integer, Msg As String)
Dim a As Integer

If IsNumeric(Msg) = True Then
   a = Msg
ElseIf IsNumeric(Msg) = False Then
   Exit Sub
End If

If a < LBound(User(Index).Item) Or _
   a > UBound(User(Index).Item) Then
      Exit Sub
End If

If User(Index).Item(a) = -1 Then
   Exit Sub
End If

Select Case Item(User(Index).Item(a)).IType
   Case C_Phone
      Call UsePhone(Index)
      Exit Sub
   Case C_MedStick
      Call UseMedStick(Index, a)
      Exit Sub
End Select


frmMain.wsk(Index).SendData Chr$(2) & "You don't see any specific way you could use this item." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub
Public Sub TravelMenu(Index As Integer)
Dim Msg As String

'Check to see if the player is at an airport first
If City(User(Index).Location).AirPort = False Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to find an Airport before you can travel anywhere." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

Msg = Chr$(255) & Chr$(2) & User(Index).CurrTown & Chr$(1)
Msg = Msg & NY_Price & Chr$(1) & LA_Price & Chr$(1) & _
HO_Price & Chr$(1) & MI_Price & Chr$(1) & _
CH_Price & Chr$(1) & NJ_Price & Chr$(1) & Chr$(0)

frmMain.wsk(Index).SendData Msg
DoEvents

End Sub
Public Sub Travel(Index As Integer, Msg As String)

'Check to see if the player is in combat
If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

Select Case LCase$(Msg)
   
   'Fly to New York
   Case "new york"
      If User(Index).Cash < NY_Price Then
         frmMain.wsk(Index).SendData Chr$(2) & "Start flapping your arms buddy, you don't have the cash to do fly anywhere." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf User(Index).Cash >= NY_Price Then
         User(Index).Cash = User(Index).Cash - NY_Price
         Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " jump on a plane and jet out of the city." & vbCrLf & vbCrLf & Chr$(0))
         User(Index).Location = NY_Location
         User(Index).CurrTown = City(User(Index).Location).CName
         Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " just arrived on a plane from another city." & vbCrLf & vbCrLf & Chr$(0))
         frmMain.wsk(Index).SendData Chr$(2) & "You jump on a plane and jet to New York." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Call UpdateGeneralInfo(Index)
         Call ShowCity(Index)
         Exit Sub
      End If
   
   'Fly to Los Angeles
   Case "los angeles"
      If User(Index).Cash < LA_Price Then
         frmMain.wsk(Index).SendData Chr$(2) & "Start flapping your arms buddy, you don't have the cash to do fly anywhere." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf User(Index).Cash >= LA_Price Then
         User(Index).Cash = User(Index).Cash - LA_Price
         Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " jump on a plane and jet out of the city." & vbCrLf & vbCrLf & Chr$(0))
         User(Index).Location = LA_Location
         User(Index).CurrTown = City(User(Index).Location).CName
         Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " just arrived on a plane from another city." & vbCrLf & vbCrLf & Chr$(0))
         frmMain.wsk(Index).SendData Chr$(2) & "You jump on a plane and jet to Los Angeles." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Call UpdateGeneralInfo(Index)
         Call ShowCity(Index)
         Exit Sub
      End If

   'Fly to Houston
   Case "houston"
      If User(Index).Cash < HO_Price Then
         frmMain.wsk(Index).SendData Chr$(2) & "Start flapping your arms buddy, you don't have the cash to do fly anywhere." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf User(Index).Cash >= HO_Price Then
         User(Index).Cash = User(Index).Cash - HO_Price
         Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " jump on a plane and jet out of the city." & vbCrLf & vbCrLf & Chr$(0))
         User(Index).Location = HO_Location
         User(Index).CurrTown = City(User(Index).Location).CName
         Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " just arrived on a plane from another city." & vbCrLf & vbCrLf & Chr$(0))
         frmMain.wsk(Index).SendData Chr$(2) & "You jump on a plane and jet to Houston." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Call UpdateGeneralInfo(Index)
         Call ShowCity(Index)
         Exit Sub
      End If

   'Fly to Miami
   Case "miami"
      If User(Index).Cash < MI_Price Then
         frmMain.wsk(Index).SendData Chr$(2) & "Start flapping your arms buddy, you don't have the cash to do fly anywhere." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf User(Index).Cash >= MI_Price Then
         User(Index).Cash = User(Index).Cash - MI_Price
         Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " jump on a plane and jet out of the city." & vbCrLf & vbCrLf & Chr$(0))
         User(Index).Location = MI_Location
         User(Index).CurrTown = City(User(Index).Location).CName
         Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " just arrived on a plane from another city." & vbCrLf & vbCrLf & Chr$(0))
         frmMain.wsk(Index).SendData Chr$(2) & "You jump on a plane and jet to Miami." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Call UpdateGeneralInfo(Index)
         Call ShowCity(Index)
         Exit Sub
      End If

   'Fly to Chicago
   Case "chicago"
      If User(Index).Cash < CH_Price Then
         frmMain.wsk(Index).SendData Chr$(2) & "Start flapping your arms buddy, you don't have the cash to do fly anywhere." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf User(Index).Cash >= CH_Price Then
         User(Index).Cash = User(Index).Cash - CH_Price
         Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " jump on a plane and jet out of the city." & vbCrLf & vbCrLf & Chr$(0))
         User(Index).Location = CH_Location
         User(Index).CurrTown = City(User(Index).Location).CName
         Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " just arrived on a plane from another city." & vbCrLf & vbCrLf & Chr$(0))
         frmMain.wsk(Index).SendData Chr$(2) & "You jump on a plane and jet to Chicago." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Call UpdateGeneralInfo(Index)
         Call ShowCity(Index)
         Exit Sub
      End If

   'Fly to New Jersey
   Case "new jersey"
      If User(Index).Cash < NJ_Price Then
         frmMain.wsk(Index).SendData Chr$(2) & "Start flapping your arms buddy, you don't have the cash to do fly anywhere." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf User(Index).Cash >= NJ_Price Then
         User(Index).Cash = User(Index).Cash - NJ_Price
         Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " jump on a plane and jet out of the city." & vbCrLf & vbCrLf & Chr$(0))
         User(Index).Location = NJ_Location
         User(Index).CurrTown = City(User(Index).Location).CName
         Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " just arrived on a plane from another city." & vbCrLf & vbCrLf & Chr$(0))
         frmMain.wsk(Index).SendData Chr$(2) & "You jump on a plane and jet to New Jersey." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Call UpdateGeneralInfo(Index)
         Call ShowCity(Index)
         Exit Sub
      End If

'On Data Error Run This
frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Select

End Sub
Public Function PlayerIsTarget(Index As Integer) As Boolean
Dim a As Integer 'Counter

'Check to see if the index user is a player/npc target
For a = 1 To MaxUsers
   If User(a).Status = "Playing" And _
         Index <> a And _
         User(a).TargetNum = Index And _
         User(a).TargetGUID = User(Index).UserGUID And _
         User(a).Location = User(Index).Location Then
         PlayerIsTarget = True
         frmMain.wsk(Index).SendData Chr$(2) & User(a).UName & " has taken aim on you, the only way you execute this action is to kill " & User(a).UName & " or flee the area.  If you choose to flee, you will lose a fair amount of rank and possibly drop an item or two in the scramble to get away." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Function
   End If
Next a

For a = 0 To 9
   If City(User(Index).Location).CNpc(a) <> -1 Then
      If Npc(City(User(Index).Location).CNpc(a)).NTargetID = Index And _
         Npc(City(User(Index).Location).CNpc(a)).NTargetGUID = User(Index).UserGUID And _
         Npc(City(User(Index).Location).CNpc(a)).NLocation = User(Index).Location Then
         Npc(City(User(Index).Location).CNpc(a)).CanMove = GetTickCount()
         PlayerIsTarget = True
         frmMain.wsk(Index).SendData Chr$(2) & Npc(City(User(Index).Location).CNpc(a)).NName & " has taken aim on you, the only way you can execute this actions is to kill " & Npc(City(User(Index).Location).CNpc(a)).NName & " or flee the area.  If you choose to flee, you will lose a fair amount of rank and possibly drop an item or two in the scramble to get away." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Function
      End If
   End If
Next a
      
PlayerIsTarget = False
User(Index).TargetNum = -1
User(Index).TargetGUID = ""

End Function
Public Sub PawnShopMenu(Index As Integer)
Dim a As Integer 'Counter
Dim b As Integer 'Counter
Dim Msg As String 'String

If City(User(Index).Location).PawnShop = False Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to find a pawn shop first." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

Call NoHiding(Index)

Msg = Chr$(255) & Chr$(3)

For a = 0 To UBound(ItemDB)
   If ItemDB(a).ForSale = True Then
      Msg = Msg & ItemDB(a).IName & Chr$(1)
   End If
Next a

Msg = Msg & Chr$(2)

For a = 0 To 19
   If User(Index).Item(a) <> -1 Then
      Msg = Msg & Item(User(Index).Item(a)).IName & Chr$(1)
   ElseIf User(Index).Item(a) = -1 Then
      Msg = Msg & "<Empty>" & Chr$(1)
   End If
Next a

frmMain.wsk(Index).SendData Msg & Chr$(0)
DoEvents

frmMain.wsk(Index).SendData Chr$(255) & Chr$(4) & User(Index).Cash & Chr$(1) & User(Index).Rank & Chr$(1) & Chr$(0)
DoEvents

End Sub

Public Sub ShopItemInfo(Index As Integer, Msg As String)
Dim a As Integer
Dim YesOrNo As String

'Check to see if the message is a number
If IsNumeric(Msg) = False Then
   Exit Sub
End If

a = Msg

If a < LBound(SlotID) Or a > UBound(SlotID) Then
   Exit Sub
End If

'Check to see if the item fits in the scope
'If User(Index).Item(a) = -1 Then
'   Exit Sub
'End If

'check to see if the item is dope
'If Item(User(Index).Item(a)).IType = C_Dope Then
'   frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & "Get Lost" & Chr$(1) & "Get Lost" & Chr$(1) & "Get Lost" & Chr$(1) & Chr$(0)
'   DoEvents
'   Exit Sub
'End If

'Send the item info to the Pawn Shop Menu
If User(Index).Reputation < ItemDB(SlotID(a)).CanBuy Then
   YesOrNo = "No"
ElseIf User(Index).Reputation >= ItemDB(SlotID(a)).CanBuy Then
   YesOrNo = "Yes"
End If

frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & ItemDB(SlotID(a)).Price & Chr$(1) & YesOrNo & Chr$(1) & ItemDB(SlotID(a)).IName & Chr$(1) & Chr$(0)
DoEvents

End Sub
Public Sub PlayerItemInfo(Index As Integer, Msg As String)
Dim a As Integer

If IsNumeric(Msg) = False Then
   Exit Sub
End If

a = Msg

If Msg < 0 Or Msg > 19 Then
   Exit Sub
End If

If User(Index).Item(a) = -1 Then
   Exit Sub
End If

If Item(User(Index).Item(a)).IType = C_Ammo And _
   Item(User(Index).Item(a)).Amount <> 10 Then
   Exit Sub
End If

frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & Int(Item(User(Index).Item(a)).Price / 2) & Chr$(1) & "N/A" & Chr$(1) & Item(User(Index).Item(a)).IName & Chr$(1) & Chr$(0)
DoEvents

End Sub

Public Sub BuyItem(Index As Integer, Msg As String)
Dim a As Integer 'Counter
Dim b As Integer
Dim c As Integer
Dim MsgX As String

For a = 0 To 19
   If User(Index).Item(a) = -1 Then
      Exit For
   ElseIf a = 19 Then
      frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & "Inventory Full" & Chr$(1) & "Inventory Full" & Chr$(1) & "Inventory Full" & Chr$(1) & Chr$(0)
      DoEvents
      Exit Sub
   End If
Next a

If IsNumeric(Msg) = False Then
   Exit Sub
End If

b = Msg

If User(Index).Cash < ItemDB(SlotID(b)).Price Then
      frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & "Lack Of Cash" & Chr$(1) & "Lack Of Cash" & Chr$(1) & "Lack Of Cash" & Chr$(1) & Chr$(0)
      DoEvents
      Exit Sub
End If

If User(Index).Reputation < ItemDB(SlotID(b)).CanBuy Then
      frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & "Lack Of Rank" & Chr$(1) & "Lack Of Rank" & Chr$(1) & "Lack Of Rank" & Chr$(1) & Chr$(0)
      DoEvents
      Exit Sub
End If

User(Index).Cash = User(Index).Cash - ItemDB(SlotID(b)).Price
ReDim Preserve Item(UBound(Item) + 1)
Item(UBound(Item)) = ItemDB(SlotID(b))
Item(UBound(Item)).Decay = -1
Item(UBound(Item)).Equip = False
Item(UBound(Item)).ForSale = False
Item(UBound(Item)).ILocation = -1
Item(UBound(Item)).ItemGUID = User(Index).UserGUID
Item(UBound(Item)).OnPlayer = True
User(Index).Item(a) = UBound(Item)
Call UpdateSingleItem(Index, a)
frmMain.wsk(Index).SendData Chr$(255) & Chr$(4) & User(Index).Cash & Chr$(1) & User(Index).Rank & Chr$(1) & Chr$(0)
DoEvents

MsgX = Chr$(255) & Chr$(6)

For a = 0 To 19
   If User(Index).Item(a) <> -1 Then
      MsgX = MsgX & Item(User(Index).Item(a)).IName & Chr$(1)
   ElseIf User(Index).Item(a) = -1 Then
      MsgX = MsgX & "<Empty>" & Chr$(1)
   End If
Next a

frmMain.wsk(Index).SendData MsgX & Chr$(0)
DoEvents

Call UpdateGeneralInfo(Index)

End Sub
Public Sub SellItem(Index As Integer, Msg As String)
On Error GoTo BadItemNo
Dim a As Integer
Dim MsgX As String

If IsNumeric(Msg) = False Then
   Exit Sub
End If

a = Msg

If a < 0 Or a > 19 Then
   Exit Sub
End If

If Item(User(Index).Item(a)).ItemGUID <> User(Index).UserGUID Then
   Exit Sub
End If

If Item(User(Index).Item(Msg)).IType = C_Dope Then
   frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & "Get Lost" & Chr$(1) & "Get Lost" & Chr$(1) & "Get Lost" & Chr$(1) & Chr$(0)
   DoEvents
   Exit Sub
End If

If Item(User(Index).Item(a)).Equip = True Then
   frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & "Item Equipted" & Chr$(1) & "Item Equipted" & Chr$(1) & "Item Equipted" & Chr$(1) & Chr$(0)
   DoEvents
   Exit Sub
End If

If Item(User(Index).Item(Msg)).IType = C_Ammo And _
   Item(User(Index).Item(Msg)).Amount < 10 Then
   frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & "Used Ammo" & Chr$(1) & "Used Ammo" & Chr$(1) & "Used Ammo" & Chr$(1) & Chr$(0)
   DoEvents
   Exit Sub
End If

User(Index).Cash = User(Index).Cash + Int((Item(User(Index).Item(a)).Price / 2))
Call ResetItem(User(Index).Item(a))
User(Index).Item(a) = -1
Call UpdateSingleItem(Index, a)

MsgX = Chr$(255) & Chr$(6)

For a = 0 To 19
   If User(Index).Item(a) <> -1 Then
      MsgX = MsgX & Item(User(Index).Item(a)).IName & Chr$(1)
   ElseIf User(Index).Item(a) = -1 Then
      MsgX = MsgX & "<Empty>" & Chr$(1)
   End If
Next a

frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & "Sold" & Chr$(1) & "Sold" & Chr$(1) & "Sold" & Chr$(1) & Chr$(0)
DoEvents

frmMain.wsk(Index).SendData MsgX & Chr$(0)
DoEvents

frmMain.wsk(Index).SendData Chr$(255) & Chr$(4) & User(Index).Cash & Chr$(1) & User(Index).Rank & Chr$(1) & Chr$(0)
DoEvents

Call UpdateGeneralInfo(Index)
Exit Sub

BadItemNo:
Dim ff As Integer
ff = FreeFile
Open App.Path & "\error.log" For Append As ff
Print #ff, "[BOE]"
Print #ff, "Bad Item Number In Sell Menu"
Print #ff, User(Index).UName & " | " & Msg
Print #ff, "[EOE]"
Close ff

End Sub
Public Function AddItem(ItemNo As Integer) As Integer
Dim a As Integer

'This adds items to NPCs who are just spawned

For a = 0 To UBound(Item)
   If Item(a).IName = "" And _
      Item(a).ItemGUID = "" Then
         Item(a) = ItemDB(ItemNo)
         Item(a).OnPlayer = True
         AddItem = a
         Exit Function
   ElseIf a = UBound(Item) Then
      ReDim Preserve Item(UBound(Item) + 1)
         Item(UBound(Item)) = ItemDB(ItemNo)
         Item(UBound(Item)).OnPlayer = True
         AddItem = UBound(Item)
         Exit Function
   End If
Next a

End Function

Public Sub AddNpcGM(Index As Integer, NpcType As String)
Dim a As Integer

'Check to make sure player access level is high enough
If User(Index).AccessLevel <> 5 Then
   frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If IsNumeric(NpcType) = True Then
   a = NpcType
ElseIf IsNumeric(NpcType) = False Then
   Exit Sub
End If

Call AddNpc(a, User(Index).Location)

End Sub

Public Sub BuyDrugMenu(Index As Integer, NpcName As String)
Dim a As Integer
Dim b As Integer
Dim Msg As String

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

Call NoHiding(Index)

'Clear users dealer tag
User(Index).NpcTrade = -1

If NpcName = "" Then
   Exit Sub
End If

If InventoryFull(Index) = True Then
   Exit Sub
End If

'Check to see if NPC Name is in room
For a = 0 To 9
   If City(User(Index).Location).CNpc(a) <> -1 Then
      If LCase$(NpcName) = LCase$(Npc(City(User(Index).Location).CNpc(a)).NName) And _
         Npc(City(User(Index).Location).CNpc(a)).NpcType = N_Dealer Then
         Exit For
      End If
   ElseIf a = 9 Then
      frmMain.wsk(Index).SendData Chr$(2) & "There is no dealer by that name here with you." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Exit Sub
   End If
Next a

'Set users NPC trading number
User(Index).NpcTrade = City(User(Index).Location).CNpc(a)
'Stop npc from walking out of room for 2 minutes
Npc(City(User(Index).Location).CNpc(a)).CanMove = GetTickCount()

'Send npc's inventory to user
Msg = Chr$(254) & Chr$(4)
For b = 0 To 19
      If Npc(City(User(Index).Location).CNpc(a)).NItem(b) = -1 Then
         Msg = Msg & "<Empty>" & Chr$(1)
      ElseIf Npc(City(User(Index).Location).CNpc(a)).NItem(b) <> -1 Then
         Msg = Msg & Item(Npc(City(User(Index).Location).CNpc(a)).NItem(b)).IName & Chr$(1)
      End If
Next b

Msg = Msg & Chr$(0)

frmMain.wsk(Index).SendData Msg
DoEvents
      
frmMain.wsk(Index).SendData Chr$(252) & Chr$(4) & City(User(Index).Location).Compass & Chr$(0)
DoEvents
   
End Sub
Public Sub DrugDealItemInfo(Index As Integer, Msg As String)
Dim a As Integer
Dim b As Single

If IsNumeric(Msg) = False Then
   Exit Sub
ElseIf IsNumeric(Msg) = True Then
   a = Msg
End If

If a < 0 Or a > 19 Then
   Exit Sub
End If

If Npc(User(Index).NpcTrade).NLocation <> User(Index).Location Then
   frmMain.wsk(Index).SendData Chr$(254) & Chr$(5) & Chr$(0)
   DoEvents
   frmMain.wsk(Index).SendData Chr$(2) & "That dealer has left the area." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   User(Index).NpcTrade = -1
   Exit Sub
End If

If Npc(User(Index).NpcTrade).NItem(Msg) = -1 Then
   frmMain.wsk(Index).SendData Chr$(254) & Chr$(7) & "Man, there is nothing in that pocket you can buy." & Chr$(0)
   DoEvents
   Call UpdateNPCInventory(Index)
   Exit Sub
End If

b = Item(Npc(User(Index).NpcTrade).NItem(a)).Price - (Item(Npc(User(Index).NpcTrade).NItem(a)).Price * 0.06)
b = Int(b)

frmMain.wsk(Index).SendData Chr$(254) & Chr$(6) & Item(Npc(User(Index).NpcTrade).NItem(Msg)).IName & Chr$(1) & b & Chr$(1) & User(Index).Cash & Chr$(1) & Chr$(0)
DoEvents

End Sub
Public Sub UpdateNPCInventory(Index As Integer)
Dim a As Integer
Dim Msg As String
Msg = Chr$(253) & Chr$(2)

For a = 0 To 19
   If Npc(User(Index).NpcTrade).NItem(a) = -1 Then
      Msg = Msg & "<Empty>" & Chr$(1)
   ElseIf Npc(User(Index).NpcTrade).NItem(a) <> -1 Then
      Msg = Msg & Item(Npc(User(Index).NpcTrade).NItem(a)).IName & Chr$(1)
   End If
Next a

frmMain.wsk(Index).SendData Msg & Chr$(0)
DoEvents

End Sub

Public Sub BuyNpcDrug(Index As Integer, Msg As String)
Dim a As Integer
Dim b As Integer
Dim c As Single

'Make sure the item is Good
If IsNumeric(Msg) = False Then
   Exit Sub
ElseIf IsNumeric(Msg) = True Then
   a = Msg
End If

'Make sure user has room in inventory
For b = 0 To 19
   If User(Index).Item(b) = -1 Then
      Exit For
   ElseIf b = 19 Then
      frmMain.wsk(Index).SendData Chr$(254) & Chr$(7) & "You ain't got the room man, try selling something." & Chr$(0)
      DoEvents
      Exit Sub
   End If
Next b

'Make sure the NPC is still in the same room as the player
If Npc(User(Index).NpcTrade).NLocation <> User(Index).Location Then
   frmMain.wsk(Index).SendData Chr$(254) & Chr$(5) & Chr$(0)
   DoEvents
   frmMain.wsk(Index).SendData Chr$(2) & "That dealer has left the area." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   User(Index).NpcTrade = -1
   Exit Sub
End If

'Make sure the dealer hasn't sold the item to another player
If Npc(User(Index).NpcTrade).NItem(Msg) = -1 Then
   frmMain.wsk(Index).SendData Chr$(254) & Chr$(7) & "Man, there is nothing in that pocket you can buy." & Chr$(0)
   DoEvents
   Call UpdateNPCInventory(Index)
   Exit Sub
End If

c = Item(Npc(User(Index).NpcTrade).NItem(Msg)).Price - (Item(Npc(User(Index).NpcTrade).NItem(Msg)).Price * 0.06)
c = Int(c)

If User(Index).Cash >= c Then
   User(Index).Cash = User(Index).Cash - c
   User(Index).Reputation = User(Index).Reputation + 1
   Call SetRank(Index)
   Call UpdateGeneralInfo(Index)
   User(Index).Item(b) = Npc(User(Index).NpcTrade).NItem(Msg)
   Npc(User(Index).NpcTrade).NItem(Msg) = -1
   Item(User(Index).Item(b)).ItemGUID = User(Index).UserGUID
   Call UpdateNPCInventory(Index)
   Call UpdateSingleItem(Index, b)
   frmMain.wsk(Index).SendData Chr$(254) & Chr$(7) & "You got it..." & Chr$(0)
   DoEvents
   frmMain.wsk(Index).SendData Chr$(254) & Chr$(6) & "Sold" & Chr$(1) & "Sold" & Chr$(1) & User(Index).Cash & Chr$(1) & Chr$(0)
   DoEvents
ElseIf User(Index).Cash < c Then
   frmMain.wsk(Index).SendData Chr$(254) & Chr$(7) & "You ain't got the cash to buy that dope from me fool." & Chr$(0)
   DoEvents
End If

End Sub

Public Sub DruggieMenu(Index As Integer, NpcName As String)
Dim a As Integer
Dim b As Integer
Dim Msg As String

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

Call NoHiding(Index)

'Clear users druggies tag
User(Index).NpcTrade = -1

If NpcName = "" Then
   Exit Sub
End If

Call NoHiding(Index)

'Check to see if NPC Name is in room
For a = 0 To 9
   If City(User(Index).Location).CNpc(a) <> -1 Then
      If LCase$(NpcName) = LCase$(Npc(City(User(Index).Location).CNpc(a)).NName) And _
         Npc(City(User(Index).Location).CNpc(a)).NpcType = N_Druggie Then
         Exit For
      End If
   ElseIf a = 9 Then
      frmMain.wsk(Index).SendData Chr$(2) & "There is no druggie by that name here with you." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Exit Sub
   End If
Next a

'Set users NPC trading number
User(Index).NpcTrade = City(User(Index).Location).CNpc(a)
'Stop npc from walking out of room for 2 minutes
Npc(City(User(Index).Location).CNpc(a)).CanMove = GetTickCount()

Msg = Chr$(253) & Chr$(3)

For b = 0 To 19
   If User(Index).Item(b) = -1 Then
      Msg = Msg & "<Empty>" & Chr$(1)
   ElseIf User(Index).Item(b) <> -1 Then
      Msg = Msg & Item(User(Index).Item(b)).IName & Chr$(1)
   End If
Next b

frmMain.wsk(Index).SendData Msg & Chr$(0)
DoEvents

frmMain.wsk(Index).SendData Chr$(252) & Chr$(5) & City(User(Index).Location).Compass & Chr$(0)
DoEvents

End Sub
Public Sub DruggieItemInfo(Index As Integer, Msg As String)
Dim a As Integer
Dim b As Single
Dim c As Integer

If IsNumeric(Msg) = True Then
   a = Msg
ElseIf IsNumeric(Msg) = False Then
   Exit Sub
End If

'Make sure the item is not empty
If User(Index).Item(Msg) = -1 Then
   Exit Sub
End If

If Item(User(Index).Item(Msg)).IType <> C_Dope Then
      frmMain.wsk(Index).SendData Chr$(253) & Chr$(5) & "I only deal in dope, try the pawn shop if you want to unload that junk." & Chr$(0)
      DoEvents
      Exit Sub
End If

'Make sure npc is still in same room as player
If Npc(User(Index).NpcTrade).NLocation <> User(Index).Location Then
   frmMain.wsk(Index).SendData Chr$(253) & Chr$(4) & Chr$(0)
   DoEvents
   frmMain.wsk(Index).SendData Chr$(2) & "That druggie has left the area." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   User(Index).NpcTrade = -1
   Exit Sub
End If

'Make sure npc has room for the item
For c = 0 To 19
   If Npc(User(Index).NpcTrade).NItem(c) = -1 Then
      Exit For
   ElseIf c = 19 Then
      frmMain.wsk(Index).SendData Chr$(253) & Chr$(5) & "I can't afford anything else right now, try me later after I unload some of this dope." & Chr$(0)
      DoEvents
      Exit Sub
   End If
Next c

'Set Price at a % Mark
b = Item(User(Index).Item(Msg)).Price + (Item(User(Index).Item(Msg)).Price * 0.06)
b = Int(b)

'Send item information to the player
frmMain.wsk(Index).SendData Chr$(253) & Chr$(6) & Item(User(Index).Item(Msg)).IName & Chr$(1) & b & Chr$(1) & User(Index).Cash & Chr$(1) & Chr$(0)
DoEvents

frmMain.wsk(Index).SendData Chr$(253) & Chr$(5) & "Well?" & Chr$(0)
DoEvents

End Sub

Public Sub ListNPCs(Index As Integer)
Dim a As Integer
Dim Msg As String


'Check to make sure player access level is high enough
If User(Index).AccessLevel <> 5 Then
   frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

For a = 0 To UBound(Npc)
   If Npc(a).NpcGUID <> "" And _
      Npc(a).NName <> "" And _
      Npc(a).NCity = City(User(Index).Location).CName Then
         Msg = Msg & "   (" & Npc(a).NName & " " & Npc(a).NameTag & " | " & City(Npc(a).NLocation).Compass & ")   "
   End If
Next a

frmMain.wsk(Index).SendData Chr$(2) & Msg & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub

Public Sub SellDruggieItem(Index As Integer, Msg As String)
Dim a As Integer
Dim b As Integer
Dim c As Single
Dim d As Integer
Dim MsgX As String

'Make sure index number is a real number
If IsNumeric(Msg) = True Then
   a = Msg
ElseIf IsNumeric(Msg) = False Then
   Exit Sub
End If

'make sure index number is not out of scope
If a < 0 Or a > 19 Then
   Exit Sub
End If

'make sure the item exists
If User(Index).Item(a) = -1 Then
   Exit Sub
End If

'make sure the druggie is still in the area
If Npc(User(Index).NpcTrade).NLocation <> User(Index).Location Then
   frmMain.wsk(Index).SendData Chr$(253) & Chr$(4) & Chr$(0)
   DoEvents
   frmMain.wsk(Index).SendData Chr$(2) & "That druggie has left the area." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   User(Index).NpcTrade = -1
   Exit Sub
End If

'make sure the item is a DOPE type
If Item(User(Index).Item(a)).IType <> C_Dope Then
   frmMain.wsk(Index).SendData Chr$(253) & Chr$(5) & "I told you once bitch, I don't buy that kinda merchandise!  Dope Only!" & Chr$(0)
   DoEvents
   Exit Sub
End If

'Make sure the npc has room to buy the item
For d = 0 To 19
   If Npc(User(Index).NpcTrade).NItem(d) = -1 Then
      Exit For
   ElseIf d = 19 Then
      frmMain.wsk(Index).SendData Chr$(253) & Chr$(5) & "I can't afford anything else right now, try me later after I unload some of this dope." & Chr$(0)
      DoEvents
      Exit Sub
   End If
Next d

'Set Price at a % Mark
c = Item(User(Index).Item(a)).Price + (Item(User(Index).Item(a)).Price * 0.06)
c = Int(c)

'Do Transaction
User(Index).Cash = User(Index).Cash + c
Npc(User(Index).NpcTrade).NItem(d) = User(Index).Item(a)
User(Index).Item(a) = -1
Item(Npc(User(Index).NpcTrade).NItem(d)).ItemGUID = Npc(User(Index).NpcTrade).NpcGUID
Call UpdateSingleItem(Index, a)
User(Index).Reputation = User(Index).Reputation + 1
Call SetRank(Index)
Call UpdateGeneralInfo(Index)

'Update the druggie menu inventory list
MsgX = Chr$(253) & Chr$(7)
For b = 0 To 19
   If User(Index).Item(b) = -1 Then
      MsgX = MsgX & "<Empty>" & Chr$(1)
   ElseIf User(Index).Item(b) <> -1 Then
      MsgX = MsgX & Item(User(Index).Item(b)).IName & Chr$(1)
   End If
Next b
frmMain.wsk(Index).SendData MsgX & Chr$(0)
DoEvents

frmMain.wsk(Index).SendData Chr$(253) & Chr$(6) & "Sold" & Chr$(1) & "Sold" & Chr$(1) & User(Index).Cash & Chr$(1) & Chr$(0)
DoEvents

frmMain.wsk(Index).SendData Chr$(253) & Chr$(5) & "Ok, anything else?" & Chr$(0)
DoEvents

End Sub

Public Sub SendMap(Index As Integer)

Select Case User(Index).CurrTown
   Case "New York"
      frmMain.wsk(Index).SendData Chr$(252) & Chr$(2) & NYMap & Chr$(0)
      DoEvents
   Case "Houston"
      frmMain.wsk(Index).SendData Chr$(252) & Chr$(2) & HOMap & Chr$(0)
      DoEvents
   Case "Miami"
      frmMain.wsk(Index).SendData Chr$(252) & Chr$(2) & MIMap & Chr$(0)
      DoEvents
   Case "Los Angeles"
      frmMain.wsk(Index).SendData Chr$(252) & Chr$(2) & LAMap & Chr$(0)
      DoEvents
   Case "New Jersey"
      frmMain.wsk(Index).SendData Chr$(252) & Chr$(2) & NJMap & Chr$(0)
      DoEvents
   Case "Chicago"
      frmMain.wsk(Index).SendData Chr$(252) & Chr$(2) & CHMap & Chr$(0)
      DoEvents
End Select

End Sub

Public Sub Aim(Index As Integer, PlayerName As String)
Dim a As Integer

For a = 1 To MaxUsers
   If LCase$(PlayerName) = LCase$(Left$(User(a).UName, Len(PlayerName))) And _
      User(Index).Location = User(a).Location And _
      a <> Index And User(a).Status = "Playing" And _
      User(a).IsHiding = False Then
      Call NoHiding(Index)
      User(Index).TargetNum = a
      User(Index).TargetGUID = User(a).UserGUID
      frmMain.wsk(Index).SendData Chr$(2) & "You take aim on " & User(a).UName & "." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      frmMain.wsk(a).SendData Chr$(2) & User(Index).UName & " takes aim on you." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Call CombatMessage(Index, a, Chr$(2) & "You see " & User(Index).UName & " take aim on " & User(a).UName & "." & vbCrLf & vbCrLf & Chr$(0))
      Exit Sub
   End If
Next a

For a = 0 To 9
   If City(User(Index).Location).CNpc(a) <> -1 Then
   If LCase$(Left$(Npc(City(User(Index).Location).CNpc(a)).NName, Len(PlayerName))) = _
      LCase$(PlayerName) Then
      Call NoHiding(Index)
      User(Index).TargetNum = City(User(Index).Location).CNpc(a)
      User(Index).TargetGUID = Npc(City(User(Index).Location).CNpc(a)).NpcGUID
      Npc(City(User(Index).Location).CNpc(a)).NTargetID = Index
      Npc(City(User(Index).Location).CNpc(a)).NTargetGUID = User(Index).UserGUID
      Npc(City(User(Index).Location).CNpc(a)).CanMove = GetTickCount()
      frmMain.wsk(Index).SendData Chr$(2) & "You take aim on " & Npc(City(User(Index).Location).CNpc(a)).NName & "." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " take aim on " & Npc(City(User(Index).Location).CNpc(a)).NName & "." & vbCrLf & vbCrLf & Chr$(0))
      Exit Sub
   End If
   End If
Next a

User(Index).TargetGUID = ""
User(Index).TargetNum = -1
frmMain.wsk(Index).SendData Chr$(2) & "You look around but one matches that name..." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub

Public Function IsHiding(Index As Integer) As Boolean

If User(Index).IsHiding = True Then
   IsHiding = False
   User(Index).IsHiding = False
   frmMain.wsk(Index).SendData Chr$(2) & "You come out of hiding." & vbCrLf & vbCrLf & Chr$(2)
   DoEvents
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " come out from the shadows." & vbCrLf & vbCrLf & Chr$(0))
   Exit Function
ElseIf User(Index).IsHiding = False Then
   IsHiding = False
   Exit Function
End If

End Function

Public Sub GotoPlayer(Index As Integer, Msg As String)
Dim a As Integer

'Check to make sure player access level is high enough
If User(Index).AccessLevel <> 5 Then
   frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

For a = 1 To MaxUsers
   If LCase$(Msg) = LCase$(Left$(User(a).UName, Len(Msg))) Then
      Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " slowly fade away." & vbCrLf & vbCrLf & Chr$(0))
      User(Index).Location = User(a).Location
      User(Index).CurrTown = City(User(Index).Location).CName
      Call UpdateGeneralInfo(Index)
      Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " slowly fade into your view." & vbCrLf & vbCrLf & Chr$(0))
   End If
Next a

End Sub

Public Sub Punch(Index As Integer)
Dim a As Integer

If SkillDelay(Index) = True Then
   Exit Sub
End If

Call NoHiding(Index)

If User(Index).TargetNum = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to take aim on someone before you can punch them." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).TargetNum >= 1 And _
   User(Index).TargetNum <= MaxUsers Then
      If User(Index).TargetGUID = User(User(Index).TargetNum).UserGUID And _
         User(User(Index).TargetNum).Status = "Playing" And _
         User(User(Index).TargetNum).Location = User(Index).Location Then
            If RunAccuracy(Index) = True Then
               User(User(Index).TargetNum).Health = User(User(Index).TargetNum).Health - 2
               If PlayerKillPlayer(Index, User(Index).TargetNum) = True Then
                  Exit Sub
               End If
               Call CombatMessage(Index, User(Index).TargetNum, Chr$(2) & "You see " & User(Index).UName & " throw a hard punch hitting " & User(User(Index).TargetNum).UName & " square in the head." & vbCrLf & vbCrLf & Chr$(0))
               frmMain.wsk(Index).SendData Chr$(2) & "You land a solid punch on " & User(User(Index).TargetNum).UName & "." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               frmMain.wsk(User(Index).TargetNum).SendData Chr$(2) & User(Index).UName & " lands a damaging blow on you." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call UpdateGeneralInfo(User(Index).TargetNum)
               Exit Sub
            Else
               frmMain.wsk(Index).SendData Chr$(2) & "You take a swing at " & User(User(Index).TargetNum).UName & " but miss." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               frmMain.wsk(User(Index).TargetNum).SendData Chr$(2) & User(Index).UName & " takes a swing at you but misses." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call CombatMessage(Index, User(Index).TargetNum, Chr$(2) & "You see " & User(Index).UName & " take a swing at " & User(User(Index).TargetNum).UName & " but miss." & vbCrLf & vbCrLf & Chr$(0))
               Exit Sub
            End If
      End If
End If

If Npc(User(Index).TargetNum).NLocation = User(Index).Location And _
   User(Index).TargetGUID = Npc(User(Index).TargetNum).NpcGUID And _
   Npc(User(Index).TargetNum).NpcGUID <> "" And _
   Npc(User(Index).TargetNum).NHealth > 0 Then
      Npc(User(Index).TargetNum).CanMove = GetTickCount()
      Npc(User(Index).TargetNum).NTargetID = Index
      Npc(User(Index).TargetNum).NTargetGUID = User(Index).UserGUID
         If RunAccuracy(Index) = True Then
            Npc(User(Index).TargetNum).NHealth = Npc(User(Index).TargetNum).NHealth - 2
               If PlayerKillNpc(Index) = True Then
                  Exit Sub
               End If
               Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " throw a hard punch hitting " & Npc(User(Index).TargetNum).NName & " square in the head." & vbCrLf & vbCrLf & Chr$(0))
               frmMain.wsk(Index).SendData Chr$(2) & "You land a solid punch on " & Npc(User(Index).TargetNum).NName & "." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Exit Sub
         Else
               frmMain.wsk(Index).SendData Chr$(2) & "You take a swing at " & Npc(User(Index).TargetNum).NName & " but miss." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " take a swing at " & Npc(User(Index).TargetNum).NName & " but miss." & vbCrLf & vbCrLf & Chr$(0))
               Exit Sub
         End If
End If
            
User(Index).TargetNum = -1
User(Index).TargetGUID = ""
frmMain.wsk(Index).SendData Chr$(2) & "You need to take aim on someone before you can punch them." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub
Public Sub Fire(Index As Integer)
Dim a As Integer

If SkillDelay(Index) = True Then
   Exit Sub
End If

Call NoHiding(Index)

If User(Index).TargetNum = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to take aim on someone before you can shoot them." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).Weapon = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You don't have a weapon equipped." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If Item(User(Index).Weapon).IType <> C_Gun Then
   frmMain.wsk(Index).SendData Chr$(2) & "Your equipped weapon is not a firearm." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).Ammo = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You don't have any ammunition loaded." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

For a = 0 To 19
   If User(Index).Item(a) = User(Index).Ammo Then
      Exit For
   ElseIf a = 19 Then
      Exit Sub
   End If
Next a

If Item(User(Index).Ammo).Amount <= 0 Then
   User(Index).Item(a) = -1
   Call UpdateSingleItem(Index, a)
   Call ResetItem(User(Index).Ammo)
   User(Index).Ammo = -1
   Call UpdateGearInfo(Index)
   frmMain.wsk(Index).SendData Chr$(2) & "Click, Click...  Sounds like your out of ammunition." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).TargetNum >= 1 And _
   User(Index).TargetNum <= MaxUsers Then
      If User(Index).TargetGUID = User(User(Index).TargetNum).UserGUID And _
         User(User(Index).TargetNum).Status = "Playing" And _
         User(User(Index).TargetNum).Location = User(Index).Location Then
            'subtract ammo
            Item(User(Index).Ammo).Amount = Item(User(Index).Ammo).Amount - 1
            Call UpdateSingleItem(Index, a)
            If RunAccuracy(Index) = True Then
               Call GunDamage(Index)
               If PlayerKillPlayer(Index, User(Index).TargetNum) = True Then
                  Exit Sub
               End If
               Call CombatMessage(Index, User(Index).TargetNum, Chr$(2) & "You see " & User(Index).UName & " fire a " & Item(User(Index).Weapon).IName & " at " & User(User(Index).TargetNum).UName & " doing massive damage." & vbCrLf & vbCrLf & Chr$(0))
               frmMain.wsk(Index).SendData Chr$(2) & "You fire your " & Item(User(Index).Weapon).IName & " at " & User(User(Index).TargetNum).UName & " and it's a direct hit." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               frmMain.wsk(User(Index).TargetNum).SendData Chr$(2) & User(Index).UName & " fires a " & Item(User(Index).Weapon).IName & " at you, it's a direct hit!" & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call UpdateGeneralInfo(User(Index).TargetNum)
               Exit Sub
            Else
               frmMain.wsk(Index).SendData Chr$(2) & "You fire your " & Item(User(Index).Weapon).IName & " at " & User(User(Index).TargetNum).UName & " but miss." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               frmMain.wsk(User(Index).TargetNum).SendData Chr$(2) & User(Index).UName & " fires a " & Item(User(Index).Weapon).IName & " at you and misses." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call CombatMessage(Index, User(Index).TargetNum, Chr$(2) & "You see " & User(Index).UName & " fire a " & Item(User(Index).Weapon).IName & " at " & User(User(Index).TargetNum).UName & " but miss." & vbCrLf & vbCrLf & Chr$(0))
               Exit Sub
            End If
      End If
End If

If Npc(User(Index).TargetNum).NLocation = User(Index).Location And _
   User(Index).TargetGUID = Npc(User(Index).TargetNum).NpcGUID And _
   Npc(User(Index).TargetNum).NpcGUID <> "" And _
   Npc(User(Index).TargetNum).NHealth > 0 Then
      Npc(User(Index).TargetNum).CanMove = GetTickCount()
      Npc(User(Index).TargetNum).NTargetID = Index
      Npc(User(Index).TargetNum).NTargetGUID = User(Index).UserGUID
         Item(User(Index).Ammo).Amount = Item(User(Index).Ammo).Amount - 1
         Call UpdateSingleItem(Index, a)
         If RunAccuracy(Index) = True Then
            Npc(User(Index).TargetNum).NHealth = Npc(User(Index).TargetNum).NHealth - (Item(User(Index).Weapon).Damage + Item(User(Index).Ammo).Damage)
               If PlayerKillNpc(Index) = True Then
                  Exit Sub
               End If
               Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " fire a " & Item(User(Index).Weapon).IName & " at " & Npc(User(Index).TargetNum).NName & " and it's a direct hit." & vbCrLf & vbCrLf & Chr$(0))
               frmMain.wsk(Index).SendData Chr$(2) & "You fire your " & Item(User(Index).Weapon).IName & " at " & Npc(User(Index).TargetNum).NName & " doing massive damage." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Exit Sub
            Else
               frmMain.wsk(Index).SendData Chr$(2) & "You fire your " & Item(User(Index).Weapon).IName & " at " & Npc(User(Index).TargetNum).NName & " and miss." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " fire a " & Item(User(Index).Weapon).IName & " at " & Npc(User(Index).TargetNum).NName & " and miss." & vbCrLf & vbCrLf & Chr$(0))
               Exit Sub
         End If
End If
            
User(Index).TargetNum = -1
User(Index).TargetGUID = ""
frmMain.wsk(Index).SendData Chr$(2) & "You need to take aim on someone before you can punch them." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub



Public Sub GunDamage(Index As Integer)
Dim a As Integer
a = 0

If User(User(Index).TargetNum).Armor <> -1 Then
   a = Item(User(User(Index).TargetNum).Armor).Armor
End If

User(User(Index).TargetNum).Health = (User(User(Index).TargetNum).Health - (Item(User(Index).Weapon).Damage + Item(User(Index).Ammo).Damage)) + a

End Sub

Public Sub Strike(Index As Integer)

If SkillDelay(Index) = True Then
   Exit Sub
End If

Call NoHiding(Index)

If User(Index).TargetNum = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to take aim on someone before you can strike them." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).Weapon = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You don't have a weapon equipped." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If Item(User(Index).Weapon).IType <> C_Melee Then
   frmMain.wsk(Index).SendData Chr$(2) & "Your equipped weapon can not be used to strike someone." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).TargetNum >= 1 And _
   User(Index).TargetNum <= MaxUsers Then
      If User(Index).TargetGUID = User(User(Index).TargetNum).UserGUID And _
         User(User(Index).TargetNum).Status = "Playing" And _
         User(User(Index).TargetNum).Location = User(Index).Location Then
            If RunAccuracy(Index) = True Then
               User(User(Index).TargetNum).Health = User(User(Index).TargetNum).Health - Item(User(Index).Weapon).Damage
               If PlayerKillPlayer(Index, User(Index).TargetNum) = True Then
                  Exit Sub
               End If
               Call CombatMessage(Index, User(Index).TargetNum, Chr$(2) & "You see " & User(Index).UName & " strike " & User(User(Index).TargetNum).UName & " with a " & Item(User(Index).Weapon).IName & " doing massive damage." & vbCrLf & vbCrLf & Chr$(0))
               frmMain.wsk(Index).SendData Chr$(2) & "You strike " & User(User(Index).TargetNum).UName & " with your " & Item(User(Index).Weapon).IName & " doing massive damage." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               frmMain.wsk(User(Index).TargetNum).SendData Chr$(2) & User(Index).UName & " strikes you with a " & Item(User(Index).Weapon).IName & " doing massive damage." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call UpdateGeneralInfo(User(Index).TargetNum)
               Exit Sub
            Else
               frmMain.wsk(Index).SendData Chr$(2) & "You strike at " & User(User(Index).TargetNum).UName & " with your " & Item(User(Index).Weapon).IName & " but miss." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               frmMain.wsk(User(Index).TargetNum).SendData Chr$(2) & User(Index).UName & " strikes at you with a " & Item(User(Index).Weapon).IName & " but misses." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call CombatMessage(Index, User(Index).TargetNum, Chr$(2) & "You see " & User(Index).UName & " take a strike at " & User(User(Index).TargetNum).UName & " with a " & Item(User(Index).Weapon).IName & " but miss." & vbCrLf & vbCrLf & Chr$(0))
               Exit Sub
            End If
      End If
End If

If Npc(User(Index).TargetNum).NLocation = User(Index).Location And _
   User(Index).TargetGUID = Npc(User(Index).TargetNum).NpcGUID And _
   Npc(User(Index).TargetNum).NpcGUID <> "" And _
   Npc(User(Index).TargetNum).NHealth > 0 Then
      Npc(User(Index).TargetNum).CanMove = GetTickCount()
      Npc(User(Index).TargetNum).NTargetID = Index
      Npc(User(Index).TargetNum).NTargetGUID = User(Index).UserGUID
         If RunAccuracy(Index) = True Then
            Npc(User(Index).TargetNum).NHealth = Npc(User(Index).TargetNum).NHealth - Item(User(Index).Weapon).Damage
               If PlayerKillNpc(Index) = True Then
                  Exit Sub
               End If
               Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " strike at " & Npc(User(Index).TargetNum).NName & " with a " & Item(User(Index).Weapon).IName & " doing massive damage." & vbCrLf & vbCrLf & Chr$(0))
               frmMain.wsk(Index).SendData Chr$(2) & "You strike " & Npc(User(Index).TargetNum).NName & " with your " & Item(User(Index).Weapon).IName & " doing massive damage." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Exit Sub
            Else
               frmMain.wsk(Index).SendData Chr$(2) & "You strike " & Npc(User(Index).TargetNum).NName & " with your " & Item(User(Index).Weapon).IName & " but miss." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " strike at " & Npc(User(Index).TargetNum).NName & " with a " & Item(User(Index).Weapon).IName & " but miss." & vbCrLf & vbCrLf & Chr$(0))
               Exit Sub
         End If
End If
            
User(Index).TargetNum = -1
User(Index).TargetGUID = ""
frmMain.wsk(Index).SendData Chr$(2) & "You need to take aim on someone before you can punch them." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub

Public Sub Hide(Index As Integer)

If SkillDelay(Index) = True Then
   Exit Sub
End If

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

User(Index).IsHiding = False
If RunHiding(Index) = True Then
   User(Index).IsHiding = True
   Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " slips into the shadows." & vbCrLf & vbCrLf & Chr$(0))
   frmMain.wsk(Index).SendData Chr$(2) & "You manage to slip into the shadows." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
Else
   User(Index).IsHiding = False
   frmMain.wsk(Index).SendData Chr$(2) & "You failed to slip into the shadows." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If
   
End Sub

Public Sub Flee(Index As Integer, Msg As String)
Dim DropItems As Boolean
Dim a As Integer
Dim b As Integer
DropItems = False

Dim i(2) As Integer

For a = 0 To 2
   i(a) = -1
Next a

'Flee North
Select Case Msg
   Case "n"
      If City(User(Index).Location).North = -1 Then
         frmMain.wsk(Index).SendData Chr$(2) & "You can't flee to the north." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf City(User(Index).Location).North <> 1 Then
      Randomize
      b = Int(19 - 0) * Rnd
      i(0) = b
      Do
      Randomize
      b = Int(19 - 0) * Rnd
      i(1) = b
      Loop Until i(1) <> i(0)
      Do
      Randomize
      b = Int(19 - 0) * Rnd
      i(2) = b
      Loop Until i(2) <> i(0) And i(2) <> i(1)
      For a = 0 To 2
         If User(Index).Item(i(a)) <> -1 Then
            If Item(User(Index).Item(i(a))).Equip = False Then
            DropItems = True
            For b = 0 To UBound(City(User(Index).Location).CItem)
               If City(User(Index).Location).CItem(b) = -1 Then
                  City(User(Index).Location).CItem(b) = User(Index).Item(i(a))
                  Item(User(Index).Item(i(a))).Equip = False
                  Item(User(Index).Item(i(a))).Decay = GetTickCount()
                  Item(User(Index).Item(i(a))).OnPlayer = False
                  Item(User(Index).Item(i(a))).ItemGUID = ""
                  Item(User(Index).Item(i(a))).ILocation = User(Index).Location
                  User(Index).Item(i(a)) = -1
                  Exit For
               ElseIf b = UBound(City(User(Index).Location).CItem) Then
                  With City(User(Index).Location)
                  ReDim Preserve .CItem(UBound(.CItem) + 1)
                  .CItem(UBound(.CItem)) = User(Index).Item(i(a))
                  End With
                  Item(User(Index).Item(i(a))).Equip = False
                  Item(User(Index).Item(i(a))).Decay = GetTickCount()
                  Item(User(Index).Item(i(a))).OnPlayer = False
                  Item(User(Index).Item(i(a))).ItemGUID = ""
                  Item(User(Index).Item(i(a))).ILocation = User(Index).Location
                  User(Index).Item(i(a)) = -1
                  Exit For
               End If
            Next b
          End If
         End If
      Next a
         If DropItems = False Then
            Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " has managed to escape the area from battle." & vbCrLf & vbCrLf & Chr$(0))
            User(Index).Reputation = User(Index).Reputation - 100
            frmMain.wsk(Index).SendData Chr$(2) & "You managed to escape to the north from battle.  Such a cowardly act costs you quite a bit of fame and reputation." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            User(Index).Location = City(User(Index).Location).North
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " stumble into the area avoiding the battle to the south." & vbCrLf & vbCrLf & Chr$(0))
            Call ShowCity(Index)
            Call FullInventoryUpdate(Index)
            Call UpdateGeneralInfo(Index)
            Exit Sub
         ElseIf DropItems = True Then
            Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " has managed to escape the area from battle but dropped some items in the process of escaping." & vbCrLf & vbCrLf & Chr$(0))
            User(Index).Reputation = User(Index).Reputation - 100
            frmMain.wsk(Index).SendData Chr$(2) & "You managed to escape to the north from battle.  Such a cowardly act costs you quite a bit of fame and reputation.  In your scuttle to escape, you drop some items." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            User(Index).Location = City(User(Index).Location).North
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " stumble into the area avoiding the battle to the south." & vbCrLf & vbCrLf & Chr$(0))
            Call ShowCity(Index)
            Call FullInventoryUpdate(Index)
            Call UpdateGeneralInfo(Index)
            Exit Sub
         End If
      End If

'Flee East
   Case "e"
      If City(User(Index).Location).East = -1 Then
         frmMain.wsk(Index).SendData Chr$(2) & "You can't flee to the east." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf City(User(Index).Location).East <> 1 Then
      Randomize
      b = Int(19 - 0) * Rnd
      i(0) = b
      Do
      Randomize
      b = Int(19 - 0) * Rnd
      i(1) = b
      Loop Until i(1) <> i(0)
      Do
      Randomize
      b = Int(19 - 0) * Rnd
      i(2) = b
      Loop Until i(2) <> i(0) And i(2) <> i(1)
      For a = 0 To 2
         If User(Index).Item(i(a)) <> -1 Then
            If Item(User(Index).Item(i(a))).Equip = False Then
            DropItems = True
            For b = 0 To UBound(City(User(Index).Location).CItem)
               If City(User(Index).Location).CItem(b) = -1 Then
                  City(User(Index).Location).CItem(b) = User(Index).Item(i(a))
                  Item(User(Index).Item(i(a))).Equip = False
                  Item(User(Index).Item(i(a))).Decay = GetTickCount()
                  Item(User(Index).Item(i(a))).OnPlayer = False
                  Item(User(Index).Item(i(a))).ItemGUID = ""
                  Item(User(Index).Item(i(a))).ILocation = User(Index).Location
                  User(Index).Item(i(a)) = -1
                  Exit For
               ElseIf b = UBound(City(User(Index).Location).CItem) Then
                  With City(User(Index).Location)
                  ReDim Preserve .CItem(UBound(.CItem) + 1)
                  .CItem(UBound(.CItem)) = User(Index).Item(i(a))
                  End With
                  Item(User(Index).Item(i(a))).Equip = False
                  Item(User(Index).Item(i(a))).Decay = GetTickCount()
                  Item(User(Index).Item(i(a))).OnPlayer = False
                  Item(User(Index).Item(i(a))).ItemGUID = ""
                  Item(User(Index).Item(i(a))).ILocation = User(Index).Location
                  User(Index).Item(i(a)) = -1
                  Exit For
               End If
            Next b
          End If
         End If
      Next a
         If DropItems = False Then
            Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " has managed to escape the area from battle." & vbCrLf & vbCrLf & Chr$(0))
            User(Index).Reputation = User(Index).Reputation - 100
            frmMain.wsk(Index).SendData Chr$(2) & "You managed to escape to the east from battle.  Such a cowardly act costs you quite a bit of fame and reputation." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            User(Index).Location = City(User(Index).Location).East
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " stumble into the area avoiding the battle to the west." & vbCrLf & vbCrLf & Chr$(0))
            Call ShowCity(Index)
            Call FullInventoryUpdate(Index)
            Call UpdateGeneralInfo(Index)
            Exit Sub
         ElseIf DropItems = True Then
            Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " has managed to escape the area from battle but dropped some items in the process of escaping." & vbCrLf & vbCrLf & Chr$(0))
            User(Index).Reputation = User(Index).Reputation - 100
            frmMain.wsk(Index).SendData Chr$(2) & "You managed to escape to the east from battle.  Such a cowardly act costs you quite a bit of fame and reputation.  In your scuttle to escape, you drop some items." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            User(Index).Location = City(User(Index).Location).East
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " stumble into the area avoiding the battle to the west." & vbCrLf & vbCrLf & Chr$(0))
            Call ShowCity(Index)
            Call FullInventoryUpdate(Index)
            Call UpdateGeneralInfo(Index)
            Exit Sub
         End If
      End If

'Flee South
   Case "s"
      If City(User(Index).Location).South = -1 Then
         frmMain.wsk(Index).SendData Chr$(2) & "You can't flee to the south." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf City(User(Index).Location).South <> 1 Then
      Randomize
      b = Int(19 - 0) * Rnd
      i(0) = b
      Do
      Randomize
      b = Int(19 - 0) * Rnd
      i(1) = b
      Loop Until i(1) <> i(0)
      Do
      Randomize
      b = Int(19 - 0) * Rnd
      i(2) = b
      Loop Until i(2) <> i(0) And i(2) <> i(1)
      For a = 0 To 2
         If User(Index).Item(i(a)) <> -1 Then
            If Item(User(Index).Item(i(a))).Equip = False Then
            DropItems = True
            For b = 0 To UBound(City(User(Index).Location).CItem)
               If City(User(Index).Location).CItem(b) = -1 Then
                  City(User(Index).Location).CItem(b) = User(Index).Item(i(a))
                  Item(User(Index).Item(i(a))).Equip = False
                  Item(User(Index).Item(i(a))).Decay = GetTickCount()
                  Item(User(Index).Item(i(a))).OnPlayer = False
                  Item(User(Index).Item(i(a))).ItemGUID = ""
                  Item(User(Index).Item(i(a))).ILocation = User(Index).Location
                  User(Index).Item(i(a)) = -1
                  Exit For
               ElseIf b = UBound(City(User(Index).Location).CItem) Then
                  With City(User(Index).Location)
                  ReDim Preserve .CItem(UBound(.CItem) + 1)
                  .CItem(UBound(.CItem)) = User(Index).Item(i(a))
                  End With
                  Item(User(Index).Item(i(a))).Equip = False
                  Item(User(Index).Item(i(a))).Decay = GetTickCount()
                  Item(User(Index).Item(i(a))).OnPlayer = False
                  Item(User(Index).Item(i(a))).ItemGUID = ""
                  Item(User(Index).Item(i(a))).ILocation = User(Index).Location
                  User(Index).Item(i(a)) = -1
                  Exit For
               End If
            Next b
          End If
         End If
      Next a
         If DropItems = False Then
            Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " has managed to escape the area from battle." & vbCrLf & vbCrLf & Chr$(0))
            User(Index).Reputation = User(Index).Reputation - 100
            frmMain.wsk(Index).SendData Chr$(2) & "You managed to escape to the south from battle.  Such a cowardly act costs you quite a bit of fame and reputation." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            User(Index).Location = City(User(Index).Location).South
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " stumble into the area avoiding the battle to the north." & vbCrLf & vbCrLf & Chr$(0))
            Call ShowCity(Index)
            Call FullInventoryUpdate(Index)
            Call UpdateGeneralInfo(Index)
            Exit Sub
         ElseIf DropItems = True Then
            Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " has managed to escape the area from battle but dropped some items in the process of escaping." & vbCrLf & vbCrLf & Chr$(0))
            User(Index).Reputation = User(Index).Reputation - 100
            frmMain.wsk(Index).SendData Chr$(2) & "You managed to escape to the south from battle.  Such a cowardly act costs you quite a bit of fame and reputation.  In your scuttle to escape, you drop some items." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            User(Index).Location = City(User(Index).Location).South
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " stumble into the area avoiding the battle to the north." & vbCrLf & vbCrLf & Chr$(0))
            Call ShowCity(Index)
            Call FullInventoryUpdate(Index)
            Call UpdateGeneralInfo(Index)
            Exit Sub
         End If
      End If

'Flee West
   Case "w"
      If City(User(Index).Location).West = -1 Then
         frmMain.wsk(Index).SendData Chr$(2) & "You can't flee to the west." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf City(User(Index).Location).West <> 1 Then
      Randomize
      b = Int(19 - 0) * Rnd
      i(0) = b
      Do
      Randomize
      b = Int(19 - 0) * Rnd
      i(1) = b
      Loop Until i(1) <> i(0)
      Do
      Randomize
      b = Int(19 - 0) * Rnd
      i(2) = b
      Loop Until i(2) <> i(0) And i(2) <> i(1)
      For a = 0 To 2
         If User(Index).Item(i(a)) <> -1 Then
            If Item(User(Index).Item(i(a))).Equip = False Then
            DropItems = True
            For b = 0 To UBound(City(User(Index).Location).CItem)
               If City(User(Index).Location).CItem(b) = -1 Then
                  City(User(Index).Location).CItem(b) = User(Index).Item(i(a))
                  Item(User(Index).Item(i(a))).Equip = False
                  Item(User(Index).Item(i(a))).Decay = GetTickCount()
                  Item(User(Index).Item(i(a))).OnPlayer = False
                  Item(User(Index).Item(i(a))).ItemGUID = ""
                  Item(User(Index).Item(i(a))).ILocation = User(Index).Location
                  User(Index).Item(i(a)) = -1
                  Exit For
               ElseIf b = UBound(City(User(Index).Location).CItem) Then
                  With City(User(Index).Location)
                  ReDim Preserve .CItem(UBound(.CItem) + 1)
                  .CItem(UBound(.CItem)) = User(Index).Item(i(a))
                  End With
                  Item(User(Index).Item(i(a))).Equip = False
                  Item(User(Index).Item(i(a))).Decay = GetTickCount()
                  Item(User(Index).Item(i(a))).OnPlayer = False
                  Item(User(Index).Item(i(a))).ItemGUID = ""
                  Item(User(Index).Item(i(a))).ILocation = User(Index).Location
                  User(Index).Item(i(a)) = -1
                  Exit For
               End If
            Next b
          End If
         End If
      Next a
         If DropItems = False Then
            Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " has managed to escape the area from battle." & vbCrLf & vbCrLf & Chr$(0))
            User(Index).Reputation = User(Index).Reputation - 100
            frmMain.wsk(Index).SendData Chr$(2) & "You managed to escape to the west from battle.  Such a cowardly act costs you quite a bit of fame and reputation." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            User(Index).Location = City(User(Index).Location).West
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " stumble into the area avoiding the battle to the east." & vbCrLf & vbCrLf & Chr$(0))
            Call ShowCity(Index)
            Call FullInventoryUpdate(Index)
            Call UpdateGeneralInfo(Index)
            Exit Sub
         ElseIf DropItems = True Then
            Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " has managed to escape the area from battle but dropped some items in the process of escaping." & vbCrLf & vbCrLf & Chr$(0))
            User(Index).Reputation = User(Index).Reputation - 100
            frmMain.wsk(Index).SendData Chr$(2) & "You managed to escape to the west from battle.  Such a cowardly act costs you quite a bit of fame and reputation.  In your scuttle to escape, you drop some items." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            User(Index).Location = City(User(Index).Location).West
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " stumble into the area avoiding the battle to the east." & vbCrLf & vbCrLf & Chr$(0))
            Call ShowCity(Index)
            Call FullInventoryUpdate(Index)
            Call UpdateGeneralInfo(Index)
            Exit Sub
         End If
      End If

End Select

End Sub

Public Sub HealInfo(Index As Integer)
Dim a As Integer
Dim b As Integer

If City(User(Index).Location).Hospital = False Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to find a hospital first." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

Call NoHiding(Index)

If User(Index).Cash < HealPrice Then
   frmMain.wsk(Index).SendData Chr$(2) & "A Nurse says: Your current financial situation won't do you any good here." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).Health >= 100 Then
   frmMain.wsk(Index).SendData Chr$(2) & "A Nurse says: You don't need our services, you're in perfect health." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
ElseIf User(Index).Health < 100 Then
   a = (100 - User(Index).Health) * HealPrice
   frmMain.wsk(Index).SendData Chr$(2) & "A Nurse says: Your health needs some attention,  It will take you " & 100 - User(Index).Health & " days in the hospital and $" & a & " dollars to have perfect health." & vbCrLf & vbCrLf & "To use our services, type  healme <amount>,  You currently can afford " & Int(User(Index).Cash / HealPrice) & " days in the hospital." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

End Sub

Public Sub HealMe(Index As Integer, Msg As String)
Dim a As Single


If IsNumeric(Msg) = False Then
   Exit Sub
ElseIf IsNumeric(Msg) = True Then
   a = Int(Msg)
End If

If Int(a) < 1 Or Int(a) > 99 Then
   Exit Sub
End If

If City(User(Index).Location).Hospital = False Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to find a hospital first." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

Call NoHiding(Index)

If User(Index).Cash < HealPrice Then
   frmMain.wsk(Index).SendData Chr$(2) & "A Nurse says: Your current financial situation won't do you any good here." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).Health >= 100 Then
   frmMain.wsk(Index).SendData Chr$(2) & "A Nurse says: You don't need our services, you're in perfect health." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If Int(a) > (100 - User(Index).Health) Then
   frmMain.wsk(Index).SendData Chr$(2) & "A Nurse says: You don't need to stay here that long." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).Health < 100 Then
   If (Int(a) * HealPrice) > User(Index).Cash Then
      frmMain.wsk(Index).SendData Chr$(2) & "A Nurse says: you do not have enough money to stay that long." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Exit Sub
   ElseIf (Int(a) * HealPrice) <= User(Index).Cash Then
      User(Index).Cash = User(Index).Cash - Int(a) * HealPrice
      User(Index).Health = User(Index).Health + Int(a)
      Call UpdateGeneralInfo(Index)
      frmMain.wsk(Index).SendData Chr$(2) & "You hand over $" & a * HealPrice & " and the doctors fix you right up." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Exit Sub
   End If
End If

frmMain.wsk(Index).SendData Chr$(2) & "The Nurse looks at you strangley." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub

Public Sub ShowSkills(Index As Integer)
On Error Resume Next
Dim a As Integer
Dim Msg As String

Msg = "Your current skills:" & vbCrLf
Msg = Msg & "Accuracy:    " & Format$(User(Index).Accuracy, "#0.0") & vbCrLf
Msg = Msg & "Hiding:     " & Format$(User(Index).Hiding, "#0.0") & vbCrLf
'Msg = Msg & "Searching:     " & Format$(User(Index).Search, "#0.0") & vbCrLf
Msg = Msg & "Tracking:     " & Format$(User(Index).Tracking, "#0.0") & vbCrLf
'Msg = Msg & "Pimping:     " & Format$(User(Index).Pimping, "#0.0") & vbCrLf
'Msg = Msg & "Chemistry:     " & Format$(User(Index).Chemistry, "#0.0") & vbCrLf
'Msg = Msg & "Snooping:     " & Format$(User(Index).Snooping, "#0.0") & vbCrLf
'Msg = Msg & "Stealing:     " & Format$(User(Index).Stealing, "#0.0") & vbCrLf

frmMain.wsk(Index).SendData Chr$(2) & Msg & vbCrLf & Chr$(0)
DoEvents

End Sub
Public Sub Deposit(Index As Integer, Msg As String)
Dim a As Single

If City(User(Index).Location).Bank = False Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to find a bank before you can deposit any cash." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If City(User(Index).Location).CName <> User(Index).HomeTown Then
   frmMain.wsk(Index).SendData Chr$(2) & "You can only bank in your home town." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).Reputation <= 200 And _
   User(Index).Reputation >= -4000 Then
   frmMain.wsk(Index).SendData Chr$(2) & "Your current rank doesn't allow you to open a bank account." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If
   
If IsNumeric(Msg) = False Then
   Exit Sub
ElseIf IsNumeric(Msg) = True Then
   a = Int(Msg)
End If

If Int(a) < 1 Then
   Exit Sub
End If

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

Call NoHiding(Index)

If Int(a) > User(Index).Cash Then
   frmMain.wsk(Index).SendData Chr$(2) & "You don't have that much cash." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
ElseIf Int(a) <= User(Index).Cash Then
   User(Index).Cash = User(Index).Cash - Int(a)
   User(Index).Bank = User(Index).Bank + Int(a)
   Call UpdateGeneralInfo(Index)
   frmMain.wsk(Index).SendData Chr$(2) & "You deposit $" & Int(a) & " into your bank account." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

frmMain.wsk(Index).SendData Chr$(2) & "The bank teller looks strangley at you." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub

Public Sub Withdraw(Index As Integer, Msg As String)
Dim a As Single

If City(User(Index).Location).Bank = False Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to find a bank before you can deposit any cash." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If City(User(Index).Location).CName <> User(Index).HomeTown Then
   frmMain.wsk(Index).SendData Chr$(2) & "You can only bank in your home town." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If
   
If IsNumeric(Msg) = False Then
   Exit Sub
ElseIf IsNumeric(Msg) = True Then
   a = Int(Msg)
End If

If Int(a) < 1 Then
   Exit Sub
End If

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

Call NoHiding(Index)

If Int(a) > User(Index).Bank Then
   frmMain.wsk(Index).SendData Chr$(2) & "You don't have that much cash in your bank." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
ElseIf Int(a) <= User(Index).Bank Then
   User(Index).Bank = User(Index).Bank - Int(a)
   User(Index).Cash = User(Index).Cash + Int(a)
   Call UpdateGeneralInfo(Index)
   frmMain.wsk(Index).SendData Chr$(2) & "You withdraw $" & Int(a) & " from your bank account." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

frmMain.wsk(Index).SendData Chr$(2) & "The bank teller looks strangley at you." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub

Public Sub TrackPlayer(Index As Integer, xMsg As String)
Dim a As Integer

'If Len(xMsg) <= 0 Then Exit Sub

'For a = 1 To MaxUsers
'   If User(a).Status = "Playing" And _
'      Trim$(LCase$(xMsg)) = Trim$(LCase$(User(a).UName)) Then
'         If RunTracking(Index) = True Then


End Sub
