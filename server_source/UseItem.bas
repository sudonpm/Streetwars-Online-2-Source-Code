Attribute VB_Name = "UseItem"
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


Public Sub UsePhone(Index As Integer)
Dim a As Integer
Dim Msg As String

If User(Index).Cash < 15 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You don't have enough cash to make the calls needed to locate any connections." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
ElseIf User(Index).Cash >= 15 Then
   User(Index).Cash = User(Index).Cash - 15
   Call UpdateGeneralInfo(Index)
   For a = 0 To UBound(Npc)
      If Npc(a).NCity = City(User(Index).Location).CName And _
         Npc(a).NpcType = N_Dealer Then
         Msg = Msg & Npc(a).NName & " " & Npc(a).NameTag & " - (" & City(Npc(a).NLocation).Compass & ")" & vbCrLf
      ElseIf Npc(a).NCity = City(User(Index).Location).CName And _
         Npc(a).NpcType = N_Druggie Then
         Msg = Msg & Npc(a).NName & " " & Npc(a).NameTag & " - (" & City(Npc(a).NLocation).Compass & ")" & vbCrLf
      End If
   Next a
   Msg = Msg & vbCrLf & "Cha-Ching...  Only $15.00 bucks, what a deal!" & vbCrLf
   frmMain.wsk(Index).SendData Chr$(2) & Msg & vbCrLf & Chr$(0)
   DoEvents
End If

End Sub

Public Sub UseMedStick(Index As Integer, ByVal ItemNo As Integer)

User(Index).Health = User(Index).Health + Int(10 - 7) * Rnd + 7

If User(Index).Health > 100 Then
   User(Index).Health = 100
End If

Call ResetItem(User(Index).Item(ItemNo))
User(Index).Item(ItemNo) = -1
Call UpdateSingleItem(Index, ItemNo)
Call UpdateGeneralInfo(Index)

frmMain.wsk(Index).SendData Chr$(2) & "You shove the medstick syringe in your arm and administer yourself a healthy dose of medicine." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub
