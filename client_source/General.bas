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

'Server Address
'Public Const ServerAddy As String = "your address/ip here"

'Server Port
Public Const ServerPort = 5002

'Version ID
Public Const ClientVer = 7000

'Text Input Delay
Public Const InputDelayTick = 150
Public InputDelayNew As Long
Public InputDelayOld As Long

'Move Delay
Public Const MoveDelayTick = 150
Public MoveDelayNew As Long
Public MoveDelayOld As Long

'Key Down/UP Boolean
Public KeyUsed As Boolean

'Disable X Control Box
Public Const MF_BYPOSITION = &H400
Public Const MF_REMOVE = &H1000
Public Declare Function DrawMenuBar Lib "user32" _
(ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" _
(ByVal hMenu As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" _
(ByVal hWnd As Long, _
ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32" _
(ByVal hMenu As Long, _
ByVal nPosition As Long, _
ByVal wFlags As Long) As Long '--end block--'

Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNORMAL = 1

Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hWnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

'Get Tick Count
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub DrugDealMessage(Msg As String)
On Error Resume Next

frmBuyDrugs.lblMessage.Caption = Msg

End Sub
Public Sub ShowText(Msg As String)
Const iMaxChar = 2500 'Maximum text scrollback
Dim iRemove As Integer

'Check to make sure the maximum scrollback of
'the text window doesn't exceed 2500

With frmMain.txtOutput
  If Len(.Text) + Len(Msg) > iMaxChar Then
     iRemove = (Len(.Text) + Len(Msg)) - iMaxChar
     .Text = Mid$(.Text & Msg, iRemove)
     .SelStart = Len(.Text)
  Else
  .Text = .Text & Msg
  .SelStart = Len(.Text)
  End If
End With

End Sub


Public Sub NewAccount()

frmMain.Enabled = False
frmNewAccount.Show
DoEvents

End Sub

Public Sub DupeName()

frmNewAccount.lblMessage.Caption = "That name is already in use by another player, please choose a different one."
frmNewAccount.txtName.Enabled = True
frmNewAccount.txtPassOne.Enabled = True
frmNewAccount.txtPassTwo.Enabled = True
frmNewAccount.cboCity.Enabled = True
frmNewAccount.cmdCreate.Enabled = True
frmNewAccount.txtName.SetFocus

End Sub

Public Sub AccountCreated()
On Error Resume Next

Unload frmNewAccount
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Public Sub UpdateFullInventory(Msg As String)
Dim a As Integer 'Counter
Dim SplitMsg() As String

SplitMsg = Split(Msg, Chr$(1))

'Update inventory slots
For a = 0 To UBound(SplitMsg) - 1
   frmMain.lstInventory.List(a) = SplitMsg(a)
Next a

End Sub

Public Sub UpdateSingleItem(Msg As String)
Dim a As Integer, b As Integer, c As Integer

a = InStr(1, Msg, Chr$(1))
If IsNumeric(Left$(Msg, a - 1)) = True Then
   b = Left$(Msg, a - 1)
Else
   Exit Sub
End If

frmMain.lstInventory.List(b) = Mid$(Msg, a + 1)
DoEvents

End Sub

Public Sub TravelMenu(Msg As String)
Dim a As Integer, b As Integer, c As Integer
Dim d As Integer, e As Integer, f As Integer
Dim TempCity As String, g As Integer

a = InStr(1, Msg, Chr$(1))
TempCity = Left$(Msg, a - 1)

frmMain.Enabled = False
frmTravel.Show

Select Case LCase$(TempCity)
   Case "new york"
      frmTravel.cmdNewYork.Enabled = False
   Case "los angeles"
      frmTravel.cmdLosAngeles.Enabled = False
   Case "houston"
      frmTravel.cmdHouston.Enabled = False
   Case "miami"
      frmTravel.cmdMiami.Enabled = False
   Case "chicago"
      frmTravel.cmdChicago.Enabled = False
   Case "new jersey"
      frmTravel.cmdNewJersey.Enabled = False
End Select

b = InStr(a + 1, Msg, Chr$(1))
frmTravel.lblNewYork.Caption = "$" & Mid$(Msg, a + 1, b - a - 1) & ".00"

c = InStr(b + 1, Msg, Chr$(1))
frmTravel.lblLosAngeles.Caption = "$" & Mid$(Msg, b + 1, c - b - 1) & ".00"

d = InStr(c + 1, Msg, Chr$(1))
frmTravel.lblHouston.Caption = "$" & Mid$(Msg, c + 1, d - c - 1) & ".00"

e = InStr(d + 1, Msg, Chr$(1))
frmTravel.lblMiami.Caption = "$" & Mid$(Msg, d + 1, e - d - 1) & ".00"

f = InStr(e + 1, Msg, Chr$(1))
frmTravel.lblChicago.Caption = "$" & Mid$(Msg, e + 1, f - e - 1) & ".00"

g = InStr(f + 1, Msg, Chr$(1))
frmTravel.lblNewJersey.Caption = "$" & Mid$(Msg, f + 1, g - f - 1) & ".00"

frmTravel.cmdForgetIt.SetFocus
DoEvents

End Sub

Public Sub PawnShopMenu(Msg As String)
Dim a As Integer 'Counter
Dim b As Integer 'Counter
Dim c As Integer 'Right Side of String
Dim LeftString As String
Dim RightString As String
Dim LeftInv() As String
Dim RightInv() As String
b = 0

c = InStr(1, Msg, Chr$(2))
LeftString = Left$(Msg, c - 1)
RightString = Mid$(Msg, c + 1)

LeftInv = Split(LeftString, Chr$(1))
RightInv = Split(RightString, Chr$(1))

frmMain.Enabled = False
frmPawnShop.Show
DoEvents

For a = 0 To UBound(LeftInv) - 1
   frmPawnShop.lstShop.AddItem LeftInv(a), b
   b = b + 1
Next a

b = 0

For a = 0 To UBound(RightInv) - 1
   frmPawnShop.lstInv.AddItem RightInv(a), b
   b = b + 1
Next a

frmPawnShop.cmdBuy.Enabled = False
frmPawnShop.cmdSell.Enabled = False
frmPawnShop.cmdExit.SetFocus

DoEvents

End Sub

Public Sub UpdateCashRank(Msg As String)
On Error GoTo Failed
Dim a As Integer
Dim b As Integer

a = InStr(1, Msg, Chr$(1))
frmPawnShop.lblCash.Caption = "$" & Left$(Msg, a - 1)

b = InStr(a + 1, Msg, Chr$(1))
frmPawnShop.lblRank.Caption = Mid$(Msg, a + 1, b - a - 1)


Exit Sub

Failed:
Unload frmPawnShop
frmMain.Enabled = True
frmMain.txtInput.SetFocus
End Sub

Public Sub PawnShopItemInfo(Msg As String)
On Error GoTo Failed
Dim a As Integer
Dim b As Integer
Dim c As Integer

a = InStr(1, Msg, Chr$(1))
frmPawnShop.lblPrice.Caption = Left$(Msg, a - 1)

b = InStr(a + 1, Msg, Chr$(1))
frmPawnShop.lblCanBuy.Caption = Mid$(Msg, a + 1, b - a - 1)

c = InStr(b + 1, Msg, Chr$(1))
frmPawnShop.lblItem.Caption = Mid$(Msg, b + 1, c - b - 1)

Exit Sub

Failed:
Unload frmPawnShop
frmMain.Enabled = True
frmMain.txtInput.SetFocus
End Sub



Public Sub PawnShopPlayerInventoryUpdate(Msg As String)
'On Error GoTo Failed
Dim a As Integer
Dim b As Integer
Dim SplitMsg() As String
b = 0

SplitMsg = Split(Msg, Chr$(1))

frmPawnShop.lstInv.Clear

For a = 0 To UBound(SplitMsg) - 1
   frmPawnShop.lstInv.AddItem SplitMsg(a), b
   b = b + 1
Next a
DoEvents
   
Exit Sub

Failed:
Unload frmPawnShop
frmMain.Enabled = True
frmMain.txtInput.SetFocus
End Sub

Public Sub UpdateGeneralInfo(Msg As String)
Dim a As Integer, b As Integer, c As Integer
Dim d As Integer, e As Integer, f As Integer
Dim g As Integer, h As Integer

a = InStr(1, Msg, Chr$(1))
frmMain.lblName.Caption = Left$(Msg, a - 1)

b = InStr(a + 1, Msg, Chr$(1))
frmMain.lblHealth.Caption = Mid$(Msg, a + 1, b - a - 1)

c = InStr(b + 1, Msg, Chr$(1))
frmMain.lblCash.Caption = "$" & Mid$(Msg, b + 1, c - b - 1)

d = InStr(c + 1, Msg, Chr$(1))
frmMain.lblBank.Caption = "$" & Mid$(Msg, c + 1, d - c - 1)

e = InStr(d + 1, Msg, Chr$(1))
frmMain.lblHomeTown.Caption = Mid$(Msg, d + 1, e - d - 1)

f = InStr(e + 1, Msg, Chr$(1))
frmMain.lblLocation.Caption = Mid$(Msg, e + 1, f - e - 1)

g = InStr(f + 1, Msg, Chr$(1))
frmMain.lblRank.Caption = Mid$(Msg, f + 1, g - f - 1)

h = InStr(g + 1, Msg, Chr$(1))
frmMain.lblKills.Caption = Mid$(Msg, g + 1, h - g - 1)

End Sub

Public Sub UpdateGearInfo(Msg As String)
Dim a As Integer, b As Integer, c As Integer

a = InStr(1, Msg, Chr$(1))
frmMain.lblWeapon.Caption = Left$(Msg, a - 1)

b = InStr(a + 1, Msg, Chr$(1))
frmMain.lblArmor.Caption = Mid$(Msg, a + 1, b - a - 1)

c = InStr(b + 1, Msg, Chr$(1))
frmMain.lblAmmo.Caption = Mid$(Msg, b + 1, c - b - 1)

End Sub

Public Sub UpdatePlayerList(Msg As String)
Dim a As Integer
Dim SplitMsg() As String

SplitMsg = Split(Msg, Chr$(1))
frmMain.lstUsers.Clear

For a = 0 To UBound(SplitMsg) - 1
   frmMain.lstUsers.AddItem SplitMsg(a)
Next a

End Sub

Public Sub BuyDrugMenu(Msg As String)
On Error Resume Next
Dim a As Integer
Dim b As Integer
Dim SplitMsg() As String
b = 0

frmMain.Enabled = False
frmBuyDrugs.Show
frmBuyDrugs.cmdBuyDrug.Enabled = False
DoEvents

SplitMsg = Split(Msg, Chr$(1))

frmBuyDrugs.lstBuyDrug.Clear
For a = 0 To UBound(SplitMsg) - 1
   frmBuyDrugs.lstBuyDrug.AddItem SplitMsg(a), b
   b = b + 1
Next a

frmBuyDrugs.cmdExit.SetFocus
DoEvents

End Sub

Public Sub CloseDrugDealMenu()
On Error Resume Next

Unload frmBuyDrugs
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Public Sub DrugDealItemInfo(Msg As String)
On Error Resume Next
Dim a As Integer
Dim b As Integer
Dim c As Integer

a = InStr(1, Msg, Chr$(1))
frmBuyDrugs.lblDrug.Caption = Left$(Msg, a - 1)

b = InStr(a + 1, Msg, Chr$(1))
frmBuyDrugs.lblPrice.Caption = Mid$(Msg, a + 1, b - a - 1)

c = InStr(b + 1, Msg, Chr$(1))
frmBuyDrugs.lblCash.Caption = "$" & Mid$(Msg, b + 1, c - b - 1)


End Sub

Public Sub UpdateDealerInventory(Msg As String)
On Error Resume Next
Dim a As Integer
Dim b As Integer
b = 0

Dim SplitMsg() As String

SplitMsg = Split(Msg, Chr$(1))
frmBuyDrugs.lstBuyDrug.Clear

For a = 0 To UBound(SplitMsg) - 1
   frmBuyDrugs.lstBuyDrug.AddItem SplitMsg(a), b
   b = b + 1
Next a

DoEvents

End Sub





Public Sub SellDrugMenu(Msg As String)
On Error Resume Next
Dim a As Integer
Dim b As Integer
Dim SplitMsg() As String
b = 0


frmMain.Enabled = False
frmSellDrugs.Show
frmSellDrugs.cmdSell.Enabled = False
DoEvents

SplitMsg = Split(Msg, Chr$(1))

frmSellDrugs.lstInventory.Clear
For a = 0 To UBound(SplitMsg) - 1
   frmSellDrugs.lstInventory.AddItem SplitMsg(a), b
   b = b + 1
Next a

frmSellDrugs.cmdExit.SetFocus
DoEvents

End Sub

Public Sub CloseDruggieMenu()
On Error Resume Next

Unload frmSellDrugs
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Public Sub DruggieMenuMessage(Msg As String)
On Error Resume Next

frmSellDrugs.lblMessage.Caption = Msg

End Sub

Public Sub DruggieMenuItemInfo(Msg As String)
On Error Resume Next
Dim a As Integer
Dim b As Integer
Dim c As Integer

a = InStr(1, Msg, Chr$(1))
frmSellDrugs.lblDrug.Caption = Left$(Msg, a - 1)

b = InStr(a + 1, Msg, Chr$(1))
frmSellDrugs.lblPrice.Caption = Mid$(Msg, a + 1, b - a - 1)

c = InStr(b + 1, Msg, Chr$(1))
frmSellDrugs.lblCash.Caption = Mid$(Msg, b + 1, c - b - 1)

End Sub

Public Sub ReUpdateDruggieInventory(Msg As String)
On Error Resume Next
Dim a As Integer
Dim b As Integer
Dim SplitMsg() As String
b = 0

frmSellDrugs.cmdSell.Enabled = False
DoEvents

SplitMsg = Split(Msg, Chr$(1))

frmSellDrugs.lstInventory.Clear
For a = 0 To UBound(SplitMsg) - 1
   frmSellDrugs.lstInventory.AddItem SplitMsg(a), b
   b = b + 1
Next a

frmSellDrugs.cmdExit.SetFocus
DoEvents

End Sub

Public Sub ShowMap(Msg As String)
On Error GoTo Failed
Dim a As Integer
Dim b As Integer
b = 0

Dim SplitMsg() As String

frmMain.Enabled = False
frmMap.Show
DoEvents

frmMap.lstCity.Clear
SplitMsg = Split(Msg, Chr$(1))

For a = 0 To UBound(SplitMsg) - 1
   frmMap.lstCity.AddItem SplitMsg(a), b
   b = b + 1
Next a
Exit Sub

Failed:
'Error
End Sub

Public Sub UpdateNews(Msg As String)

If frmMain.txtNews.ForeColor = vbWhite Then
   frmMain.txtNews.ForeColor = vbRed
ElseIf frmMain.txtNews.ForeColor = vbRed Then
   frmMain.txtNews.ForeColor = vbGreen
ElseIf frmMain.txtNews.ForeColor = vbGreen Then
   frmMain.txtNews.ForeColor = vbYellow
ElseIf frmMain.txtNews.ForeColor = vbYellow Then
   frmMain.txtNews.ForeColor = vbMagenta
ElseIf frmMain.txtNews.ForeColor = vbMagenta Then
   frmMain.txtNews.ForeColor = vbCyan
ElseIf frmMain.txtNews.ForeColor = vbCyan Then
   frmMain.txtNews.ForeColor = vbWhite
End If

frmMain.txtNews.Text = ""
frmMain.txtNews.Text = Msg

End Sub

Public Function InputDelay() As Boolean

InputDelayNew = GetTickCount()

If InputDelayNew - InputDelayOld > InputDelayTick Then
   InputDelay = False
   InputDelayOld = GetTickCount()
   Exit Function
Else
   InputDelay = True
End If

End Function

Public Function MoveDelay() As Boolean

MoveDelayNew = GetTickCount()

If MoveDelayNew - MoveDelayOld > MoveDelayTick Then
   MoveDelay = False
   MoveDelayOld = GetTickCount()
   Exit Function
Else
   MoveDelay = True
End If

End Function

Public Function OpenLocation(URL As String, WindowState As Long) As Long
    
    Dim lHWnd As Long
    Dim lAns As Long

    lAns = ShellExecute(lHWnd, "open", URL, vbNullString, _
    vbNullString, WindowState)
   
    OpenLocation = lAns

    'ALTERNATIVE: if not interested in module handle or error
    'code change return value to boolean; then the above line
    'becomes:

    'OpenLocation = (lAns < 32)

End Function
