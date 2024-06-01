VERSION 5.00
Begin VB.Form frmSymbols 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Symbols"
   ClientHeight    =   7005
   ClientLeft      =   7935
   ClientTop       =   5010
   ClientWidth     =   8100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   8100
   Begin VB.Frame frmHighAlarm 
      ForeColor       =   &H00808080&
      Height          =   1005
      Left            =   3270
      TabIndex        =   39
      Top             =   5130
      Width           =   4620
      Begin VB.CheckBox chkShow 
         Caption         =   "Percent"
         Height          =   195
         Index           =   14
         Left            =   3075
         TabIndex        =   22
         ToolTipText     =   "Treat the value as a percentage of the base cost"
         Top             =   360
         Width           =   1230
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Enable High Alarm"
         Height          =   195
         Index           =   13
         Left            =   360
         TabIndex        =   20
         ToolTipText     =   "Enable an alarm to trigger if price rises above a threshold"
         Top             =   30
         Width           =   1695
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Sound Alarm"
         Height          =   195
         Index           =   15
         Left            =   3075
         TabIndex        =   23
         ToolTipText     =   "Sound an audible alert when the alarm is triggered"
         Top             =   630
         Width           =   1395
      End
      Begin VB.TextBox txtSymbol 
         Height          =   285
         Index           =   7
         Left            =   1620
         TabIndex        =   21
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "If Price Rises to"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   40
         ToolTipText     =   "Value of stock above which will trigger an alarm"
         Top             =   420
         Width           =   1305
      End
   End
   Begin VB.Frame frmLowAlarm 
      ForeColor       =   &H00808080&
      Height          =   1005
      Left            =   3270
      TabIndex        =   37
      Top             =   4050
      Width           =   4620
      Begin VB.TextBox txtSymbol 
         Height          =   285
         Index           =   6
         Left            =   1620
         TabIndex        =   17
         Top             =   360
         Width           =   885
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Sound Alarm"
         Height          =   195
         Index           =   12
         Left            =   3075
         TabIndex        =   19
         ToolTipText     =   "Sound an audible alert when the alarm is triggered"
         Top             =   630
         Width           =   1395
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Enable Low Alarm"
         Height          =   195
         Index           =   10
         Left            =   360
         TabIndex        =   16
         ToolTipText     =   "Enable an alarm to trigger if price drops below a threshold"
         Top             =   30
         Width           =   1695
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Percent"
         Height          =   195
         Index           =   11
         Left            =   3075
         TabIndex        =   18
         ToolTipText     =   "Treat the value as a percentage of the base cost"
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "If Price Drops to"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   38
         ToolTipText     =   "Value of stock below which will trigger an alarm"
         Top             =   420
         Width           =   1215
      End
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "Disabled"
      Height          =   195
      Index           =   6
      Left            =   6975
      TabIndex        =   2
      ToolTipText     =   "Do not use this symbol"
      Top             =   285
      Width           =   1500
   End
   Begin VB.TextBox txtSymbol 
      Height          =   285
      Index           =   3
      Left            =   4395
      TabIndex        =   6
      ToolTipText     =   "Currency e.g. GBP, USD etc"
      Top             =   1410
      Width           =   900
   End
   Begin VB.TextBox txtSymbol 
      Height          =   285
      Index           =   5
      Left            =   7005
      TabIndex        =   7
      ToolTipText     =   "Currency symbol to use when showing the stock values"
      Top             =   1410
      Width           =   405
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   28
      Top             =   6405
      Width           =   1125
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "Add"
      Height          =   255
      Index           =   1
      Left            =   210
      TabIndex        =   9
      ToolTipText     =   "Click here to add a stock holding to the list"
      Top             =   6210
      Width           =   765
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   1050
      TabIndex        =   10
      ToolTipText     =   "Delete the currently selected item in the list"
      Top             =   6210
      Width           =   765
   End
   Begin VB.CommandButton cmdMain 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Index           =   0
      Left            =   6750
      TabIndex        =   29
      Top             =   6405
      Width           =   1125
   End
   Begin VB.TextBox txtSymbol 
      Height          =   285
      Index           =   4
      Left            =   7005
      TabIndex        =   5
      Top             =   1020
      Width           =   885
   End
   Begin VB.TextBox txtSymbol 
      Height          =   285
      Index           =   2
      Left            =   4395
      TabIndex        =   4
      Top             =   1020
      Width           =   885
   End
   Begin VB.TextBox txtSymbol 
      Height          =   285
      Index           =   1
      Left            =   4395
      TabIndex        =   3
      ToolTipText     =   "The name you want to appear on the ticker for this stock instead of the actual symbol"
      Top             =   630
      Width           =   3465
   End
   Begin VB.TextBox txtSymbol 
      Height          =   285
      Index           =   0
      Left            =   4395
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.ListBox lstSymbols 
      Height          =   5910
      Left            =   195
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   210
      Width           =   2775
   End
   Begin VB.Frame frmShow 
      Caption         =   "Show"
      ForeColor       =   &H00808080&
      Height          =   1935
      Left            =   3270
      TabIndex        =   30
      Top             =   2040
      Width           =   4620
      Begin VB.CheckBox chkShow 
         Caption         =   "Day Change"
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   14
         ToolTipText     =   "The change in value between the current price and days starting price for this stock"
         Top             =   1290
         Width           =   1905
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Day Change %"
         Height          =   195
         Index           =   8
         Left            =   2625
         TabIndex        =   24
         ToolTipText     =   "The change in percent between the current price and days starting price for this stock"
         Top             =   1290
         Width           =   1845
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Day Up/Down"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   15
         ToolTipText     =   "Show a symbol to indicate if the current price is higher or lower the days starting price"
         Top             =   1590
         Width           =   1995
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Hide Summary"
         Height          =   195
         Index           =   5
         Left            =   2625
         TabIndex        =   25
         ToolTipText     =   "Hide this symbol from the summary display and any effect it might have on it"
         Top             =   300
         Width           =   1800
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Price"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   11
         ToolTipText     =   "Current price that this stock is being traded at"
         Top             =   300
         Width           =   1695
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Profit && Loss"
         Height          =   195
         Index           =   3
         Left            =   2625
         TabIndex        =   27
         ToolTipText     =   "Show the amount of money you are up or down on the stock"
         Top             =   900
         Width           =   1845
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Up/Down"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   13
         ToolTipText     =   "Show a symbol to indicate if the current price is higher or lower the price you bought at"
         Top             =   900
         Width           =   1695
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Change %"
         Height          =   195
         Index           =   1
         Left            =   2625
         TabIndex        =   26
         ToolTipText     =   "The change in percent between the current price and the price you paid for this stock"
         Top             =   600
         Width           =   1845
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Change"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   12
         ToolTipText     =   "The change in value between the current price and the price you paid for this stock"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   150
         X2              =   4470
         Y1              =   1170
         Y2              =   1170
      End
   End
   Begin VB.Label lblDateAdded 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   4380
      TabIndex        =   0
      Top             =   1800
      Width           =   3465
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Currency Symbol"
      Height          =   195
      Index           =   0
      Left            =   5640
      TabIndex        =   36
      Top             =   1470
      Width           =   1275
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Currency Code"
      Height          =   195
      Index           =   6
      Left            =   3180
      TabIndex        =   35
      Top             =   1470
      Width           =   1125
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Shares Bought"
      Height          =   195
      Index           =   5
      Left            =   5415
      TabIndex        =   34
      Top             =   1080
      Width           =   1530
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Price Paid"
      Height          =   195
      Index           =   4
      Left            =   3540
      TabIndex        =   33
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Display Name"
      Height          =   195
      Index           =   3
      Left            =   3270
      TabIndex        =   32
      Top             =   690
      Width           =   1035
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Symbol"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   3540
      TabIndex        =   31
      ToolTipText     =   "Stock symbol for company"
      Top             =   300
      Width           =   765
   End
End
Attribute VB_Name = "frmSymbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Copyright (c) 2024, Pivotal Solutions and/or its affiliates. All rights reserved.
' Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
'
' Contains the interface for adding/editing stock transactions
'
    Dim mobjReg As New cRegistry
    Public mobjSymbols As New Collection
    Dim mbDirty As Boolean
    Dim mbDisabled As Boolean
    
    
Private Sub chkShow_Click(Index As Integer)

    ' Save the changes
    txtSymbol_Change 0

End Sub

Private Sub cmdList_Click(Index As Integer)

Dim objSymbol As New cSymbol
Dim iIndex%

    ' Remove the selected symbol
    If Index = 0 Then
        If lstSymbols.ListIndex > -1 Then
            mobjSymbols.Remove Format(lstSymbols.ItemData(lstSymbols.ListIndex))
            iIndex = lstSymbols.ListIndex
            lstSymbols.RemoveItem iIndex
            If iIndex < lstSymbols.ListCount Then
                lstSymbols.ListIndex = iIndex
            ElseIf lstSymbols.ListCount > 0 Then
                lstSymbols.ListIndex = iIndex - 1
            Else
                Z_DisplaySymbol "ZZZZ"
            End If
        End If
    
    ' Create a new symbol and add it to the list
    ElseIf Index = 1 Then
        objSymbol.RegKey = Format(DateDiff("s", DateSerial(2008, 1, 1), Now) + Format(lstSymbols.ListCount))
        objSymbol.Code = "YHOO"
        objSymbol.CurrencySymbol = "$"
        objSymbol.CurrencyName = "USD"
        objSymbol.ShowChangePercent = True
        mobjSymbols.Add objSymbol, objSymbol.RegKey
        lstSymbols.AddItem objSymbol.DisplayName
        lstSymbols.ItemData(lstSymbols.NewIndex) = CLng(objSymbol.RegKey)
        lstSymbols.ListIndex = lstSymbols.NewIndex
    End If
    Z_SetDirty True
    lstSymbols.SetFocus
    
End Sub

Private Sub cmdMain_Click(Index As Integer)

    If Index = 1 Then
        frmMain.mbForceRefresh = True
        WriteSymbolsToRegistry mobjSymbols
        mbDirty = False
        Set frmMain.mobjCurrentSymbols = Nothing
        Unload Me
    
    ElseIf Index = 0 Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    ' Position the form in the middle of the display
    CentreForm Me
    
    ' Display the symbols and select the first one
    Set mobjSymbols = ReadSymbolsFromRegistry
    Z_DisplaySymbolList
    If lstSymbols.ListCount > 0 Then lstSymbols.ListIndex = 0
    Z_SetDirty False

    'Subclass the "Form", to Capture the Listbox Notification Messages ...
    glPrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SubClassedList)
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then
        Call PSGEN_LaunchBrowser("file://" + App.path + "/user guide/index.htm#Editing_Symbols")
        KeyCode = 0
    End If
    
End Sub

Private Sub Z_DisplaySymbolList()

Dim objSymbol As cSymbol

    On Error Resume Next
    lstSymbols.Clear
    For Each objSymbol In Z_SortSymbols(mobjSymbols)
        lstSymbols.AddItem objSymbol.DisplayName
        lstSymbols.ItemData(lstSymbols.NewIndex) = CLng(objSymbol.RegKey)
    Next
    
End Sub

Private Sub Z_DisplaySymbol(ByVal sSymbolKey$)

Dim objSymbol As New cSymbol

    ' Get the symbol from the local store
    On Error Resume Next
    Set objSymbol = mobjSymbols.Item(sSymbolKey)
    
    ' Display the values to the screen
    txtSymbol(0).Text = objSymbol.Code
    txtSymbol(3).Text = objSymbol.CurrencyName
    txtSymbol(1).Text = objSymbol.Alias
    txtSymbol(2).Text = objSymbol.Price
    txtSymbol(4).Text = objSymbol.Shares
    txtSymbol(5).Text = objSymbol.CurrencySymbol
    chkShow(4).Value = IIf(objSymbol.ShowPrice, vbChecked, vbUnchecked)
    chkShow(0).Value = IIf(objSymbol.ShowChange, vbChecked, vbUnchecked)
    chkShow(1).Value = IIf(objSymbol.ShowChangePercent, vbChecked, vbUnchecked)
    chkShow(2).Value = IIf(objSymbol.ShowChangeUpDown, vbChecked, vbUnchecked)
    chkShow(3).Value = IIf(objSymbol.ShowProfitLoss, vbChecked, vbUnchecked)
    chkShow(5).Value = IIf(objSymbol.ExcludeFromSummary, vbChecked, vbUnchecked)
    chkShow(6).Value = IIf(objSymbol.Disabled, vbChecked, vbUnchecked)
    chkShow(7).Value = IIf(objSymbol.ShowDayChange, vbChecked, vbUnchecked)
    chkShow(8).Value = IIf(objSymbol.ShowDayChangePercent, vbChecked, vbUnchecked)
    chkShow(9).Value = IIf(objSymbol.ShowDayChangeUpDown, vbChecked, vbUnchecked)
    
    chkShow(10).Value = IIf(objSymbol.LowAlarmEnabled, vbChecked, vbUnchecked)
    chkShow(11).Value = IIf(objSymbol.LowAlarmIsPercent, vbChecked, vbUnchecked)
    chkShow(12).Value = IIf(objSymbol.LowAlarmSoundEnabled, vbChecked, vbUnchecked)
    txtSymbol(6).Text = objSymbol.LowAlarmValue
    
    chkShow(13).Value = IIf(objSymbol.HighAlarmEnabled, vbChecked, vbUnchecked)
    chkShow(14).Value = IIf(objSymbol.HighAlarmIsPercent, vbChecked, vbUnchecked)
    chkShow(15).Value = IIf(objSymbol.HighAlarmSoundEnabled, vbChecked, vbUnchecked)
    txtSymbol(7).Text = objSymbol.HighAlarmValue
    
    lblDateAdded.Caption = "Added: " + Format(DateAdd("s", CDbl(sSymbolKey), DateSerial(2008, 1, 1)), "d mmmm yyyy  hh:nn")
    
End Sub

Private Sub Z_SaveSymbol()

Dim objSymbol As New cSymbol
Dim iIndex%

    ' Save the values from the screen
    On Error Resume Next
    If lstSymbols.ListIndex > -1 Then
        objSymbol.RegKey = lstSymbols.ItemData(lstSymbols.ListIndex)
        objSymbol.Code = txtSymbol(0).Text
        objSymbol.CurrencyName = txtSymbol(3).Text
        objSymbol.Alias = txtSymbol(1).Text
        objSymbol.Price = PSGEN_GetLocaleValue(txtSymbol(2).Text)
        objSymbol.Shares = PSGEN_GetLocaleValue(txtSymbol(4).Text)
        objSymbol.CurrencySymbol = txtSymbol(5).Text
        objSymbol.ShowPrice = chkShow(4).Value = vbChecked
        objSymbol.ShowChange = chkShow(0).Value = vbChecked
        objSymbol.ShowChangePercent = chkShow(1).Value = vbChecked
        objSymbol.ShowChangeUpDown = chkShow(2).Value = vbChecked
        objSymbol.ShowProfitLoss = chkShow(3).Value = vbChecked
        objSymbol.ExcludeFromSummary = chkShow(5).Value = vbChecked
        objSymbol.Disabled = chkShow(6).Value = vbChecked
        objSymbol.ShowDayChange = chkShow(7).Value = vbChecked
        objSymbol.ShowDayChangePercent = chkShow(8).Value = vbChecked
        objSymbol.ShowDayChangeUpDown = chkShow(9).Value = vbChecked
        
        objSymbol.LowAlarmEnabled = chkShow(10).Value = vbChecked
        objSymbol.LowAlarmIsPercent = chkShow(11).Value = vbChecked
        objSymbol.LowAlarmSoundEnabled = chkShow(12).Value = vbChecked
        objSymbol.LowAlarmValue = PSGEN_GetLocaleValue(txtSymbol(6).Text)
        
        objSymbol.HighAlarmEnabled = chkShow(13).Value = vbChecked
        objSymbol.HighAlarmIsPercent = chkShow(14).Value = vbChecked
        objSymbol.HighAlarmSoundEnabled = chkShow(15).Value = vbChecked
        objSymbol.HighAlarmValue = PSGEN_GetLocaleValue(txtSymbol(7).Text)
        
        ' Put them into the local storage
        mobjSymbols.Remove objSymbol.RegKey
        mobjSymbols.Add objSymbol, objSymbol.RegKey
        
        ' Rearrange the list to reflect any change names
        lstSymbols.List(lstSymbols.ListIndex) = objSymbol.DisplayName
    End If
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If mbDirty Then Cancel = MsgBox("You have made changes which will be lost if you continue" + vbCrLf + vbCrLf + "Continue and lose changes ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo

End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Release the SubClassing, Very Important to Prevent Crashing!
    Call SetWindowLong(hWnd, GWL_WNDPROC, glPrevWndProc)

End Sub

Private Sub lstSymbols_Click()

    mbDisabled = True
    Debug.Print lstSymbols.ItemData(lstSymbols.ListIndex)
    Z_DisplaySymbol lstSymbols.ItemData(lstSymbols.ListIndex)
    mbDisabled = False

End Sub

Private Sub lstSymbols_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then cmdList_Click 0
    
End Sub

Private Sub txtSymbol_Change(Index As Integer)

    If Visible And Not mbDisabled And lstSymbols.ListIndex > -1 Then
        Z_SaveSymbol
        Z_SetDirty True
    End If

End Sub

Private Sub txtSymbol_GotFocus(Index As Integer)
    
    txtSymbol(Index).SelStart = 0
    txtSymbol(Index).SelLength = Len(txtSymbol(Index).Text)

End Sub

Private Sub Z_SetDirty(ByVal bValue As Boolean)

Dim iCnt%

    mbDirty = bValue
    For iCnt = 0 To txtSymbol.UBound
        txtSymbol(iCnt).Enabled = lstSymbols.ListIndex > -1
        txtSymbol(iCnt).BackColor = IIf(txtSymbol(iCnt).Enabled, vbWhite, &HE0E0E0)
    Next
    For iCnt = 0 To chkShow.UBound
        chkShow(iCnt).Enabled = lstSymbols.ListIndex > -1
    Next
    cmdList(0).Enabled = lstSymbols.ListIndex > -1
    cmdMain(1).Enabled = mbDirty

End Sub

Private Function Z_SortSymbols(col As Collection) As Collection
   Dim colNew As Collection
   Dim oCurrent As cSymbol
   Dim oCompare As cSymbol
   Dim lCompareIndex As Long
   Dim bolGreaterValueFound As Boolean

   Set colNew = New Collection

   For Each oCurrent In col

      bolGreaterValueFound = False
      lCompareIndex = 0

      For Each oCompare In colNew
         lCompareIndex = lCompareIndex + 1

        If StrComp(oCurrent.Code, oCompare.Code, vbTextCompare) < 0 Or _
           (StrComp(oCurrent.Code, oCompare.Code, vbTextCompare) = 0 And CLng(oCurrent.RegKey) > CLng(oCompare.RegKey)) Then
           bolGreaterValueFound = True
           colNew.Add oCurrent, , lCompareIndex
           Exit For
        End If
      Next oCompare
      
      If bolGreaterValueFound = False Then
         colNew.Add oCurrent
      End If

   Next oCurrent

   Set Z_SortSymbols = colNew

End Function


