VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   9945
   ClientLeft      =   30795
   ClientTop       =   2730
   ClientWidth     =   6585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9945
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdColours 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Index           =   4
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   1260
      Width           =   345
   End
   Begin VB.CommandButton cmdColours 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Index           =   5
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   1560
      Width           =   345
   End
   Begin VB.TextBox txtTwelveDataKey 
      Height          =   285
      Left            =   1770
      TabIndex        =   55
      ToolTipText     =   "The token to enable the use of the Alpha Vantage cloud API"
      Top             =   8835
      Width           =   4635
   End
   Begin VB.TextBox txtMarketStackKey 
      Height          =   285
      Left            =   1770
      TabIndex        =   53
      ToolTipText     =   "The token to enable the use of the Alpha Vantage cloud API"
      Top             =   8490
      Width           =   4635
   End
   Begin VB.TextBox txtAlphaVantageKey 
      Height          =   285
      Left            =   1770
      TabIndex        =   26
      ToolTipText     =   "The token to enable the use of the Alpha Vantage cloud API"
      Top             =   8130
      Width           =   4635
   End
   Begin VB.TextBox txtIexKey 
      Height          =   285
      Left            =   1770
      TabIndex        =   25
      ToolTipText     =   "The token to enable the use of the IEX cloud API"
      Top             =   7770
      Width           =   4635
   End
   Begin VB.CheckBox chkSummarise 
      Caption         =   "Show average for stocks with multiple buys"
      Height          =   195
      Left            =   690
      TabIndex        =   49
      ToolTipText     =   "Will show a single summary for stocks of the same name and average their cost"
      Top             =   5460
      Width           =   5910
   End
   Begin VB.CommandButton cmdLowAlarm 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6030
      TabIndex        =   13
      ToolTipText     =   "Select a wave file"
      Top             =   2670
      Width           =   315
   End
   Begin VB.CommandButton cmdHighAlarm 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6030
      TabIndex        =   11
      ToolTipText     =   "Select a wave file"
      Top             =   2340
      Width           =   315
   End
   Begin VB.TextBox txtLowalarm 
      Height          =   285
      Left            =   1770
      TabIndex        =   12
      ToolTipText     =   "An alternative wave file to use when a low alarm is triggered"
      Top             =   2670
      Width           =   4155
   End
   Begin VB.TextBox txtHighAlarm 
      Height          =   285
      Left            =   1770
      TabIndex        =   10
      ToolTipText     =   "An alternative wave file to use when a high alarm is triggered"
      Top             =   2340
      Width           =   4155
   End
   Begin VB.CheckBox chkShowTotalPercent 
      Caption         =   "Show Total Profit && Loss as a Percentage"
      Height          =   195
      Left            =   690
      TabIndex        =   15
      ToolTipText     =   "Set this if you would like an overall position to be displayed as a percentage of profit/loss over totla cost"
      Top             =   3660
      Width           =   5910
   End
   Begin VB.CheckBox chkAlwaysOnTop 
      Caption         =   "Always on top"
      Height          =   195
      Left            =   4170
      TabIndex        =   2
      ToolTipText     =   "Make sure that the ticker is always shown over the top of other windows"
      Top             =   540
      Width           =   2220
   End
   Begin VB.TextBox txtFrequency 
      Height          =   285
      Left            =   1770
      TabIndex        =   1
      ToolTipText     =   "How often to check for data changes (1-600 seconds)"
      Top             =   480
      Width           =   645
   End
   Begin VB.TextBox txtProxy 
      Height          =   285
      Left            =   1770
      TabIndex        =   0
      ToolTipText     =   "The address of a proxy server to use e.g. www.myproxy.com:8989 or 192.168.0.1:3180 etc"
      Top             =   150
      Width           =   4635
   End
   Begin VB.TextBox txtMargin 
      Height          =   285
      Left            =   4185
      TabIndex        =   24
      ToolTipText     =   "Amount of money you have in your debit (margin) account"
      Top             =   6975
      Width           =   1080
   End
   Begin VB.TextBox txtTotal 
      Height          =   285
      Left            =   1800
      TabIndex        =   23
      ToolTipText     =   "Total amount you have invested in stocks in your local currency"
      Top             =   6975
      Width           =   1080
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Restore"
      Height          =   390
      Index           =   3
      Left            =   1350
      TabIndex        =   28
      ToolTipText     =   "Restore your settings and symbols from a backup file"
      Top             =   9330
      Width           =   1050
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Backup"
      Height          =   390
      Index           =   2
      Left            =   180
      TabIndex        =   27
      ToolTipText     =   "Backup your settings and symbols to a file"
      Top             =   9330
      Width           =   1050
   End
   Begin VB.CheckBox chkShowTotalValue 
      Caption         =   "Show Total Current Value of Investments"
      Height          =   195
      Left            =   690
      TabIndex        =   17
      ToolTipText     =   "Set this if you want to see the total current value of all stocks"
      Top             =   4215
      Width           =   5910
   End
   Begin VB.CheckBox chkShowTotalCost 
      Caption         =   "Show Total Cost of Investments"
      Height          =   195
      Left            =   690
      TabIndex        =   16
      ToolTipText     =   "Set this if you want to see the total amount invested in all stocks"
      Top             =   3930
      Width           =   5910
   End
   Begin VB.TextBox txtCurrencySymbol 
      Height          =   285
      Left            =   4200
      TabIndex        =   22
      Top             =   6180
      Width           =   435
   End
   Begin VB.TextBox txtCurrency 
      Height          =   285
      Left            =   1800
      TabIndex        =   21
      ToolTipText     =   "Currency to convert summary values into e.g. GBP, USD etc"
      Top             =   6180
      Width           =   645
   End
   Begin VB.CheckBox chkItalic 
      Caption         =   "Italic"
      Height          =   195
      Left            =   4980
      TabIndex        =   9
      Top             =   2010
      Width           =   960
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Bold"
      Height          =   195
      Left            =   4140
      TabIndex        =   8
      Top             =   2010
      Width           =   960
   End
   Begin VB.ComboBox lstFont 
      Height          =   315
      Left            =   1785
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1950
      Width           =   2070
   End
   Begin VB.CommandButton cmdColours 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Index           =   3
      Left            =   3495
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1575
      Width           =   345
   End
   Begin VB.CommandButton cmdColours 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Index           =   2
      Left            =   3495
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1275
      Width           =   345
   End
   Begin VB.CommandButton cmdColours 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   210
      Index           =   1
      Left            =   1845
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1605
      Width           =   345
   End
   Begin VB.CommandButton cmdColours 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Index           =   0
      Left            =   1845
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1275
      Width           =   345
   End
   Begin VB.CheckBox chkShowPercent 
      Caption         =   "Show the Difference of Base Cost Against Current Price as Percentage"
      Height          =   195
      Left            =   690
      TabIndex        =   20
      Top             =   5175
      Width           =   5910
   End
   Begin VB.CheckBox chkShowPrice 
      Caption         =   "Show the Current Price of each Stock"
      Height          =   195
      Left            =   690
      TabIndex        =   18
      ToolTipText     =   "Set this if you would like an overall position to be displayed"
      Top             =   4590
      Width           =   5910
   End
   Begin VB.CheckBox chkShowCostBase 
      Caption         =   "Show the Average Cost of each Stock (Cost Base)"
      Height          =   195
      Left            =   690
      TabIndex        =   19
      ToolTipText     =   "Set this if you would like an overall position to be displayed"
      Top             =   4875
      Width           =   5910
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   390
      Index           =   1
      Left            =   4020
      TabIndex        =   29
      Top             =   9375
      Width           =   1125
   End
   Begin VB.CommandButton cmdMain 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Index           =   0
      Left            =   5280
      TabIndex        =   30
      Top             =   9360
      Width           =   1125
   End
   Begin VB.CheckBox chkShowTotal 
      Caption         =   "Show Total Profit && Loss"
      Height          =   195
      Left            =   690
      TabIndex        =   14
      ToolTipText     =   "Set this if you would like an overall position to be displayed"
      Top             =   3405
      Width           =   5910
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Up Arrow Colour"
      Height          =   225
      Index           =   24
      Left            =   3750
      TabIndex        =   60
      Top             =   1290
      Width           =   1755
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Down Arrow Colour"
      Height          =   225
      Index           =   23
      Left            =   3945
      TabIndex        =   59
      Top             =   1590
      Width           =   1560
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TwelveData Token"
      Height          =   225
      Index           =   22
      Left            =   -60
      TabIndex        =   56
      Top             =   8865
      Width           =   1755
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MarketStack Token"
      Height          =   225
      Index           =   21
      Left            =   -60
      TabIndex        =   54
      Top             =   8520
      Width           =   1755
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "AlphaVantage Token"
      Height          =   225
      Index           =   20
      Left            =   -60
      TabIndex        =   52
      Top             =   8160
      Width           =   1755
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IEX Token"
      Height          =   225
      Index           =   19
      Left            =   -60
      TabIndex        =   51
      Top             =   7800
      Width           =   1755
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  'Center
      Caption         =   "API Keys"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   0
      Left            =   390
      TabIndex        =   50
      Top             =   7470
      Width           =   1095
   End
   Begin VB.Line ctlLine 
      BorderColor     =   &H00808080&
      Index           =   8
      X1              =   150
      X2              =   6385
      Y1              =   7575
      Y2              =   7575
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Low Alarm"
      Height          =   225
      Index           =   18
      Left            =   -60
      TabIndex        =   48
      Top             =   2730
      Width           =   1755
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "High Alarm"
      Height          =   225
      Index           =   17
      Left            =   -60
      TabIndex        =   47
      Top             =   2370
      Width           =   1755
   End
   Begin VB.Label lblLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Seconds"
      Height          =   225
      Index           =   16
      Left            =   2520
      TabIndex        =   46
      Top             =   540
      Width           =   825
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Update Every"
      Height          =   225
      Index           =   15
      Left            =   -60
      TabIndex        =   45
      Top             =   540
      Width           =   1755
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Proxy Server"
      Height          =   225
      Index           =   14
      Left            =   -60
      TabIndex        =   44
      Top             =   180
      Width           =   1755
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  'Center
      Caption         =   "Investment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   13
      Left            =   390
      TabIndex        =   43
      Top             =   6690
      Width           =   1095
   End
   Begin VB.Line ctlLine 
      BorderColor     =   &H00FFFFFF&
      Index           =   7
      X1              =   150
      X2              =   6385
      Y1              =   6795
      Y2              =   6795
   End
   Begin VB.Line ctlLine 
      BorderColor     =   &H00808080&
      Index           =   6
      X1              =   150
      X2              =   6385
      Y1              =   6765
      Y2              =   6765
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Margin"
      Height          =   225
      Index           =   12
      Left            =   2355
      TabIndex        =   42
      Top             =   7035
      Width           =   1755
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Investment"
      Height          =   225
      Index           =   11
      Left            =   -30
      TabIndex        =   41
      Top             =   7035
      Width           =   1755
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Currency Symbol"
      Height          =   225
      Index           =   10
      Left            =   2370
      TabIndex        =   40
      Top             =   6240
      Width           =   1755
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  'Center
      Caption         =   "Currency Conversion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   9
      Left            =   405
      TabIndex        =   39
      Top             =   5865
      Width           =   1935
   End
   Begin VB.Line ctlLine 
      BorderColor     =   &H00808080&
      Index           =   5
      X1              =   165
      X2              =   6400
      Y1              =   5940
      Y2              =   5940
   End
   Begin VB.Line ctlLine 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   165
      X2              =   6400
      Y1              =   5970
      Y2              =   5970
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Currency Code"
      Height          =   225
      Index           =   8
      Left            =   -30
      TabIndex        =   38
      Top             =   6240
      Width           =   1755
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Font"
      Height          =   225
      Index           =   7
      Left            =   540
      TabIndex        =   37
      Top             =   1995
      Width           =   1155
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Down Colour"
      Height          =   225
      Index           =   6
      Left            =   2205
      TabIndex        =   36
      Top             =   1605
      Width           =   1155
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Up Colour"
      Height          =   225
      Index           =   5
      Left            =   1605
      TabIndex        =   35
      Top             =   1305
      Width           =   1755
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Normal Text"
      Height          =   195
      Index           =   4
      Left            =   555
      TabIndex        =   34
      Top             =   1635
      Width           =   1155
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  'Center
      Caption         =   "Colours, Fonts && Sounds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   3
      Left            =   495
      TabIndex        =   33
      Top             =   960
      Width           =   2190
   End
   Begin VB.Line ctlLine 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   180
      X2              =   6400
      Y1              =   1065
      Y2              =   1065
   End
   Begin VB.Line ctlLine 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   180
      X2              =   6400
      Y1              =   1035
      Y2              =   1035
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Background"
      Height          =   225
      Index           =   2
      Left            =   -45
      TabIndex        =   32
      Top             =   1305
      Width           =   1755
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  'Center
      Caption         =   "Summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   1
      Left            =   450
      TabIndex        =   31
      Top             =   3075
      Width           =   870
   End
   Begin VB.Line ctlLine 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   165
      X2              =   6400
      Y1              =   3180
      Y2              =   3180
   End
   Begin VB.Line ctlLine 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   165
      X2              =   6400
      Y1              =   3150
      Y2              =   3150
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim mobjReg As New cRegistry
    Dim mbDirty As Boolean
    
Private Sub Check1_Click()

End Sub

Private Sub chkAlwaysOnTop_Click()
    
    Z_SetDirty True

End Sub

Private Sub chkBold_Click()

    Z_SetDirty True
    
End Sub


Private Sub chkItalic_Click()

    Z_SetDirty True
    
End Sub

Private Sub chkShowCostBase_Click()

    Z_SetDirty True
    
End Sub

Private Sub chkShowPercent_Click()

    Z_SetDirty True
    
End Sub

Private Sub chkShowPrice_Click()

    Z_SetDirty True
    
End Sub

Private Sub chkShowTotal_Click()

    Z_SetDirty True
    
End Sub

Private Sub chkShowTotalCost_Click()

    Z_SetDirty True

End Sub

Private Sub chkShowTotalPercent_Click()

    Z_SetDirty True
    
End Sub

Private Sub chkShowTotalValue_Click()

    Z_SetDirty True

End Sub

Private Sub chkSummarise_Click()

    Z_SetDirty True
    
End Sub

Private Sub cmdColours_Click(Index As Integer)

Dim lColour&
    
    lColour = PSGEN_ChooseColor(hWnd, cmdColours(Index).BackColor)
    If cmdColours(Index).BackColor <> lColour Then
        cmdColours(Index).BackColor = lColour
        Z_SetDirty True
    End If
    
End Sub

Private Sub cmdHighAlarm_Click()

Dim sFilename$

        sFilename = PSGEN_SelectOpenFile(cmdHighAlarm.hWnd, "Wave Files" + vbNullChar + "*.wav", OFN_EXPLORER + OFN_OVERWRITEPROMPT + OFN_PATHMUSTEXIST + OFN_SHAREAWARE, "Select a Wave File")
        If sFilename <> "" Then
            txtHighAlarm.Text = sFilename
        End If

End Sub

Private Sub cmdLowAlarm_Click()

Dim sFilename$

        sFilename = PSGEN_SelectOpenFile(cmdHighAlarm.hWnd, "Wave Files" + vbNullChar + "*.wav", OFN_EXPLORER + OFN_OVERWRITEPROMPT + OFN_PATHMUSTEXIST + OFN_SHAREAWARE, "Select a Wave File")
        If sFilename <> "" Then
            txtLowalarm.Text = sFilename
        End If

End Sub

Private Sub cmdMain_Click(Index As Integer)

Dim sFilename$
Dim lTmp&

    On Error Resume Next
    If Index = 1 Then
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_PROXY, txtProxy.Text
        lTmp = REG_FREQUENCY_DEF
        If IsNumeric(txtFrequency.Text) Then
            lTmp = CLng(txtFrequency.Text)
            If lTmp < 1 Or lTmp > 30 Then lTmp = REG_FREQUENCY_DEF
        End If
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_FREQUENCY, Format(lTmp)
        frmMain.timData.Interval = lTmp * 1000
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_SUMMARY_CURRENCY, txtCurrency.Text
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_SUMMARY_CURRENCY_SYMBOL, txtCurrencySymbol.Text
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_SUMMARY_TOTAL, txtTotal.Text
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_SUMMARY_MARGIN, txtMargin.Text
        
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_HIGH_ALARM_WAVE, txtHighAlarm.Text
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_HIGH_LOW_WAVE, txtLowalarm.Text
        
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_PROFIT_LOSS, chkShowTotal.Value = vbChecked
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_PROFIT_LOSS_PERCENT, chkShowTotalPercent.Value = vbChecked
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_TOTAL_COST, chkShowTotalCost.Value = vbChecked
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_TOTAL_VALUE, chkShowTotalValue.Value = vbChecked
        
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_COST_BASE, chkShowCostBase.Value = vbChecked
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_PRICE, chkShowCostBase.Value = vbChecked
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_PERCENT, chkShowPercent.Value = vbChecked
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_SUMMARISE, chkSummarise.Value = vbChecked
        
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_BACK_COLOUR, cmdColours(0).BackColor
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_TEXT_COLOUR, cmdColours(1).BackColor
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_UP_COLOUR, cmdColours(2).BackColor
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_DOWN_COLOUR, cmdColours(3).BackColor
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_UP_ARROW_COLOUR, cmdColours(4).BackColor
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_DOWN_ARROW_COLOUR, cmdColours(5).BackColor
        
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_BOLD, chkBold.Value = vbChecked
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_ITALIC, chkItalic.Value = vbChecked
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_FONT, lstFont.Text
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_ALWAYS_ON_TOP, chkAlwaysOnTop.Value = vbChecked
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_IEX_KEY, txtIexKey.Text
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_ALPHA_VANTAGE_KEY, txtAlphaVantageKey.Text
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_MARKET_STACK_KEY, txtMarketStackKey.Text
        mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_TWELVE_DATA_KEY, txtTwelveDataKey.Text
        mbDirty = False
        Unload Me
    
    ElseIf Index = 0 Then
        Unload Me
    
    '
    ' Backup the registry to a file
    '
    ElseIf Index = 2 Then
        sFilename = PSGEN_SelectSaveFile(cmdMain(Index).hWnd, "Backup Files" + vbNullChar + "*.bck", OFN_EXPLORER + OFN_OVERWRITEPROMPT + OFN_PATHMUSTEXIST + OFN_SHAREAWARE, "Save to File")
        If sFilename <> "" Then
            Call Shell("regedit /E """ + sFilename + """ ""HKEY_LOCAL_MACHINE\SOFTWARE\Pivotal\" + App.Title + """", vbHide)
        End If
    
    ElseIf Index = 3 Then
        sFilename = PSGEN_SelectOpenFile(cmdMain(Index).hWnd, "Backup Files" + vbNullChar + "*.bck", OFN_EXPLORER + OFN_OVERWRITEPROMPT + OFN_PATHMUSTEXIST + OFN_SHAREAWARE, "Save to File")
        If sFilename <> "" Then
            If MsgBox("Are you sure you want to restore these settings and lose your current values?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                
                '
                ' Backup the original
                '
                Call Shell("regedit /E """ + App.Path + "\backup_" + Format(Now, "ddmmyy_hhNNss") + ".bck"" ""HKEY_LOCAL_MACHINE\SOFTWARE\Pivotal\" + App.Title + """", vbHide)
                
                '
                ' Now delete the original and load the new values
                '
                mobjReg.DeleteSetting App.Title, REG_SETTINGS
                mobjReg.DeleteSetting App.Title, REG_SYMBOLS
                mobjReg.DeleteSettingEx RegLocalMachine, "SOFTWARE\Pivotal\" + App.Title
                Call Shell("regedit /s """ + sFilename + """", vbHide)
            End If
        End If
    
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then
        Call PSGEN_LaunchBrowser("file://" + App.Path + "/user guide/index.htm#Editing_Settings")
        KeyCode = 0
    End If
    
End Sub

Private Sub Form_Load()

Dim sFont$
Dim iCnt%

    '
    ' Position the form in the middle of the display
    '
    CentreForm Me
    
    '
    ' Display the normal stuff
    '
    txtProxy.Text = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_PROXY)
    txtFrequency.Text = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_FREQUENCY, Format(REG_FREQUENCY_DEF))
    txtCurrency.Text = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_CURRENCY)
    txtCurrencySymbol.Text = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_CURRENCY_SYMBOL)
    txtTotal.Text = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_TOTAL)
    txtMargin.Text = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_MARGIN)
    
    txtHighAlarm.Text = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_HIGH_ALARM_WAVE)
    txtLowalarm.Text = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_LOW_ALARM_WAVE)
    
    chkShowTotal.Value = IIf(CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_PROFIT_LOSS, "0")), vbChecked, vbUnchecked)
    chkShowTotalPercent.Value = IIf(CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_PROFIT_LOSS_PERCENT, "0")), vbChecked, vbUnchecked)
    chkShowTotalCost.Value = IIf(CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_TOTAL_COST, "0")), vbChecked, vbUnchecked)
    chkShowTotalValue.Value = IIf(CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_TOTAL_VALUE, "0")), vbChecked, vbUnchecked)
    
    chkShowCostBase.Value = IIf(CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_COST_BASE, "0")), vbChecked, vbUnchecked)
    chkShowPrice.Value = IIf(CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_PRICE, "0")), vbChecked, vbUnchecked)
    chkShowPercent.Value = IIf(CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_PERCENT, "0")), vbChecked, vbUnchecked)
    chkSummarise.Value = IIf(CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_SUMMARISE, "0")), vbChecked, vbUnchecked)
    
    cmdColours(0).BackColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_BACK_COLOUR, Format(vbBlack)))
    cmdColours(1).BackColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_TEXT_COLOUR, Format(vbWhite)))
    cmdColours(2).BackColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_UP_COLOUR, Format(vbGreen)))
    cmdColours(3).BackColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_DOWN_COLOUR, Format(vbRed)))
    cmdColours(4).BackColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_UP_ARROW_COLOUR, Format(vbGreen)))
    cmdColours(5).BackColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_DOWN_ARROW_COLOUR, Format(vbRed)))
    chkBold.Value = IIf(CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_BOLD, "0")), vbChecked, vbUnchecked)
    chkAlwaysOnTop.Value = IIf(CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_ALWAYS_ON_TOP, "-1")), vbChecked, vbUnchecked)
    chkItalic.Value = IIf(CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_ITALIC, "0")), vbChecked, vbUnchecked)
    txtIexKey.Text = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_IEX_KEY)
    txtAlphaVantageKey.Text = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_ALPHA_VANTAGE_KEY)
    txtMarketStackKey.Text = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_MARKET_STACK_KEY)
    txtTwelveDataKey.Text = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_TWELVE_DATA_KEY)
    
    sFont = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_FONT, frmMain.Font.Name)
    For iCnt = 0 To Screen.FontCount - 1
        lstFont.AddItem Screen.Fonts(iCnt)
        If lstFont.List(lstFont.NewIndex) = sFont Then lstFont.ListIndex = lstFont.NewIndex
    Next
    
    Z_SetDirty False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If mbDirty Then Cancel = MsgBox("You have made changes which will be lost if you continue" + vbCrLf + vbCrLf + "Continue and lose changes ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo

End Sub


Private Sub lstFont_Click()

    Z_SetDirty True
    
End Sub

Private Sub txtAlphaVantageKey_Change()

    Z_SetDirty True

End Sub

Private Sub txtCurrency_Change()

    Z_SetDirty True

End Sub

Private Sub txtCurrencySymbol_Change()

    Z_SetDirty True

End Sub

Private Sub txtHighAlarm_Change()

    Z_SetDirty True

End Sub

Private Sub txtIexKey_Change()

    Z_SetDirty True

End Sub

Private Sub txtLowalarm_Change()

    Z_SetDirty True

End Sub

Private Sub txtMargin_Change()

    Z_SetDirty True

End Sub

Private Sub txtMarketStackKey_Change()

    Z_SetDirty True

End Sub

Private Sub txtProxy_Change()

    Z_SetDirty True

End Sub

Private Sub txtFrequency_Change()

    Z_SetDirty True

End Sub

Private Sub txtTotal_Change()

    Z_SetDirty True

End Sub

Private Sub Z_SetDirty(ByVal bValue As Boolean)

Dim iCnt%

    mbDirty = bValue
    cmdMain(1).Enabled = mbDirty

End Sub

Private Sub txtTwelveDataKey_Change()

    Z_SetDirty True

End Sub
