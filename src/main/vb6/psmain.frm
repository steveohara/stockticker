VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   2130
   ClientLeft      =   9810
   ClientTop       =   2190
   ClientWidth     =   5880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000C0C0&
   Icon            =   "psmain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   142
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   392
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timMouse 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3615
      Top             =   1185
   End
   Begin VB.PictureBox picData 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   330
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1380
      Width           =   1935
   End
   Begin VB.PictureBox picSize 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   495
      Index           =   0
      Left            =   150
      MousePointer    =   9  'Size W E
      ScaleHeight     =   495
      ScaleWidth      =   75
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   75
   End
   Begin VB.PictureBox picSize 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   495
      Index           =   1
      Left            =   4440
      MousePointer    =   9  'Size W E
      ScaleHeight     =   495
      ScaleWidth      =   75
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   75
   End
   Begin VB.PictureBox picText 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1035
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   91
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   1365
   End
   Begin VB.Timer timData 
      Interval        =   60000
      Left            =   4755
      Top             =   450
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuEditSymbols 
         Caption         =   "Edit Symbols"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Edit Settings"
      End
      Begin VB.Menu mnuSpc2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFontSize 
         Caption         =   "Font Size"
         Begin VB.Menu mnuFontSize7pt 
            Caption         =   "7pt"
         End
         Begin VB.Menu mnuFontSize8pt 
            Caption         =   "8pt"
         End
         Begin VB.Menu mnuFontSize9pt 
            Caption         =   "9pt"
         End
         Begin VB.Menu mnuFontSize10pt 
            Caption         =   "10pt"
         End
         Begin VB.Menu mnuFontSize12pt 
            Caption         =   "12pt"
         End
      End
      Begin VB.Menu mnuSpc1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScroll 
         Caption         =   "Scroll"
         Begin VB.Menu mnuScrollSlow 
            Caption         =   "Slow"
         End
         Begin VB.Menu mnuScrollMedium 
            Caption         =   "Medium"
         End
         Begin VB.Menu mnuScrollFast 
            Caption         =   "Fast"
         End
      End
      Begin VB.Menu mnuSpc7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDock 
         Caption         =   "Dock"
         Begin VB.Menu mnuDockNone 
            Caption         =   "No Docking"
         End
         Begin VB.Menu mnuDockTop 
            Caption         =   "Top"
         End
         Begin VB.Menu mnuDockBottom 
            Caption         =   "Bottom"
         End
         Begin VB.Menu mnuSpc8 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDockAutoHide 
            Caption         =   "Auto Hide"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuSpc5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuSpc3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTopMost 
         Caption         =   "Keep the ticker on top of other windows"
      End
      Begin VB.Menu mnuRunAtStartup 
         Caption         =   "Run at Startup"
      End
      Begin VB.Menu mnuSpc4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
      Begin VB.Menu mnuUpgrade 
         Caption         =   "Check for new version"
      End
      Begin VB.Menu mnuSpc6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export"
         Begin VB.Menu mnuExportAll 
            Caption         =   "Export to CSV (All)"
         End
         Begin VB.Menu mnuExportLive 
            Caption         =   "Export to CSV (Live)"
         End
         Begin VB.Menu mnuExportSummarised 
            Caption         =   "Export to CSV (Summarised)"
         End
      End
      Begin VB.Menu mnuSpc9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Copyright (c) 2024, Pivotal Solutions and/or its affiliates. All rights reserved.
' Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
'
' This the main ticker 'bar' and contains both the summarisation
' and scrolling portion of the ticker display
'
Option Explicit

    ' Registry entries
    Const REG_POSITION = "Position"
    Const REG_FONTSIZE = "Font Size"
    Const REG_SCROLLSPEED = "Scroll Speed"
    Const REG_SCROLLSPEED_SLOW = "1,40"
    Const REG_SCROLLSPEED_MEDIUM = "1,30"
    Const REG_SCROLLSPEED_FAST = "2,20"
    
    Const ALPHA_VANTAGE_KEY = "7856O2Q3ZWKM0EOH"
    
    Dim mstPoint As POINTAPI
    Dim mbCapturing As Boolean
    Dim mobjReg As New cRegistry
    Dim msURL$
    Dim mDataHasChanged As Boolean
    Dim mlCurrentWidth&
    Dim mbScrolling As Boolean
    
    Dim mlDockHandle&
    Dim mstDock As APPBARDATA
    
    Public mobjTotal As New cStock
    
    Public mobjSummaryStocks As New Collection
    Public mobjCurrentSymbols As New Collection
    Public mobjExchangeRates As New Collection
    Public mbForceRefresh As Boolean
    
    Private mlTimer&
    Public miScrollPosition%
    Private miScrollInterval%
    Private miScrollMovement%
    
    Public mrSummaryRegionStartX As Single
    Public mrSummaryRegionEndX As Single
    
    Public mrDaySummaryRegionStartX As Single
    Public mrDaySummaryRegionEndX As Single
    
    Dim mfrmPrevaiew As frmPreview
    Dim mlChartLagger&
    Dim picSelectedSizer As PictureBox
    
    
Private Sub Z_DisplaySymbols()
Attribute Z_DisplaySymbols.VB_Description = "Displays symbols for the scrolling portion"
'
' Displays symbols for the scrolling portion
'
Dim iCnt%
Dim lBmp&
Dim lBufferDC&
Dim stRect As RECT
Dim lBrush&

    ' Output the data and scroll if neccersary
    On Error Resume Next
    'Debug.Print Format(GetTickCount) + " Displaying symbols"
    
    lBufferDC = CreateCompatibleDC(picText.hDC)
    lBmp = CreateCompatibleBitmap(picText.hDC, 5000, picText.ScaleHeight)
    Call SelectObject(lBufferDC, lBmp)
    
    ' Copy the data image into the buffer
    BitBlt lBufferDC, 0, 0, 5000, picText.ScaleHeight, picData.hDC, 0, 0, SRCCOPY
    If mbScrolling Then BitBlt lBufferDC, picData.CurrentX, 0, 5000, picText.ScaleHeight, picData.hDC, 0, 0, SRCCOPY
    
    ' Copy the buffer to the screen
    BitBlt picText.hDC, 0, 0, 5000, picText.ScaleHeight, lBufferDC, miScrollPosition, 0, SRCCOPY
    Call DeleteDC(lBufferDC)
    Call DeleteObject(lBmp)
    
    ' Adjust the scroll position
    If Not mbScrolling Then
        miScrollPosition = 0
    Else
        iCnt = miScrollPosition + miScrollMovement
        If iCnt >= picData.CurrentX Then iCnt = 0
        miScrollPosition = iCnt
    End If
    
End Sub


Private Sub Z_ShowYahooChart(ByVal sSymbols)
'
' Shows the admin page of choice
'
Dim sURL$

    ' Initialise error vector
    On Error Resume Next
    sURL = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_LAUNCH_URL)
    If sURL = "" Then sURL = REG_LAUNCH_URL_DEF
    sURL = sURL + "/" + sSymbols + "?p=" + sSymbols
    Call PSGEN_LaunchBrowser(sURL)

End Sub

Private Sub Z_DisplayData(Optional ByVal sData = "")
'
'                     sData$             - Data to display
'                     bFlash As Boolean            - Flash the display
'
' Displays data on the screen
'
Dim objStock As cStock
Dim iCnt%, iImageWidth%, iDisplayWidth%
Dim bShowTotal As Boolean
Dim bShowTotalPercent As Boolean
Dim bShowCostBase As Boolean
Dim bShowPrice As Boolean
Dim bShowPercent As Boolean
Dim bShowSummary As Boolean
Dim bShowTotalCost As Boolean
Dim bShowTotalValue As Boolean
Dim bShowDailyChange As Boolean
Dim bBold As Boolean
Dim bItalic As Boolean
Dim bAlwaysOnTop As Boolean
Dim bShown As Boolean
Dim sLeader$, sFont$, sCurrencyName$, sCurrencySymbol$
Dim lBackColor&, lTextColor&, lUpColor&, lDownColor&
Dim rTotalInvestment#, rMargin#, rRate#, rTotalChange#
Dim objSymbol As cSymbol

    ' Draw the text on the display
    On Error Resume Next
    Set picData.Font = Font
    bShowTotal = CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_PROFIT_LOSS, "0"))
    bShowTotalPercent = CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_PROFIT_LOSS_PERCENT, "0"))
    bShowTotalValue = CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_TOTAL_VALUE, "0"))
    bShowDailyChange = CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_DAILY_CHANGE, "0"))
    bShowTotalCost = CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_TOTAL_COST, "0"))
    bShowCostBase = CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_COST_BASE, "0"))
    bShowPrice = CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_PRICE, "0"))
    bShowPercent = CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_PERCENT, "0"))
    bShowSummary = bShowTotal Or bShowTotalPercent Or bShowTotalCost Or bShowTotalValue Or bShowCostBase Or bShowPrice Or bShowPercent
    
    lBackColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_BACK_COLOUR, Format(vbBlack)))
    lTextColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_TEXT_COLOUR, Format(vbWhite)))
    lUpColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_UP_COLOUR, Format(vbGreen)))
    lDownColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_DOWN_COLOUR, Format(vbRed)))
    
    If PSGEN_IsCommaLocale Then
        rTotalInvestment = CDbl("0" + Replace(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_TOTAL, "0"), ".", ","))
        rMargin = CDbl("0" + Replace(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_MARGIN, "0"), ".", ","))
    Else
        rTotalInvestment = CDbl("0" + Replace(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_TOTAL, "0"), ",", "."))
        rMargin = CDbl("0" + Replace(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_MARGIN, "0"), ",", "."))
    End If
    
    bBold = CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_BOLD, "0"))
    bItalic = CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_ITALIC, "0"))
    sFont = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_FONT, Font.Name)
    bAlwaysOnTop = CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_ALWAYS_ON_TOP, "-1"))
    sCurrencySymbol = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_CURRENCY_SYMBOL, "�")
    
    ' Position the display elements
    Cls
    Font.Name = sFont
    Font.Bold = bBold
    Font.Italic = bItalic
    BackColor = lBackColor
    picData.BackColor = lBackColor
    picData.Font.Name = sFont
    picData.Font.Bold = bBold
    picData.Font.Italic = bItalic
    iDisplayWidth = ScaleWidth - (2 * picSize(0).Width) - 4
    CurrentX = picSize(0).Width + 2 + iImageWidth
    CurrentY = 1
    
    ' Draw the summary (non-scrolling)
    mrSummaryRegionStartX = 0
    mrSummaryRegionEndX = 0
    If bShowSummary Then
        mrSummaryRegionStartX = CurrentX
        ForeColor = &HC0C0&
        Print "Summary:";
        If bShowTotal Then
            ForeColor = IIf(mobjTotal.TotalValue > mobjTotal.TotalCost, lUpColor, IIf(mobjTotal.TotalValue < mobjTotal.TotalCost, lDownColor, lTextColor))
            Print "  " + mobjTotal.FormattedLoss;
            If rTotalInvestment > 0 Then
                ForeColor = IIf(mobjTotal.TotalValue > (rTotalInvestment + rMargin), lUpColor, IIf(mobjTotal.TotalValue < (rTotalInvestment + rMargin), lDownColor, lTextColor))
                Print " (" + mobjTotal.FormattedLossAdjusted(rTotalInvestment + rMargin) + ")";
            End If
        End If
        
        If bShowTotalPercent Then
            ForeColor = IIf(mobjTotal.TotalValue > mobjTotal.TotalCost, lUpColor, IIf(mobjTotal.TotalValue < mobjTotal.TotalCost, lDownColor, lTextColor))
            Print "  " + mobjTotal.FormattedLossPercent;
        End If
        
        ForeColor = lTextColor
        If bShowTotalCost Then Print "  Cost:" + mobjTotal.FormattedCost;
        If bShowTotalValue Then Print "  Val:" + mobjTotal.FormattedValue;
        
        If bShowTotal Or bShowTotalCost Or bShowTotalValue Then
            CurrentX = CurrentX + 4
            Line (CurrentX, 0)-(CurrentX, ScaleHeight), &H808080
            CurrentY = 1
        End If
        mrSummaryRegionEndX = CurrentX
        
        If bShowPrice Or bShowCostBase Or bShowPercent Or bShowDailyChange Then
            bShown = False
            rTotalChange = 0
            For Each objStock In mobjSummaryStocks
                If bShowPrice Or bShowPercent Then
                    objStock.Position.Left = CurrentX
                    ForeColor = IIf(objStock.CurrentPrice > objStock.AverageCost, lUpColor, IIf(objStock.CurrentPrice < objStock.AverageCost, lDownColor, lTextColor))
                    CurrentX = CurrentX + IIf(bShown, 10, 6)
                    Print objStock.DisplayName;
                    If bShowPrice Then Print " " + objStock.FormattedPrice;
                    If bShowPercent Then Print " " + objStock.FormattedBasePercent;
                    objStock.Position.Right = CurrentX
                    bShown = True
                End If
                ForeColor = lTextColor
                If bShowCostBase Then
                    Print " " + IIf(bShowPrice Or bShowPercent, "", "   " + objStock.Code + " ") + objStock.FormattedAverageCost;
                End If
                
                ' Get the total daily change
                Set objSymbol = mobjCurrentSymbols.Item(objStock.Code)
                rTotalChange = rTotalChange + (ConvertCurrency(objSymbol, objSymbol.DayChange) * objStock.NumberOfShares)
            Next
            mrDaySummaryRegionStartX = 0
            mrDaySummaryRegionEndX = 0
            If bShowDailyChange Then
                mrDaySummaryRegionStartX = CurrentX
                CurrentX = CurrentX + IIf(bShown, 10, 6)
                ForeColor = &HC0C0&
                Print "Today:";
                ForeColor = IIf(rTotalChange > 0, lUpColor, IIf(rTotalChange < 0, lDownColor, lTextColor))
                CurrentX = CurrentX + 6
                Print FormatCurrencyValue(sCurrencySymbol, rTotalChange);
                CurrentX = CurrentX + 6
                Print "(" + Format(rTotalChange / (mobjTotal.TotalValue - rTotalChange), "0.00%") + ")";
                mrDaySummaryRegionEndX = CurrentX
                bShown = True
            End If
            
            If bShown Then
                CurrentX = CurrentX + 4
                Line (CurrentX, 0)-(CurrentX, ScaleHeight), &H808080
            End If
        End If
        CurrentX = CurrentX + 6
    End If
    picText.Move CurrentX, 0, iDisplayWidth - CurrentX + 8, ScaleHeight
        
    ' Set the display flashing
    If sData <> "" Then
        picData.Visible = False
        Visible = True
        For iCnt = 1 To 3
            Cls
            picText.Cls
            picData.Cls
            Refresh
            Sleep 200
            ForeColor = IIf(iCnt = 3, &HE0E0&, &HE0&)
            CurrentX = picSize(0).Width + 2 + iImageWidth
            CurrentY = 1
            Print sData;
            Refresh
            Sleep 100
        Next iCnt
        miScrollPosition = 0
        picText.Visible = True
        
    ' Display the symbols
    Else
        Z_DrawSymbolText
        Z_DisplaySymbols
        frmPreview.RefreshDisplay
    End If
    
End Sub

Private Sub Form_DblClick()

Dim objSymbol As Object

    ' Show the browser window for this application
    Set objSymbol = Z_GetSymbolUnderMouse
    If Not objSymbol Is Nothing Then
        frmPreview.HideChart True
        Call Z_ShowYahooChart(objSymbol.Code)
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 And Shift = vbAltMask Then
        mnuExit_Click
    ElseIf KeyCode = vbKeyF1 Then
        mnuHelp_Click
    End If
    
End Sub

Private Sub Form_Load()

    ' Remove the border
    SetWindowLong Me.hWnd, GWL_STYLE, 100663296
    Set mobjCurrentSymbols = Nothing
    Set mobjExchangeRates = Nothing
    Set mobjSummaryStocks = Nothing
    
    ' Position the display based upon the docking
    Z_SetupDisplay
    
    ' Display the data
    Call Z_DisplayData("Loading....")
    Z_GetSymbolData
    Z_DisplayData
    timData.Enabled = True
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim stRect As RECT
Dim stPoint As POINTAPI

    ' Show the context menu
    If Button = vbRightButton Then
        frmPreview.HideChart True
        PopupMenu mnuPopup
    
    ' Set the capture so that we can catch the mouse up
    ElseIf Not mbCapturing Then
        Call SetCapture(hWnd)
        Call GetCursorPos(stPoint)
        Call GetCursorPos(mstPoint)
        Call ScreenToClient(hWnd, mstPoint)
        stRect.Left = mstPoint.X
        stRect.Right = (GetSystemMetrics(SM_CXVIRTUALSCREEN) - (ScaleWidth - mstPoint.X - 1))
        stRect.Top = mstPoint.Y
        stRect.Bottom = (GetSystemMetrics(SM_CYVIRTUALSCREEN) - (ScaleHeight - mstPoint.Y + 1))
        Call ClipCursor(stRect)
        Call GetWindowRect(hWnd, stRect)
        mbCapturing = True
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Z_MouseMove False

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim stRect As RECT

    If mbCapturing Then
        
        ' Release the mouse capture and save the current position
        ReleaseCapture
        Call ClipCursor(ByVal 0&)
        mbCapturing = False
        Call GetWindowRect(hWnd, stRect)
        If stRect.Top = 0 And stRect.Right = GetSystemMetrics(SM_CXVIRTUALSCREEN) Then
            Move Left, 0
        ElseIf stRect.Left = 0 And stRect.Bottom = GetSystemMetrics(SM_CYVIRTUALSCREEN) Then
            Move Left, Top + Screen.TwipsPerPixelY
        End If
        Call GetWindowRect(hWnd, stRect)
        Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_POSITION, Format(stRect.Left) + "," + Format(stRect.Top) + "," + Format(stRect.Right - stRect.Left))
        
        ' If we have in some way changed the docking
        If Not mnuDockNone.Checked Then
            If stRect.Top <> mstDock.rc.Top Then mnuDockNone_Click
        End If
    End If
    
End Sub

Private Sub Form_Resize()

    If WindowState = vbNormal Then
        picSize(0).Move 0, 0, picSize(0).Width, ScaleHeight
        picSize(1).Move ScaleWidth - picSize(0).Width, 0, picSize(0).Width, ScaleHeight
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Initialise before unloading
    timData.Enabled = False
    Set mobjCurrentSymbols = Nothing
    Set mobjExchangeRates = Nothing
    Set mobjSummaryStocks = Nothing
    
    ' Undock ourselves
    Z_UnloadDock

End Sub

Private Sub mnuAbout_Click()

    MsgBox VERSION_NAME + vbCrLf + vbCrLf + "Version:" + VERSION_NUMBER + " Built:" + VERSION_TIMESTAMP
    
End Sub

Private Sub mnuDockBottom_Click()

    Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_DOCK_TYPE, "Bottom")
    Z_SetupDisplay

End Sub

Private Sub mnuDockTop_Click()

    Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_DOCK_TYPE, "Top")
    Z_SetupDisplay

End Sub

Private Sub mnuDockNone_Click()

    Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_DOCK_TYPE, "None")
    Z_SetupDisplay

End Sub

Private Sub mnuDockAutoHide_Click()

    Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_DOCK_AUTOHIDE, IIf(mnuDockAutoHide.Checked, "false", "true"))
    Z_SetupDisplay

End Sub

Private Sub mnuEditSymbols_Click()

    mbForceRefresh = False
    frmSymbols.Show vbModal, Me
    DoEvents
    If mobjCurrentSymbols.Count = 0 Or mbForceRefresh Then
        mnuRefresh_Click
    End If

End Sub

Private Sub mnuExportAll_Click()

Dim sFilename$
Dim objSymbol As cSymbol
Dim objSymbols As New Collection
Dim objStock As cStock

    On Error Resume Next
    sFilename = PSGEN_SelectSaveFile(frmMain.hWnd, "Export Files" + vbNullChar + "*.csv", OFN_EXPLORER + OFN_OVERWRITEPROMPT + OFN_PATHMUSTEXIST + OFN_SHAREAWARE, "Save to File")
    If sFilename <> "" Then
        Set objSymbols = Z_GetDisplaySymbols(False, True)
        
        Open sFilename For Output As #1
        
        Print #1, "Date," + Format(Now, "yyyy-mm-dd hh:nn:ss")
        Print #1, "Code,Display Name,Trade Date,Disabled,Currency Name,Currency Symbol,Shares,Cost,Current Price"
        For Each objSymbol In objSymbols
            Print #1, objSymbol.Code + ",";
            Print #1, objSymbol.DisplayName + ",";
            Print #1, Format(DateAdd("s", CDbl(objSymbol.RegKey), DateSerial(2008, 1, 1)), "yyyy-mm-dd hh:nn:ss") + ",";
            Print #1, IIf(objSymbol.Disabled, "true", "false") + ",";
            Print #1, objSymbol.CurrencyName + ",";
            Print #1, objSymbol.CurrencySymbol + ",";
            Print #1, CStr(objSymbol.Shares) + ",";
            Print #1, CStr(objSymbol.Price) + ",";
            
            Set objStock = Nothing
            Set objStock = mobjSummaryStocks.Item(objSymbol.Code)
            If objStock Is Nothing Then
                Print #1, ""
            Else
                Print #1, CStr(objStock.CurrentPrice)
            End If
        Next objSymbol
        Close #1
    End If
    
End Sub

Private Sub mnuExportLive_Click()

Dim sFilename$
Dim objSymbol As cSymbol
Dim objSymbols As New Collection
Dim objStock As cStock

    On Error Resume Next
    sFilename = PSGEN_SelectSaveFile(frmMain.hWnd, "Export Files" + vbNullChar + "*.csv", OFN_EXPLORER + OFN_OVERWRITEPROMPT + OFN_PATHMUSTEXIST + OFN_SHAREAWARE, "Save to File")
    If sFilename <> "" Then
        Set objSymbols = Z_GetDisplaySymbols(False, True)
        
        Open sFilename For Output As #1
        
        Print #1, "Date," + Format(Now, "yyyy-mm-dd hh:nn:ss")
        Print #1, "Code,Display Name,Currency Name,Currency Symbol,Shares,Cost,Current Price"
        For Each objSymbol In objSymbols
            If Not objSymbol.Disabled Then
                Print #1, objSymbol.Code + ",";
                Print #1, objSymbol.DisplayName + ",";
                Print #1, objSymbol.CurrencyName + ",";
                Print #1, objSymbol.CurrencySymbol + ",";
                Print #1, CStr(objSymbol.Shares) + ",";
                Print #1, CStr(objSymbol.Price) + ",";
                
                Set objStock = Nothing
                Set objStock = mobjSummaryStocks.Item(objSymbol.Code)
                If objStock Is Nothing Then
                    Print #1, ""
                Else
                    Print #1, CStr(objStock.CurrentPrice)
                End If
            End If
        Next objSymbol
        Close #1
    End If
    
End Sub

Private Sub mnuExportSummarised_Click()

Dim sFilename$
Dim objSymbol As cSymbol
Dim objSymbols As New Collection
Dim objStock As cStock

    On Error Resume Next
    sFilename = PSGEN_SelectSaveFile(frmMain.hWnd, "Export Files" + vbNullChar + "*.csv", OFN_EXPLORER + OFN_OVERWRITEPROMPT + OFN_PATHMUSTEXIST + OFN_SHAREAWARE, "Save to File")
    If sFilename <> "" Then
        Set objSymbols = Z_GetDisplaySymbols(True, True)
        
        Open sFilename For Output As #1
        
        Print #1, "Date," + Format(Now, "yyyy-mm-dd hh:nn:ss")
        Print #1, "Code,Display Name,Currency Name,Currency Symbol,Shares,Cost,Current Price"
        For Each objSymbol In objSymbols
            If Not objSymbol.Disabled Then
                Print #1, objSymbol.Code + ",";
                Print #1, objSymbol.DisplayName + ",";
                Print #1, objSymbol.CurrencyName + ",";
                Print #1, objSymbol.CurrencySymbol + ",";
                Print #1, CStr(objSymbol.Shares) + ",";
                Print #1, CStr(objSymbol.Price) + ",";
                
                Set objStock = Nothing
                Set objStock = mobjSummaryStocks.Item(objSymbol.Code)
                If objStock Is Nothing Then
                    Print #1, ""
                Else
                    Print #1, CStr(objStock.CurrentPrice)
                End If
            End If
        Next objSymbol
        Close #1
    End If
    
End Sub

Private Sub mnuHelp_Click()

    Call PSGEN_LaunchBrowser("https://github.com/steveohara/stockticker/wiki")

End Sub

Private Sub mnuSettings_Click()

    mbForceRefresh = False
    frmSettings.Show vbModal, Me
    DoEvents

    If mbForceRefresh Then
        ' Set top most if required
        If CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_ALWAYS_ON_TOP, "-1")) Then
            Call SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
        Else
            Call SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
        End If
    
        ' Refresh the display
        mnuRefresh_Click
    End If

End Sub

Private Sub mnuExit_Click()

    End
    
End Sub

Private Sub mnuFontSize10pt_Click()

    Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_FONTSIZE, "10")
    Form_Load

End Sub

Private Sub mnuFontSize12pt_Click()

    Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_FONTSIZE, "12")
    Form_Load

End Sub

Private Sub mnuFontSize7pt_Click()

    Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_FONTSIZE, "7")
    Form_Load

End Sub


Private Sub mnuFontSize8pt_Click()

    Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_FONTSIZE, "8")
    Form_Load

End Sub

Private Sub mnuFontSize9pt_Click()

    Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_FONTSIZE, "9")
    Form_Load

End Sub

Private Sub mnuRefresh_Click()

    timData.Enabled = False
    Set mobjCurrentSymbols = Nothing
    If Visible Then Call Z_DisplayData("Loading....")
    Z_GetSymbolData
    Z_DisplayData
    timData.Enabled = True
    
End Sub

Private Sub mnuRunAtStartup_Click()

Dim objShell As Object
Dim objLink As Object
  
    On Error Resume Next
    If mnuRunAtStartup.Checked Then
        Kill PSGEN_GetSpecialFolderLocation(CSIDL_STARTUP) + "\" + App.ProductName + ".lnk"
    Else
        Set objShell = CreateObject("WScript.Shell")
        Set objLink = objShell.CreateShortcut(PSGEN_GetSpecialFolderLocation(CSIDL_STARTUP) + "\" + App.ProductName + ".lnk")
        objLink.Description = App.FileDescription
        objLink.TargetPath = App.path + "\" + App.EXEName + ".exe"
        objLink.WindowStyle = 1
        objLink.Save
    End If
    mnuRunAtStartup.Checked = PSGEN_FileExists(PSGEN_GetSpecialFolderLocation(CSIDL_STARTUP) + "\" + App.ProductName + ".lnk")
    
End Sub

Private Sub mnuScrollFast_Click()

    Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_SCROLLSPEED, REG_SCROLLSPEED_FAST)
    Z_SetScrollInterval

End Sub

Private Sub mnuScrollMedium_Click()

    Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_SCROLLSPEED, REG_SCROLLSPEED_MEDIUM)
    Z_SetScrollInterval

End Sub

Private Sub mnuScrollSlow_Click()

    Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_SCROLLSPEED, REG_SCROLLSPEED_SLOW)
    Z_SetScrollInterval

End Sub

Private Sub mnuTopMost_Click()

    mobjReg.SaveSetting App.Title, REG_SETTINGS, REG_ALWAYS_ON_TOP, Not mnuTopMost.Checked
    Z_SetTopMost

End Sub

Private Sub mnuUpgrade_Click()

    PSMAIN_CheckForUpgrade
    
End Sub





Private Sub picSize_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim stRect As RECT

    ' Set the capture so that we can catch the mouse up
    If Not mbCapturing Then
        Call SetCapture(picSize(Index).hWnd)
        Call GetCursorPos(mstPoint)
        Call ScreenToClient(picSize(Index).hWnd, mstPoint)
        mbCapturing = True
    End If
    
End Sub

Private Sub picSize_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim lScreenWidth&, lScreenHeight&
Dim stPoint As POINTAPI
Dim stRect As RECT
Dim stRectGrabs As RECT
Dim iDisplayWidth%

    ' If we are capturing then move the display
    On Error Resume Next
    If mbCapturing Then
        Call GetWindowRect(hWnd, stRect)
        Call GetWindowRect(picSize(0).hWnd, stRectGrabs)
        lScreenWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN)
        lScreenHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN)
        Call GetCursorPos(stPoint)
        If Index = 0 Then
            If stPoint.X < mstPoint.X Then stPoint.X = mstPoint.X
            If stPoint.X > stRect.Right - (stRectGrabs.Right - stRectGrabs.Left + mstPoint.X + 1) Then stPoint.X = stRect.Right - (stRectGrabs.Right - stRectGrabs.Left + mstPoint.X + 1)
            Call SetWindowPos(hWnd, 0, stPoint.X - mstPoint.X, stRect.Top, stRect.Right - stPoint.X + mstPoint.X, stRect.Bottom - stRect.Top, 0)
        Else
            If stPoint.X > lScreenWidth - (stRectGrabs.Right - stRectGrabs.Left) + mstPoint.X + 1 Then stPoint.X = lScreenWidth - (stRectGrabs.Right - stRectGrabs.Left) + mstPoint.X + 1
            If stPoint.X < stRect.Left + (stRectGrabs.Right - stRectGrabs.Left) + mstPoint.X + 1 Then stPoint.X = stRect.Left + (stRectGrabs.Right - stRectGrabs.Left) + mstPoint.X + 1
            Call SetWindowPos(hWnd, 0, stRect.Left, stRect.Top, stPoint.X - mstPoint.X + (stRectGrabs.Right - stRectGrabs.Left) - stRect.Left, stRect.Bottom - stRect.Top, 0)
        End If
        iDisplayWidth = ScaleWidth - (2 * picSize(0).Width) - 4
        picText.Move picText.Left, 0, iDisplayWidth - picText.Left + 8, ScaleHeight
        mbScrolling = picData.CurrentX > picText.ScaleWidth + 8
    ElseIf Not timMouse.Enabled Then
        picSize(Index).BackColor = &H808080
        timMouse.Tag = CStr(Index)
        timMouse.Enabled = True
    End If
    frmPreview.HideChart True
    
End Sub

Private Sub picSize_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim stRect As RECT

    If mbCapturing Then
        
        ' Release the mouse capture and save the current position
        ReleaseCapture
        mbCapturing = False
        Call GetWindowRect(hWnd, stRect)
        Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_POSITION, Format(stRect.Left) + "," + Format(stRect.Top) + "," + Format(stRect.Right - stRect.Left))
        miScrollPosition = 0
        mbScrolling = False
        Z_DisplayData
    End If
    
End Sub


Private Sub picText_DblClick()

    Form_DblClick

End Sub

Private Sub picText_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Form_KeyDown(KeyCode, Shift)
    
End Sub

Private Sub picText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call Form_MouseDown(Button, Shift, X, Y)
    
End Sub

Private Sub picText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Z_MouseMove True

End Sub

Private Sub timData_Timer()
    
    'Debug.Print Format(Now, "hh:nn:ss") + " Data timer"
    Z_GetSymbolData

End Sub

Public Sub TimerEvent()

    ' Draw the text on the display if required
    If Visible Then
        If mDataHasChanged Then
            Z_DisplayData
        
        ElseIf Width <> mlCurrentWidth Or mbScrolling Then
            Z_DisplaySymbols
        End If
        
        mDataHasChanged = False
        mlCurrentWidth = Width
    End If

End Sub



Private Function Z_GetSymbolData()
'
'                          ) As String
'
' Returns the user counts display string
'
Static objOldValues As New Collection
Static objOldPercents As New Collection

Dim objSymbol As cSymbol
Dim objStock As cStock
Dim objSymsToLookup As New Collection
Dim objSummaryStocks As New Collection
Dim objSymbolsWithData As New Collection

Dim sSymbol As Variant
Dim sSymbolInfo As Variant
Dim rTotalInvested#, rTotalValue#, rOldPrice#, rOldPercent#
Dim sSummaryCurrencySymbol$, sSummaryCurrencyName$, sCurrencySymbol$
Dim objAlarm As frmAlarm

    ' Get the useful stuff
    On Error Resume Next
    sSummaryCurrencyName = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_CURRENCY, "GBP")
    sSummaryCurrencySymbol = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_CURRENCY_SYMBOL)
    If mobjCurrentSymbols.Count = 0 Then
        Set mobjCurrentSymbols = ReadSymbolsFromRegistry
    End If
    
    ' Decide if we need to adjust the symbols
    Set mobjCurrentSymbols = Z_GetDisplaySymbols(CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SHOW_SUMMARY_SUMMARISE, "0")), False)
    
    ' Loop through all the symbols getting the lists
    If mobjCurrentSymbols.Count > 0 Then
        
        ' Only attempt to get the exchanges rates every 10 minutes after the first successful get
        ' Do this by using the tag as a counter of seconds
        If Val(timData.Tag) >= (600 / timData.Interval) Or timData.Tag = "" Then
            If Not PSDATA_GetExchangeRates(mobjCurrentSymbols, mobjExchangeRates) Then
                timData.Tag = ""
            Else
                timData.Tag = "600"
            End If
        End If
        timData.Tag = IIf(timData.Tag = "", 1, Format(Val(timData.Tag) + 1))
        
        ' Get all the enabled symbols to lookup
        For Each objSymbol In mobjCurrentSymbols
            If Not objSymbol.Disabled Then
                objSymsToLookup.Add objSymbol.Code, objSymbol.Code
            End If
        Next
        
        ' Now get a list of all the symbols from IEX if we have a key
        Call PSDATA_GetPrices(objSymsToLookup, objSymbolsWithData)
        
        ' Get the values from each stock
        For Each objSymbol In mobjCurrentSymbols
            If Not objSymbol.Disabled Then
                Err.Clear
                Set objStock = Nothing
                Set objStock = objSymbolsWithData.Item(objSymbol.Code)
                If objStock Is Nothing Then
                    sSymbolInfo = mobjReg.GetSetting(App.Title, REG_LAST_GOOD_VALUES, objSymbol.Code, "")
                    If sSymbolInfo = "" Then
                        PSGEN_Log "Couldn't get data for " + objSymbol.Code + " - and no historic values to use", LogEventTypes.LogError, EventIdTypes.StockPrices
                        objSymbol.ErrorDescription = "Couldn't refresh value and no historic data to retrieve"
                    Else
                        Set objStock = New cStock
                        Call objStock.initFromString(sSymbolInfo)
                        PSGEN_Log "Couldn't get data for " + objSymbol.Code + " - using historic value", LogEventTypes.LogWarning, EventIdTypes.StockPrices
                        objSymbol.ErrorDescription = "Couldn't refresh value"
                    End If
                Else
                    objSymbol.ErrorDescription = ""
                End If
                
                If Not objStock Is Nothing Then
                    ' Work out the running stuff
                    Call mobjReg.SaveSetting(App.Title, REG_LAST_GOOD_VALUES, objSymbol.Code, objStock.toString())
                    objSymbol.CurrentPrice = objStock.CurrentPrice
                    objSymbol.DayLow = objStock.DayStart
                    objSymbol.DayHigh = objStock.DayHigh
                    objSymbol.DayChange = objStock.DayChange
                    objSymbol.DayStart = objSymbol.CurrentPrice - objSymbol.DayChange
                    objSymbol.LastUpdate = Now
                    objSymbol.Source = objStock.Source
                End If
                    
                If Not objSymbol.ExcludeFromSummary Then
                    rTotalInvested = rTotalInvested + ConvertCurrency(objSymbol, objSymbol.Price * objSymbol.Shares)
                    rTotalValue = rTotalValue + ConvertCurrency(objSymbol, objSymbol.CurrentPrice * objSymbol.Shares)
                End If
                
                ' Check for any problems
                If objSymbol.CurrentPrice = 0 Then objSymbol.ErrorDescription = "Bad Symbol"
                sCurrencySymbol = sCurrencySymbol + objSymbol.CurrencySymbol
                
                ' Create the summary stock values
                If Not objSymbol.ExcludeFromSummary Then
                    Set objStock = objSummaryStocks.Item(objSymbol.DisplayName)
                    objStock.Code = objSymbol.Code
                    objStock.DisplayName = objSymbol.DisplayName
                    objStock.CurrentPrice = objSymbol.CurrentPrice
                    objStock.CurrencyName = objSymbol.CurrencyName
                    objStock.DayStart = objSymbol.DayStart
                    objStock.DayChange = objSymbol.DayChange
                    objStock.Source = objSymbol.Source
                    objStock.AddStock objSymbol.Shares, objSymbol.Price
                    objSummaryStocks.Add objStock, objStock.DisplayName
                    If objStock.CurrencySymbol = "" And objSymbol.CurrencySymbol <> "" Then objStock.CurrencySymbol = objSymbol.CurrencySymbol
                End If
                
                ' Check for an alarm condition
                If objOldValues Is Nothing Then Set objOldValues = New Collection
                If objOldPercents Is Nothing Then Set objOldPercents = New Collection
                rOldPrice = 0
                rOldPrice = objOldValues.Item(objSymbol.Code)
                rOldPercent = 0
                rOldPercent = objOldValues.Item(objSymbol.Code)
                If objSymbol.LowAlarmEnabled Then
                    
                    ' If this is a percentage change then base it on the last recorded value
                    If objSymbol.LowAlarmIsPercent Then
                        If objSymbol.PercentChange <= objSymbol.LowAlarmValue And (rOldPercent = 0 Or rOldPercent > objSymbol.LowAlarmValue) Then
                            Set objAlarm = New frmAlarm
                            Call objAlarm.ShowLowAlarm(objSymbol)
                            Call objOldValues.Add(objSymbol.CurrentPrice, objSymbol.Code)
                        End If
                    Else
                        If objSymbol.CurrentPrice <= objSymbol.LowAlarmValue And (rOldPrice = 0 Or rOldPrice > objSymbol.LowAlarmValue) Then
                            Set objAlarm = New frmAlarm
                            Call objAlarm.ShowLowAlarm(objSymbol)
                        End If
                        Call objOldValues.Add(objSymbol.CurrentPrice, objSymbol.Code)
                    End If
                End If
                
                If objSymbol.HighAlarmEnabled And Not objSymbol.AlarmShowing Then
                    If objSymbol.HighAlarmIsPercent Then
                        If objSymbol.PercentChange >= objSymbol.HighAlarmValue And (rOldPercent = 0 Or rOldPercent < objSymbol.HighAlarmValue) Then
                            Set objAlarm = New frmAlarm
                            Call objAlarm.ShowHighAlarm(objSymbol)
                            Call objOldValues.Add(objSymbol.CurrentPrice, objSymbol.Code)
                        End If
                    Else
                        If objSymbol.CurrentPrice >= objSymbol.HighAlarmValue And (rOldPrice = 0 Or rOldPrice < objSymbol.HighAlarmValue) Then
                            Set objAlarm = New frmAlarm
                            Call objAlarm.ShowHighAlarm(objSymbol)
                        End If
                        Call objOldValues.Add(objSymbol.CurrentPrice, objSymbol.Code)
                    End If
                End If
                Call objOldValues.Add(objSymbol.CurrentPrice, objSymbol.Code)
                Call objOldPercents.Add(objSymbol.PercentChange, objSymbol.Code)
            End If
        Next
    End If

    ' Return value to caller
    mDataHasChanged = True
    Set mobjSummaryStocks = objSummaryStocks
    mobjTotal.TotalCost = rTotalInvested
    mobjTotal.TotalValue = rTotalValue
    mobjTotal.CurrencySymbol = ""
    
    ' Set the currency symbol for conversion if specified by the user
    If sSummaryCurrencySymbol <> "" Then
        mobjTotal.CurrencySymbol = sSummaryCurrencySymbol
        mobjTotal.CurrencyName = sSummaryCurrencyName
        
    ' Use the first symbol in the list
    ElseIf sCurrencySymbol <> "" Then
        If sCurrencySymbol = String(Len(sCurrencySymbol), Left(sCurrencySymbol, 1)) Then mobjTotal.CurrencySymbol = Left(sCurrencySymbol, 1)
    End If
    
End Function
    
Private Function Z_GetDisplaySymbols(bSummarise As Boolean, bGetAll As Boolean)

Dim objSymbols As New Collection
Dim objSymbolsToUse As New Collection
Dim objSymbol As cSymbol
Dim objBaseSymbol As cSymbol

    If bGetAll Then
        Set objSymbolsToUse = ReadSymbolsFromRegistry
    Else
        Set objSymbolsToUse = mobjCurrentSymbols
    End If
    If bSummarise Then
        On Error Resume Next
        For Each objSymbol In objSymbolsToUse
            If bGetAll Or Not objSymbol.Disabled Then
            
                ' Check if we have this symbol already
                Set objBaseSymbol = Nothing
                Set objBaseSymbol = objSymbols.Item(objSymbol.Code)
                If objBaseSymbol Is Nothing Then
                    Set objBaseSymbol = objSymbol
                    objSymbols.Add objBaseSymbol, objBaseSymbol.Code
                Else
                    
                    ' Update the original with the new values
                    objBaseSymbol.Price = ((objBaseSymbol.Price * objBaseSymbol.Shares) + (objSymbol.Price * objSymbol.Shares)) / (objBaseSymbol.Shares + objSymbol.Shares)
                    objBaseSymbol.Shares = objBaseSymbol.Shares + objSymbol.Shares
                    objBaseSymbol.ShowPrice = IIf(objSymbol.ShowPrice, True, objBaseSymbol.ShowPrice)
                    objBaseSymbol.ShowChange = IIf(objSymbol.ShowChange, True, objBaseSymbol.ShowChange)
                    objBaseSymbol.ShowChangePercent = IIf(objSymbol.ShowChangePercent, True, objBaseSymbol.ShowChangePercent)
                    objBaseSymbol.ShowChangeUpDown = IIf(objSymbol.ShowChangeUpDown, True, objBaseSymbol.ShowChangeUpDown)
                    objBaseSymbol.ShowProfitLoss = IIf(objSymbol.ShowProfitLoss, True, objBaseSymbol.ShowProfitLoss)
                    objBaseSymbol.ShowDayChange = IIf(objSymbol.ShowDayChange, True, objBaseSymbol.ShowDayChange)
                    objBaseSymbol.ShowDayChangePercent = IIf(objSymbol.ShowDayChangePercent, True, objBaseSymbol.ShowDayChangePercent)
                    objBaseSymbol.ShowDayChangeUpDown = IIf(objSymbol.ShowDayChangeUpDown, True, objBaseSymbol.ShowDayChangeUpDown)
                End If
            End If
        Next objSymbol
    Else
        Set objSymbols = objSymbolsToUse
    End If
    Set Z_GetDisplaySymbols = objSymbols

End Function


Private Sub Z_DrawSymbolText()

Dim asSymbols$()
Dim i%, iLeft%, iTop%
Dim rSpace#
Dim lBackColor&, lTextColor&, lUpColor&, lDownColor&, lTmp&
Dim bNotFirst As Boolean
Dim objSymbol As cSymbol
Dim bShownOtherData As Boolean
Dim bShownBraces As Boolean
Dim objSymbols As Collection

    ' Get the colours from the registry
    'Debug.Print Format(Now, "hh:nn:ss") + " Drawing symbol text"
    lTextColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_TEXT_COLOUR, Format(vbWhite)))
    lUpColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_UP_COLOUR, Format(vbGreen)))
    lDownColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_DOWN_COLOUR, Format(vbRed)))
    
    ' Loop through all the symbols getting a sorted list
    picData.ScaleMode = vbPixels
    picData.Width = Int(GetSystemMetrics(SM_CXVIRTUALSCREEN)) * Screen.TwipsPerPixelX
    picData.Cls
    picData.CurrentX = 0
    picData.CurrentY = 1
    picData.BackColor = BackColor
    
    ' Work out the spacing
    rSpace = (10# * (picData.FontSize / 10#))
    If (rSpace > 10) Then rSpace = 10
        
    If mobjCurrentSymbols.Count > 0 Then
    
        ' Split the text into it's constituents
        On Error Resume Next
        For Each objSymbol In mobjCurrentSymbols

            If Not objSymbol.Disabled Then
                bShownOtherData = False
                
                ' Display all the bits starting with the correct colour
                If bNotFirst Then picData.CurrentX = picData.CurrentX + rSpace
                objSymbol.Position.Left = CurrentX + picData.CurrentX
                If objSymbol.ShowChangeUpDown Or objSymbol.ShowChange Or objSymbol.ShowChangePercent Or objSymbol.ShowProfitLoss Then
                    picData.ForeColor = IIf(objSymbol.CurrentPrice > objSymbol.Price, lUpColor, IIf(objSymbol.CurrentPrice < objSymbol.Price, lDownColor, lTextColor))
                    picData.Print objSymbol.DisplayName;
                
                    ' Show the price
                    If objSymbol.ShowPrice Then
                        picData.CurrentX = picData.CurrentX + 3
                        picData.Print objSymbol.FormattedValue;
                    End If
                    
                    ' Show the price difference
                    If objSymbol.ShowChange Then
                        bShownOtherData = True
                        picData.CurrentX = picData.CurrentX + 3
                        picData.Print FormatCurrencyValue(objSymbol.CurrencySymbol, objSymbol.CurrentPrice - objSymbol.Price);
                    End If
                    
                    ' Show the change in percent
                    If objSymbol.ShowChangePercent Then
                        bShownOtherData = True
                        picData.CurrentX = picData.CurrentX + 3
                        If objSymbol.Price <> 0 Then
                            picData.Print Format((objSymbol.CurrentPrice - objSymbol.Price) / objSymbol.Price, "0.00%");
                        Else
                            picData.Print "0.00%";
                        End If
                    End If
                    
                    ' Show the profit/loss
                    If objSymbol.ShowProfitLoss Then
                        picData.CurrentX = picData.CurrentX + 3
                        picData.Print FormatCurrencyValueWithSymbol(objSymbol.CurrencySymbol, objSymbol.CurrencyName, (objSymbol.CurrentPrice * objSymbol.Shares) - (objSymbol.Price * objSymbol.Shares));
                    End If
                
                    ' Show the up/down arrows
                    If objSymbol.ShowChangeUpDown Then
                        DrawUpDownArrow objSymbol, picData, ScaleHeight, picData.CurrentX + 4, picData.CurrentY
                    End If
                Else
                    picData.ForeColor = IIf(objSymbol.CurrentPrice > objSymbol.DayStart, lUpColor, IIf(objSymbol.CurrentPrice < objSymbol.Price, lDownColor, lTextColor))
                    picData.Print objSymbol.DisplayName;
                    
                    ' Show the price
                    If objSymbol.ShowPrice Then
                        picData.CurrentX = picData.CurrentX + 3
                        picData.Print objSymbol.FormattedValue;
                    End If
                End If
                
                ' Show day changes
                If objSymbol.ShowDayChange Or objSymbol.ShowDayChangePercent Or objSymbol.ShowDayChangeUpDown Then
                    bShownBraces = (bShownOtherData And (objSymbol.ShowDayChange Or objSymbol.ShowDayChangePercent)) Or (objSymbol.ShowChangeUpDown And objSymbol.ShowDayChangeUpDown)
                    picData.ForeColor = IIf(objSymbol.CurrentPrice > objSymbol.DayStart, lUpColor, IIf(objSymbol.CurrentPrice < objSymbol.DayStart, lDownColor, lTextColor))
                    picData.CurrentX = picData.CurrentX + 3
                    If bShownBraces Then picData.Print "(";
                
                    ' Show the Day price difference
                    If objSymbol.ShowDayChange Then
                        picData.CurrentX = picData.CurrentX + 3
                        picData.Print FormatCurrencyValue(objSymbol.CurrencySymbol, objSymbol.CurrentPrice - objSymbol.DayStart);
                    End If
                    
                    ' Show the Day change in percent
                    If objSymbol.ShowDayChangePercent Then
                        picData.CurrentX = picData.CurrentX + 3
                        If objSymbol.DayStart <> 0 Then
                            picData.Print Format(objSymbol.DayChange / objSymbol.DayStart, "0.00%");
                        Else
                            picData.Print "0.00%";
                        End If
                    End If
                    
                    ' Show the day profit/loss
                    If objSymbol.ShowProfitLoss And objSymbol.ShowDayChange Then
                        picData.CurrentX = picData.CurrentX + 3
                        picData.Print FormatCurrencyValueWithSymbol(objSymbol.CurrencySymbol, objSymbol.CurrencyName, (objSymbol.CurrentPrice * objSymbol.Shares) - (objSymbol.DayStart * objSymbol.Shares));
                    End If
                    
                    ' Show the Day up/down arrows
                    If objSymbol.ShowDayChangeUpDown Then
                        DrawUpDownArrow objSymbol, picData, ScaleHeight, picData.CurrentX + 1, picData.CurrentY
                    End If
                
                    If bShownBraces Then picData.Print ")";
                End If
                
                objSymbol.Position.Right = CurrentX + picData.CurrentX
                bNotFirst = True
            End If
        Next
        If bNotFirst Then picData.CurrentX = picData.CurrentX + 10
    End If
    mbScrolling = picData.CurrentX > picText.ScaleWidth + 8

End Sub

Private Function Z_GetSymbolUnderMouse() As Object

Dim stPoint As POINTAPI
Dim objSymbol As cSymbol
Dim objStock As cStock
Dim objReturn As Object
    
    ' Find the symbol we're over - check the symbols first
    Call GetCursorPos(stPoint)
    Call ScreenToClient(hWnd, stPoint)
    If stPoint.Y > 0 And stPoint.Y < ScaleHeight Then
        stPoint.X = stPoint.X + miScrollPosition
        For Each objSymbol In mobjCurrentSymbols
            If stPoint.X > objSymbol.Position.Left And stPoint.X < objSymbol.Position.Right Then
                Set objReturn = objSymbol
                Exit For
            End If
        Next objSymbol
        
        ' If we haven't found one then now try the summary stocks
        If objReturn Is Nothing Then
            For Each objStock In mobjSummaryStocks
                If stPoint.X > objStock.Position.Left And stPoint.X < objStock.Position.Right Then
                    Set objReturn = objStock
                    Exit For
                End If
            Next objStock
        End If
    End If
    
    Set Z_GetSymbolUnderMouse = objReturn

End Function

Private Sub Z_UnloadDock()

    If mstDock.hWnd <> 0 Then
        mstDock.lParam = False
        Call SHAppBarMessage(ABM_REMOVE, mstDock)
        mstDock.hWnd = 0
    End If

End Sub


Private Sub Z_SetupDisplay()

Dim lTop&, lLeft&, lWidth&, lHeight&
Dim sPos$, sDocking$, sTmp$
Dim lScreenWidth&, lScreenHeight&
Dim stRect As RECT
    
    ' Set the scroll and data intervals
    mnuRunAtStartup.Checked = PSGEN_FileExists(PSGEN_GetSpecialFolderLocation(CSIDL_STARTUP) + "\" + App.ProductName + ".lnk")
    Z_SetScrollInterval
    timData.Interval = CInt(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_FREQUENCY, Format(REG_FREQUENCY_DEF))) * 1000
    
    ' Set the docking stuff
    sPos = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_DOCK_TYPE)
    mnuDockTop.Checked = PSGEN_IsSameText(sPos, "Top")
    mnuDockBottom.Checked = PSGEN_IsSameText(sPos, "Bottom")
    mnuDockNone.Checked = PSGEN_IsSameText(sPos, "None") Or sPos = ""
    mnuDockAutoHide.Checked = PSGEN_IsSameText(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_DOCK_AUTOHIDE), "true")
    
    ' Set the font
    sPos = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_FONTSIZE, "10")
    mnuFontSize7pt.Checked = (sPos = "7")
    mnuFontSize8pt.Checked = (sPos = "8")
    mnuFontSize9pt.Checked = (sPos = "9")
    mnuFontSize10pt.Checked = (sPos = "10")
    mnuFontSize12pt.Checked = (sPos = "12")
    Font.size = Val(sPos)
    Font.Bold = CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_BOLD, "0"))
    Font.Italic = CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_ITALIC, "0"))
    If mnuFontSize7pt.Checked Then
        Font.Name = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_FONT, "Small Fonts")
    Else
        Font.Name = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_FONT, "Calibri")
    End If
    
    ' Set the background colours
    BackColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_BACK_COLOUR, BackColor))
    picText.BackColor = BackColor
    picData.BackColor = BackColor
    
    ' Get the URL
    msURL = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_URL, REG_URL_DEF)
    
    ' Position the display
    lScreenWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN)
    lScreenHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN)
    sPos = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_POSITION)
    If sPos = "" Then
        lLeft = (ScaleX(Screen.Width, vbTwips, vbPixels) - ScaleWidth) / 2
        lTop = ScaleY(Screen.Height / 5, vbTwips, vbPixels)
        lWidth = ScaleWidth
    Else
        lLeft = Val(PSGEN_GetItem(1, ",", sPos))
        lTop = Val(PSGEN_GetItem(2, ",", sPos))
        lWidth = Val(PSGEN_GetItem(3, ",", sPos))
    End If
    lHeight = TextHeight("|") + 2
    
    ' Check that the parameters are ok
    If lWidth > lScreenWidth Then lWidth = lScreenWidth
    If lLeft < 0 Then lLeft = 0
    If lLeft > (lScreenWidth - lWidth) Then lLeft = (lScreenWidth - lWidth)
    If lTop < 0 Then lTop = 0
    If lTop > (lScreenHeight - lHeight) Then lTop = (lScreenHeight - lHeight)
    
    ' Position the window or dock it
    Z_UnloadDock
    If mnuDockNone.Checked Then
        Call SetWindowPos(hWnd, 0, lLeft, lTop, lWidth, lHeight, 0)
    Else
        mstDock.cbSize = Len(mstDock)
        mstDock.hWnd = hWnd
        Select Case True
            Case mnuDockTop.Checked
                mstDock.rc.Left = lLeft
                mstDock.rc.Top = 0
                mstDock.rc.Bottom = lHeight
                mstDock.rc.Right = lLeft + lWidth
                mstDock.uEdge = ABE_TOP
                
            Case mnuDockBottom.Checked
                mstDock.rc.Left = lLeft
                mstDock.rc.Top = ScaleX(Screen.Height, vbTwips, vbPixels) - lHeight
                mstDock.rc.Bottom = ScaleX(Screen.Height, vbTwips, vbPixels)
                mstDock.rc.Right = lLeft + lWidth
                mstDock.uEdge = ABE_BOTTOM
        
        End Select
        Call SHAppBarMessage(ABM_NEW, mstDock)
        Call SHAppBarMessage(ABM_SETPOS, mstDock)
        Do
            Call SetWindowPos(hWnd, 0, lLeft, mstDock.rc.Top, lWidth, lHeight, 0)
            DoEvents
            Call GetWindowRect(hWnd, stRect)
        Loop Until stRect.Top = mstDock.rc.Top
    End If
    
    ' Set top most if required
    Z_SetTopMost
    
End Sub

Private Sub Z_SetScrollInterval()

Dim sSpeed$
Dim iInterval%

    ' Get the scroll speed and adjust the timer if required
    sSpeed = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SCROLLSPEED, REG_SCROLLSPEED_MEDIUM)
    If sSpeed <> REG_SCROLLSPEED_FAST And sSpeed <> REG_SCROLLSPEED_MEDIUM And sSpeed <> REG_SCROLLSPEED_SLOW Then sSpeed = REG_SCROLLSPEED_MEDIUM
    
    ' If the values are different, then change the timer
    miScrollMovement = Val(Split(sSpeed, ",")(0))
    iInterval = Val(Split(sSpeed, ",")(1))
    If iInterval <> miScrollInterval Or mlTimer = 0 Then
        miScrollInterval = iInterval
        If mlTimer <> 0 Then Call KillTimer(hWnd, mlTimer)
        mlTimer = SetTimer(hWnd, 0, iInterval, AddressOf TimerProc)
    End If
    
    ' Indicate to the user which speed is selected
    mnuScrollFast.Checked = (sSpeed = REG_SCROLLSPEED_FAST)
    mnuScrollMedium.Checked = (sSpeed = REG_SCROLLSPEED_MEDIUM)
    mnuScrollSlow.Checked = (sSpeed = REG_SCROLLSPEED_SLOW)

End Sub

Private Sub Z_SetTopMost()

    If CBool(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_ALWAYS_ON_TOP, "-1")) Then
        Call SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
        mnuTopMost.Checked = True
    Else
        Call SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
        mnuTopMost.Checked = False
    End If

End Sub


Private Sub timMouse_Timer()
    Dim stPoint As POINTAPI
    Dim lHwnd&
    Dim iIndex%
    
    If timMouse.Tag <> "" And Not mbCapturing Then
        iIndex = Val(timMouse.Tag)
        Call GetCursorPos(stPoint)
        lHwnd = WindowFromPoint(stPoint.X, stPoint.Y)
        If lHwnd <> picSize(iIndex).hWnd And lHwnd <> picData.hWnd And lHwnd <> frmMain.hWnd Then
            picSize(iIndex).BackColor = &H0
            timMouse.Enabled = False
            timMouse.Tag = ""
        End If
    End If
    
End Sub

Private Sub Z_MouseMove(ByVal bOverData As Boolean)

Dim stPoint As POINTAPI

    ' If we are capturing (dragging) then move the display
    Call GetCursorPos(stPoint)
    If mbCapturing Then
        frmPreview.HideChart True
        Call SetWindowPos(hWnd, 0, stPoint.X - mstPoint.X, stPoint.Y - mstPoint.Y, 0, 0, SWP_NOSIZE)
    
    ' Show the preview window for this application
    ElseIf bOverData Then
        frmPreview.ShowChart Z_GetSymbolUnderMouse
    
    Else
        Call ScreenToClient(hWnd, stPoint)
        ' Show the preview window for the sumary
        If stPoint.X > mrSummaryRegionStartX And stPoint.X < mrSummaryRegionEndX Then
            frmPreview.ShowSummary
            
        ' Show the preview window for the daily sumary
        ElseIf stPoint.X > mrDaySummaryRegionStartX And stPoint.X < mrDaySummaryRegionEndX Then
            frmPreview.ShowDaySummary
        End If
    End If
End Sub


