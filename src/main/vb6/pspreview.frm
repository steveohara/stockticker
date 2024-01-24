VERSION 5.00
Begin VB.Form frmPreview 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5115
   ClientLeft      =   6840
   ClientTop       =   3495
   ClientWidth     =   9780
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "pspreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   341
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   652
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picSort 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4440
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2280
      Width           =   495
   End
   Begin VB.PictureBox picGraph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   690
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   0
      Top             =   1830
      Width           =   615
   End
   Begin VB.Timer timOpen 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   600
   End
   Begin VB.Timer timClose 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   645
      Top             =   660
   End
   Begin VB.Label lblSummaryHeader 
      Caption         =   "Gain/Loss"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   6720
      MousePointer    =   7  'Size N S
      TabIndex        =   11
      ToolTipText     =   "Sort by gain/loss"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblSummaryHeader 
      Caption         =   "Value"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   6840
      MousePointer    =   7  'Size N S
      TabIndex        =   10
      ToolTipText     =   "Sort by total value"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblSummaryHeader 
      Caption         =   "Shares"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6720
      MousePointer    =   7  'Size N S
      TabIndex        =   9
      ToolTipText     =   "Sort by number of shares"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblSummaryHeader 
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6720
      MousePointer    =   7  'Size N S
      TabIndex        =   8
      ToolTipText     =   "Sort by current price"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblSummaryHeader 
      Caption         =   "Cost"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6720
      MousePointer    =   7  'Size N S
      TabIndex        =   7
      ToolTipText     =   "Sort by base average cost"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblSummaryHeader 
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6840
      MousePointer    =   7  'Size N S
      TabIndex        =   6
      ToolTipText     =   "Sort by symbol"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblHeader 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2520
      MousePointer    =   7  'Size N S
      TabIndex        =   4
      ToolTipText     =   "Sort by change (percent)"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblHeader 
      Caption         =   "Value"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2520
      MousePointer    =   7  'Size N S
      TabIndex        =   3
      ToolTipText     =   "Sort by change (value)"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblHeader 
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2520
      MousePointer    =   7  'Size N S
      TabIndex        =   2
      ToolTipText     =   "Sort by current price"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblHeader 
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2520
      MousePointer    =   7  'Size N S
      TabIndex        =   1
      ToolTipText     =   "Sort by symbol"
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2008
'
'****************************************************************************
'
' LANGUAGE:             Microsoft Visual Basic V6.00
'
' MODULE NAME:          Pivotal_Preview
'
' MODULE TYPE:          BASIC Form
'
' FILE NAME:            PSPREVIEW.FRM
'
' MODIFICATION HISTORY: Steve O'Hara    28 August 2008   First created for StockTicker
'
' PURPOSE:              Provides a way of showing a preview of data from
'                           Yahoo
'
'
'****************************************************************************
'
'****************************************************
' MODULE VARIABLE DECLARATIONS
'****************************************************
'
Option Explicit

    Const LEFT_MARGIN = 520
    Const LEFT_MARGIN_VALUE = 600

    Dim mobjReg As New cRegistry

    Private mobjStockSymbol As Object
    Private mlLeftPos&
    Private mbOneDay As Boolean
    Private mbSummary As Boolean
    Private mbDaySummary As Boolean
    Private mobjChartCache As Collection


Private Sub Form_Activate()

    PSGEN_SetTopMost hWnd

End Sub

Private Sub Form_DblClick()

    picGraph_DblClick
    
End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblHeader_Click(Index As Integer)

Dim sSortOrder$
Dim iSortColumn%

    sSortOrder = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_DAY_SUMMARY_SORT_ORDER, "asc")
    iSortColumn = CInt(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_DAY_SUMMARY_SORT_COLUMN, "0"))
    If Index = iSortColumn Then
        Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_DAY_SUMMARY_SORT_ORDER, IIf(sSortOrder = "asc", "desc", "asc"))
    Else
        Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_DAY_SUMMARY_SORT_ORDER, "asc")
        Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_DAY_SUMMARY_SORT_COLUMN, Format(Index))
    End If
    Z_ShowDaySummary

End Sub

Private Sub lblSummaryHeader_Click(Index As Integer)

Dim sSortOrder$
Dim iSortColumn%

    sSortOrder = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_SORT_ORDER, "asc")
    iSortColumn = CInt(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_SORT_COLUMN, "0"))
    If Index = iSortColumn Then
        Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_SUMMARY_SORT_ORDER, IIf(sSortOrder = "asc", "desc", "asc"))
    Else
        Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_SUMMARY_SORT_ORDER, "asc")
        Call mobjReg.SaveSetting(App.Title, REG_SETTINGS, REG_SUMMARY_SORT_COLUMN, Format(Index))
    End If
    Z_ShowSummary

End Sub

Private Sub picGraph_DblClick()

    If Not mbDaySummary And Not mbSummary Then
        mbOneDay = Not mbOneDay
        Z_ShowChart mobjStockSymbol
    End If

End Sub

Private Sub Form_Load()

    PSGEN_SetTopMost hWnd
    Set mobjChartCache = New Collection

End Sub

Private Sub timClose_Timer()
    
Dim stPoint As POINTAPI
Dim stRect As RECT
Dim stRectParent As RECT
    
    If Visible Or timOpen.Enabled Then
        Call GetCursorPos(stPoint)
        Call GetWindowRect(hWnd, stRect)
        Call GetWindowRect(frmMain.hWnd, stRectParent)
        If (stPoint.X < stRect.Left Or stPoint.X > stRect.Right Or stPoint.Y < stRect.Top Or stPoint.Y > stRect.Bottom) And _
           (stPoint.X < stRectParent.Left Or stPoint.X > stRectParent.Right Or stPoint.Y < stRectParent.Top Or stPoint.Y > stRectParent.Bottom) Then
           HideChart True
        End If
    Else
        timClose.Enabled = False
    End If

End Sub

Private Sub timOpen_Timer()

    timOpen.Enabled = False
    If mbDaySummary Then
        Z_ShowDaySummary
    ElseIf mbSummary Then
        Z_ShowSummary
    Else
        Z_ShowChart mobjStockSymbol
    End If

End Sub

Public Sub HideChart(Optional ByVal bImmediately As Boolean)
    
    If bImmediately Then
        timOpen.Enabled = False
        timClose.Enabled = False
        Set mobjStockSymbol = Nothing
        Hide
    Else
        timClose.Enabled = True
    End If

End Sub

Public Sub ShowChart(objStockSymbol As Object)

Dim stRect As RECT
Dim stPoint As POINTAPI

    mbDaySummary = False
    mbSummary = False
    If objStockSymbol Is Nothing Then
        HideChart True
        
    ElseIf Not objStockSymbol Is mobjStockSymbol Then
        timOpen.Enabled = False
        mbOneDay = True
        Set mobjStockSymbol = objStockSymbol
        GetWindowRect frmMain.hWnd, stRect
        Call GetCursorPos(stPoint)
        Call ScreenToClient(frmMain.hWnd, stPoint)
        mlLeftPos = mobjStockSymbol.Position.Left + stRect.Left - frmMain.miScrollPosition
        timOpen.Enabled = True
        timClose.Enabled = True
    End If

End Sub

Public Sub ShowDaySummary()

Dim stRect As RECT

    mbSummary = False
    mbDaySummary = True
    If Not timClose.Enabled Then
        GetWindowRect frmMain.hWnd, stRect
        timOpen.Enabled = False
        mlLeftPos = stRect.Left + frmMain.mrDaySummaryRegionStartX
        timOpen.Enabled = True
        timClose.Enabled = True
    End If

End Sub

Public Sub ShowSummary()

Dim stRect As RECT

    mbSummary = True
    mbDaySummary = False
    If Not timClose.Enabled Then
        GetWindowRect frmMain.hWnd, stRect
        timOpen.Enabled = False
        mlLeftPos = stRect.Left + frmMain.mrSummaryRegionStartX
        timOpen.Enabled = True
        timClose.Enabled = True
    End If

End Sub

Private Sub Z_ShowChart(objStockSymbol As Object)
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2008
'
'****************************************************************************
'
'                     NAME: Sub Z_ShowChart
'
'                     sSymbol$           - Symbol to show chart for
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    28 August 2008   First created for StockTicker
'
'                  PURPOSE: Shows a chart window
'
'****************************************************************************
'
'
Dim sTmp$, sSymbol$
Dim lWidth&, lToken&, lTmp&, lLeft&, lTop&, lHeight&
Dim lUpColor&, lDownColor&
Dim objStock As cStock
Dim objSymbol As cSymbol
Dim rRate#
Dim sCurrencyName$, sProxy$, sChartUrl$
Dim asAttempts$()
Dim iCnt%
Dim objChart As cChart
Dim bLoaded As Boolean

    '
    ' Determine the type of symbol
    '
    On Error Resume Next
    timClose.Enabled = False
    If TypeOf objStockSymbol Is cSymbol Then
        Set objSymbol = objStockSymbol
        Set objStock = frmMain.mobjSummaryStocks.Item(objSymbol.Code)
        sSymbol = objSymbol.Code
    Else
        Set objStock = objStockSymbol
        Set objSymbol = Nothing
        sSymbol = objStock.Code
    End If

    '
    ' Load and position the display
    '
    Z_ShowHeaders
    lLeft = mlLeftPos
    lTop = (frmMain.Top + frmMain.Height) / Screen.TwipsPerPixelY
    lWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN)
    lHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN)
    If lLeft + 700 > lWidth Then lLeft = lWidth - 700
    If lTop + 342 > lHeight Then lTop = (frmMain.Top / Screen.TwipsPerPixelY) - 342
    Move lLeft * Screen.TwipsPerPixelX, lTop * Screen.TwipsPerPixelY, 690 * Screen.TwipsPerPixelX, 342 * Screen.TwipsPerPixelY
    
    Set picGraph.Picture = Nothing
    picGraph.Move 0, 0, 512, 342
    picGraph.CurrentX = 180
    picGraph.CurrentY = 110
    picGraph.ForeColor = vbBlack
    picGraph.FontBold = False
    picGraph.Print "Retrieving " + IIf(mbOneDay, "week", "day") + " graph for  ";
    picGraph.FontBold = True
    picGraph.Print sSymbol
    DoEvents
    
    '
    ' Draw the useful text
    '
    Cls
    CurrentX = LEFT_MARGIN
    CurrentY = 5
    ForeColor = vbWhite
    lUpColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_UP_COLOUR, Format(vbGreen)))
    lDownColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_DOWN_COLOUR, Format(vbRed)))
    
    ' ******************************************
    ' Day Position
    ' ******************************************
    FontSize = 11
    FontBold = True
    Print "Overall Position"
    If objSymbol.FromNasdaqRealTime Then
        CurrentX = LEFT_MARGIN
        lTmp = ForeColor
        ForeColor = vbYellow
        Print "Real Time"
        ForeColor = lTmp
    End If
    CurrentY = CurrentY + 2
    FontSize = 9
    FontBold = False
    CurrentX = LEFT_MARGIN
    If Not objSymbol Is Nothing Then
        sCurrencyName = objSymbol.CurrencySymbol
        rRate = frmMain.mobjExchangeRates.Item(objSymbol.CurrencyName)
        
        Print "Price:";
        CurrentX = LEFT_MARGIN_VALUE
        Print objSymbol.FormattedValue
        CurrentX = LEFT_MARGIN
        CurrentY = CurrentY + 1
        
        Print "Cost:";
        CurrentX = LEFT_MARGIN_VALUE
        Print objSymbol.FormattedCost
        CurrentX = LEFT_MARGIN_VALUE
        Print objSymbol.FormattedTotalCost
        CurrentX = LEFT_MARGIN
        CurrentY = CurrentY + 1
        
        Print "Shares:";
        CurrentX = LEFT_MARGIN_VALUE
        Print Format(objSymbol.Shares, "#,###,##0")
        CurrentX = LEFT_MARGIN
        CurrentY = CurrentY + 1
        
        Print "Value:";
        CurrentX = LEFT_MARGIN_VALUE
        Print Format(objSymbol.FormattedTotalValue)
        CurrentX = LEFT_MARGIN
        CurrentY = CurrentY + 1
        
        Print "Gain/Loss:";
        ForeColor = IIf(objSymbol.CurrentPrice > objSymbol.Price, lUpColor, IIf(objSymbol.CurrentPrice < objSymbol.Price, lDownColor, ForeColor))
        CurrentX = LEFT_MARGIN_VALUE
        Print FormatCurrencyValueWithSymbol(objSymbol.CurrencySymbol, objSymbol.CurrencyName, objSymbol.Shares * (objSymbol.CurrentPrice - objSymbol.Price))
        CurrentX = LEFT_MARGIN_VALUE
        If rRate <> 1 And rRate > 0 Then
            CurrentY = CurrentY + 1
            Print "(" + FormatCurrencyValueWithSymbol(frmMain.mobjTotal.CurrencySymbol, frmMain.mobjTotal.CurrencyName, objSymbol.Shares * (objSymbol.CurrentPrice - objSymbol.Price) * rRate) + ")"
        End If
        ForeColor = vbWhite
        CurrentX = LEFT_MARGIN
        
        If objSymbol.CurrentPrice < objSymbol.Price Then
            Print "Required Chg:";
            ForeColor = lDownColor
            CurrentX = LEFT_MARGIN_VALUE
            Print FormatCurrencyValueWithSymbol(objSymbol.CurrencySymbol, objSymbol.CurrencyName, objSymbol.Price - objSymbol.CurrentPrice)
            CurrentX = LEFT_MARGIN_VALUE
            If rRate <> 1 And rRate > 0 Then
                CurrentY = CurrentY + 1
                Print "(" + FormatCurrencyValueWithSymbol(frmMain.mobjTotal.CurrencySymbol, frmMain.mobjTotal.CurrencyName, (objSymbol.Price - objSymbol.CurrentPrice) * rRate) + ")"
            End If
            ForeColor = vbWhite
            CurrentX = LEFT_MARGIN
        End If
        
        ' ******************************************
        ' Day Position
        ' ******************************************
        CurrentY = CurrentY + 5
        FontSize = 11
        FontBold = True
        Print "Day Position"
        FontSize = 9
        CurrentY = CurrentY + 2
        FontBold = False
        CurrentX = LEFT_MARGIN
        
        Print "Start:";
        CurrentX = LEFT_MARGIN_VALUE
        Print FormatCurrencyValue(objSymbol.CurrencySymbol, objSymbol.DayStart)
        CurrentX = LEFT_MARGIN
        CurrentY = CurrentY + 1
        
        Print "Low:";
        CurrentX = LEFT_MARGIN_VALUE
        Print FormatCurrencyValue(objSymbol.CurrencySymbol, objSymbol.DayLow)
        CurrentX = LEFT_MARGIN
        CurrentY = CurrentY + 1
        
        Print "High:";
        CurrentX = LEFT_MARGIN_VALUE
        Print FormatCurrencyValue(objSymbol.CurrencySymbol, objSymbol.DayHigh)
        CurrentX = LEFT_MARGIN
        CurrentY = CurrentY + 1
        
        ForeColor = vbWhite
        Print "Change: ";
        ForeColor = IIf(objSymbol.DayChange > 0, lUpColor, IIf(objSymbol.DayChange < 0, lDownColor, ForeColor))
        CurrentX = LEFT_MARGIN_VALUE
        Print FormatCurrencyValue(objSymbol.CurrencySymbol, objSymbol.DayChange) + " " + Format(objSymbol.DayChange / IIf(objSymbol.DayStart = 0, 1, objSymbol.DayStart), "0.00%")
        CurrentX = LEFT_MARGIN
        CurrentY = CurrentY + 1
        
        ForeColor = vbWhite
        Print "Gain/Loss: ";
        ForeColor = IIf(objSymbol.DayChange > 0, lUpColor, IIf(objSymbol.DayChange < 0, lDownColor, ForeColor))
        CurrentX = LEFT_MARGIN_VALUE
        Print FormatCurrencyValueWithSymbol(objSymbol.CurrencySymbol, objSymbol.CurrencyName, (objSymbol.CurrentPrice * objSymbol.Shares) - (objSymbol.DayStart * objSymbol.Shares))
    
        If rRate <> 1 And rRate > 0 Then
            CurrentY = CurrentY + 1
            CurrentX = LEFT_MARGIN_VALUE
            Print "(" + FormatCurrencyValueWithSymbol(frmMain.mobjTotal.CurrencySymbol, frmMain.mobjTotal.CurrencyName, objSymbol.Shares * (objSymbol.CurrentPrice - objSymbol.DayStart) * rRate) + ")"
            CurrentY = CurrentY - 1
        End If
        ForeColor = vbWhite
    Else
        sCurrencyName = objStock.CurrencySymbol
        rRate = frmMain.mobjExchangeRates.Item(objStock.CurrencyName)
        
        Print "Price: " + objStock.FormattedPrice
        CurrentX = LEFT_MARGIN
        CurrentY = CurrentY + 1
        
        Print "Cost Base: " + objStock.FormattedAverageCost
        CurrentX = LEFT_MARGIN
        CurrentY = CurrentY + 1
        
        Print "Shares: " + Format(objStock.NumberOfShares)
        CurrentX = LEFT_MARGIN
        CurrentY = CurrentY + 1
        
        Print "Gain/Loss: ";
        lTmp = CurrentX
        ForeColor = IIf(objStock.CurrentPrice > objStock.AverageCost, lUpColor, IIf(objStock.CurrentPrice < objStock.AverageCost, lDownColor, vbBlack))
        Print FormatCurrencyValueWithSymbol(objStock.CurrencySymbol, objStock.CurrencyName, objStock.NumberOfShares * (objStock.CurrentPrice - objStock.AverageCost))
        CurrentX = lTmp - TextWidth("(")
        If rRate <> 1 And rRate > 0 Then Print "(" + FormatCurrencyValueWithSymbol(frmMain.mobjTotal.CurrencySymbol, objStock.CurrencyName, objStock.NumberOfShares * (objStock.CurrentPrice - objStock.AverageCost) * rRate) + ")"
        CurrentX = LEFT_MARGIN
        ForeColor = vbWhite
    End If
    
    ' ******************************************
    ' Exchange rates
    ' ******************************************
    CurrentX = LEFT_MARGIN
    CurrentY = CurrentY + 5
    If rRate <> 1 And rRate > 0 Then
        FontBold = True
        FontSize = 11
        Print "FX Rate"
        FontSize = 9
        FontBold = False
        CurrentY = CurrentY + 1
        CurrentX = LEFT_MARGIN
        Print sCurrencyName + "1" + " = " + frmMain.mobjTotal.CurrencySymbol + Format(rRate, "0.00")
        CurrentX = LEFT_MARGIN
        Print frmMain.mobjTotal.CurrencySymbol + "1" + " = " + sCurrencyName + Format(1 / rRate, "0.00")
    End If
    
    FontBold = False
    CurrentX = LEFT_MARGIN
    CurrentY = 324
    FontSize = 8
    ForeColor = vbGrayText
    Print "Updated:" + Format(objSymbol.LastUpdate)
    Line (512, 0)-(687, 339), vbWhite, B
    
    '
    ' Get the graph
    '
    bLoaded = False
    Set objChart = mobjChartCache.Item(sSymbol + "-" + Format(mbOneDay))
    If Not objChart Is Nothing Then
        bLoaded = DateDiff("s", objChart.Timestamp, Now) < 30
    End If
    If bLoaded Then
        picGraph.Picture = objChart.Chart
    Else
        Show
        DoEvents
        sProxy = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_PROXY)
        
        sChartUrl = "https://www.reuters.wallst.com/enhancements/chartapi/index_chart_api.asp"
        Call PSINET_GetHTTPFile(sChartUrl + "?symbol=" + sSymbol + "&duration=" + IIf(mbOneDay, "1", "5") + "&headerType=quote&width=500&height=342", sTmp, sProxyName:=sProxy, lConnectionTimeout:=2000, lReadTimeout:=10000)
        If sTmp = "" Then
            Call PSINET_GetHTTPFile(sChartUrl + "?symbol=" + sSymbol + ".OQ&duration=" + IIf(mbOneDay, "1", "5") + "&headerType=quote&width=500&height=342", sTmp, sProxyName:=sProxy, lConnectionTimeout:=2000, lReadTimeout:=10000)
        End If
        If sTmp = "" Then
            Call PSINET_GetHTTPFile(sChartUrl + "?symbol=" + sSymbol + ".N&duration=" + IIf(mbOneDay, "1", "5") + "&headerType=quote&width=500&height=342", sTmp, sProxyName:=sProxy, lConnectionTimeout:=2000, lReadTimeout:=10000)
        End If
        
        '
        ' Draw the graph
        ' Initialise GDI+
        '
        lToken = InitGDIPlus
        
        ' Load pictures
        picGraph.Picture = LoadPictureFromStringGDIPlus(sTmp, 500, 342)
        
        ' Free GDI+
        FreeGDIPlus lToken
        
        Set objChart = New cChart
        Set objChart.Chart = picGraph.Picture
        objChart.Timestamp = Now
        mobjChartCache.Remove sSymbol + "-" + Format(mbOneDay)
        mobjChartCache.Add objChart, sSymbol + "-" + Format(mbOneDay)
    End If
    Show
    timClose.Enabled = True

End Sub

Private Sub Z_ShowDaySummary()
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2008
'
'****************************************************************************
'
'                     NAME: Sub Z_ShowDaySummary
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    28 January 2024   First created for StockTicker
'
'                  PURPOSE: Shows a summary of the day position
'
'****************************************************************************
'
'
Dim lWidth&, lLeft&, lTop&, lHeight&, lWidest&, lTableTop&
Dim lTextColor&, lUpColor&, lDownColor&, lUpArrowColor&, lDownArrowColor&
Dim objStock As cStock
Dim rRate#, rTotalChange#, rChangeValue#
Dim sCurrencyName$, sCurrencySymbol$, sSortOrder$
Dim i%, iSortColumn%

    On Error Resume Next
    timClose.Enabled = False
        
    '
    ' Draw the useful text
    '
    Cls
    ForeColor = vbWhite
    FontSize = 11
    lTextColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_TEXT_COLOUR, Format(vbWhite)))
    lUpColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_UP_COLOUR, Format(vbGreen)))
    lDownColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_DOWN_COLOUR, Format(vbRed)))
    lUpArrowColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_UP_ARROW_COLOUR, Format(vbGreen)))
    lDownArrowColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_DOWN_ARROW_COLOUR, Format(vbRed)))
    sSortOrder = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_DAY_SUMMARY_SORT_ORDER, "asc")
    iSortColumn = CInt(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_DAY_SUMMARY_SORT_COLUMN, "0"))
    sCurrencySymbol = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_CURRENCY_SYMBOL, "£")
        
    Z_ShowHeaders
    lblHeader(0).Move 10, 8
    lblHeader(1).Move 65, lblHeader(0).Top
    lblHeader(2).Move 140, lblHeader(0).Top
    lblHeader(3).Move 210, lblHeader(0).Top
    lTableTop = lblHeader(0).Top + lblHeader(0).Height + 2
    
    '
    ' Show the sort indicator
    '
    Z_ShowSortIndicator sSortOrder = "asc", lblHeader(iSortColumn)
    
    '
    ' Each stock sumarised for the day
    '
    CurrentY = lTableTop + 2
    CurrentX = 10
    rTotalChange = 0
    For Each objStock In Z_SortDaySummaryCollection(frmMain.mobjSummaryStocks, sSortOrder = "asc", iSortColumn)
        ForeColor = lTextColor
        CurrentX = 10
        Print objStock.DisplayName;
        CurrentX = 65
        
        ForeColor = IIf(objStock.CurrentPrice > objStock.DayStart, lUpColor, IIf(objStock.CurrentPrice < objStock.DayStart, lDownColor, lTextColor))
        Print " " & objStock.FormattedPrice;
        CurrentX = 140
        rChangeValue = ConvertCurrency(objStock, objStock.DayChange) * objStock.NumberOfShares
        Print FormatCurrencyValue(sCurrencySymbol, rChangeValue);
        CurrentX = 210
        Print "(" & FormatCurrencyValue(objStock.CurrencySymbol, objStock.DayChange);
        CurrentX = 260
        Print Format(objStock.DayChange / objStock.DayStart, "0.00%)");
        
        rTotalChange = rTotalChange + rChangeValue
        lWidest = IIf(CurrentX > lWidest, CurrentX, lWidest)
        Print ""
    Next
    lWidest = lWidest + 5
    
    '
    ' Add a line under the title
    '
    ForeColor = lTextColor
    lTop = CurrentY
    Line (10, lTableTop)-(lWidest, lTableTop)
    CurrentY = lTop
    
    '
    ' Add the total value
    '
    FontBold = True
    CurrentY = CurrentY + 5
    Line (10, CurrentY)-(lWidest, CurrentY)
    CurrentY = CurrentY + 5
    CurrentX = 10
    FontBold = True
    Print "Total: ";
    ForeColor = IIf(rTotalChange > 0, lUpColor, IIf(rTotalChange < 0, lDownColor, lTextColor))
    Print FormatCurrencyValue(sCurrencySymbol, rTotalChange) & " (" + Format(rTotalChange / (frmMain.mobjTotal.TotalValue - rTotalChange), "0.00%") + ")"
    
    '
    ' Resize the window to match the content
    '
    lLeft = mlLeftPos
    lTop = frmMain.Top + frmMain.Height
    Move lLeft * Screen.TwipsPerPixelX, lTop, (lWidest + 10) * Screen.TwipsPerPixelX, (CurrentY + 10) * Screen.TwipsPerPixelY
    
    Show
    timClose.Enabled = True

End Sub

Private Sub Z_ShowSummary()
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2008
'
'****************************************************************************
'
'                     NAME: Sub Z_ShowSummary
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    28 January 2024   First created for StockTicker
'
'                  PURPOSE: Shows a summary of the position
'
'****************************************************************************
'
'
Dim lWidth&, lLeft&, lTop&, lHeight&, lWidest&, lTableTop&
Dim lTextColor&, lUpColor&, lDownColor&, lUpArrowColor&, lDownArrowColor&
Dim objStock As cStock
Dim rMargin#, rTotalInvestment#
Dim sCurrencyName$, sCurrencySymbol$, sSortOrder$
Dim i%, iSortColumn%

    On Error Resume Next
    timClose.Enabled = False
        
    '
    ' Draw the useful text
    '
    Cls
    ForeColor = vbWhite
    FontSize = 11
    lTextColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_TEXT_COLOUR, Format(vbWhite)))
    lUpColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_UP_COLOUR, Format(vbGreen)))
    lDownColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_DOWN_COLOUR, Format(vbRed)))
    lUpArrowColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_UP_ARROW_COLOUR, Format(vbGreen)))
    lDownArrowColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_DOWN_ARROW_COLOUR, Format(vbRed)))
    sSortOrder = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_SORT_ORDER, "asc")
    iSortColumn = CInt(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_SORT_COLUMN, "0"))
    sCurrencySymbol = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_CURRENCY_SYMBOL, "£")
    If PSGEN_IsCommaLocale Then
        rTotalInvestment = CDbl("0" + Replace(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_TOTAL, "0"), ".", ","))
        rMargin = CDbl("0" + Replace(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_MARGIN, "0"), ".", ","))
    Else
        rTotalInvestment = CDbl("0" + Replace(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_TOTAL, "0"), ",", "."))
        rMargin = CDbl("0" + Replace(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_MARGIN, "0"), ",", "."))
    End If
        
    Z_ShowHeaders
    lblSummaryHeader(0).Move 10, 8
    lblSummaryHeader(1).Move 70, lblSummaryHeader(0).Top
    lblSummaryHeader(2).Move 140, lblSummaryHeader(0).Top
    lblSummaryHeader(3).Move 210, lblSummaryHeader(0).Top
    lblSummaryHeader(4).Move 280, lblSummaryHeader(0).Top
    lblSummaryHeader(5).Move 370, lblSummaryHeader(0).Top
    
    lTableTop = lblSummaryHeader(0).Top + lblSummaryHeader(0).Height + 2
    
    '
    ' Show the sort indicator
    '
    Z_ShowSortIndicator sSortOrder = "asc", lblSummaryHeader(iSortColumn)
    
    '
    ' Each stock sumarised for the day
    '
    CurrentY = lTableTop + 2
    CurrentX = 10
    For Each objStock In Z_SortSummaryCollection(frmMain.mobjSummaryStocks, sSortOrder = "asc", iSortColumn)
        ForeColor = lTextColor
        CurrentX = 10
        Print objStock.DisplayName;
        
        CurrentX = 70
        ForeColor = lTextColor
        Print objStock.FormattedAverageCost;
        
        CurrentX = 140
        ForeColor = IIf(objStock.CurrentPrice > objStock.AverageCost, lUpColor, IIf(objStock.CurrentPrice < objStock.AverageCost, lDownColor, lTextColor))
        Print " " & objStock.FormattedPrice;
        
        CurrentX = 210
        ForeColor = lTextColor
        Print Format(objStock.NumberOfShares, "#,###,##0");
        
        CurrentX = 280
        Print FormatCurrencyValue(sCurrencySymbol, ConvertCurrency(objStock, objStock.CurrentPrice) * objStock.NumberOfShares);
        
        CurrentX = 370
        ForeColor = IIf(objStock.CurrentPrice > objStock.AverageCost, lUpColor, IIf(objStock.CurrentPrice < objStock.AverageCost, lDownColor, lTextColor))
        Print objStock.FormattedLossPercent;
        
        lWidest = IIf(CurrentX > lWidest, CurrentX, lWidest)
        Print ""
    Next
    lWidest = lWidest + 30
    
    '
    ' Add a line under the title
    '
    ForeColor = lTextColor
    lTop = CurrentY
    Line (10, lTableTop)-(lWidest, lTableTop)
    CurrentY = lTop
    
    '
    ' Add the total value
    '
    CurrentY = CurrentY + 5
    Line (10, CurrentY)-(lWidest, CurrentY)
    CurrentY = CurrentY + 5
    
    CurrentX = 10
    ForeColor = lTextColor
    If rTotalInvestment > 0 Then
        Print "Investment: " + FormatCurrencyValue(sCurrencySymbol, rTotalInvestment + rMargin) + " ";
    End If
    Print "Cost: " + frmMain.mobjTotal.FormattedCost;
    Print " Val: " + frmMain.mobjTotal.FormattedValue
    
    CurrentX = 10
    Print "Summary: ";
    ForeColor = IIf(frmMain.mobjTotal.TotalValue > frmMain.mobjTotal.TotalCost, lUpColor, IIf(frmMain.mobjTotal.TotalValue < frmMain.mobjTotal.TotalCost, lDownColor, lTextColor))
    Print frmMain.mobjTotal.FormattedLoss;
    If rTotalInvestment > 0 Then
        ForeColor = IIf(frmMain.mobjTotal.TotalValue > (rTotalInvestment + rMargin), lUpColor, IIf(frmMain.mobjTotal.TotalValue < (rTotalInvestment + rMargin), lDownColor, lTextColor))
        Print " (" + frmMain.mobjTotal.FormattedLossAdjusted(rTotalInvestment + rMargin) + ")";
    End If
        
    ForeColor = IIf(frmMain.mobjTotal.TotalValue > frmMain.mobjTotal.TotalCost, lUpColor, IIf(frmMain.mobjTotal.TotalValue < frmMain.mobjTotal.TotalCost, lDownColor, lTextColor))
    Print "  " + frmMain.mobjTotal.FormattedLossPercent
            
    '
    ' Resize the window to match the content
    '
    lLeft = mlLeftPos
    lTop = frmMain.Top + frmMain.Height
    Move lLeft * Screen.TwipsPerPixelX, lTop, (lWidest + 10) * Screen.TwipsPerPixelX, (CurrentY + 10) * Screen.TwipsPerPixelY
    
    Show
    timClose.Enabled = True

End Sub

Private Function Z_SortDaySummaryCollection(ByVal objStocks As Collection, ByVal bAscending As Boolean, ByVal iColumn%)

    Dim objReturn As Collection
    Dim i%, j%
    Dim bFoundSlot As Boolean
    Dim objSlot As cStock
    Dim objNew As cStock

    Set objReturn = New Collection
    For i = 1 To objStocks.Count
        Set objNew = objStocks.Item(i)
        If i = 1 Then
            objReturn.Add objNew
        
        ElseIf bAscending Then
            For j = 1 To objReturn.Count
                Set objSlot = objReturn(j)
                If iColumn = 3 Then
                    bFoundSlot = ConvertCurrency(objNew, objNew.DayChange / objNew.DayStart) < ConvertCurrency(objSlot, objSlot.DayChange / objSlot.DayStart)
                ElseIf iColumn = 2 Then
                    bFoundSlot = ConvertCurrency(objNew, objNew.DayChange * objNew.NumberOfShares) < ConvertCurrency(objSlot, objSlot.DayChange * objSlot.NumberOfShares)
                ElseIf iColumn = 1 Then
                    bFoundSlot = ConvertCurrency(objNew, objNew.CurrentPrice) < ConvertCurrency(objSlot, objSlot.CurrentPrice)
                Else
                    bFoundSlot = StrComp(objNew.DisplayName, objSlot.DisplayName) < 0
                End If
                If bFoundSlot Then
                    Exit For
                End If
            Next j
            If j > objReturn.Count Then
                objReturn.Add objNew
            Else
                objReturn.Add objNew, Before:=j
            End If
        
        Else
            For j = objReturn.Count To 1 Step -1
                Set objSlot = objReturn(j)
                If iColumn = 3 Then
                    bFoundSlot = ConvertCurrency(objNew, objNew.DayChange / objNew.DayStart) < ConvertCurrency(objSlot, objSlot.DayChange / objSlot.DayStart)
                ElseIf iColumn = 2 Then
                    bFoundSlot = ConvertCurrency(objNew, objNew.DayChange * objNew.NumberOfShares) < ConvertCurrency(objSlot, objSlot.DayChange * objSlot.NumberOfShares)
                ElseIf iColumn = 1 Then
                    bFoundSlot = ConvertCurrency(objNew, objNew.CurrentPrice) < ConvertCurrency(objSlot, objSlot.CurrentPrice)
                Else
                    bFoundSlot = StrComp(objNew.DisplayName, objSlot.DisplayName) < 0
                End If
                If bFoundSlot Then
                    Exit For
                End If
            Next j
            If j < 1 Then
                objReturn.Add objNew, Before:=1
            Else
                objReturn.Add objNew, After:=j
            End If
        End If
        
    Next i

    Set Z_SortDaySummaryCollection = objReturn
End Function

Private Function Z_SortSummaryCollection(ByVal objStocks As Collection, ByVal bAscending As Boolean, ByVal iColumn%)

    Dim objReturn As Collection
    Dim i%, j%
    Dim bFoundSlot As Boolean
    Dim objSlot As cStock
    Dim objNew As cStock

    Set objReturn = New Collection
    For i = 1 To objStocks.Count
        Set objNew = objStocks.Item(i)
        If i = 1 Then
            objReturn.Add objNew
        
        ElseIf bAscending Then
            For j = 1 To objReturn.Count
                Set objSlot = objReturn(j)
                If iColumn = 5 Then
                    bFoundSlot = ConvertCurrency(objNew, objNew.LossPercent) < ConvertCurrency(objSlot, objSlot.LossPercent)
                ElseIf iColumn = 4 Then
                    bFoundSlot = ConvertCurrency(objNew, objNew.TotalValue) < ConvertCurrency(objSlot, objSlot.TotalValue)
                ElseIf iColumn = 3 Then
                    bFoundSlot = objNew.NumberOfShares < objSlot.NumberOfShares
                ElseIf iColumn = 2 Then
                    bFoundSlot = ConvertCurrency(objNew, objNew.CurrentPrice) < ConvertCurrency(objSlot, objSlot.CurrentPrice)
                ElseIf iColumn = 1 Then
                    bFoundSlot = ConvertCurrency(objNew, objNew.AverageCost) < ConvertCurrency(objSlot, objSlot.AverageCost)
                Else
                    bFoundSlot = StrComp(objNew.DisplayName, objSlot.DisplayName) < 0
                End If
                If bFoundSlot Then
                    Exit For
                End If
            Next j
            If j > objReturn.Count Then
                objReturn.Add objNew
            Else
                objReturn.Add objNew, Before:=j
            End If
        
        Else
            For j = objReturn.Count To 1 Step -1
                Set objSlot = objReturn(j)
                If iColumn = 5 Then
                    bFoundSlot = ConvertCurrency(objNew, objNew.LossPercent) < ConvertCurrency(objSlot, objSlot.LossPercent)
                ElseIf iColumn = 4 Then
                    bFoundSlot = ConvertCurrency(objNew, objNew.TotalValue) < ConvertCurrency(objSlot, objSlot.TotalValue)
                ElseIf iColumn = 3 Then
                    bFoundSlot = objNew.NumberOfShares < objSlot.NumberOfShares
                ElseIf iColumn = 2 Then
                    bFoundSlot = ConvertCurrency(objNew, objNew.CurrentPrice) < ConvertCurrency(objSlot, objSlot.CurrentPrice)
                ElseIf iColumn = 1 Then
                    bFoundSlot = ConvertCurrency(objNew, objNew.AverageCost) < ConvertCurrency(objSlot, objSlot.AverageCost)
                Else
                    bFoundSlot = StrComp(objNew.DisplayName, objSlot.DisplayName) < 0
                End If
                If bFoundSlot Then
                    Exit For
                End If
            Next j
            If j < 1 Then
                objReturn.Add objNew, Before:=1
            Else
                objReturn.Add objNew, After:=j
            End If
        End If
        
    Next i

    Set Z_SortSummaryCollection = objReturn
End Function

Private Sub Z_ShowHeaders()

Dim i%
Dim lTextColor&

    lTextColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_TEXT_COLOUR, Format(vbWhite)))
    For i = 0 To lblSummaryHeader.UBound
        lblSummaryHeader(i).BackColor = BackColor
        lblSummaryHeader(i).ForeColor = lTextColor
        Set lblSummaryHeader(i).Font = Font
        lblSummaryHeader(i).Height = TextHeight(lblSummaryHeader(i).Caption)
        lblSummaryHeader(i).Visible = mbSummary
    Next i
    
    For i = 0 To lblHeader.UBound
        lblHeader(i).BackColor = BackColor
        lblHeader(i).ForeColor = lTextColor
        Set lblHeader(i).Font = Font
        lblHeader(i).Height = TextHeight(lblHeader(i).Caption)
        lblHeader(i).Visible = mbDaySummary
    Next i
    
    picGraph.Visible = Not mbSummary And Not mbDaySummary
    picSort.Visible = mbSummary Or mbDaySummary

End Sub

Private Sub Z_ShowSortIndicator(ByVal bAscending As Boolean, ByVal objLabel As Label)

Dim i%
Dim lTextColor&

    lTextColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_TEXT_COLOUR, Format(vbWhite)))
    picSort.BackColor = BackColor
    picSort.ForeColor = ForeColor
    picSort.DrawWidth = 1
    If bAscending Then
        For i = 0 To picSort.ScaleHeight \ 4
           picSort.Line ((picSort.ScaleHeight \ 4) - i, picSort.ScaleHeight - i - 3)-((picSort.ScaleHeight \ 4) + i, picSort.ScaleHeight - i - 3), lTextColor
        Next i
        picSort.DrawWidth = 2
        picSort.Line ((picSort.ScaleHeight \ 4), 3)-((picSort.ScaleHeight \ 4), 3 * picSort.ScaleHeight \ 4), lTextColor
    Else
        For i = 0 To picSort.ScaleHeight \ 4
           picSort.Line ((picSort.ScaleHeight \ 4) - i, i + 3)-((picSort.ScaleHeight \ 4) + i, i + 3), lTextColor
        Next i
        picSort.DrawWidth = 2
        picSort.Line ((picSort.ScaleHeight \ 4), picSort.ScaleHeight \ 4)-((picSort.ScaleHeight \ 4), picSort.ScaleHeight - 4), lTextColor
    End If
    picSort.Move objLabel.Left + TextWidth(objLabel.Caption) + 4, objLabel.Top, 10, objLabel.Height
    picSort.Visible = True
    picSort.ZOrder 1000
    
End Sub
