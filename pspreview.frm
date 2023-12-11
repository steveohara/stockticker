VERSION 5.00
Begin VB.Form frmPreview 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5115
   ClientLeft      =   7830
   ClientTop       =   4245
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
    Private mobjChartCache As Collection


Private Sub Form_Activate()

    PSGEN_SetTopMost hWnd

End Sub

Private Sub Form_DblClick()

    picGraph_DblClick
    
End Sub

Private Sub picGraph_DblClick()

    mbOneDay = Not mbOneDay
    Z_ShowChart mobjStockSymbol

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
    Z_ShowChart mobjStockSymbol

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

