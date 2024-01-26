VERSION 5.00
Begin VB.Form frmTooltip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1200
   ClientLeft      =   11265
   ClientTop       =   9075
   ClientWidth     =   6180
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000017&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   80
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   412
   Visible         =   0   'False
   Begin VB.Timer timClose 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   960
      Top             =   480
   End
End
Attribute VB_Name = "frmTooltip"
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


    Private msCaption$
    Private mstPoint As POINTAPI

Private Sub Form_Activate()

    PSGEN_SetTopMost hWnd

End Sub

Public Sub ShowToolTip(ByVal objFont As Font, ByVal sText$)

Dim lWidth&, lHeight&, lLeft&, lTop&
Dim rTextWidth#, rTextHeight#

    '
    ' Check if we need to do anything
    '
    timClose.Enabled = False
    If Not Visible Or sText <> msCaption Then

        '
        ' Get all the metrics
        '
        Cls
        Font.Charset = objFont.Charset
        Font.size = Font.size
        Font.Bold = False
        lWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN)
        lHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN)
        rTextHeight = TextHeight(sText) + 4
        rTextWidth = TextWidth(sText) + 6
        
        '
        ' Position the tooltip
        '
        Call GetCursorPos(mstPoint)
        lLeft = mstPoint.X - (rTextWidth / 3)
        If lLeft + rTextWidth > lWidth Then
            lLeft = lWidth - rTextWidth
        End If
        
        lTop = mstPoint.Y + 30
        If lTop + rTextHeight > lHeight Then
            lTop = mstPoint.Y - rTextHeight - 30
        End If
        Move lLeft * Screen.TwipsPerPixelX, lTop * Screen.TwipsPerPixelY, rTextWidth * Screen.TwipsPerPixelX, rTextHeight * Screen.TwipsPerPixelY
        
        '
        ' Print the text
        '
        CurrentX = 2
        CurrentY = 1
        Print sText;
        msCaption = sText
        Visible = True
    End If
    
    If Visible And sText <> "" Then
        Call GetCursorPos(mstPoint)
        timClose.Enabled = True
        PSGEN_SetTopMost hWnd
    End If
    
End Sub

Public Sub HideTooltip()
    
    Visible = False
    timClose.Enabled = False
    
End Sub

Private Sub timClose_Timer()
    
Dim stPoint As POINTAPI
    
    If Visible Then
        Call GetCursorPos(stPoint)
        If stPoint.X <> mstPoint.X Or stPoint.Y <> mstPoint.Y Then
            HideTooltip
        End If
    Else
        HideTooltip
    End If

End Sub
