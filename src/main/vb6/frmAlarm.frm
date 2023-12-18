VERSION 5.00
Begin VB.Form frmAlarm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alarm"
   ClientHeight    =   3630
   ClientLeft      =   4800
   ClientTop       =   3105
   ClientWidth     =   6555
   FillColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCommand 
      Caption         =   "Accept && Disable"
      Height          =   405
      Index           =   2
      Left            =   4710
      TabIndex        =   3
      ToolTipText     =   "Accept this alarm"
      Top             =   2940
      Width           =   1485
   End
   Begin VB.CommandButton cmdCommand 
      Caption         =   "Accept"
      Default         =   -1  'True
      Height          =   405
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      ToolTipText     =   "Accept this alarm"
      Top             =   2940
      Width           =   1485
   End
   Begin VB.CommandButton cmdCommand 
      Caption         =   "Reset Alarm to New Value"
      Height          =   405
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   2940
      Width           =   2385
   End
   Begin VB.Frame frmAlarm 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Day Position"
      ForeColor       =   &H00E0E0E0&
      Height          =   1965
      Index           =   1
      Left            =   3390
      TabIndex        =   5
      Top             =   210
      Width           =   2805
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Day Position"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   11
         Left            =   60
         TabIndex        =   26
         Top             =   60
         Width           =   2265
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gain / Loss"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   15
         Top             =   1590
         Width           =   1005
      End
      Begin VB.Label lblValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   1410
         TabIndex        =   14
         Top             =   1590
         Width           =   1275
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   13
         Top             =   390
         Width           =   1005
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Low"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "High"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   990
         Width           =   1005
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Change"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1290
         Width           =   1005
      End
      Begin VB.Label lblValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   1410
         TabIndex        =   9
         Top             =   390
         Width           =   1275
      End
      Begin VB.Label lblValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   1410
         TabIndex        =   8
         Top             =   690
         Width           =   1275
      End
      Begin VB.Label lblValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   1410
         TabIndex        =   7
         Top             =   990
         Width           =   1275
      End
      Begin VB.Label lblValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   1410
         TabIndex        =   6
         Top             =   1290
         Width           =   1275
      End
   End
   Begin VB.Frame frmAlarm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Overall Position"
      ForeColor       =   &H00FFFFFF&
      Height          =   1965
      Index           =   0
      Left            =   330
      TabIndex        =   4
      Top             =   210
      Width           =   2805
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Overall Position"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   10
         Left            =   90
         TabIndex        =   25
         Top             =   60
         Width           =   2265
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   390
         Width           =   1005
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cost"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Shares"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   990
         Width           =   1005
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gain / Loss"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   1290
         Width           =   1005
      End
      Begin VB.Label lblValue 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   0
         Left            =   1410
         TabIndex        =   20
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label lblValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   1410
         TabIndex        =   19
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label lblValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   1410
         TabIndex        =   18
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label lblValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   1410
         TabIndex        =   17
         Top             =   1290
         Width           =   1215
      End
   End
   Begin VB.TextBox txtSymbol 
      Height          =   285
      Left            =   1470
      TabIndex        =   0
      Text            =   "wqewee"
      Top             =   2430
      Width           =   945
   End
   Begin VB.Label lblPercent 
      BackStyle       =   0  'Transparent
      Caption         =   "Percent"
      Height          =   195
      Left            =   2580
      TabIndex        =   27
      Top             =   2490
      Width           =   1005
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Alarm Level"
      Height          =   195
      Index           =   9
      Left            =   330
      TabIndex        =   16
      Top             =   2490
      Width           =   1005
   End
End
Attribute VB_Name = "frmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim mbytSound() As Byte ' Always store binary data in byte arrays!
    
    Dim mobjReg As New cRegistry
    Dim mobjSymbol As cSymbol
    Dim mbIsLow As Boolean


Private Sub cmdCommand_Click(Index As Integer)

    If Index = 0 Then
        If mbIsLow Then
            mobjSymbol.LowAlarmValue = txtSymbol.Text
        Else
            mobjSymbol.HighAlarmValue = txtSymbol.Text
        End If
        mobjSymbol.Save
    
    ElseIf Index = 2 Then
        If mbIsLow Then
            mobjSymbol.LowAlarmEnabled = False
        Else
            mobjSymbol.HighAlarmEnabled = False
        End If
        mobjSymbol.Save
    
    End If
    Unload Me

End Sub

Private Sub Form_Load()

    '
    ' Position the form in the middle of the display
    '
    CentreForm Me
    PSGEN_SetTopMost hWnd

End Sub


Public Sub ShowLowAlarm(ByVal objSymbol As cSymbol)

Dim sFilename$

    '
    ' Set the display
    '
    mbIsLow = True
    Caption = "Low Alarm " + objSymbol.Code
    txtSymbol.Text = objSymbol.LowAlarmValue
    If objSymbol.LowAlarmSoundEnabled Then
        sFilename = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_LOW_ALARM_WAVE)
        If sFilename <> "" And PSGEN_FileExists(sFilename) Then
            Call Z_PlayWave(sFilename)
        Else
            Call Z_PlayWaveRes("LOWALARM")
        End If
    End If
    lblPercent.Visible = objSymbol.LowAlarmIsPercent
    Call Z_DisplaySymbol(objSymbol)

End Sub


Public Sub ShowHighAlarm(ByVal objSymbol As cSymbol)

Dim sFilename$

    '
    ' Set the display
    '
    Caption = "High Alarm " + objSymbol.Code
    txtSymbol.Text = objSymbol.HighAlarmValue
    If objSymbol.HighAlarmSoundEnabled Then
        sFilename = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_HIGH_ALARM_WAVE)
        If sFilename <> "" And PSGEN_FileExists(sFilename) Then
            Call Z_PlayWave(sFilename)
        Else
            Call Z_PlayWaveRes("HIGHALARM")
        End If
    End If
    lblPercent.Visible = objSymbol.HighAlarmIsPercent
    Call Z_DisplaySymbol(objSymbol)

End Sub


Private Sub Z_DisplaySymbol(ByVal objSymbol As cSymbol)

Dim lUpColor&, lDownColor&

    If objSymbol.AlarmShowing Then
        Unload Me
    Else
        Set mobjSymbol = objSymbol
        mobjSymbol.AlarmShowing = True
    
        '
        ' Get the up down colours
        '
        lUpColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_UP_COLOUR, Format(vbGreen)))
        lDownColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_DOWN_COLOUR, Format(vbRed)))
        
        '
        ' Set the display
        '
        lblValue(0).Caption = objSymbol.FormattedValue
        lblValue(1).Caption = objSymbol.FormattedCost
        lblValue(2).Caption = Format(objSymbol.Shares)
        lblValue(3).ForeColor = IIf(objSymbol.CurrentPrice > objSymbol.Price, lUpColor, IIf(objSymbol.CurrentPrice < objSymbol.Price, lDownColor, frmPreview.ForeColor))
        lblValue(3).Caption = FormatCurrencyValueWithSymbol(objSymbol.CurrencySymbol, objSymbol.CurrencyName, objSymbol.Shares * (objSymbol.CurrentPrice - objSymbol.Price))
    
        lblValue(4).Caption = FormatCurrencyValue(objSymbol.CurrencySymbol, objSymbol.DayStart)
        lblValue(5).Caption = FormatCurrencyValue(objSymbol.CurrencySymbol, objSymbol.DayLow)
        lblValue(6).Caption = FormatCurrencyValue(objSymbol.CurrencySymbol, objSymbol.DayHigh)
        
        lblValue(7).ForeColor = IIf(objSymbol.DayChange > 0, lUpColor, IIf(objSymbol.DayChange < 0, lDownColor, frmPreview.ForeColor))
        lblValue(7).Caption = FormatCurrencyValue(objSymbol.CurrencySymbol, objSymbol.DayChange) + " " + Format(objSymbol.DayChange / IIf(objSymbol.DayStart = 0, 1, objSymbol.DayStart), "0.00%")
        lblValue(8).ForeColor = lblValue(7).ForeColor
        lblValue(8).Caption = FormatCurrencyValueWithSymbol(objSymbol.CurrencySymbol, objSymbol.CurrencyName, (objSymbol.CurrentPrice * objSymbol.Shares) - (objSymbol.DayStart * objSymbol.Shares))
    
        Visible = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    mobjSymbol.AlarmShowing = False
    Z_PlayWaveRes

End Sub


Private Sub Z_PlayWaveRes(Optional vntResourceID As Variant, Optional vntFlags)
    If Not IsMissing(vntResourceID) Then
        mbytSound = LoadResData(vntResourceID, "WAVE")
        If IsMissing(vntFlags) Then vntFlags = SND_NODEFAULT Or SND_ASYNC Or SND_MEMORY Or SND_LOOP
        If (vntFlags And SND_MEMORY) = 0 Then vntFlags = vntFlags Or SND_MEMORY
        sndPlaySound mbytSound(0), vntFlags
    Else
        sndPlaySound ByVal vbNullString, 0&
    End If
End Sub

Private Sub Z_PlayWave(ByVal sFilename$, Optional vntFlags)
    If IsMissing(vntFlags) Then vntFlags = SND_NODEFAULT Or SND_ASYNC Or SND_LOOP
    sndPlaySound ByVal sFilename, vntFlags
End Sub


