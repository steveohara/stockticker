VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   ClientHeight    =   3345
   ClientLeft      =   6090
   ClientTop       =   5355
   ClientWidth     =   8520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "psabout.frx":0000
   ScaleHeight     =   3345
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblWebSite 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.pivotal-solutions.co.uk/stockticker"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   420
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      ToolTipText     =   "Click here to vist the web site"
      Top             =   2580
      Width           =   4755
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   405
      TabIndex        =   1
      Top             =   1470
      Width           =   3495
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   405
      TabIndex        =   0
      Top             =   1080
      Width           =   3495
   End
End
Attribute VB_Name = "frmAbout"
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
' MODULE NAME:          Pivotal_about
'
' MODULE TYPE:          BASIC Form
'
' FILE NAME:            PSABOUT.FRM
'
' MODIFICATION HISTORY: Steve O'Hara    29 August 2008   First created for StockTicker
'
' PURPOSE:              The about box
'
'
'****************************************************************************
'
'****************************************************
' MODULE VARIABLE DECLARATIONS
'****************************************************
'
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then
        Call PSGEN_LaunchBrowser("file://" + App.Path + "/user guide/index.htm")
        KeyCode = 0
    End If
    Unload Me
    
End Sub

Private Sub Form_Load()

    lblTitle.Caption = App.Title
    lblVersion = "Version " + Format(App.Major) + "." + Format(App.Minor) + "." + Format(App.Revision)
    
    CentreForm Me
        

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Unload Me
    
End Sub

Private Sub lblWebSite_Click()

    Call PSGEN_LaunchBrowser(lblWebSite.Caption)
    Unload Me
    
End Sub

Private Sub lblWebSite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblWebSite.ForeColor = &HFF0000
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblWebSite.ForeColor = &H800000
    
End Sub

