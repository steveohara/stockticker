VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1545
   ClientLeft      =   16065
   ClientTop       =   5805
   ClientWidth     =   4815
   HasDC           =   0   'False
   Icon            =   "psprogress.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox aniProgress 
      Height          =   885
      Left            =   360
      ScaleHeight     =   825
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   60
      Width           =   4095
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "lblCaption"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   4095
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2001
'
'****************************************************************************
'
' LANGUAGE:             Microsoft Visual Basic V6.00
'
' MODULE NAME:          Pivotal_PrintPreview
'
' MODULE TYPE:          BASIC Form
'
' FILE NAME:            PSPROGRESS.FRM
'
' MODIFICATION HISTORY: Steve O'Hara    09 May 2001   First created for TeamPlayer
'
' PURPOSE:              Provides an animated progress dialog
'
'****************************************************************************
'
'****************************************************
' MODULE VARIABLE DECLARATIONS
'****************************************************
'
Option Explicit


    '
    ' Type enumerations
    '
    Enum AnimationTypes
        AnimateFileCopy = 1
        AnimateFileMove = 2
        AnimateFileDelete = 3
        AnimateFileDeleteR = 4
        AnimateFileNuke = 5
        AnimateFileFind = 6
        AnimateFileComp = 7
        AnimateSearch = 8
    End Enum
    
    '
    ' Associated IDs
    '
    Const ANIMATION_IDS = "FILECOPY,FILEMOVE,FILEDELETE,FILEDELETER,FILENUKE,FINDFILE,FILECOMP,SEARCH"

    '
    ' Session specific files
    '
    Dim msFilename$
    Dim mbProgressVisible As Boolean
    
Public Property Get ProgressText$()
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2001
'
'****************************************************************************
'
'                     NAME: Property GET ProgressText
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    16 December 1998   First created for Barclays MIS
'
'                  PURPOSE: Caption to display
'
'****************************************************************************
'
'

    '
    ' Initialise error vector
    '
    On Error Resume Next
    ProgressText = lblCaption.Caption

End Property

Public Property Let ProgressText(ByVal sValue$)
Attribute ProgressText.VB_Description = "Caption to display"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2001
'
'****************************************************************************
'
'                     NAME: Property LET ProgressText
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    16 December 1998   First created for Barclays MIS
'
'                  PURPOSE: Caption to display
'
'****************************************************************************
'
'


    '
    ' Initialise error vector
    '
    On Error Resume Next
    lblCaption.Caption = sValue
    lblCaption.Refresh

End Property

Public Property Let AnimationType(ByVal iValue As AnimationTypes)
Attribute AnimationType.VB_Description = "Type of animation to show"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2001
'
'****************************************************************************
'
'                     NAME: Property LET AnimationType
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    16 December 1998   First created for Barclays MIS
'
'                  PURPOSE: Type of animation to show
'
'****************************************************************************
'
'

Dim abData() As Byte
Dim sID$
Dim iFile%

    '
    ' Initialise error vector and load the required resource
    '
    On Error Resume Next
'    aniProgress.Close
'    If PSGEN_FileExists(msFilename) Then Kill msFilename
'    msFilename = PSGEN_GetTempPathFilename("avi")
    
    '
    ' Get the resource from the file
    '
'    sID = PSVBUTLS_GetItem(iValue, ",", ANIMATION_IDS)
'    If sID <> "" Then
'        abData = LoadResData(sID, "AVI")
'        If Err = 0 Then
'            If UBound(abData) > 0 Then
        
                '
                ' Copy the data to the file and use to drive the animation
                '
'                iFile = FreeFile
'                Open msFilename For Binary Access Write As #iFile
'                Put #iFile, , abData
'                Close #iFile
'                Call aniProgress.Open(msFilename)
'            End If
'        End If
'    End If
'    DoEvents
    
End Property

    

Private Sub Form_Activate()

    Call PSGEN_SetTopMost(Me.hWnd)
    
End Sub

Private Sub Form_Load()
    
    '
    ' Position the form
    '
    #If Not NoMdiMain Then
        If PSGEN_FormLoaded(MdiMain) Then
            Me.Move MdiMain.Left + (MdiMain.Width / 2 - Me.Width / 2), MdiMain.Top + (MdiMain.Height / 3 - Me.Height / 2)
        Else
            Me.Move (Screen.Width / 2 - Me.Width / 2), (Screen.Height / 3 - Me.Height / 2)
        End If
    #Else
        Me.Move (Screen.Width / 2 - Me.Width / 2), (Screen.Height / 3 - Me.Height / 2)
    #End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    '
    ' Don't allow the user to unload it
    '
    Cancel = (UnloadMode = vbFormControlMenu)

End Sub


Private Sub Form_Unload(Cancel As Integer)

    '
    ' Remove the temporary avi file
    '
    On Error Resume Next
'    aniProgress.Close
    If PSGEN_FileExists(msFilename) Then Kill msFilename

End Sub


