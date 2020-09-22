VERSION 5.00
Begin VB.Form frmDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Directory"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmDir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create"
      Height          =   435
      Left            =   1875
      TabIndex        =   7
      Top             =   4710
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   3180
      TabIndex        =   4
      Top             =   4710
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   435
      Left            =   570
      TabIndex        =   3
      Top             =   4710
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3795
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   4320
      Begin VB.DirListBox dirMain 
         Height          =   3015
         Left            =   150
         TabIndex        =   2
         Top             =   660
         Width           =   4005
      End
      Begin VB.DriveListBox drvMain 
         Height          =   315
         Left            =   165
         TabIndex        =   1
         Top             =   255
         Width           =   4020
      End
   End
   Begin VB.Label lblDir 
      Caption         =   "Selected Directory  :"
      Height          =   480
      Left            =   90
      TabIndex        =   6
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Selected Directory  :"
      Height          =   285
      Left            =   105
      TabIndex        =   5
      Top             =   90
      Width           =   1530
   End
End
Attribute VB_Name = "frmDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================================================================
'
' Developed by Anoop. M
' anoopj12 @ yahoo.com
'
' Anoop M, Govindanikethan, Nedumkunnam P.O, Kottayam,
' Kerala, India - 686 542
'
' Hey sir, Kindly rate this code, if you like it.
'
' READ THIS BEFORE USING THE CODE:
'
' You can study and view the source code for creating your
' own apps, but do not reproduce/release Icon Hunter fully
' or partially for any commercial and/or personal purposes. All
' rights of this product is related to it's author. Any violation
' of above conditions will be treated seriously and is punishable.
'
' I do not have full time to add complete explanation, read the help
' file (click Help->Contents) in Icon Hunter. Contact me for
' additional help/suggestions
'
' I recently inveted a technology for streaming audio, and is
' now looking promoters/investors to invest in a web-phone network
' project.
'
' VISIT MY WEBSITE : http://www.geocities.com/streamingaudio for details

'=============================================================================================================================

Public TheDir As String

Private Sub cmdCancel_Click()
TheDir = ""
Unload Me
End Sub

Private Sub cmdCreate_Click()
frmDirCreate.CreateDir dirMain.Path
dirMain.Refresh

End Sub

Private Sub cmdOK_Click()
TheDir = dirMain.Path
Unload Me
End Sub

Private Sub dirMain_Change()
TheDir = dirMain.Path
lblDir = TheDir
End Sub

Private Sub drvMain_Change()
dirMain.Path = drvMain.Drive
End Sub

Private Sub Form_Load()
dirMain.Path = CurDir
TheDir = CurDir
lblDir = CurDir
End Sub

Public Function ShowDir(Optional Cap As String = "Select Folder") As String

    Me.Caption = Cap
    Me.Show vbModal
    ShowDir = TheDir
    
End Function
