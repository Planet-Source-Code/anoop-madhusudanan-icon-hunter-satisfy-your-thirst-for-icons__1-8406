VERSION 5.00
Begin VB.Form frmDirCreate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Directory"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmDirCreate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   3120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Create"
      Height          =   405
      Left            =   1710
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   330
      Left            =   1515
      TabIndex        =   2
      Top             =   300
      Width           =   2835
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   150
      Picture         =   "frmDirCreate.frx":000C
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   540
   End
   Begin VB.Label lblName 
      Caption         =   "Name  :"
      Height          =   255
      Left            =   825
      TabIndex        =   1
      Top             =   345
      Width           =   675
   End
End
Attribute VB_Name = "frmDirCreate"
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

Dim ParDir As String

Function CreateDir(CurDirn As String)
txt.Text = "New Folder"
Me.Caption = "Create In " & CurDirn
ParDir = CurDirn
Me.Show vbModal
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo Handler
    If Right(ParDir, 1) <> "\" Then
     MkDir ParDir & "\" & txt.Text
    frmDir.dirMain.Path = ParDir & "\" & txt.Text
    Else
     MkDir ParDir & txt.Text
     frmDir.dirMain.Path = ParDir & txt.Text
    End If

    Unload Me
    Exit Sub
    
Handler:
    MsgBox "Unable to create folder. Please check that whether the name you entered is valid and contains no invalid characters", vbInformation + vbOKOnly, "Cannot Create"
End Sub


Private Sub txt_GotFocus()
txt.SelStart = 0
txt.SelLength = Len(txt.Text)
End Sub
