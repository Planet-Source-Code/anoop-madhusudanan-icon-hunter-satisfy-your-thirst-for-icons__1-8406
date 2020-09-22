VERSION 5.00
Begin VB.Form frmAuthor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About The Author"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   Icon            =   "frmAuthor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   4425
      Left            =   45
      ScaleHeight     =   4365
      ScaleWidth      =   6390
      TabIndex        =   0
      Top             =   60
      Width           =   6450
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAuthor.frx":000C
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
         Height          =   465
         Index           =   5
         Left            =   135
         TabIndex        =   7
         Top             =   2910
         Width           =   6135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAuthor.frx":00AA
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
         Height          =   735
         Index           =   4
         Left            =   135
         TabIndex        =   6
         Top             =   3480
         Width           =   5865
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAuthor.frx":017F
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
         Height          =   870
         Index           =   3
         Left            =   135
         TabIndex        =   5
         Top             =   1965
         Width           =   6105
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Hi,"
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
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   4
         Top             =   1680
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Anoop M, anoopj12@yahoo.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   750
         Width           =   2625
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Govindanikethan, Nedumkunnam, Kottayam, Kerala, India - 686 542"
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
         Height          =   480
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   1020
         Width           =   3705
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   165
         X2              =   3990
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Label lblMain 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ICON HUNTER 1.0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   150
         TabIndex        =   1
         Top             =   195
         Width           =   3000
      End
      Begin VB.Image Image1 
         Height          =   1335
         Left            =   4125
         Picture         =   "frmAuthor.frx":0295
         Top             =   105
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmAuthor"
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

