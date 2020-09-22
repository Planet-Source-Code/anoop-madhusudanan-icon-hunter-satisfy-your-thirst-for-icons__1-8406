VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4095
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5835
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   90
      ScaleHeight     =   390
      ScaleWidth      =   4560
      TabIndex        =   4
      Top             =   105
      Width           =   4560
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Developed by Anoop M, anoopj12@yahoo.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   75
         Width           =   3390
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   705
      Picture         =   "frmSplash.frx":0000
      Top             =   1185
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   135
      X2              =   5655
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   3990
      Left            =   45
      Top             =   75
      Width           =   5745
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "See Help about Disclaimer and other details"
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
      Index           =   3
      Left            =   165
      TabIndex        =   3
      Top             =   3720
      Width           =   3540
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0 May 2000 Release"
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
      Index           =   1
      Left            =   3960
      TabIndex        =   2
      Top             =   3720
      Width           =   1785
   End
   Begin VB.Label lblMain 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ICON HUNTER "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   1350
      TabIndex        =   1
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "For Extracting And Saving Icons From Executables"
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
      Height          =   240
      Left            =   900
      TabIndex        =   0
      Top             =   2085
      Width           =   4185
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   120
      X2              =   5640
      Y1              =   2010
      Y2              =   2010
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   2190
      Index           =   2
      Left            =   3150
      Top             =   195
      Width           =   2145
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   5
      Height          =   1950
      Index           =   1
      Left            =   1950
      Top             =   1860
      Width           =   2205
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   5
      Height          =   2190
      Index           =   3
      Left            =   225
      Top             =   630
      Width           =   2145
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   5
      Height          =   1950
      Index           =   0
      Left            =   585
      Top             =   3270
      Width           =   2460
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0E0FF&
      BorderWidth     =   5
      Height          =   1980
      Index           =   4
      Left            =   3600
      Top             =   2700
      Width           =   2505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0FF&
      BorderWidth     =   5
      Height          =   855
      Index           =   5
      Left            =   -525
      Top             =   180
      Width           =   1620
   End
End
Attribute VB_Name = "frmSplash"
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

