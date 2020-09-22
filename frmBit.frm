VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBit 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Bitmap View"
   ClientHeight    =   3720
   ClientLeft      =   165
   ClientTop       =   690
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMain 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   870
      ScaleHeight     =   1815
      ScaleWidth      =   1905
      TabIndex        =   1
      Top             =   735
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   3450
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   476
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Image img 
      Height          =   3090
      Left            =   30
      Stretch         =   -1  'True
      Top             =   135
      Width           =   4860
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save As"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuNext 
         Caption         =   "&Next Icon                   Spacebar"
      End
      Begin VB.Menu mnuPrev 
         Caption         =   "&Previous Icon             Backspace"
      End
      Begin VB.Menu mnuVB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFirst 
         Caption         =   "&First Icon"
      End
      Begin VB.Menu mnuLast 
         Caption         =   "&Last Icon"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuKeep 
         Caption         =   "&Keep On Top"
         Checked         =   -1  'True
         Shortcut        =   ^K
      End
   End
End
Attribute VB_Name = "frmBit"
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

Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public i As Integer

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
    mnuNext_Click
End If

If KeyCode = vbKeyBack Then
    mnuPrev_Click
End If

End Sub

Private Sub Form_Load()
frmMain.BitLoaded = True
Me.Height = GetSetting(App.EXEName, "Bitmapview", "Height", 200)
Me.Width = GetSetting(App.EXEName, "Bitmapview", "Width", 200)
Form_Resize
ToggleState

End Sub

Private Sub Form_Paint()
ToggleState
End Sub

Private Sub Form_Resize()
img.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - sbMain.Height
Me.Caption = "Enhanced View - [Height=" & img.Height & ", Width=" & img.Width & "]"
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting App.EXEName, "Bitmapview", "Height", Me.Height
SaveSetting App.EXEName, "Bitmapview", "Width", Me.Width
frmMain.BitLoaded = False
End Sub

Private Sub mnuCopy_Click()
'On Error Resume Next
With frmMain
Clipboard.Clear
.picLarge.Picture = img.Picture
Clipboard.SetData .picLarge.Image
End With

End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuFirst_Click()
i = 1
img.Picture = frmMain.imgLarge.ListImages(i).ExtractIcon
SetStatus i

End Sub

Private Sub mnuKeep_Click()
mnuKeep.Checked = Not mnuKeep.Checked
ToggleState

End Sub

Sub ToggleState()
        If mnuKeep.Checked Then
           SetWindowPos Me.hWnd, HWND_TOPMOST, Me.Left / 15, _
                        Me.Top / 15, Me.Width / 15, _
                        Me.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
        Else
           SetWindowPos Me.hWnd, HWND_NOTOPMOST, Me.Left / 15, _
                        Me.Top / 15, Me.Width / 15, _
                        Me.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
        End If

End Sub

Private Sub mnuLast_Click()
i = frmMain.imgLarge.ListImages.Count
img.Picture = frmMain.imgLarge.ListImages(i).ExtractIcon
SetStatus i

End Sub

Private Sub mnuNext_Click()

On Error Resume Next
If i = frmMain.imgLarge.ListImages.Count Then i = 1

i = i + 1
img.Picture = frmMain.imgLarge.ListImages(i).ExtractIcon
SetStatus i

End Sub

Private Sub mnuPrev_Click()
On Error Resume Next

If i = 1 Then i = frmMain.imgLarge.ListImages.Count

i = i - 1
img.Picture = frmMain.imgLarge.ListImages(i).ExtractIcon
SetStatus i

End Sub

Private Sub mnuSave_Click()

Dim clOpen As CommonDialog

Set clOpen = frmMain.cdlOpen

On Error GoTo nosave
clOpen.CancelError = True
clOpen.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist
clOpen.DialogTitle = "Save Icon"
clOpen.Filter = "Icon File|*.ico"
clOpen.DefaultExt = "ico"
clOpen.ShowSave


SavePicture img.Picture, clOpen.FileTitle

nosave:

End Sub


Public Sub SetStatus(cIndex As Integer)
sbMain.SimpleText = ""
    
    sbMain.SimpleText = frmMain.lvIcons.ListItems(cIndex).Text
    If sbMain.SimpleText <> "" Then sbMain.SimpleText = sbMain.SimpleText & " - "
    sbMain.SimpleText = sbMain.SimpleText & frmMain.FPathFromKey(frmMain.imgLarge.ListImages(cIndex).Key)
    
End Sub
