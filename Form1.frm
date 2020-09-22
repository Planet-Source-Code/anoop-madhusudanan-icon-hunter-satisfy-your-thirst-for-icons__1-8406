VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   Caption         =   "Icon Hunter"
   ClientHeight    =   5460
   ClientLeft      =   5865
   ClientTop       =   2700
   ClientWidth     =   7920
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frIn 
      Height          =   630
      Left            =   45
      TabIndex        =   23
      Top             =   1740
      Width           =   5835
      Begin MSComCtl2.UpDown udMain 
         Height          =   330
         Left            =   3226
         TabIndex        =   27
         Top             =   195
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         BuddyControl    =   "txtMax"
         BuddyDispid     =   196610
         OrigLeft        =   3480
         OrigTop         =   225
         OrigRight       =   3720
         OrigBottom      =   540
         Max             =   2000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtMax 
         Height          =   330
         Left            =   2370
         TabIndex        =   25
         Text            =   "500"
         Top             =   195
         Width           =   855
      End
      Begin VB.Label lblInfo 
         Caption         =   "icons to save memory"
         Height          =   270
         Index           =   3
         Left            =   3645
         TabIndex        =   26
         Top             =   255
         Width           =   1650
      End
      Begin VB.Label lblInfo 
         Caption         =   "&Automatically stop after listing"
         Height          =   270
         Index           =   2
         Left            =   165
         TabIndex        =   24
         Top             =   240
         Width           =   2145
      End
   End
   Begin VB.PictureBox picSide 
      BorderStyle     =   0  'None
      Height          =   1680
      Left            =   5970
      ScaleHeight     =   1680
      ScaleWidth      =   1845
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   90
      Width           =   1845
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Start Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   45
         TabIndex        =   4
         Top             =   0
         Width           =   1530
      End
      Begin MSComCtl2.Animation anMain 
         Height          =   795
         Left            =   480
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   690
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   1402
         _Version        =   393216
         Center          =   -1  'True
         FullWidth       =   45
         FullHeight      =   53
      End
   End
   Begin VB.PictureBox picLarge 
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   -60
      ScaleHeight     =   510
      ScaleWidth      =   495
      TabIndex        =   7
      Top             =   1470
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox lstFoundFiles 
      Height          =   2205
      Left            =   4110
      TabIndex        =   20
      Top             =   585
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.PictureBox picInvisible 
      Height          =   3480
      Left            =   1680
      ScaleHeight     =   3420
      ScaleWidth      =   3075
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   3135
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   555
         ScaleHeight     =   2895
         ScaleWidth      =   3855
         TabIndex        =   14
         Top             =   2085
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   135
         ScaleHeight     =   2895
         ScaleWidth      =   3420
         TabIndex        =   15
         Top             =   285
         Width           =   3420
         Begin VB.FileListBox filList 
            Height          =   2040
            Left            =   120
            Pattern         =   "*.exe"
            TabIndex        =   18
            Top             =   480
            Width           =   1815
         End
         Begin VB.DirListBox dirList 
            Height          =   1665
            Left            =   2040
            TabIndex        =   17
            Top             =   960
            Width           =   1575
         End
         Begin VB.DriveListBox drvList 
            Height          =   315
            Left            =   2040
            TabIndex        =   16
            Top             =   480
            Width           =   1575
         End
      End
   End
   Begin MSComctlLib.StatusBar SbMain 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   5130
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      Style           =   1
      SimpleText      =   "Click the Start button to start search"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13467
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSmall 
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   450
      ScaleHeight     =   510
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   3195
      Top             =   4020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FilterIndex     =   1
   End
   Begin MSComctlLib.ImageList imgLarge 
      Left            =   2280
      Top             =   4035
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgSmall 
      Left            =   1545
      Top             =   4005
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin VB.Frame fr 
      Height          =   1740
      Left            =   45
      TabIndex        =   10
      Top             =   0
      Width           =   5835
      Begin VB.CommandButton cmdDir 
         Caption         =   "&Browse"
         Height          =   330
         Left            =   4470
         TabIndex        =   3
         Top             =   1230
         Width           =   1125
      End
      Begin VB.CheckBox chkSub 
         Caption         =   "Include Subfolders"
         Height          =   240
         Left            =   1050
         TabIndex        =   2
         Top             =   1320
         Width           =   1620
      End
      Begin VB.TextBox txtFolder 
         Height          =   330
         Left            =   1350
         TabIndex        =   1
         Top             =   735
         Width           =   4320
      End
      Begin VB.TextBox txtLook 
         Height          =   330
         Left            =   1365
         TabIndex        =   0
         Text            =   "*.exe"
         Top             =   270
         Width           =   4305
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   4320
         TabIndex        =   19
         Top             =   915
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         Caption         =   "&From Folder  :"
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   765
         Width           =   990
      End
      Begin VB.Label lblInfo 
         Caption         =   "&Look In        :"
         Height          =   315
         Index           =   0
         Left            =   195
         TabIndex        =   11
         Top             =   285
         Width           =   990
      End
   End
   Begin MSComctlLib.ListView lvIcons 
      Height          =   2595
      Left            =   45
      TabIndex        =   28
      Top             =   2490
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4577
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgLarge"
      SmallIcons      =   "imgSmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Icon"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Label lblIcons 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblIcon 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuStart 
         Caption         =   "&Start Search"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSSfolder 
         Caption         =   "&Set Start Folder"
      End
      Begin VB.Menu mnuB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "&Save As"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuBatch 
         Caption         =   "&Batch Save"
         Begin VB.Menu mnuSaveSel 
            Caption         =   "&Save Selected Icons"
            Shortcut        =   ^I
         End
         Begin VB.Menu mnuSaveAll 
            Caption         =   "&Save All Icons In List"
            Shortcut        =   ^A
         End
      End
      Begin VB.Menu mnuB2 
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
      Begin VB.Menu mnuB3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRem 
         Caption         =   "&Invert Selection"
         Shortcut        =   +{INSERT}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuLarge 
         Caption         =   "&Normal View"
      End
      Begin VB.Menu mnuSmall 
         Caption         =   "&List View"
      End
      Begin VB.Menu mnuVB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDetails 
         Caption         =   "&Details View"
      End
      Begin VB.Menu mnuVB2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBitmap 
         Caption         =   "&Enhanced View"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuOP 
      Caption         =   "&Options"
      Begin VB.Menu mnuShowLabel 
         Caption         =   "&Display Labels"
         Checked         =   -1  'True
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuOB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIcLabel 
         Caption         =   "&Icon Labels"
         Begin VB.Menu mnuIL1 
            Caption         =   "&Continuous Naming"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuIL2 
            Caption         =   "&With Respect To File"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuCont 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   ^{F1}
      End
   End
End
Attribute VB_Name = "frmMain"
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
'
'=============================================================================================================================

Dim glLargeIcons() As Long
Dim glSmallIcons() As Long
Dim lIcons         As Long
Dim sExeName       As String
Dim ExitFlag As Boolean


Public BitLoaded As Boolean

Dim SearchFlag As Boolean

Const LARGE_ICON As Integer = 32
Const SMALL_ICON As Integer = 16
Const DI_NORMAL = 3

Private Declare Function DrawIconEx Lib "User32" _
    (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, _
    ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, _
    ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, _
    ByVal diFlags As Long) As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function ExtractIconEx Lib "shell32.dll" _
    Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, _
    phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
    
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociateIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long

Private Sub AddIcons()

Dim FirstPath As String, DirCount As Integer


' ========== Now Adding Files ==========

Dim i
lvIcons.ListItems.Clear

' Initialize labels. Clear the picture boxes.

picSmall.Picture = LoadPicture("")
picLarge.Picture = LoadPicture("")

Set lvIcons.Icons = Nothing
Set lvIcons.SmallIcons = Nothing



'Setting large imagelist
With imgLarge
    .ListImages.Clear
    .ImageHeight = LARGE_ICON
    .ImageWidth = LARGE_ICON
End With


'Small Imagelist
With imgSmall
    .ListImages.Clear
    .ImageHeight = SMALL_ICON
    .ImageWidth = SMALL_ICON
End With

sbMain.SimpleText = "Wait..., Searching " & txtFolder.Text


    FirstPath = dirList.Path
    DirCount = dirList.ListCount
    NumFiles = 0    ' Reset found files indicator.
    lstFoundFiles.Clear
    result = DirDiver(FirstPath, DirCount, "")
    filList.Path = dirList.Path
    
' ========== Now adding Icons ==========

'On Error Resume Next
Set lvIcons.Icons = imgLarge
Set lvIcons.SmallIcons = imgLarge
'Set lvIcons.SmallIcons = imgSmall


For i = 1 To imgLarge.ListImages.Count
sbMain.SimpleText = "Adding " & i & "th icon to list"
Dim ic As ListItem
    Set ic = lvIcons.ListItems.Add(, , "", i, i)
    ic.SubItems(1) = FPathFromKey(imgLarge.ListImages(i).Key)
Next i

ToggleCaption mnuShowLabel.Checked


sbMain.SimpleText = lvIcons.ListItems.Count & " Icons Added. Double click an icon for enhanced view."

End Sub

Private Sub cmdDir_Click()
txtFolder.Text = frmDir.ShowDir()

On Error Resume Next
dirList.Path = txtFolder.Text

If Trim(txtFolder.Text) = "" Then
    txtFolder.Text = dirList.Path
End If


End Sub



Private Sub cmdStart_Click()

On Error Resume Next
If BitLoaded Then Unload frmBit

If cmdStart.Caption = "&Stop Search" Then
    ExitFlag = True
    SearchFlag = False
    cmdStart.Caption = "&Start Search"
    anMain.Stop
    mnuStart.Caption = cmdStart.Caption
    Exit Sub
End If

If cmdStart.Caption = "&Start Search" Then
    ExitFlag = False
    cmdStart.Caption = "&Stop Search"
    mnuStart.Caption = cmdStart.Caption
    lblCount.Caption = "0"
    SearchFlag = True
    anMain.Play
    txtFolder_Change
    AddIcons
    mnuStart.Caption = cmdStart.Caption
    txtFolder_Change
    EnableDisable
    cmdStart.Caption = "&Start Search"
    anMain.Stop
    sbMain.SimpleText = lvIcons.ListItems.Count & " Icons added"
    Exit Sub
End If

End Sub

 Private Sub dirList_Change()
filList.Path = dirList.Path
End Sub

Private Sub drvList_Change()
dirList.Path = drvList.Drive
End Sub

Private Sub Form_Load()

' Set the dimensions of the PictureBox controls where the
' icons will be drawn.  We r using 32x32 and 16x16 icons.
' Each size uses its own PictureBox.

ExitFlag = False

'Hey, I am adding an easter egg here.
If Month(Date) = 7 And Day(Date) = 12 Then
    Dim Str As String
    Str = " Icon Hunter Easter Egg " + vbCrLf
    Str = Str + " ===================="
    Str = Str + vbCrLf + vbCrLf
    Str = Str + "Today is July 12th. July 12th is the birthday of Icon Hunter's author, Mr Anoop M." + vbCrLf + vbCrLf
    Str = Str + "You opened Icon Hunter in this day, and cracked the Easter Egg in it. Do you want to see the Easter Egg dialog box?"
    ret = MsgBox(Str, vbInformation + vbYesNo, "Icon Hunter")
    If ret = vbYes Then frmAuthor.Show vbModal
End If

picLarge.Height = LARGE_ICON * Screen.TwipsPerPixelY
picLarge.Width = LARGE_ICON * Screen.TwipsPerPixelX
picSmall.Height = SMALL_ICON * Screen.TwipsPerPixelY
picSmall.Width = SMALL_ICON * Screen.TwipsPerPixelX
anMain.Open App.Path & "\search.avi"

dirList.Path = drvList.Drive
filList.Path = dirList.Path
txtFolder.Text = dirList.Path
Form_Resize
BitLoaded = False
CodeUnload = False

LoadPrefs
EnableDisable



End Sub



Public Sub pGetIcons(sExeName As String)
Dim l As Long

lIcons = ExtractIconEx(sExeName, -1, 0, 0, 0)

 If lIcons < 0 Then Exit Sub

' Dimension the arrays to the number of icons.
' Get the icons' handles.
'
ReDim glLargeIcons(lIcons)
ReDim glSmallIcons(lIcons)

Dim lIndex

For lIndex = 0 To lIcons - 1

Call ExtractIconEx(sExeName, lIndex, glLargeIcons(lIndex), glSmallIcons(lIndex), 1)
'
' Draw the icon to respective picturebox control.
'
With picLarge
    Set .Picture = LoadPicture("")
     .AutoRedraw = True
    Call DrawIconEx(.hdc, 0, 0, glLargeIcons(lIndex), LARGE_ICON, LARGE_ICON, 0, 0, DI_NORMAL)
     .Refresh
End With

On Error GoTo stopThis


mykey = sExeName & "(" & lIndex & ")"

If Val(txtMax.Text) = imgLarge.ListImages.Count Then
SearchFlag = False
Else
imgLarge.ListImages.Add , mykey, picLarge.Image
End If

With picSmall
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    Call DrawIconEx(.hdc, 0, 0, glSmallIcons(lIndex), SMALL_ICON, SMALL_ICON, 0, 0, DI_NORMAL)
    .Refresh
End With

'mykey = sExeName & "(" & lIndex & ")"
    'imgSmall.ListImages.Add , mykey, picSmall.Image

nextIcon:
Next lIndex


Exit Sub

stopThis:

txtMax.Text = imgLarge.ListImages.Count

End Sub


'===================================================
' Helper Functions
'===================================================

'===================================================
Private Function DirDiver(NewPath As String, DirCount As Integer, BackUp As String) As Integer
'===================================================

'  Recursively search directories from NewPath down...
'  NewPath is searched on this recursion.
'  BackUp is origin of this recursion.
'  DirCount is number of subdirectories in this directory.
Static FirstErr As Integer
Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, entry As String

Dim retval As Integer

    If ExitFlag = True Then
        DirDiver = True
        Exit Function
    End If
    
    
    SearchFlag = True           ' Set flag so the user can interrupt.
    DirDiver = False            ' Set to True if there is an error.
    
    'If imgLarge.ListImages.Count = Val(txtMax.Text) Then
    'SearchFlag = False
    'DirDiver = True
    'Exit Function
    'End If
   
    If SearchFlag = False Then
        DirDiver = True
        Exit Function
    End If
    
    On Local Error GoTo DirDriverHandler
    
    If chkSub.Value = 1 Then
        DirsToPeek = dirList.ListCount                  ' How many directories below this?
    Else
        DirsToPeek = 0
    End If
    
    Do While DirsToPeek > 0 And SearchFlag = True
    
        OldPath = dirList.Path                      ' Save old path for next recursion.
        dirList.Path = NewPath
        If dirList.ListCount > 0 Then
            ' Get to the node bottom.
            dirList.Path = dirList.List(DirsToPeek - 1)
            retval = DoEvents()         ' Check for events (for instance, if the user chooses Cancel).
            AbandonSearch = DirDiver((dirList.Path), DirCount%, OldPath)
        End If
        ' Go up one level in directories.
        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop
    ' Call function to enumerate files.
    If filList.ListCount Then
        If Len(dirList.Path) <= 3 Then             ' Check for 2 bytes/character
            ThePath = dirList.Path                  ' If at root level, leave as is...
        Else
            ThePath = dirList.Path + "\"            ' Otherwise put "\" before the filename.
        End If
        
        For ind = 0 To filList.ListCount - 1        ' Add conforming files in this directory to the list box.
            entry = ThePath + filList.List(ind)
            'lstFoundFiles.AddItem entry
            pGetIcons entry
            sbMain.SimpleText = "Adding from " & entry
            lblCount.Caption = Str(Val(lblCount.Caption) + 1)
        Next ind
    End If
    If BackUp <> "" Then        ' If there is a superior directory, move it.
        dirList.Path = BackUp
    End If
    Exit Function
DirDriverHandler:
    If Err = 7 Then             ' If Out of Memory error occurs, assume the list box just got full.
        DirDiver = True         ' Create Msg and set return value AbandonSearch.
        'MsgBox "You've filled the list box. Abandoning search..."
        Exit Function           ' Note that the exit procedure resets Err to 0.
    Else                        ' Otherwise display error message and quit.
        'MsgBox Error
        End
    End If
    
'===================================================
End Function
'===================================================



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

ret = MsgBox("Are you sure that you want to exit Icon Hunter?", vbQuestion + vbYesNo, "Exit Icon Hunter")
If ret = vbNo Then Cancel = 1

End Sub

Private Sub Form_Resize()
On Error Resume Next
lvIcons.Move 0, lvIcons.Top, Me.ScaleWidth, (Me.ScaleHeight - lvIcons.Top - sbMain.Height)

lvIcons.ColumnHeaders(2).Width = lvIcons.Width - lvIcons.ColumnHeaders(1).Width * -80

fr.Width = Me.ScaleWidth - (2 * fr.Left + picSide.Width)
frIn.Width = fr.Width

picSide.Left = fr.Width + fr.Left + 100
txtLook.Width = fr.Width - txtLook.Left - 100
txtFolder.Width = txtLook.Width

cmdDir.Left = txtLook.Width + txtLook.Left - cmdDir.Width

If cmdDir.Left < chkSub.Left + chkSub.Width Then
    cmdDir.Visible = False
    Else
    cmdDir.Visible = True
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next



SavePrefs

Set lvIcons.Icons = Nothing

lvIcons.ListItems.Clear

imgLarge.ListImages.Clear
imgSmall.ListImages.Clear

End


End Sub

Private Sub lvIcons_DblClick()
If lvIcons.ListItems.Count > 0 Then
frmBit.Show
frmBit.img.Picture = imgLarge.ListImages(lvIcons.SelectedItem.Icon).ExtractIcon
frmBit.SetStatus lvIcons.SelectedItem.Icon
End If
End Sub

Private Sub lvIcons_ItemClick(ByVal Item As MSComctlLib.ListItem)
sbMain.SimpleText = Item.Text & ": This icon is from file " & FPathFromKey(imgLarge.ListImages(Item.Icon).Key)

If BitLoaded Then
    frmBit.i = lvIcons.SelectedItem.Icon
    frmBit.img.Picture = imgLarge.ListImages(Item.Icon).ExtractIcon
End If

End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal

End Sub

Private Sub mnuBitmap_Click()
lvIcons_DblClick

End Sub

Private Sub mnuCont_Click()
X = HyperJump(App.HelpFile)

End Sub

Private Sub mnuCopy_Click()
On Error Resume Next
Clipboard.Clear
picLarge.Picture = imgLarge.ListImages(lvIcons.SelectedItem.Icon).ExtractIcon
Clipboard.SetData picLarge.Image

End Sub

Private Sub mnuDetails_Click()
lvIcons.View = lvwReport
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub


Private Sub mnuIL1_Click()
mnuIL2.Checked = Not mnuIL2.Checked
mnuIL1.Checked = Not mnuIL1.Checked
ToggleCaption mnuShowLabel.Checked
End Sub

Private Sub mnuIL2_Click()
mnuIL1_Click
End Sub

Private Sub mnuLarge_Click()
lvIcons.View = lvwIcon
End Sub

Private Sub mnuRem_Click()
On Error Resume Next
For i = 1 To lvIcons.ListItems.Count
    lvIcons.ListItems(i).Selected = Not lvIcons.ListItems(i).Selected
Next i

End Sub

Private Sub mnuSaveAll_Click()
SaveAllIcons
End Sub

Private Sub mnuSaveAs_Click()
On Error GoTo nosave
cdlOpen.CancelError = True
cdlOpen.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist
cdlOpen.DialogTitle = "Save Icon"
cdlOpen.Filter = "Icon File|*.ico"
cdlOpen.DefaultExt = "ico"
cdlOpen.ShowSave

SavePicture imgLarge.ListImages(lvIcons.ListItems(i).Icon).ExtractIcon, cdlOpen.FileTitle

nosave:


End Sub

Private Sub mnuSaveSel_Click()
SaveSelectedIcons
End Sub


Private Sub mnuShowLabel_Click()
mnuShowLabel.Checked = Not mnuShowLabel.Checked
ToggleCaption mnuShowLabel.Checked
End Sub

Private Sub mnuSmall_Click()
lvIcons.View = lvwList
End Sub

Private Sub mnuSSfolder_Click()
cmdDir_Click
End Sub


Private Sub mnuStart_Click()
cmdStart_Click
End Sub

Private Sub txtFolder_Change()
On Error Resume Next
dirList.Path = txtFolder.Text
End Sub

Private Sub txtFolder_GotFocus()
txtFolder.SelStart = 0
txtFolder.SelLength = Len(txtFolder.Text)
End Sub

Private Sub txtFolder_Validate(Cancel As Boolean)
On Error GoTo Handle
dirList.Path = txtFolder.Text
Exit Sub

Handle:

mstr = "The folder you entered is not valid. Please enter a valid folder, or click the button next to it for selecting a folder"

MsgBox mstr, vbOKOnly + vbInformation, "Invalid Directory"
txtFolder.Text = CurDir
dirList.Path = CurDir
Cancel = True
End Sub

Private Sub txtLook_Change()
On Error Resume Next
filList.Pattern = txtLook.Text
End Sub

Private Sub txtLook_GotFocus()
txtLook.SelStart = 0
txtLook.SelLength = Len(txtLook.Text)
End Sub

Private Sub txtLook_Validate(Cancel As Boolean)
On Error GoTo Handle
filList.Pattern = txtLook.Text
Exit Sub

Handle:
Dim Str As String
mstr = "The file pattern you entered is invalid. Please enter a valid criteria. Use wild cards if necessary"
mstr = mstr + vbCrLf + vbCrLf
mstr = mstr + "Examples:" + vbCrLf
mstr = mstr + "   1) *.exe       - Searches in all EXE files" + vbCrLf
mstr = mstr + "   2) *.exe;*.dll - Searches in all EXE files and DLL files"


MsgBox mstr, vbOKOnly + vbInformation, "Invalid Filetype"

filList.Pattern = "*.exe"

End Sub

Sub SaveAllIcons()
'For saving all icons

    Dim Getpath As String
    Getpath = frmDir.ShowDir("Store To Directory..")
    If Getpath = "" Then Exit Sub
    If Right(Getpath, 1) <> "\" Then Getpath = Getpath & "\"
    Cap = sbMain.SimpleText
    sbMain.SimpleText = "Saving.."
    For i = 1 To lvIcons.ListItems.Count
    On Error Resume Next
        SavePicture imgLarge.ListImages(lvIcons.ListItems(i).Icon).ExtractIcon, Getpath & "\Icon " & i & ".ico"
    Next i
    sbMain.SimpleText = Cap
End Sub

Sub SaveSelectedIcons()
'For saving selected icons

    Dim Getpath As String
    Getpath = frmDir.ShowDir("Store To Directory..")
    If Getpath = "" Then Exit Sub
    If Right(Getpath, 1) <> "\" Then Getpath = Getpath & "\"
    Cap = sbMain.SimpleText
    sbMain.SimpleText = "Saving.."
    For i = 1 To lvIcons.ListItems.Count
    On Error Resume Next
    If lvIcons.ListItems(i).Selected = True Then _
        SavePicture imgLarge.ListImages(lvIcons.ListItems(i).Icon).ExtractIcon, Getpath & "\Icon " & i & ".ico"
    Next i
    sbMain.SimpleText = Cap
End Sub

Function FNameFromPath(FullFile As String) As String
'Obtains filename from path


Dim LastPos
LastPos = -1

For i = 1 To Len(FullFile)
    If Right(VBA.Left(FullFile, i), 1) = "\" Then
        LastPos = i
    End If
Next i
        
If LastPos > 0 Then
        FNameFromPath = Right(FullFile, Len(FullFile) - LastPos)
        Exit Function
End If

End Function

Public Function FPathFromKey(FullFile As String) As String
'Obtains filepath from icon key


Dim LastPos
LastPos = -1

For i = 1 To Len(FullFile)
    If Right(VBA.Left(FullFile, i), 1) = "(" Then
        LastPos = i
    End If
Next i
        
If LastPos > 0 Then
        FPathFromKey = Left(FullFile, LastPos - 1)
        Exit Function
End If

End Function

Function ToggleCaption(TState As Boolean)
'Toggles caption


On Error Resume Next

cucap = sbMain.SimpleText

    If TState = True Then
        For i = 1 To imgLarge.ListImages.Count
            sbMain.SimpleText = "Wait,Setting Captions.."
            If mnuIL1.Checked = True Then
            lvIcons.ListItems(i).Text = "Icon " & i
            Else
            lvIcons.ListItems(i).Text = FNameFromPath(imgLarge.ListImages(i).Key)
            End If
            
         Next i
    Else
        For i = 1 To imgLarge.ListImages.Count
            sbMain.SimpleText = "Wait,Removing Captions.."
            lvIcons.ListItems(i).Text = ""
         Next i
        
    End If
    
sbMain.SimpleText = cucap
End Function


Sub SavePrefs()
SaveSetting App.EXEName, "Options", "View", lvIcons.View
SaveSetting App.EXEName, "Options", "Lookin", txtLook.Text
SaveSetting App.EXEName, "Options", "Folder", txtFolder.Text
SaveSetting App.EXEName, "Options", "Maximum", txtMax.Text
SaveSetting App.EXEName, "Options", "Sub", chkSub.Value

End Sub

Sub LoadPrefs()
On Error Resume Next

lvIcons.View = GetSetting(App.EXEName, "Options", "View", 0)

txtLook.Text = GetSetting(App.EXEName, "Options", "Lookin", "*.exe")
filList.Pattern = txtLook.Text

txtFolder.Text = GetSetting(App.EXEName, "Options", "SearchFolder", WinDir())
dirList.Path = txtFolder.Text

txtMax.Text = GetSetting(App.EXEName, "Options", "Maximum", "300")

chkSub.Value = GetSetting(App.EXEName, "Options", "Sub", 1)

End Sub

Sub EnableDisable()
If lvIcons.ListItems.Count < 1 Then
    mnuSaveAll.Enabled = False
    mnuSaveSel.Enabled = False
    mnuSaveAs.Enabled = False
Else
    mnuSaveAll.Enabled = True
    mnuSaveSel.Enabled = True
    mnuSaveAs.Enabled = True
End If

mnuCopy.Enabled = mnuSaveAs.Enabled
mnuRem.Enabled = mnuCopy.Enabled
mnuBitmap.Enabled = mnuRem.Enabled


End Sub

Private Sub txtMax_Validate(Cancel As Boolean)
If Not IsNumeric(txtMax.Text) Then
MsgBox "The value you entered for maximum icons is invalid. Please enter a valid value", vbInformation + vbOKOnly, "Invalid Entry"
    Cancel = 1
    txtMax.Text = udMain.Value
End If
End Sub

Function WinDir() As String
Dim WinPath As String
    WinPath = String(145, Chr(0))
    WinDir = Left(WinPath, GetWindowsDirectory(WinPath, 145))
End Function

