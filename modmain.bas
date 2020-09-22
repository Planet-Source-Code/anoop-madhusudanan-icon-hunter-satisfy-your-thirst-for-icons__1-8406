Attribute VB_Name = "Module1"
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


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Sub Main()
On Error Resume Next
App.HelpFile = App.Path & "\icon hunter help.chm"

frmSplash.Show
frmSplash.Refresh
Load frmMain
frmSplash.Refresh
frmMain.Show
frmSplash.Refresh
Unload frmSplash

End Sub

Public Function HyperJump(ByVal URL As String) As Long
On Error Resume Next
   HyperJump = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function

