VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmPlayer 
   Caption         =   "Avisoft Media Player"
   ClientHeight    =   4545
   ClientLeft      =   6270
   ClientTop       =   4395
   ClientWidth     =   9180
   Icon            =   "Player.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   9180
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8520
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin WMPLibCtl.WindowsMediaPlayer Player1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   16113
      _cy             =   8070
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu MnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuOpURL 
         Caption         =   "Open URL"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "FrmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CVideoPlayer21_GotFocus()

End Sub

Private Sub CmdOpen_Click()
Dim Openfile As String
CommonDialog1.Flags = cd10filenamemustexist + cd10FNhideReadonly
CommonDialog1.Filter = "All Files (*.*) |*.*|"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowOpen
Openfile = CommonDialog1.FileName
Player1.URL = Openfile

End Sub


Private Sub Form_Terminate()
End
End Sub

Private Sub mnuAbout_Click()
FrmPlayer.Enabled = False
frmAbout.Show
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub MnuOpen_Click()
Dim Openfile As String
CommonDialog1.Flags = cd10filenamemustexist + cd10FNhideReadonly
CommonDialog1.Filter = "All Files (*.*) |*.*|"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowOpen
Openfile = CommonDialog1.FileName
Player1.URL = Openfile
End Sub

Private Sub mnuOpURL_Click()
FrmPlayer.Enabled = False
frmURL.Show
End Sub

Private Sub MnuSav_Click()

End Sub
