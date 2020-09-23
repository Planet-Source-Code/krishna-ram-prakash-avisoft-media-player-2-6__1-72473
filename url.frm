VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form frmURL 
   Caption         =   "Open URL"
   ClientHeight    =   1560
   ClientLeft      =   7935
   ClientTop       =   5925
   ClientWidth     =   6060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   6060
   Begin XPControls.XPText XPText1 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   360
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdOK 
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdCancel 
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "URL:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frmURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
FrmPlayer.Enabled = True
Unload Me
End Sub

Private Sub cmdOK_Click()
FrmPlayer.Enabled = True
FrmPlayer.Player1.URL = XPText1.Text
Unload Me
End Sub
