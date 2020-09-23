VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "FLASH.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Huggs' Shockwave Flash Movie Player 1.0 BETA "
   ClientHeight    =   7215
   ClientLeft      =   150
   ClientTop       =   510
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin Project1.UserControlButton btnRewind 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Rewind"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16315377
      FCOL            =   0
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   0
   End
   Begin VB.CheckBox chkLoop 
      Caption         =   "loop"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   5880
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flash 
      Height          =   5595
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7005
      _cx             =   12356
      _cy             =   9869
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin Project1.UserControlButton btnBack 
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Back"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16315377
      FCOL            =   0
   End
   Begin Project1.UserControlButton btnPlay 
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Play"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16315377
      FCOL            =   0
   End
   Begin Project1.UserControlButton btnStop 
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Stop"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16315377
      FCOL            =   0
   End
   Begin Project1.UserControlButton btnForward 
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Forward"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16315377
      FCOL            =   0
   End
   Begin Project1.UserControlButton btnDown 
      Height          =   375
      Left            =   6120
      TabIndex        =   14
      Top             =   6360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "High Quality"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16315377
      FCOL            =   0
   End
   Begin Project1.UserControlButton btnUp 
      Height          =   375
      Left            =   6120
      TabIndex        =   15
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Low Quality"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   16315377
      FCOL            =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Please Vote!   E-mail: huggsnkisses_2003@yahoo.com   Made in the Philippines"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   6840
      Width           =   7215
   End
   Begin VB.Label Label4 
      Caption         =   "Frame :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label lblCurrFrame 
      Caption         =   "00000000"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   7
      ToolTipText     =   "Current frame"
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   6480
      Width           =   135
   End
   Begin VB.Label lblTotalFrame 
      Caption         =   "00000000"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      ToolTipText     =   "Total frames"
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label lblLoadedLbl 
      Caption         =   "Progress:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label lblLoaded 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   2
      Top             =   6480
      Width           =   135
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu open 
         Caption         =   "Open"
      End
      Begin VB.Menu bar 
         Caption         =   "-"
      End
      Begin VB.Menu quit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu about 
      Caption         =   "Help"
      Begin VB.Menu aboutme 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1


Private Sub aboutme_Click()
frmAbout.Show vbModal
End Sub


Private Sub btnDown_Click()
Dim Index As Integer
flash.Quality = 1
    btnDown.Enabled = False
    btnUp.Enabled = True
End Sub

Private Sub btnUp_Click()
flash.Quality = 0
    btnUp.Enabled = False
    btnDown.Enabled = True
End Sub

Private Sub Form_Load()
If Command <> "" Then
flash.Movie = Command
Me.Caption = Command & " - Huggs' Flash Movie Player"

End If
Timer1.Interval = 250
    Timer1.Enabled = True
flash.Quality = 1
End Sub




Private Sub open_Click()
CD.Filter = "Shockwave Flash Files (*.swf)|*.swf"
CD.ShowOpen
If CD.FileName <> "" Then
flash.Movie = CD.FileName
Me.Caption = CD.FileTitle & " - Huggs' Flash Movie Player"
lblTotalFrame.Caption = flash.TotalFrames
End If
End Sub

Private Sub quit_Click()
Unload Me
End Sub
Private Sub Timer1_Timer()
    If flash.CurrentFrame <> -1 Then
        lblCurrFrame.Caption = flash.CurrentFrame
    End If
    
    lblLoaded = flash.PercentLoaded
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub
Private Sub chkLoop_Click()
    If chkLoop = 1 Then
        flash.SetVariable "loop", True
        Else
        flash.SetVariable "loop", False
    End If
End Sub
Private Sub btnRewind_Click()
    flash.Rewind
End Sub
Private Sub btnBack_Click()
    flash.Back
End Sub
Private Sub btnPlay_Click()
    flash.Play
End Sub
Private Sub btnStop_Click()
    flash.Stop
End Sub
Private Sub btnForward_Click()
    flash.Forward
End Sub

