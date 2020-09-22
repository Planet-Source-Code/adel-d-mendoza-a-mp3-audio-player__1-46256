VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMP3Player 
   Appearance      =   0  'Flat
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADMP3 2003"
   ClientHeight    =   6885
   ClientLeft      =   3000
   ClientTop       =   2340
   ClientWidth     =   3975
   Icon            =   "frmMP3Player.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6885
   ScaleWidth      =   3975
   Begin Project1.lvButtons_H lvButtons_About 
      Height          =   495
      Left            =   3360
      TabIndex        =   28
      ToolTipText     =   "About "
      Top             =   1440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "ABT"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial monospaced for SAP"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   -2147483634
      cFHover         =   255
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   24
      cBack           =   0
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   960
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      TickStyle       =   3
      TickFrequency   =   5
   End
   Begin Project1.lvButtons_H lvButtons_Close 
      Height          =   495
      Left            =   2760
      TabIndex        =   17
      ToolTipText     =   "Close "
      Top             =   1440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "CLS"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial monospaced for SAP"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   -2147483634
      cFHover         =   255
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      ImgSize         =   32
      cBack           =   0
   End
   Begin Project1.lvButtons_H lvButtons_Pause 
      Height          =   495
      Left            =   720
      TabIndex        =   16
      ToolTipText     =   "Pause"
      Top             =   1440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMP3Player.frx":0442
      ImgSize         =   32
      cBack           =   0
   End
   Begin Project1.lvButtons_H lvButtons_Stop 
      Height          =   495
      Left            =   1320
      TabIndex        =   15
      ToolTipText     =   "Stop"
      Top             =   1440
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMP3Player.frx":0A40
      ImgSize         =   32
      cBack           =   0
   End
   Begin Project1.lvButtons_H lvButtons_Play 
      Height          =   495
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Play"
      Top             =   1440
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMP3Player.frx":103E
      ImgSize         =   24
      cBack           =   0
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      Caption         =   "VOLUME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   1560
      TabIndex        =   12
      Top             =   120
      Width           =   2295
      Begin MSComctlLib.Slider Slider2 
         Height          =   255
         Left            =   1560
         TabIndex        =   23
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         Min             =   -5000
         Max             =   5000
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Max             =   6000
         TickStyle       =   3
         TickFrequency   =   5
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00404000&
         BeginProperty Font 
            Name            =   "Moderne"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00404000&
         BeginProperty Font 
            Name            =   "Moderne"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   1080
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   4680
      Width           =   1455
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   1470
      Left            =   1800
      Pattern         =   "*.mp3;*.cda;*.mid;*.wav"
      TabIndex        =   7
      Top             =   4680
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   1140
      Left            =   240
      TabIndex        =   6
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   2160
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   1515
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      Caption         =   "FILE  LIST"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   3735
      Begin Project1.lvButtons_H lvButtons_Command3 
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         Caption         =   "ADD DIR"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   255
         LockHover       =   2
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   0
      End
      Begin Project1.lvButtons_H lvButtons_Command2 
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         Caption         =   "ADD FILE"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   -2147483634
         cFHover         =   255
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   2
         ImgSize         =   32
         cBack           =   0
      End
      Begin Project1.lvButtons_H lvButtons_Command1 
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         Caption         =   "PLAY FILE"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   -2147483634
         cFHover         =   255
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   2
         ImgSize         =   32
         cBack           =   0
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   3735
      Begin Project1.lvButtons_H lvButtons_Remove 
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         ToolTipText     =   "Remove From Play List"
         Top             =   2160
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         Caption         =   "REM"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12640511
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16761024
      End
      Begin Project1.lvButtons_H lvButtons_Empty 
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         ToolTipText     =   "Empty Play List"
         Top             =   2160
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         Caption         =   "EMP"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12640511
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16761024
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404000&
         Caption         =   "SIN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Single Play"
         Top             =   2160
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton opCont 
         BackColor       =   &H00404000&
         Caption         =   "CON"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         ToolTipText     =   "Continous Play"
         Top             =   2160
         Width           =   615
      End
      Begin VB.OptionButton opRepeat 
         BackColor       =   &H00404000&
         Caption         =   "REP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   720
         TabIndex        =   3
         ToolTipText     =   "Repeat Play"
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404000&
         Caption         =   "SELECTED:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      Caption         =   "NOW PLAYING"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1455
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Moderne"
            Size            =   20.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   375
      Left            =   3480
      TabIndex        =   27
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -200
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "frmMP3Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################
'#          Coded by Adel D. Mendoza          #
'#        Designed by Adel D. Mendoza         #
'#                MP3 Player                  #
'#                                            #
'#        area    :  frmMP3Player             #
'#    description :  Code File Mp3 Player     #
'#        E-mail  :  adm@rfm.com.ph           #
'#        url     :  http://www.rfm.com.ph    #
'#                                            #
'#         Special Thanks to LaVolpe          #
'#              for the Buttons               #
'##############################################

Dim allow_play As String
Dim paused As Boolean
Dim allow_pause As Boolean

Private Sub lvButtons_About_Click()
   frmAbout.Show
End Sub

Private Sub lvButtons_Close_Click()
   MediaPlayer1.Visible = False
   Unload Me
End Sub

Private Sub lvButtons_Remove_Click()
   Dim sTemp As String
   If List1.ListCount <> 0 Then
      If Text1.Text <> "" Then
         List1.RemoveItem (List1.ListIndex)
         Text1.Text = ""
         If FileExists("c:PlayList.txt") Then
            Kill ("c:/PlayList.txt")
         End If
         Open "c:/PlayList.txt" For Output As #1
         Close #1
         Open "c:/PlayList.txt" For Append As #1
         For I = 0 To List1.ListCount
             sTemp = List1.List(I)
             If sTemp <> "" Then
                Print #1, List1.List(I)
             End If
         Next
         Close #1
      End If
   End If
End Sub

Private Sub lvButtons_Empty_Click()
   List1.Clear
   Text1.Text = ""
   If FileExists("c:/PlayList.txt") Then
      Kill ("c:/PlayList.txt")
   End If
End Sub

Private Sub lvButtons_Pause_Click()
   If allow_pause = True Then
      On Error Resume Next
      If paused = False Then
         MediaPlayer1.Pause
         paused = True
         allow_play = "no"
         Exit Sub
      End If
   End If
End Sub

Private Sub lvButtons_Play_Click()
   If Text1.Text <> "" Then
      If paused = True Then
         MediaPlayer1.Play
         Slider3.Max = MediaPlayer1.Duration
         paused = False
         allow_play = "yes"
         Exit Sub
      End If
      If paused = False Then
         MediaPlayer1.FileName = Text1.Text
         MediaPlayer1.Play
         Slider3.Max = MediaPlayer1.Duration
         Exit Sub
      End If
   End If
End Sub

Private Sub lvButtons_Stop_Click()
   MediaPlayer1.Stop
   allow_play = "no"
   allow_pause = False
   Call Reset_Timer
   Me.Width = 4080
End Sub

Private Sub lvButtons_Command1_Click()
   If File1.FileName <> "" Then
      MediaPlayer1.FileName = File1.Path & "\" & File1.FileName
      MediaPlayer1.Play
      Slider3.Max = MediaPlayer1.Duration
   End If
End Sub

Private Sub lvButtons_Command2_Click()
   If File1.FileName <> "" Then
      Call Add_To_PlayList
   End If
End Sub

Private Sub lvButtons_Command3_Click()
   File1.Path = Dir1.Path
   If File1.ListCount <> 0 Then
      For tel = 1 To File1.ListCount
          File1.ListIndex = tel - 1
          If Len(Dir1.Path) > 3 Then
             Call Add_To_PlayList
          Else
             Call Add_To_PlayList
          End If
      Next tel
      Exit Sub
   End If
End Sub

Private Sub Dir1_Change()
   On Error Resume Next
   File1.Path = Dir1.Path
   Exit Sub
End Sub

Private Sub Drive1_Change()
   On Error Resume Next
   Dir1.Path = Drive1.Drive
   Exit Sub
End Sub

Private Sub Form_Load()
   Me.Top = 0
   Me.Left = 0
   Me.Width = 4080
   allow_pause = False
   Dir1.Path = Drive1.Drive
   paused = False
   On Error Resume Next
   If Not FileExists("c:/PlayList.txt") Then
      Open "c:/PlayList.txt" For Output As #1
      Close #1
   End If
   Open "c:/PlayList.txt" For Input As #1
   Do Until EOF(1)
      Input #1, playlistitem
      List1.AddItem UCase(playlistitem)
   Loop
   Close #1
   allow_play = "no"
   '--------------------------------------
   'set volume setting for the mediaplayer
   '--------------------------------------
   MediaPlayer1.Volume = -3000
   Slider1.Value = 3000
   Label2.Caption = "50 %"
   '--------------------------------------
   'check the headset setting
   '--------------------------------------
   Call Slider2_Scroll
   '--------------------------------------
   'reset timer
   '--------------------------------------
   Call Reset_Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Width = 4080
  End
End Sub

Private Sub List1_Click()
   Text1.Text = List1.Text
End Sub

Private Sub List1_DblClick()
   Text1.Text = List1.Text
   MediaPlayer1.Stop
   MediaPlayer1.FileName = Text1.Text
   MediaPlayer1.Play
   Slider3.Max = MediaPlayer1.Duration
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
   Text1.Text = List1.Text
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)
   allow_play = "no"
   allow_pause = False
   If opCont.Value = True Then
      On Error GoTo error1
      allow_pause = True
      List1.ListIndex = List1.ListIndex + 1
      MediaPlayer1.Open Text1.Text
      Exit Sub
   End If
   
error1:
If Text1.Text <> "" Then
   If List1.ListCount <> 0 Then
      List1.ListIndex = 0
   End If
   MediaPlayer1.Open Text1.Text
Else
   MediaPlayer1.Stop
End If

If opRepeat.Value = True Then
   If Text1.Text <> "" Then
      allow_pause = True
      MediaPlayer1.Open Text1.Text
   Else
      MediaPlayer1.Stop
   End If
End If
End Sub

Private Sub MediaPlayer1_NewStream()
   allow_play = "yes"
   allow_pause = True
End Sub

Private Sub Slider1_Scroll()
   Dim sha
   Dim per As Integer
   sha = Slider1.Value - 6000
   MediaPlayer1.Volume = sha
   On Error GoTo hell
   per = Slider1.Value
   Label2.Caption = per \ 60 & " %"
hell:
Exit Sub
End Sub

Private Sub Slider2_Scroll()
   On Error GoTo DamnYou
   If Slider2.Value > -500 And Slider2.Value < 500 Then
      Label3.Caption = "C"
   End If
   If Slider2.Value < -500 Then
      Label3.Caption = "L"
   End If
   If Slider2.Value > 500 Then
      Label3.Caption = "R"
   End If
   MediaPlayer1.Balance = Slider2.Value
   Exit Sub
   
DamnYou:
MsgBox "Err"
End Sub

Private Sub Slider3_Scroll()
   MediaPlayer1.CurrentPosition = Slider3.Value
End Sub

Private Sub Timer1_Timer()
   If allow_play = "yes" Then
      Slider3.Value = MediaPlayer1.CurrentPosition
      tinseconden = MediaPlayer1.CurrentPosition
      Dim min As Integer
      Dim sec As Integer
      min = tinseconden \ 60
      sec = tinseconden - (min * 60)
      If sec = "-1" Then
         sec = "0"
      End If
      lblTime.Caption = Format(min, "0#") & ":" & Format(sec, "0#")
      Exit Sub
   End If
End Sub

Private Sub Reset_Timer()
   min = 0
   sec = 0
   lblTime.Caption = Format(min, "0#") & ":" & Format(sec, "0#")
   Slider3.Value = 0
End Sub

Private Sub Add_To_PlayList()
   oldsongs = ""
   List1.AddItem UCase(File1.Path & "\" & File1.FileName)
   newsong = UCase(File1.Path & "\" & File1.FileName)
   On Error Resume Next
   Open "c:/PlayList.txt" For Append As #1
   Print #1, "" & newsong & ""
   Close #1
   Exit Sub
End Sub

Private Function FileExists(FullFileName As String) As Boolean
   On Error GoTo MakeF
   'If file does Not exist, there will be an Error
   Open FullFileName For Input As #1
   Close #1
   'no error, file exists
   FileExists = True
   Exit Function
   
MakeF:
   'error, file does Not exist
   FileExists = False
   Exit Function
End Function



