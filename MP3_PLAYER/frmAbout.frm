VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin Project1.lvButtons_H lvButtons_Close 
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   4320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&Close"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   8421504
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "frmAbout.frx":0442
      Top             =   2400
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000A&
         Caption         =   " ADMP3  2003 "
         Height          =   1695
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   2655
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "adm@rfm.com.ph"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "admendoza@swiftfoods.com.ph"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "E-mails:"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "by ADEL D. MENDOZA"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Copyright © 2003"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1530
         Left            =   240
         Picture         =   "frmAbout.frx":05F3
         Top             =   360
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Me.Top = frmMP3Player.Top
  Me.Left = frmMP3Player.Left + frmMP3Player.Width
End Sub

Private Sub lvButtons_Close_Click()
   Unload Me
End Sub
