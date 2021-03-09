VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About HunTeR's CD Player"
   ClientHeight    =   3150
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2174.186
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":030A
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2625
      Width           =   1260
   End
   Begin VB.Label Label4 
      Caption         =   "the_dragon@programmer.net"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "defender@mad.scientist.com"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Copyright (C) HunTeR Softwares ULtd."
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Windows 98"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblTitle 
      Caption         =   "HunTeR's (R) CD Player"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1050
      TabIndex        =   2
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
    Unload Me
End Sub
