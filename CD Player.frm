VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CD Player"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3525
   Icon            =   "CD Player.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   3525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton cmdEject 
      Caption         =   "Eject"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "Previous"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Current track:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblCurrent 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Dim CurrentTrack As Integer
Dim NumTracks As Integer
Dim i As Integer
Dim DoorOpen As Boolean

Private Sub cmdEject_Click()
    If DoorOpen = False Then
        mciSendString "set cd door open", 0, 0, 0
        DoorOpen = True
    Else
        If DoorOpen = True Then
            mciSendString "set cd door closed", 0, 0, 0
            DoorOpen = False
        End If
    End If
End Sub
Private Sub cmdNext_Click()
    If CurrentTrack < NumTracks Then
        CurrentTrack = CurrentTrack + 1
        lblCurrent.Caption = CurrentTrack & "/" & NumTracks
        mciSendString "stop cd wait", 0, 0, 0
        mciSendString "seek cd to " & CurrentTrack, 0, 0, 0
        mciSendString "play cd", 0, 0, 0
    End If
End Sub
Private Sub cmdPlay_Click()
        mciSendString "stop cd wait", 0, 0, 0
        mciSendString "seek cd to " & CurrentTrack, 0, 0, 0
        mciSendString "play cd", 0, 0, 0
End Sub
Private Sub cmdPrev_Click()
    If CurrentTrack > 1 Then
        CurrentTrack = CurrentTrack - 1
        lblCurrent.Caption = CurrentTrack & "/" & NumTracks
        mciSendString "stop cd wait", 0, 0, 0
        mciSendString "seek cd to " & CurrentTrack, 0, 0, 0
        mciSendString "play cd", 0, 0, 0
    End If
End Sub
Private Sub cmdStop_Click()
    mciSendString "stop cd wait", 0, 0, 0
End Sub
Private Sub Form_Load()
    mciSendString "close all", 0, 0, 0
    mciSendString "open cdaudio alias cd wait shareable", 0, 0, 0
    mciSendString "set cd time format tmsf wait", 0, 0, 0
    NumberOfTracks
    CurrentTrack = 1
    DoorOpen = False
End Sub
Private Sub NumberOfTracks()
    On Error GoTo Kraj
    Dim Tracks As String * 30
    
    mciSendString "status cd number of tracks wait", Tracks, Len(Tracks), 0
    NumTracks = CInt(Mid$(Tracks, 1, 2))
    lblCurrent.Caption = "1/" & NumTracks
    GoTo Kraj1
Kraj: MsgBox "No CD inserted", vbOKOnly + vbCritical, "CD Player"
Kraj1:
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mciSendString "stop cd wait", 0, 0, 0
    mciSendString "close all", 0, 0, 0
End Sub
