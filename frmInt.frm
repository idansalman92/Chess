VERSION 5.00
Begin VB.Form frmInt 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Instructions"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmInt.frx":0000
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "frmInt.frx":5037F
      ScaleHeight     =   495
      ScaleWidth      =   1815
      TabIndex        =   1
      Top             =   11040
      Width           =   1815
   End
   Begin VB.Timer TimerSound 
      Interval        =   1
      Left            =   14760
      Top             =   120
   End
   Begin VB.Label MainMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        frmMain.ClickCheck = True
        frmDirections.Show
        Exit Sub
    End If
    If KeyCode = vbKeyEscape Then
        'frmMenu.Show
        'Unload Me
    End If
End Sub
Private Sub MainMenu_Click()
    frmInt.Picture = LoadPicture("Instructions\Instructions12.jpg")
    frmInt.Refresh
    frmInt.Picture = LoadPicture("Instructions\Instructions11.jpg")
    frmInt.Refresh
    frmInt.Picture = LoadPicture("Instructions\Instructions10.jpg")
    frmInt.Refresh
    frmInt.Picture = LoadPicture("Instructions\Instructions9.jpg")
    frmInt.Refresh
    frmInt.Picture = LoadPicture("Instructions\Instructions8.jpg")
    frmInt.Refresh
    frmInt.Picture = LoadPicture("Instructions\Instructions7.jpg")
    frmInt.Refresh
    frmInt.Picture = LoadPicture("Instructions\Instructions6.jpg")
    frmInt.Refresh
    frmInt.Picture = LoadPicture("Instructions\Instructions5.jpg")
    frmInt.Refresh
    frmInt.Picture = LoadPicture("Instructions\Instructions4.jpg")
    frmInt.Refresh
    frmInt.Picture = LoadPicture("Instructions\Instructions3.jpg")
    frmInt.Refresh
    frmInt.Picture = LoadPicture("Instructions\Instructions2.jpg")
    frmInt.Refresh
    frmInt.Picture = LoadPicture("Instructions\Instructions1.jpg")
    frmInt.Refresh
    frmInt.Picture = LoadPicture("Instructions\Instructions1.jpg")
    frmInt.Refresh
    frmInt.Picture = LoadPicture("main/index15.jpg")
    frmInt.Refresh
    frmInt.Picture = LoadPicture("main/index16.jpg")
    frmInt.Refresh
    frmInt.Picture = LoadPicture("main/index17.jpg")
    frmInt.Refresh
    frmInt.Picture = LoadPicture("main/index18.jpg")
    frmInt.Refresh
    frmInt.Picture = LoadPicture("main/index19.jpg")
    frmInt.Refresh
    frmInt.Picture = LoadPicture("main/index20.jpg")
    frmInt.Refresh
    frmMenu.Show
    Unload Me
End Sub
Private Sub Form_Load()
    If Not frmMenu.SoundCheck Then
        Picture1.Picture = LoadPicture("Sound/Mute.jpg")
        Picture1.Refresh
        TimerSound.Enabled = False
    End If
End Sub
Private Sub Picture1_Click()
    If frmMenu.SoundCheck Then
        Picture1.Picture = LoadPicture("Sound/Mute.jpg")
        Picture1.Refresh
        frmMenu.SoundCheck = False
        TimerSound.Enabled = False
    Else
        TimerSound.Enabled = True
        frmMenu.SoundCheck = True
    End If
End Sub

Private Sub TimerSound_Timer()
    Dim t As Integer
    Dim s As Integer
    s = 50
        If Tzlil = 0 Then
            For t = 0 To s
                Picture1.Picture = LoadPicture("Sound\Sound1.jpg")
                Picture1.Refresh
            Next t
            For t = 0 To s
                Picture1.Picture = LoadPicture("Sound\Sound2.jpg")
                Picture1.Refresh
            Next t
            For t = 0 To s
                Picture1.Picture = LoadPicture("Sound\Sound3.jpg")
                Picture1.Refresh
            Next t
            For t = 0 To s
                Picture1.Picture = LoadPicture("Sound\Sound4.jpg")
                Picture1.Refresh
            Next t
        End If
End Sub
