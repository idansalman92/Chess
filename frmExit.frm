VERSION 5.00
Begin VB.Form frmExit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   Picture         =   "frmExit.frx":0000
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
      Picture         =   "frmExit.frx":49EF3
      ScaleHeight     =   495
      ScaleWidth      =   1815
      TabIndex        =   4
      Top             =   11040
      Width           =   1815
   End
   Begin VB.Timer TimerSound 
      Interval        =   1
      Left            =   14760
      Top             =   120
   End
   Begin VB.Label Back 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label No 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label1 
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
Attribute VB_Name = "frmExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Back_Click()
    frmExit.Picture = LoadPicture("Exit/Exit19.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit18.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit17.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit16.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit15.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit14.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit13.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit12.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit11.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit10.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit9.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit8.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit7.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit6.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit5.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit4.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit3.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit2.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit1.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index15.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index14.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index13.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index12.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index11.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index10.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index9.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index8.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index7.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index6.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index5.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index4.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index3.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index3.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index3.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index2.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index1.jpg")
    frmExit.Refresh
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
Private Sub Label2_Click()
    If frmMenu.SoundCheck Then
        PlayWaveSound "byebye.wav"
    End If
    frmExit.Picture = LoadPicture("Exit/Exit19.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit18.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit17.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit16.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit15.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit14.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit13.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit12.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit11.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit10.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit9.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit8.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit7.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit6.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit5.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit4.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit3.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit2.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit1.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit1.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit1.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit1.jpg")
    frmExit.Refresh
    End
End Sub
Private Sub No_Click()
    frmExit.Picture = LoadPicture("Exit/Exit19.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit18.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit17.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit16.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit15.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit14.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit13.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit12.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit11.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit10.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit9.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit8.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit7.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit6.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit5.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit4.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit3.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit2.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("Exit/Exit1.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index15.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index14.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index13.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index12.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index11.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index10.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index9.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index8.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index7.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index6.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index5.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index4.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index3.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index3.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index3.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index2.jpg")
    frmExit.Refresh
    frmExit.Picture = LoadPicture("main/index1.jpg")
    frmExit.Refresh
    frmMenu.Show
    Unload Me
End Sub
