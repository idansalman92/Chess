VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   Picture         =   "frmAbout.frx":0000
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
      Picture         =   "frmAbout.frx":4C0E7
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
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        frmMain.ClickCheck = True
        frmDirections.Show
        Exit Sub
    End If
    If KeyCode = vbKeyEscape Then
        'frmMenu.Show
        'Unload Me
        'Exit Sub
    End If
End Sub
Private Sub Label1_Click()
    frmAbout.Picture = LoadPicture("About\About14.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("About\About13.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("About\About12.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("About\About11.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("About\About10.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("About\About9.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("About\About8.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("About\About7.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("About\About6.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("About\About5.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("About\About4.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("About\About3.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("About\About2.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("About\About1.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("main/index15.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("main/index14.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("main/index13.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("main/index12.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("main/index11.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("main/index10.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("main/index9.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("main/index8.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("main/index7.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("main/index7.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("main/index5.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("main/index4.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("main/index3.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("main/index2.jpg")
    frmAbout.Refresh
    frmAbout.Picture = LoadPicture("main/index1.jpg")
    frmAbout.Refresh
    frmMenu.Show
    Unload Me
End Sub
