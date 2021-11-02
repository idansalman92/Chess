VERSION 5.00
Begin VB.Form frmIntro 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   Picture         =   "frmIntro.frx":0000
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Enter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   5880
      Width           =   2535
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Form_Pic()
    PlayWaveSound "welcome.wav"
    If frmMenu.EnterCheck = 1 Then
        frmMenu.Show
        Me.Hide
    Else
    frmIntro.Picture = LoadPicture("Intro\Intro1.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro2.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro3.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro4.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro5.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro6.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro7.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro8.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro9.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro10.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro11.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro12.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro13.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro14.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro15.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro16.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro17.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro18.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro19.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro20.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro21.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro22.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro23.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro24.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro25.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro26.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro27.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro28.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro29.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro30.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro31.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro32.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro33.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro34.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro35.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro36.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro37.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro38.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro39.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro40.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro41.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro42.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro43.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro44.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro45.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro46.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro47.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro48.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro49.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro50.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro51.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("Intro\Intro52.jpg")
    frmIntro.Refresh
    End If
End Sub
Private Sub Enter_Click()
    PlayWaveSound "row.wav"
    frmIntro.Picture = LoadPicture("main\index15.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("main\index14.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("main\index13.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("main\index12.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("main\index11.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("main\index10.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("main\index9.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("main\index8.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("main\index7.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("main\index6.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("main\index5.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("main\index4.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("main\index3.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("main\index2.jpg")
    frmIntro.Refresh
    frmIntro.Picture = LoadPicture("main\index1.jpg")
    frmIntro.Refresh
    frmMenu.Show
    Unload Me
End Sub
