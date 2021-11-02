VERSION 5.00
Begin VB.Form frmMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Game Menu"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   Icon            =   "frmMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmMenu.frx":1272
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerSound 
      Interval        =   1
      Left            =   14880
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "frmMenu.frx":4C5F7
      ScaleHeight     =   495
      ScaleWidth      =   1815
      TabIndex        =   8
      Top             =   11040
      Width           =   1815
   End
   Begin VB.Label AboutCopy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   13560
      TabIndex        =   7
      Top             =   11040
      Width           =   855
   End
   Begin VB.Label ExitCopy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   14520
      TabIndex        =   6
      Top             =   11040
      Width           =   855
   End
   Begin VB.Label InstructionsCopy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   11880
      TabIndex        =   5
      Top             =   11040
      Width           =   1455
   End
   Begin VB.Label NewGameCopy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   10800
      TabIndex        =   4
      Top             =   11040
      Width           =   855
   End
   Begin VB.Label Exit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7680
      TabIndex        =   3
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label About 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Instructions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6720
      TabIndex        =   1
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label NewGame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6720
      TabIndex        =   0
      Top             =   4440
      Width           =   1455
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tzlil As Integer
Public SoundCheck As Boolean
Public check As Integer
Public Enter As Boolean
Public EnterCheck As Integer
Private Sub About_Click()
    frmMenu.Picture = LoadPicture("main/index1.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index2.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index4.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index5.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index6.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index7.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index8.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index9.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index10.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index11.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index12.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index13.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index14.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index15.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About1.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About2.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About4.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About5.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About6.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About7.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About8.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About9.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About10.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About11.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About12.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About13.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About14.jpg")
    frmMenu.Refresh
    frmAbout.Show
    Unload Me
End Sub
Private Sub AboutCopy_Click()
    frmMenu.Picture = LoadPicture("main/index1.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index2.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index4.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index5.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index6.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index7.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index8.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index9.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index10.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index11.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index12.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index13.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index14.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index15.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About1.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About2.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About4.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About5.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About6.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About7.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About8.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About9.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About10.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About11.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About12.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About13.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("About\About14.jpg")
    frmMenu.Refresh
    frmAbout.Show
    Unload Me
End Sub
Private Sub Exit_Click()
    frmMenu.Picture = LoadPicture("main/index1.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index2.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index4.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index5.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index6.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index7.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index8.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index9.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index10.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index11.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index12.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index13.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index14.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index15.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit1.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit2.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit4.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit5.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit6.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit7.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit8.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit9.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit10.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit11.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit12.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit13.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit14.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit15.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit16.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit17.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit18.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit19.jpg")
    frmMenu.Refresh
    frmExit.Show
    Unload Me
End Sub
Private Sub ExitCopy_Click()
    frmMenu.Picture = LoadPicture("main/index1.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index2.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index4.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index5.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index6.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index7.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index8.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index9.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index10.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index11.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index12.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index13.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index14.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index15.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit1.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit2.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit4.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit5.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit6.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit7.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit8.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit9.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit10.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit11.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit12.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit13.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit14.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit15.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit16.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit17.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit18.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Exit/Exit19.jpg")
    frmMenu.Refresh
    frmExit.Show
    Unload Me
End Sub
Private Sub Form_Load()
    If EnterCheck <> 1 Then
        frmIntro.Show
        frmIntro.Form_Pic
        EnterCheck = 1
        Me.Hide
    End If
    Tzlil = 0
    frmMain.ClickCheck = False
    SoundCheck = Not SoundCheck
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        frmMain.ClickCheck = True
        frmDirections.Show
        Exit Sub
    End If
    If KeyCode = vbKeyEscape Then
        frmMenu.Picture = LoadPicture("main/index1.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("main/index2.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("main/index3.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("main/index3.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("main/index3.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("main/index4.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("main/index5.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("main/index6.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("main/index7.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("main/index8.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("main/index9.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("main/index10.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("main/index11.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("main/index12.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("main/index13.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("main/index14.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("main/index15.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("Exit/Exit1.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("Exit/Exit2.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("Exit/Exit3.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("Exit/Exit4.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("Exit/Exit5.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("Exit/Exit6.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("Exit/Exit7.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("Exit/Exit8.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("Exit/Exit9.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("Exit/Exit10.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("Exit/Exit11.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("Exit/Exit12.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("Exit/Exit13.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("Exit/Exit14.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("Exit/Exit15.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("Exit/Exit16.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("Exit/Exit17.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("Exit/Exit18.jpg")
        frmMenu.Refresh
        frmMenu.Picture = LoadPicture("Exit/Exit19.jpg")
        frmMenu.Refresh
        frmExit.Show
        Unload Me
        Exit Sub
    End If
End Sub
Private Sub Instructions_Click()
    frmMenu.Picture = LoadPicture("main/index1.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index2.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index4.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index5.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index6.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index7.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index8.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index9.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index10.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index11.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index12.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index13.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index14.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index15.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions1.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions2.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions4.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions5.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions6.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions7.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions8.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions9.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions10.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions11.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions12.jpg")
    frmMenu.Refresh
    frmInt.Show
    Unload Me
End Sub
Private Sub InstructionsCopy_Click()
    frmMenu.Picture = LoadPicture("main/index1.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index2.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index4.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index5.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index6.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index7.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index8.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index9.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index10.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index11.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index12.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index13.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index14.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index15.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions1.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions2.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions4.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions5.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions6.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions7.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions8.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions9.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions10.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions11.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("Instructions\Instructions12.jpg")
    frmMenu.Refresh
    frmInt.Show
    Unload Me
End Sub
Private Sub NewGame_Click()
    frmMenu.Picture = LoadPicture("main/index1.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index2.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index4.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index5.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index6.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index7.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index8.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index9.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index10.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index11.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index12.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index13.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index14.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index15.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game1.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game2.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game4.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game5.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game6.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game7.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game8.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game9.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game10.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game11.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game12.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game13.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game14.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game15.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game16.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game17.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game18.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game19.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game20.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game21.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game22.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game23.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game24.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game25.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game26.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game27.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game28.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game29.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game30.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game31.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game32.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game33.jpg")
    frmMenu.Refresh
    frmMain.Show
    Unload Me
End Sub
Private Sub NewGameCopy_Click()
    frmMenu.Picture = LoadPicture("main/index1.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index2.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index4.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index5.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index6.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index7.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index8.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index9.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index10.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index11.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index12.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index13.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index14.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("main/index15.jpg")
    frmMenu.Refresh
    
    frmMenu.Picture = LoadPicture("game/game1.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game2.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game3.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game4.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game5.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game6.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game7.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game8.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game9.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game10.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game11.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game12.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game13.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game14.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game15.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game16.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game17.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game18.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game19.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game20.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game21.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game22.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game23.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game24.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game25.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game26.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game27.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game28.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game29.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game30.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game31.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game32.jpg")
    frmMenu.Refresh
    frmMenu.Picture = LoadPicture("game/game33.jpg")
    frmMenu.Refresh
    frmMain.Show
    Unload Me
End Sub
Private Sub Picture1_Click()
    If SoundCheck Then
        Picture1.Picture = LoadPicture("Sound\Mute.jpg")
        Picture1.Refresh
        TimerSound.Enabled = False
        SoundCheck = False
    Else
        TimerSound.Enabled = True
        SoundCheck = True
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
