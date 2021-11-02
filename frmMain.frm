VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Chess"
   MouseIcon       =   "frmMain.frx":0000
   Picture         =   "frmMain.frx":1272
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
      Picture         =   "frmMain.frx":4AFA2
      ScaleHeight     =   495
      ScaleWidth      =   1815
      TabIndex        =   5
      Top             =   11040
      Width           =   1815
   End
   Begin VB.PictureBox Picscene 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   9015
      Left            =   0
      Picture         =   "frmMain.frx":51087
      ScaleHeight     =   601
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   681
      TabIndex        =   0
      Top             =   960
      Width           =   10215
   End
   Begin VB.Label Restart 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   11760
      TabIndex        =   4
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Back 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Quit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   12000
      TabIndex        =   2
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label move 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   11760
      TabIndex        =   1
      Top             =   4800
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private currM As m3Matrix
Private Xstart As Double
Private Ystart As Double
Private solids As m3SolidCollection
Private pickFrom As Boolean
Private pickTo As Boolean
Public ClickCheck As Boolean
Private CheckT As Integer
Public Sub SceneInit()
    Dim w As Integer
    Dim h As Integer
    Dim cent As m3Point
    Dim xorg As Integer
    Dim yorg As Integer
    Dim i As Integer
    Dim size As Double
    Dim p As m3Point
    w = Picscene.ScaleWidth
    h = Picscene.ScaleHeight
    xorg = w / 2
    yorg = h / 2
    Picscene.ScaleLeft = -xorg
    Picscene.ScaleHeight = -h
    Picscene.ScaleTop = yorg
    solids = GameMod.SoldierInit(50)
    'solids = GameMod.TzariahInit(50)
    'solids = GameMod.HorseInit(50)
    'solids = GameMod.RatzInit(50)
    'solids = GameMod.MalkaInit(50)
    'solids = GameMod.KingInit(50)
    MapGame.MapGameInit 50
    m3SolidCollectionApply solids, m3MatrixTranslate(-cent.x, -cent.y, -cent.z)
    m3SetDistance 700
    m3SetLightVector m3VectorInit(m3PointInit(0, 0, 0), m3PointInit(2, 5, 10))
End Sub
Public Sub SceneDraw()
     Picscene.Cls
     'm3SolidCollectionDraw Picscene, solids
     GameDraw Picscene
     'm3DrawLine Picscene, m3PointInit(-600, -600, 0), solids.solids(1).verts(18)
     Picscene.Refresh
     'txtSolid.Text = m3solidToString(sol)
End Sub
Private Sub SceneApply()
    m3SolidCollectionApply solids, currM
    MapGame.MapApply currM
End Sub
Private Sub cmdRotate_Click()
    'currM = m3LineRotate(pStart, m3VectorInit(pStart, pEnd), PI / 10)
    'SceneApply
    'SceneDraw
End Sub
Private Sub Back_Click()
    Picscene.Visible = False
    frmMain.Picture = LoadPicture("game/game33.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game32.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game31.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game30.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game29.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game28.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game27.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game26.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game25.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game24.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game23.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game22.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game21.jpg")
    frmMain.Refresh
    frmMenu.Picture = LoadPicture("game/game20.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game19.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game18.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game17.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game16.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game15.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game14.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game13.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game12.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game11.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game10.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game9.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game8.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game7.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game6.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game5.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game4.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game3.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game2.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game1.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index15.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index14.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index13.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index12.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index11.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index10.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index9.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index8.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index7.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index6.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index5.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index4.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index3.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index3.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index3.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index2.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index1.jpg")
    frmMain.Refresh
    frmMenu.Picture = LoadPicture("index.jpg")
    frmMenu.Show
    Me.Hide
End Sub
Private Sub Form_Load()
    ClickCheck = False
    Me.Show
    frmMenu.Show
    Me.Hide
    If frmMenu.SoundCheck Then
        CheckT = 1
    Else
        Picture1.Picture = LoadPicture("Sound/Mute.jpg")
        CheckT = 0
    End If
    SceneInit
    SceneDraw
End Sub
Private Sub move_Click()
    If MoveFromTo(Picscene) Then
        pickFrom = False
        pickTo = False
    End If
End Sub
Private Sub Picscene_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        ClickCheck = True
        frmDirections.Show
        Exit Sub
    End If
    If KeyCode = vbKeyEscape Then
        'Me.Hide
        'frmMenu.Show
        'Exit Sub
    End If
End Sub
Private Sub picscene_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim c As CellLocation
    Xstart = x
    Ystart = y
    If Button = vbLeftButton Then
        If Not pickFrom Then
           pickFrom = SetFrom(x, y)
        Else
            pickTo = SetTo(x, y)
        End If
         
        c = GetCellLocation(x, y)
        'Text1.Text = c.row & " " & c.col
    End If
End Sub
Private Sub picscene_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim dx As Double
    Dim dy As Double
    Dim c As m3Point
    Dim d As m3Point
    If Button <> vbRightButton Then
        Exit Sub
    End If
    Exit Sub
    dx = x - Xstart
    dy = y - Ystart
    
    Xstart = x
    Ystart = y
    'd = m3SolidCollectionCenter(solids)
    currM = m3MatrixTranslate(-d.x, -d.y, -d.z)
    currM = m3MatrixMultiply(currM, m3YRotate(dx / 100))
    currM = m3MatrixMultiply(currM, m3XRotate(-dy / 100))
    currM = m3MatrixMultiply(currM, m3MatrixTranslate(d.x, d.y, d.z))
    SceneApply
    SceneDraw
End Sub
Private Sub Command1_Click()
    If MoveFromTo(Picscene) Then
        pickFrom = False
        pickTo = False
    End If
End Sub
Private Sub Picture1_Click()
    If CheckT = 1 Then
        Picture1.Picture = LoadPicture("Sound/Mute.jpg")
        CheckT = 0
        frmMenu.SoundCheck = False
    Else
        Picture1.Picture = LoadPicture("Sound/Sound1.jpg")
        CheckT = 1
        frmMenu.SoundCheck = True
    End If
End Sub
Private Sub Quit_Click()
    Picscene.Visible = False
    frmMain.Picture = LoadPicture("game/game33.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game32.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game31.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game30.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game29.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game28.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game27.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game26.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game25.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game24.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game23.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game22.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game21.jpg")
    frmMain.Refresh
    frmMenu.Picture = LoadPicture("game/game20.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game19.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game18.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game17.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game16.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game15.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game14.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game13.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game12.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game11.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game10.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game9.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game8.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game7.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game6.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game5.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game4.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game3.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game2.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("game/game1.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index15.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index14.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index13.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index12.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index11.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index10.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index9.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index8.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index7.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index6.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index5.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index4.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index3.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index3.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index3.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index2.jpg")
    frmMain.Refresh
    frmMain.Picture = LoadPicture("main/index1.jpg")
    frmMain.Refresh
    frmMenu.Picture = LoadPicture("index.jpg")
    frmMenu.Show
    Unload Me
End Sub
Private Sub Restart_Click()
    SceneInit
    SceneInit
    SceneDraw
End Sub
