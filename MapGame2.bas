Attribute VB_Name = "MapGame"
Option Explicit
Private Const ROWS = 8
Private Const COLS = 8
Private cellSize As Double
Private location As m3Point
Private iX As m3Vector
Private jY As m3Vector
Private map() As Integer
Public Const EMPTY_CELL = -1
Public Const SOLDIER_CELL = 0 ' White soldier
Public Const TZARIAH_CELL = 1 ' White Tzariah
Public Const HORSE_CELL = 2 ' White horse
Public Const RATZR_CELL = 3 ' White ratz right
Public Const RATZL_CELL = 7 ' White ratz left
Public Const MALKA_CELL = 4 'White Queen
Public Const KING_CELL = 5 ' White king
Public Const SOLDIERB_CELL = 8 ' Black soldier
Public Const TZARIAHB_CELL = 9 ' Black Tzariah
Public Const HORSEB_CELL = 10 ' Black horse
Public Const RATZBR_CELL = 11 ' Black ratz right
Public Const RATZBL_CELL = 14 ' Black ratz left
Public Const MALKAB_CELL = 12 ' Black Queen
Public Const KINGB_CELL = 13 ' Black king
Public Const NO_IN_BOARD = -100
Private base As m3solid
Private shapes() As m3SolidCollection
'******* Location From *****************
' if location not in board then rowPick or colPick = -1
' location in board if 0 <= rowPick and rowPick <= 7
' and 0 <= colPick and colPick <= 7
Private rowPickFrom As Integer
Private colPickFrom As Integer
' update rowPick and colPick and return value of board
' if location not in board return NO_IN_BOARD
Private rowPickTo As Integer
Private colPickTo As Integer
' update rowPick and colPick and return value of board
' if location not in board return NO_IN_BOARD
Public Type CellLocation
    row As Integer
    col As Integer
End Type
Public Function SetFrom(ByVal x As Double, ByVal y As Double) As Boolean
    Dim loc As CellLocation
    loc = GetCellLocation(x, y)
    If loc.row > 7 Or loc.row < 0 Then
        rowPickFrom = -1
        colPickFrom = -1
        SetFrom = False
        Exit Function
    End If
    If loc.col > 7 Or loc.col < 0 Then
        rowPickFrom = -1
        colPickFrom = -1
        SetFrom = False
        Exit Function
    End If
    If map(loc.row, loc.col) = EMPTY_CELL Then
        rowPickFrom = -1
        colPickFrom = -1
        SetFrom = False
        Exit Function
    End If
    rowPickFrom = loc.row
    colPickFrom = loc.col
    SetFrom = True
End Function
Public Function SetTo(ByVal x As Double, ByVal y As Double) As Boolean
    Dim loc As CellLocation
    loc = GetCellLocation(x, y)
    If loc.row > 7 Or loc.row < 0 Then
        rowPickTo = -1
        colPickTo = -1
        SetTo = False
        Exit Function
    End If
    If loc.row > 7 Or loc.row < 0 Then
        rowPickTo = -1
        colPickTo = -1
        SetTo = False
        Exit Function
    End If
    
    rowPickTo = loc.row
    colPickTo = loc.col
    SetTo = True
End Function
Public Function GetCellLocation(ByVal x As Double, ByVal y As Double) As CellLocation
    Dim p As m3Point ' Point 2D Point 3D
    Dim v As m3Vector
    Dim row As Integer
    Dim col As Integer
    Dim dotX As Double
    Dim dotY As Double
    p = MousePoint2DTo3D(x, y)
    v = m3VectorInit(location, p)
    dotX = m3VectorDot(v, iX)
    dotY = m3VectorDot(v, jY)
    If dotX < 0 Or dotY < 0 Then
        GetCellLocation.row = -1
        GetCellLocation.col = -1
        Exit Function
    End If
    col = Int(dotX / (cellSize * cellSize))
    row = Int(dotY / (cellSize * cellSize))
    If row > 7 Or row < 0 Then
        GetCellLocation.row = -1
        GetCellLocation.col = -1
        Exit Function
    End If
    If col > 7 Or col < 0 Then
        GetCellLocation.row = -1
        GetCellLocation.col = -1
        Exit Function
    End If
    GetCellLocation.row = row
    GetCellLocation.col = col
End Function
Private Function MousePoint2DTo3D(ByVal x As Double, ByVal y As Double) As m3Point
    Dim c As m3Point
    Dim v As m3Vector
    Dim w As m3Vector
    Dim n As m3Vector
    Dim dot As Double
    Dim dot1 As Double
    Dim t As Double
    c = m3PointInit(0, 0, Draw3D.m3GetDistance)
    v = m3VectorInit(c, m3PointInit(x, y, 0))
    n = m3VectorCross(iX, jY)
    dot = m3VectorDot(v, n)
    w = m3VectorInit(c, location)
    dot1 = m3VectorDot(w, n)
    t = dot1 / dot
    MousePoint2DTo3D.x = c.x + t * v.x
    MousePoint2DTo3D.y = c.y + t * v.y
    MousePoint2DTo3D.z = c.z + t * v.z
    
End Function
Public Function MoveFromTo(ByRef Obj As Object) As Boolean
    Dim Eat As Boolean
    Dim fig As Integer
    
    Dim pFrom As m3Point
    Dim pTo As m3Point
    Dim s As m3SolidCollection
    Dim tX As Double
    Dim tY As Double
    Dim tZ As Double
    Dim i As Integer
    Dim nSteps As Integer
    Dim v As m3Vector
    Dim n As m3Vector
    Dim colF As Integer
    Dim rowF As Integer
    If rowPickFrom = -1 Or colPickFrom = -1 Then
        MoveFromTo = False
        Exit Function
    End If
    If rowPickTo = -1 Or colPickTo = -1 Then
        MoveFromTo = False
        Exit Function
    End If
    
    '******** Check if Move in role*********************
    
    '******** Check Soldier Moves **********************
    If (map(rowPickFrom, colPickFrom) = 0) Or (map(rowPickFrom, colPickFrom) = 8) Then
        If (rowPickFrom = rowPickTo) And (colPickFrom = colPickTo - 1) Then
            If map(rowPickTo, colPickTo) = -1 Then
                MoveFromTo = True
            Else
                MoveFromTo = False
                Exit Function
            End If
        ElseIf (colPickFrom + 1 = colPickTo) And (rowPickFrom + 1 = rowPickTo Or rowPickFrom - 1 = rowPickTo) Then
            If map(rowPickTo, colPickTo) >= 0 Then
                MoveFromTo = True
            Else
                MoveFromTo = False
                Exit Function
            End If
        Else
            MoveFromTo = False
            Exit Function
        End If
    End If
    '**************** End Soldier Check ****************
    If map(rowPickFrom, colPickFrom) = 2 Then
        If (rowPickFrom - rowPickTo = colPickFrom - colPickTo Or rowPickFrom - rowPickTo = colPickTo - colPickFrom) Then
            For rowF = rowPickFrom To rowPickTo
                For colF = colPickFrom To colPickTo
                    If map(rowF, colF) >= 0 Then
                                        'Else
                        rowPickTo = rowF
                        colPickTo = colF
                        rowF = rowPickTo - 1
                        colF = colPickTo - 1
                        'MoveFromTo = True
                    End If
                Next colF
            Next rowF
            MoveFromTo = True
        Else
            MoveFromTo = False
            Exit Function
        End If
    End If
    
    fig = map(rowPickFrom, colPickFrom)
    s = shapes(fig Mod 4)
    tX = iX.x * colPickFrom + jY.x * rowPickFrom
    tY = iX.y * colPickFrom + jY.y * rowPickFrom
    tZ = iX.z * colPickFrom + jY.z * rowPickFrom
    m3SolidCollectionApply s, m3MatrixTranslate(tX, tY, tZ)
    map(rowPickFrom, colPickFrom) = EMPTY_CELL
    map(rowPickTo, colPickTo) = EMPTY_CELL
    pFrom.x = location.x + iX.x * (colPickFrom + 0.5) + jY.x * (rowPickFrom + 0.5)
    pFrom.y = location.y + iX.y * (colPickFrom + 0.5) + jY.y * (rowPickFrom + 0.5)
    pFrom.z = location.z + iX.z * (colPickFrom + 0.5) + jY.z * (rowPickFrom + 0.5)
    pTo.x = location.x + iX.x * (colPickTo + 0.5) + jY.x * (rowPickTo + 0.5)
    pTo.y = location.y + iX.y * (colPickTo + 0.5) + jY.y * (rowPickTo + 0.5)
    pTo.z = location.z + iX.z * (colPickTo + 0.5) + jY.z * (rowPickTo + 0.5)
    v = m3VectorInit(pFrom, pTo)
    nSteps = 20
    
    Obj.Cls
    GameDraw Obj
    Obj.Refresh
    Obj.Picture = Obj.Image
    Obj.Refresh
    If fig <= 7 And fig >= 1 Then
        Obj.FillColor = vbWhite
        Obj.ForeColor = 0
    Else
        Obj.FillColor = 7735039
        Obj.ForeColor = 0
    End If
    n = m3VectorCross(iX, jY)
    Vector3D.m3VectorSetLen 3, n
    For i = 1 To nSteps
        Obj.Cls
        m3SolidCollectionApply s, m3MatrixTranslate(n.x, n.y, n.z)
        m3SolidCollectionDraw Obj, s
        Obj.Refresh
    Next i
    tX = v.x / nSteps
    tY = v.y / nSteps
    tZ = v.z / nSteps
    For i = 1 To nSteps
        Obj.Cls
        m3SolidCollectionApply s, m3MatrixTranslate(tX, tY, tZ)
        m3SolidCollectionDraw Obj, s
        Obj.Refresh
    Next i
    For i = 1 To nSteps
        Obj.Cls
        m3SolidCollectionApply s, m3MatrixTranslate(-n.x, -n.y, -n.z)
        m3SolidCollectionDraw Obj, s
        Obj.Refresh
    Next i
    Obj.Picture = LoadPicture("")
    map(rowPickTo, colPickTo) = fig
    Obj.Cls
    GameDraw Obj
End Function
Private Sub MapDrawMapCell(ByRef Obj As Object, ByVal row As Integer, ByVal col As Integer)
    Dim tmp As m3SolidCollection
    Dim tX As Double
    Dim tY As Double
    Dim tZ As Double
    If map(row, col) = EMPTY_CELL Then
        Exit Sub
    End If
    If map(row, col) <= 7 And map(row, col) >= 0 Then
        Obj.FillColor = vbWhite
        Obj.ForeColor = 0
    Else
        Obj.FillColor = 7735039
        Obj.ForeColor = 0
    End If
    
    tmp = shapes(map(row, col) Mod 8)
    tX = iX.x * col + jY.x * row
    tY = iX.y * col + jY.y * row
    tZ = iX.z * col + jY.z * row
    m3SolidCollectionApply tmp, m3MatrixTranslate(tX, tY, tZ)
    m3SolidCollectionDraw Obj, tmp
 
End Sub
Public Sub MapGameInit(ByVal size As Double)
    Dim p As m3Point
    Dim i As Integer
    Dim j As Integer
    Dim v As m3Vector
    Dim ix2 As m3Vector
    Dim center As m3Point
    Dim tX As Double
    Dim tY As Double
    Dim tZ As Double
    Dim M As m3Matrix
    MapGame.cellSize = size
    ReDim map(ROWS - 1, COLS - 1) As Integer
    For i = 0 To ROWS - 1
        For j = 0 To COLS - 1
            map(i, j) = EMPTY_CELL
        Next j
    Next i
    For i = 0 To ROWS - 1
        For j = 1 To 1
            map(i, j) = SOLDIER_CELL
        Next j
        For j = 6 To 6
            map(i, j) = SOLDIERB_CELL
        Next j
        For j = 0 To 0
            If i = 0 Or i = 7 Then
                map(i, j) = TZARIAH_CELL
            End If
            If i = 1 Or i = 6 Then
                map(i, j) = HORSE_CELL
            End If
            If i = 2 Then
                map(i, j) = RATZR_CELL
            End If
            If i = 5 Then
                map(i, j) = RATZL_CELL
            End If
            If i = 3 Then
                map(i, j) = MALKA_CELL
            End If
            If i = 4 Then
                map(i, j) = KING_CELL
            End If
        Next j
        For j = 7 To 7
            If i = 0 Or i = 7 Then
                map(i, j) = TZARIAHB_CELL
            End If
            If i = 1 Or i = 6 Then
                map(i, j) = HORSEB_CELL
            End If
            If i = 2 Then
                map(i, j) = RATZBR_CELL
            End If
            If i = 5 Then
                map(i, j) = RATZBL_CELL
            End If
            If i = 3 Then
                map(i, j) = MALKAB_CELL
            End If
            If i = 4 Then
                map(i, j) = KINGB_CELL
            End If
        Next j
    Next i
    location = m3PointInit(-COLS * cellSize / 2, -ROWS * cellSize / 2, 0)
    iX.x = cellSize
    iX.y = 0
    iX.z = 0
    jY.x = 0
    jY.y = cellSize
    jY.z = 0

    ix2 = iX
    '**************************************************
    base = m3GetCube(1)
    m3solidApply base, m3MatrixScale(COLS * cellSize, ROWS * cellSize, cellSize / 2)
    v = m3VectorInit(base.verts(0), location)
    m3solidApply base, m3MatrixTranslate(v.x, v.y, v.z)
    '******************************************************************************
    ReDim shapes(0 To 7) As m3SolidCollection
    shapes(0) = SoldierInit(cellSize / 1.5)
    shapes(1) = TzariahInit(cellSize / 1.5)
    shapes(2) = RatzInit(cellSize / 1.5)
    shapes(3) = RatzInit(cellSize / 1.5) ' horse
    shapes(4) = MalkaInit(cellSize / 1.5)
    shapes(5) = KingInit(cellSize / 1.5)
    shapes(6) = MalkaInit(cellSize / 1.5)
    shapes(7) = RatzInit(cellSize / 1.5)
    p = m3PointInit(-COLS * cellSize / 2 + cellSize / 2, -ROWS * cellSize / 2 + cellSize / 2, 0)
    tX = -COLS * cellSize / 2 + cellSize / 2
    tY = -ROWS * cellSize / 2 + cellSize / 2
    tZ = 0
    M = m3XRotate(PI / 2)
    For i = 0 To 7
         m3SolidCollectionApply shapes(i), M
         
    Next i
    M = m3MatrixTranslate(tX, tY, tZ)
    For i = 0 To 7
         m3SolidCollectionApply shapes(i), M
         
    Next i
    rowPickFrom = -1
    colPickFrom = -1
    rowPickTo = -1
    colPickTo = -1
End Sub
Public Sub MapApply(ByRef M As m3Matrix)
    Dim i As Integer
    m3PointApply location, M
    m3VectorApply iX, M
    m3VectorApply jY, M
    m3solidApply base, M
    For i = 0 To 7 '5
        m3SolidCollectionApply shapes(i), M
    
    Next i
End Sub
Public Sub GameDraw(ByRef Obj As Object)
    If Draw3D.m3PlaneIsVisible(base.verts(0), base.verts(1), base.verts(2)) Then
        MapDraw Obj
        MapDrawMap Obj
    Else
        MapDrawMap Obj
        MapDraw Obj
    End If
    'm3DrawLine obj, m3PointInit(-600, -600, 0), base.verts(2)
End Sub
Private Sub MapDraw(ByRef Obj As Object)
    Dim i As Integer
    Dim j As Integer
    Dim p As m3Point
    Dim tmpSolid As m3solid
    tmpSolid.nVerts = 4
    tmpSolid.nFaces = 1
    ReDim tmpSolid.verts(3) As m3Point
    ReDim tmpSolid.Faces(1) As m3Face
    tmpSolid.Faces(0).nLinks = 4
    ReDim tmpSolid.Faces(0).Links(3) As Integer
    tmpSolid.Faces(0).Links(0) = 0
    tmpSolid.Faces(0).Links(1) = 1
    tmpSolid.Faces(0).Links(2) = 2
    tmpSolid.Faces(0).Links(3) = 3
    Obj.FillColor = 3158064
    m3SolidFillShading Obj, base
    For i = 0 To ROWS - 1
        For j = 0 To COLS - 1
            p.x = location.x + iX.x * j + jY.x * i
            p.y = location.y + iX.y * j + jY.y * i
            p.z = location.z + iX.z * j + jY.z * i
            tmpSolid.verts(0) = p
            tmpSolid.verts(1).x = p.x + iX.x
            tmpSolid.verts(1).y = p.y + iX.y
            tmpSolid.verts(1).z = p.z + iX.z
            tmpSolid.verts(2).x = p.x + iX.x + jY.x
            tmpSolid.verts(2).y = p.y + iX.y + jY.y
            tmpSolid.verts(2).z = p.z + iX.z + jY.z
            tmpSolid.verts(3).x = p.x + jY.x
            tmpSolid.verts(3).y = p.y + jY.y
            tmpSolid.verts(3).z = p.z + jY.z
            Obj.FillColor = vbWhite
            If (i + j) Mod 2 = 0 Then
                Obj.FillColor = 3158064
                
            End If
            m3SolidFillShading Obj, tmpSolid
                        
        Next j
    Next i
    
    
    'm3DrawLine obj, m3PointInit(-600, -600, 0), shapes(0).solids(0).verts(shapes(0).solids(0).nVerts / 4)
End Sub
Public Function MapCenter() As m3Point
    Dim p As m3Point
    p.x = location.x + (iX.x * COLS + jY.x * ROWS) / 2
    p.y = location.y + (iX.y * COLS + jY.y * ROWS) / 2
    p.z = location.z + (iX.z * COLS + jY.z * ROWS) / 2
    MapCenter = p
End Function

Private Sub MapDrawMap(ByRef Obj As Object)
    Dim i As Integer
    Dim j As Integer
    Dim p1 As m3Point
    Dim P2 As m3Point
    Dim v As m3Vector
    Dim midCol As Integer
    Dim midRow As Integer
    Dim dot As Double
    midCol = COLS - 1
    p1 = m3PointInit(0, 0, Draw3D.m3GetDistance)
    For j = 0 To COLS - 2
        P2.x = location.x + iX.x * (j + 1)
        P2.y = location.y + iX.y * (j + 1)
        P2.z = location.z + iX.z * (j + 1)
        v = m3VectorInit(p1, P2)
        dot = iX.x * v.x + iX.y * v.y + iX.z * v.z
        If dot > 0 Then
            midCol = j
            Exit For
        End If
    Next j
    midRow = ROWS - 1
    For i = 0 To ROWS - 2
        P2.x = location.x + jY.x * (i + 1)
        P2.y = location.y + jY.y * (i + 1)
        P2.z = location.z + jY.z * (i + 1)
        v = m3VectorInit(p1, P2)
        dot = jY.x * v.x + jY.y * v.y + jY.z * v.z
        If dot > 0 Then
            midRow = i
            Exit For
        End If
    Next i
    '*************************************************************
    For i = 0 To midRow - 1
        For j = 0 To midCol - 1
            MapDrawMapCell Obj, i, j
            
        Next j
        For j = COLS - 1 To midCol Step -1
            MapDrawMapCell Obj, i, j
        Next j
    Next i
    For i = ROWS - 1 To midRow Step -1
        For j = 0 To midCol - 1
            MapDrawMapCell Obj, i, j
        Next j
        For j = COLS - 1 To midCol Step -1
            MapDrawMapCell Obj, i, j
        Next j
    Next i
    
    
End Sub
