Attribute VB_Name = "GameMod"
Option Explicit
'Private soldier As m3SolidCollection
Public Function SoldierInit(ByVal size As Double) As m3SolidCollection
    Dim solids As m3SolidCollection
    Dim i As Integer
    Dim r As Double
    r = size / 2
    solids.nSolids = 4
    ReDim solids.solids(solids.nSolids - 1) As m3solid
    ReDim solids.points(solids.nSolids - 2) As m3Point
    solids.normal.x = 0
    solids.normal.y = 1
    solids.normal.z = 0
    '******************************************************
    solids.solids(0) = Solid3D.m3GetPrizm(r, size / 5, 18)
    solids.points(0).y = size / 5
    
    solids.solids(1) = Solid3D.m3GetPrizm(r / 1.3, size / 5, 12)
    m3solidApply solids.solids(1), Matrix3D.m3MatrixTranslate(0, size / 5, 0)
    solids.points(1).y = (size / 5) + (size / 5)
    
    solids.solids(2) = Solid3D.m3GetPyramid(size / 3.5, size / 2.2, 12)
    m3solidApply solids.solids(2), Matrix3D.m3MatrixTranslate(0, (size / 5) + (size / 5), 0)
    solids.points(2).y = (size / 5) + (size / 5) + (size / 2.2)
    
    solids.solids(3) = Solid3D.SphareInit(size / 6.65, 12)
    m3solidApply solids.solids(3), Matrix3D.m3MatrixTranslate(0, solids.points(2).y + (size / 6.65), 0)
    SoldierInit = solids
    '******************************************************
End Function
'Private Ratz As m3SolidCollection
Public Function RatzInit(ByVal size As Double) As m3SolidCollection
    Dim solids As m3SolidCollection
    Dim i As Integer
    Dim r As Double
    r = size / 2
    solids.nSolids = 4
    ReDim solids.solids(solids.nSolids - 1) As m3solid
    ReDim solids.points(solids.nSolids - 2) As m3Point
    solids.normal.x = 0
    solids.normal.y = 1
    solids.normal.z = 0
    '******************************************************
    solids.solids(0) = Solid3D.m3GetPrizm(r, size / 5.85, 18)
    solids.points(0).y = r / 5.85
    solids.solids(1) = Solid3D.m3GetPyramid(r / 1.25, size / 1.1, 18)
    m3solidApply solids.solids(1), Matrix3D.m3MatrixTranslate(0, size / 5.85, 0)
    solids.points(1).y = (size / (4.5 * 1.3)) + (size / 1.1)
    solids.solids(2) = Solid3D.SphareInit(r / 1.75, 12)
    m3solidApply solids.solids(2), Matrix3D.m3MatrixTranslate(0, solids.solids(1).verts(18).y + r / 1.75, 0)
    solids.solids(3) = Solid3D.SphareInit(r / 6.65, 12)
    m3solidApply solids.solids(3), Matrix3D.m3MatrixTranslate(0, solids.solids(1).verts(18).y + (r * 2) / 1.75 + r / 6.65, 0)
    RatzInit = solids
    '******************************************************
End Function
'Private Horse As m3SolidCollection
Public Function HorseInit(ByVal size As Double) As m3SolidCollection
    Dim solids As m3SolidCollection
    Dim i As Integer
    Dim r As Double
    r = size / 2
    solids.nSolids = 5
    ReDim solids.solids(solids.nSolids - 1) As m3solid
    ReDim solids.points(solids.nSolids - 2) As m3Point
    For i = 0 To solids.nSolids - 1
        solids.solids(i) = Solid3D.m3GetCube(size)
        
        m3solidApply solids.solids(i), m3MatrixTranslate(0, size * i, 0)
        
    Next i
    For i = 0 To solids.nSolids - 2
        solids.points(i) = m3PointInit(0, size * (i + 1), 0)
     
    Next i
    solids.normal.x = 0
    solids.normal.y = 1
    solids.normal.z = 0
    '******************************************************
    solids.solids(0) = Solid3D.m3GetPrizm(r, size / 6.3, 18)
    solids.points(0).y = size / 6.3
    
    solids.solids(1) = Solid3D.m3GetPrizm(r / 1.7, size / 1.15, 18)
    m3solidApply solids.solids(1), Matrix3D.m3MatrixTranslate(0, size / 6.3, 0)
    solids.points(1).y = (size / 6.3) + (size / 1.15)
    
    solids.solids(2) = Solid3D.m3GetPrizm(r / 1.15, size / 7.5, 18)
    m3solidApply solids.solids(2), Matrix3D.m3MatrixTranslate(0, solids.solids(1).verts(18).y, 0)
    solids.points(2).y = (size / 6.3) + (size / 1.15) + (size / 7.5)
        
    solids.solids(3) = Solid3D.m3GetPyramid(r / 1.7, size / 2, 18)
    m3solidApply solids.solids(3), Matrix3D.m3MatrixTranslate(0, solids.solids(2).verts(18).y, 0)
    solids.points(3).y = (size / 6.3) + (size / 1.15) + (size / 7.5) + (size / 2)
        
    solids.solids(4) = Solid3D.SphareInit(size / 6.65, 12)
    m3solidApply solids.solids(4), Matrix3D.m3MatrixTranslate(0, solids.solids(1).verts(18).y + (size / 2) + (size / 7.5) + (size / 6.65), 0)
    
    HorseInit = solids
    '******************************************************
End Function
'Private Tzariah As m3SolidCollection
Public Function TzariahInit(ByVal size As Double) As m3SolidCollection
    Dim solids As m3SolidCollection
    Dim i As Integer
    Dim r As Double
    r = size / 2
    
    solids.nSolids = 4
    ReDim solids.solids(solids.nSolids - 1) As m3solid
    ReDim solids.points(solids.nSolids - 2) As m3Point
    solids.normal.x = 0
    solids.normal.y = 1
    solids.normal.z = 0
    '******************************************************
    solids.solids(0) = Solid3D.m3GetPrizm(r, size / 5.85, 18)
    solids.points(0).y = r / 5.85
    solids.solids(1) = Solid3D.m3GetPrizm(r / 2.2, size / 1.4, 18)
    m3solidApply solids.solids(1), Matrix3D.m3MatrixTranslate(0, size / 5.85, 0)
    solids.points(1).y = (size / (5.85)) + (size / 1.4)
    solids.solids(2) = Solid3D.m3GetPrizm(r / 1, size / 5.85, 18)
    m3solidApply solids.solids(2), Matrix3D.m3MatrixTranslate(0, solids.solids(1).verts(18).y, 0)
    solids.solids(3) = Solid3D.SphareInit(size / 6.65, 12)
    m3solidApply solids.solids(3), Matrix3D.m3MatrixTranslate(0, solids.solids(1).verts(18).y + size / 3.7 + size / 6.65, 0)
    TzariahInit = solids
    '******************************************************
End Function
'Private Malka As m3SolidCollection
Public Function MalkaInit(ByVal size As Double) As m3SolidCollection
    Dim solids As m3SolidCollection
    Dim i As Integer
    Dim r As Double
    r = size / 2
    solids.nSolids = 6
    ReDim solids.solids(solids.nSolids - 1) As m3solid
    ReDim solids.points(solids.nSolids - 2) As m3Point
    For i = 0 To solids.nSolids - 1
        solids.solids(i) = Solid3D.m3GetCube(size)
        
        m3solidApply solids.solids(i), m3MatrixTranslate(0, size * i, 0)
        
    Next i
    For i = 0 To solids.nSolids - 2
        solids.points(i) = m3PointInit(0, size * (i + 1), 0)
     
    Next i
    solids.normal.x = 0
    solids.normal.y = 1
    solids.normal.z = 0
    '******************************************************
    solids.solids(0) = Solid3D.m3GetPrizm(r, size / 6.3, 18)
    solids.points(0).y = size / 6.3
    
    solids.solids(1) = Solid3D.m3GetPyramid(r / 1.3, size, 18)
    m3solidApply solids.solids(1), Matrix3D.m3MatrixTranslate(0, size / 6.3, 0)
    solids.points(1).y = (size / 6.3) + (size)
    
    solids.solids(2) = Solid3D.m3GetPrizm(r / 1.3, size / 7.5, 18)
    m3solidApply solids.solids(2), Matrix3D.m3MatrixTranslate(0, solids.solids(1).verts(18).y, 0)
    solids.points(2).y = (size / 6.3) + (size) + (size / 7.5)
        
    solids.solids(3) = Solid3D.m3GetPrizm(r / 1.7, size / 7.5, 18)
    m3solidApply solids.solids(3), Matrix3D.m3MatrixTranslate(0, solids.solids(2).verts(18).y, 0)
    solids.points(3).y = (size / 6.3) + (size) + (size / 7.5) + (size / 7.5)
        
    solids.solids(4) = Solid3D.m3GetPyramid(r / 2.3, size / 3.4, 18)
    m3solidApply solids.solids(4), Matrix3D.m3MatrixTranslate(0, solids.solids(1).verts(18).y + (size / 7.5) + (size / 7.5), 0)
    solids.points(4).y = (size / 6.3) + (size) + (size / 7.5) + (size / 7.5) + (size / 3.4)
    
    solids.solids(5) = Solid3D.SphareInit(size / 6.65, 12)
    m3solidApply solids.solids(5), Matrix3D.m3MatrixTranslate(0, solids.solids(1).verts(18).y + (size / 7.5) + (size / 7.5) + (size / 3.4) + (size / 6.65), 0)
    
    MalkaInit = solids
    '******************************************************
End Function
'Private King As m3SolidCollection
Public Function KingInit(ByVal size As Double) As m3SolidCollection
    Dim solids As m3SolidCollection
    Dim i As Integer
    Dim r As Double
    r = size / 2
    solids.nSolids = 7
    ReDim solids.solids(solids.nSolids - 1) As m3solid
    ReDim solids.points(solids.nSolids - 2) As m3Point
    For i = 0 To solids.nSolids - 1
        solids.solids(i) = Solid3D.m3GetCube(size)
        
        m3solidApply solids.solids(i), m3MatrixTranslate(0, size * i, 0)
        
    Next i
    For i = 0 To solids.nSolids - 2
        solids.points(i) = m3PointInit(0, size * (i + 1), 0)
     
    Next i
    solids.normal.x = 0
    solids.normal.y = 1
    solids.normal.z = 0
    '******************************************************
    solids.solids(0) = Solid3D.m3GetPrizm(r, size / 6.3, 18)
    solids.points(0).y = size / 6.3
    
    solids.solids(1) = Solid3D.m3GetPrizm(r / 1.5, size, 18)
    m3solidApply solids.solids(1), Matrix3D.m3MatrixTranslate(0, size / 6.3, 0)
    solids.points(1).y = (size / 6.3) + (size)
    
    solids.solids(2) = Solid3D.m3GetPrizm(r / 1.05, size / 7.5, 18)
    m3solidApply solids.solids(2), Matrix3D.m3MatrixTranslate(0, solids.solids(1).verts(18).y, 0)
    solids.points(2).y = (size / 6.3) + (size) + (size / 7.5)
    
    solids.solids(3) = Solid3D.m3GetPrizm(r / 1.3, size / 7.5, 18)
    m3solidApply solids.solids(3), Matrix3D.m3MatrixTranslate(0, solids.solids(1).verts(18).y + (size / 7.5), 0)
    solids.points(3).y = (size / 6.3) + (size) + (size / 7.5) + (size / 7.5)
             
    solids.solids(4) = Solid3D.m3GetPrizm(r / 1.75, size / 7.5, 18)
    m3solidApply solids.solids(4), Matrix3D.m3MatrixTranslate(0, solids.solids(1).verts(18).y + (size / 7.5) + (size / 7.5), 0)
    solids.points(4).y = (size / 6.3) + (size) + (size / 7.5) + (size / 7.5) + (size / 7.5)
            
    solids.solids(5) = Solid3D.m3GetPrizm(r / 2.4, size / 7.5, 18)
    m3solidApply solids.solids(5), Matrix3D.m3MatrixTranslate(0, solids.solids(1).verts(18).y + (size / 7.5) + (size / 7.5) + (size / 7.5), 0)
    solids.points(5).y = (size / 6.3) + (size) + (size / 7.5) + (size / 7.5) + (size / 7.5) + (size / 7.5)
    
    solids.solids(6) = Solid3D.SphareInit(size / 6.65, 12)
    m3solidApply solids.solids(6), Matrix3D.m3MatrixTranslate(0, solids.solids(1).verts(18).y + (size / 7.5) + (size / 7.5) + (size / 7.5) + (size / 6.65) + (size / 7.5), 0)
        
    KingInit = solids
    '******************************************************
End Function
