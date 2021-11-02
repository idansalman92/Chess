Attribute VB_Name = "SolidCollection3D"
Option Explicit
Public Type m3SolidCollection
    nSolids As Integer
    solids() As m3solid
    normal As m3Vector
    points() As m3Point
End Type
Public Sub m3SolidCollectionApply(ByRef S As m3SolidCollection, ByRef m As m3Matrix)
    Dim i As Integer
    m3VectorApply S.normal, m
    For i = 0 To S.nSolids - 1
        m3solidApply S.solids(i), m
    Next i
    For i = 0 To S.nSolids - 2
        m3PointApply S.points(i), m
    Next i
End Sub
Public Sub m3SolidCollectionDraw(ByRef Obj As Object, ByRef S As m3SolidCollection)
    Dim midPlane As Integer
    Dim dot As Double
    Dim V As m3Vector
    Dim p As m3Point
    Dim i As Integer
    midPlane = S.nSolids - 1
    p = m3PointInit(0, 0, Draw3D.m3GetDistance)
    For i = 0 To S.nSolids - 2
        V = m3VectorInit(p, S.points(i))
        dot = m3VectorDot(V, S.normal)
        If dot > 0 Then
           midPlane = i
           Exit For
        End If
    Next i
    For i = 0 To midPlane - 1
        m3SolidFillShading Obj, S.solids(i)
    Next i
     For i = S.nSolids - 1 To midPlane Step -1
        m3SolidFillShading Obj, S.solids(i)
    Next i
    
End Sub
Public Function m3SolidCollectionCenter(ByRef S As m3SolidCollection) As m3Point
    Dim i As Integer
    Dim cent As m3Point
    Dim x As Double
    Dim y As Double
    Dim z As Double
    x = 0: y = 0: z = 0
    For i = 0 To S.nSolids - 1
        cent = m3solidcenter(S.solids(i))
        x = x + cent.x
        y = y + cent.y
        z = z + cent.z
    Next i
    m3SolidCollectionCenter.x = x / S.nSolids
    m3SolidCollectionCenter.y = y / S.nSolids
    m3SolidCollectionCenter.z = z / S.nSolids
End Function
