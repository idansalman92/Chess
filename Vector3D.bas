Attribute VB_Name = "Vector3D"
Option Explicit
Public Type m3Vector
    X As Double
    Y As Double
    Z As Double
    'W=0
End Type
Public Function m3VectorInit(ByRef P1 As m3Point, ByRef P2 As m3Point) As m3Vector
    Dim Pr As m3Vector
    Pr.X = P2.X - P1.X
    Pr.Y = P2.Y - P1.Y
    Pr.Z = P2.Z - P1.Z
    m3VectorInit = Pr
End Function
Public Function m3VectorLen(ByRef V As m3Vector) As Double
    m3VectorLen = Sqr(V.X * V.X + V.Y * V.Y + V.Z * V.Z)
End Function
Public Sub m3VectorSetLen(ByVal L As Double, ByRef V As m3Vector)
    L = L / m3VectorLen(V)
    V.X = V.X * L
    V.Y = V.Y * L
    V.Z = V.Z * L
End Sub
Public Function m3VectSum(ByRef V As m3Vector, ByRef U As m3Vector) As m3Vector
    Dim w As m3Vector
    w.X = V.X + U.X
    w.Y = V.Y + U.Y
    w.Z = V.Z + U.Z
    m3VectSum = w
End Function
Public Sub m3VectorApply(ByRef V As m3Vector, ByRef m As m3Matrix)
    Dim X As Double
    Dim Y As Double
    Dim Z As Double
    X = V.X
    Y = V.Y
    Z = V.Z
    V.X = X * m.m11 + Y * m.m21 + Z * m.m31
    V.Y = X * m.m12 + Y * m.m22 + Z * m.m32
    V.Z = X * m.m13 + Y * m.m23 + Z * m.m33
End Sub
Public Function m3VectorToString(ByRef V As m3Vector) As String
    Dim s As String
    s = "(" & Format(V.X, "0.00") & vbTab & Format(V.Y, "0.00") & vbTab & Format(V.Z, "0.00") & ")"
    m3VectorToString = s
End Function
Public Function m3VectorCross(ByRef V As m3Vector, ByRef U As m3Vector) As m3Vector
    Dim w As m3Vector
    w.X = V.Y * U.Z - V.Z * U.Y
    w.Y = V.Z * U.X - V.X * U.Z
    w.Z = V.X * U.Y - V.Y * U.X
    m3VectorCross = w
End Function
Public Function m3VectorDot(ByRef V As m3Vector, ByRef U As m3Vector) As Double
    m3VectorDot = V.X * U.X + V.Y * U.Y + V.Z * U.Z
End Function
