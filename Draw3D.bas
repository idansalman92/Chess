Attribute VB_Name = "Draw3D"
Option Explicit
Public Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long

Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private dist As Double
Private lightVector As m3Vector
Public Const AmbFactor As Double = 0.1
Public Sub m3SetLightVector(ByRef lVector As m3Vector)
    Dim L As Single
    L = m3VectorLen(lVector)
    If L < 0.000001 Then
        lightVector.x = 0
        lightVector.y = 0
        lightVector.z = 1
    Else
        lightVector.x = lVector.x / L
        lightVector.y = lVector.y / L
        lightVector.z = lVector.z / L
    End If
End Sub
Public Function m3GetLightVector() As m3Vector
    m3GetLightVector = lightVector
End Function
Public Sub m3SetDistance(ByVal newDist As Double)
    If newDist > 100 Then
        dist = newDist
    End If
End Sub
Public Function m3GetDistance() As Double
    m3GetDistance = dist
End Function
Public Sub m3DrawLine(ByRef Obj As Object, ByRef P1 As m3Point, ByRef P2 As m3Point)
    Dim X1 As Long
    Dim y1 As Long
    Dim X2 As Long
    Dim y2 As Long
    Dim xorg As Double
    Dim yorg As Double
    Dim f As Double
    Dim pt As POINTAPI
    xorg = -Obj.ScaleLeft
    yorg = Obj.ScaleTop
    f = dist / (dist - P1.z)
    X1 = xorg + P1.x * f
    y1 = yorg - P1.y * f
    f = dist / (dist - P2.z)
    X2 = xorg + P2.x * f
    y2 = yorg - P2.y * f
    MoveToEx Obj.hdc, X1, y1, pt
    LineTo Obj.hdc, X2, y2
End Sub
Public Function m3PlaneIsVisible(ByRef P0 As m3Point, ByRef P1 As m3Point, ByRef P2 As m3Point) As Boolean
    Dim PP0 As m3Point
    Dim PP1 As m3Point
    Dim PP2 As m3Point
    Dim f As Double
    Dim u As m3Vector
    Dim V As m3Vector
    Dim n As m3Vector
    f = dist / (dist - P0.z)
    PP0 = m3PointInit(P0.x * f, P0.y * f, P0.z)
    f = dist / (dist - P1.z)
    PP1 = m3PointInit(P1.x * f, P1.y * f, P1.z)
    f = dist / (dist - P2.z)
    PP2 = m3PointInit(P2.x * f, P2.y * f, P2.z)
    V = m3VectorInit(PP0, PP1)
    u = m3VectorInit(PP1, PP2)
    n = m3VectorCross(V, u)
    m3PlaneIsVisible = (n.z >= 0)
End Function
