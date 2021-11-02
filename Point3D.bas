Attribute VB_Name = "Point3D"
Option Explicit
Public Type m3Point
    x As Double
    y As Double
    z As Double
End Type

Public Function m3PointInit(ByVal x As Double, ByVal y As Double, ByVal z As Double) As m3Point
    Dim m As m3Point
    m.x = x
    m.y = y
    m.z = z
    m3PointInit = m
End Function
Public Function m3PointToString(ByRef m As m3Point) As String
    Dim S As String
    S = "(" & Format(m.x, "0.00") & vbTab & Format(m.y, "0.00") & vbTab & Format(m.z, "0.00") & ")"
        m3PointToString = S
End Function
Public Sub m3PointApply(ByRef p As m3Point, ByRef m As m3Matrix)
    Dim x As Double
    Dim y As Double
    Dim z As Double
    x = p.x
    y = p.y
    z = p.z
    p.x = x * m.m11 + y * m.m21 + z * m.m31 + m.m41
    p.y = x * m.m12 + y * m.m22 + z * m.m32 + m.m42
    p.z = x * m.m13 + y * m.m23 + z * m.m33 + m.m43
    '1
End Sub
