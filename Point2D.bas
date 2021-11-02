Attribute VB_Name = "Point2D"
Option Explicit
Public Type m2Point
        x As Double
        y As Double
End Type
Public Function m2PointInit(ByVal x As Double, ByVal y As Double) As m2Point
    Dim p As m2Point
    p.x = x
    p.y = y
    m2PointInit = p
End Function
Public Function m2PointToString(ByRef p As m2Point) As String
    Dim s As String
    s = "(" & Format(p.x, "0.00") & vbTab & Format(p.y, "0.00") & ")"
    m2PointToString = s
End Function


