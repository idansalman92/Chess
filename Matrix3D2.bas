Attribute VB_Name = "Matrix3D2"
Option Explicit
Public Function m3LineRotate(ByRef P As m3Point, ByRef d As m3Vector, ByVal theta As Single) As m3Matrix
    Dim M As m3Matrix
    Dim sn As Double
    Dim cs As Double
    Dim c As Double
    sn = Sin(theta)
    cs = Cos(theta)
    c = 1 - cs
    m3VectorSetLen 1, d
    M.m11 = d.X * d.X * c + cs
    M.m12 = d.X * d.Y * c + d.Z * sn
    M.m13 = d.X * d.Z * c - d.Y * sn
    M.m21 = d.Y * d.X * c - d.Z * sn
    M.m22 = d.Y * d.Y * c + cs
    M.m23 = d.Y * d.Z * c + d.X * sn
    M.m31 = d.Z * d.X * c + d.Y * sn
    M.m32 = d.Z * d.Y * c - d.X * sn
    M.m33 = d.Z * d.Z * c + cs
    M.m41 = P.X - P.X * M.m11 - P.Y * M.m21 - P.Z * M.m31
    M.m42 = P.Y - P.X * M.m12 - P.Y * M.m22 - P.Z * M.m32
    M.m43 = P.Z - P.X * M.m13 - P.Y * M.m23 - P.Z * M.m33
    m3LineRotate = M
End Function
