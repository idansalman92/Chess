Attribute VB_Name = "Matrix3D"
Option Explicit
Public Const PI = 3.14159265359
Public Type m3Matrix
    m11 As Double
    m12 As Double
    m13 As Double
    '0
    m21 As Double
    m22 As Double
    m23 As Double
    '0
    m31 As Double
    m32 As Double
    m33 As Double
    '0
    m41 As Double
    m42 As Double
    m43 As Double
    'm44 = 1
End Type
Public Function m3MatrixToString(ByRef M As m3Matrix) As String
    Dim S As String
    S = Format(M.m11, "0.00") & vbTab & _
        Format(M.m12, "0.00") & vbTab & _
        Format(M.m13, "0.00") & vbTab & "0" & vbCrLf & _
        Format(M.m21, "0.00") & vbTab & _
        Format(M.m22, "0.00") & vbTab & _
        Format(M.m23, "0.00") & vbTab & "0" & vbCrLf & _
        Format(M.m31, "0.00") & vbTab & _
        Format(M.m32, "0.00") & vbTab & _
        Format(M.m33, "0.00") & vbTab & "0" & vbCrLf & _
        Format(M.m41, "0.00") & vbTab & _
        Format(M.m42, "0.00") & vbTab & _
        Format(M.m43, "0.00") & vbTab & "1"
        m3MatrixToString = S
        

End Function

Public Function m3MatrixMultiply(ByRef a As m3Matrix, ByRef b As m3Matrix) As m3Matrix
    Dim c As m3Matrix
    c.m11 = a.m11 * b.m11 + a.m12 * b.m21 + a.m13 * b.m31
    c.m12 = a.m11 * b.m12 + a.m12 * b.m22 + a.m13 * b.m32
    c.m13 = a.m11 * b.m13 + a.m12 * b.m23 + a.m13 * b.m33
    ' C.m14 = 0 *****************************************
    c.m21 = a.m21 * b.m11 + a.m22 * b.m21 + a.m23 * b.m31
    c.m22 = a.m21 * b.m12 + a.m22 * b.m22 + a.m23 * b.m32
    c.m23 = a.m21 * b.m13 + a.m22 * b.m23 + a.m23 * b.m33
    ' C.m24 = 0 *****************************************
    c.m31 = a.m31 * b.m11 + a.m32 * b.m21 + a.m33 * b.m31
    c.m32 = a.m31 * b.m12 + a.m32 * b.m22 + a.m33 * b.m32
    c.m33 = a.m31 * b.m13 + a.m32 * b.m23 + a.m33 * b.m33
    ' C.m34 = 0 *****************************************
    c.m41 = a.m41 * b.m11 + a.m42 * b.m21 + a.m43 * b.m31 + b.m41
    c.m42 = a.m41 * b.m12 + a.m42 * b.m22 + a.m43 * b.m32 + b.m42
    c.m43 = a.m41 * b.m13 + a.m42 * b.m23 + a.m43 * b.m33 + b.m43
    ' C.m44 = 1 *****************************************
    m3MatrixMultiply = c
End Function
Public Function m3MatrixRandom(ByVal Min As Integer, ByVal Max As Integer) As m3Matrix
    Dim c As m3Matrix
    Randomize
    c.m11 = Min + Int(Rnd * (Max - Min))
    c.m12 = Min + Int(Rnd * (Max - Min))
    c.m13 = Min + Int(Rnd * (Max - Min))
    c.m21 = Min + Int(Rnd * (Max - Min))
    c.m22 = Min + Int(Rnd * (Max - Min))
    c.m23 = Min + Int(Rnd * (Max - Min))
    c.m31 = Min + Int(Rnd * (Max - Min))
    c.m32 = Min + Int(Rnd * (Max - Min))
    c.m33 = Min + Int(Rnd * (Max - Min))
    c.m41 = Min + Int(Rnd * (Max - Min))
    c.m42 = Min + Int(Rnd * (Max - Min))
    c.m43 = Min + Int(Rnd * (Max - Min))
    m3MatrixRandom = c
End Function
Public Function m3MatrixTranslate(ByVal Tx As Double, ByVal Ty As Double, ByVal Tz As Double) As m3Matrix
    
    Dim M As m3Matrix
    M.m11 = 1
    M.m12 = 0
    M.m13 = 0
    M.m21 = 0
    M.m22 = 1
    M.m23 = 0
    M.m31 = 0
    M.m32 = 0
    M.m33 = 1
    M.m41 = Tx
    M.m42 = Ty
    M.m43 = Tz
    '1
    m3MatrixTranslate = M
End Function
Public Function m3MatrixScale(ByVal Sx As Double, ByVal Sy As Double, ByVal Sz As Double) As m3Matrix
    Dim M As m3Matrix
    M.m11 = Sx
    M.m12 = 0
    M.m13 = 0
    M.m21 = 0
    M.m22 = Sy
    M.m23 = 0
    M.m31 = 0
    M.m32 = 0
    M.m33 = Sz
    M.m41 = 0
    M.m42 = 0
    M.m43 = 0
    '1
    m3MatrixScale = M
End Function
Public Function m3XRotate(ByVal theta As Double) As m3Matrix
    Dim M As m3Matrix
    Dim cs As Double
    Dim sn As Double
    cs = Cos(theta)
    sn = Sin(theta)
    M.m11 = 1
    M.m12 = 0
    M.m13 = 0
    M.m21 = 0
    M.m22 = cs
    M.m23 = sn
    M.m31 = 0
    M.m32 = -sn
    M.m33 = cs
    M.m41 = 0
    M.m42 = 0
    M.m43 = 0
    '1
    m3XRotate = M
End Function
Public Function m3YRotate(ByVal theta As Double) As m3Matrix
    Dim M As m3Matrix
    Dim cs As Double
    Dim sn As Double
    cs = Cos(theta)
    sn = Sin(theta)
    M.m11 = cs
    M.m12 = 0
    M.m13 = -sn
    M.m21 = 0
    M.m22 = 1
    M.m23 = 0
    M.m31 = sn
    M.m32 = 0
    M.m33 = cs
    M.m41 = 0
    M.m42 = 0
    M.m43 = 0
    '1
    m3YRotate = M
End Function
Public Function m3ZRotate(ByVal theta As Double) As m3Matrix
    Dim M As m3Matrix
    Dim cs As Double
    Dim sn As Double
    cs = Cos(theta)
    sn = Sin(theta)
    M.m11 = cs
    M.m12 = sn
    M.m13 = 0
    M.m21 = -sn
    M.m22 = cs
    M.m23 = 0
    M.m31 = 0
    M.m32 = 0
    M.m33 = 1
    M.m41 = 0
    M.m42 = 0
    M.m43 = 0
    '1
    m3ZRotate = M
End Function
Public Function m3Identity() As m3Matrix
    Dim M As m3Matrix
    M.m11 = 1
    M.m12 = 0
    M.m13 = 0
    M.m21 = 0
    M.m22 = 1
    M.m23 = 0
    M.m31 = 0
    M.m32 = 0
    M.m33 = 1
    M.m41 = 0
    M.m42 = 0
    M.m43 = 0
    '1
    m3Identity = M
End Function
