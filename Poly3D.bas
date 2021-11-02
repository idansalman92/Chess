Attribute VB_Name = "Poly3D"
Option Explicit
Public Const MAX_VERTS = 2000
Private Type POINTAPI
        x As Long
        y As Long
        
End Type

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private pt(MAX_VERTS) As POINTAPI
Public Type m3Poly
        verts() As m3Point
        nVerts As Integer
End Type
Public Function m3PolyInit(ByVal path As String) As m3Poly
    Dim P As m3Poly
    Dim fileNum As Integer
    Dim i As Integer
    fileNum = FreeFile
    Open path For Input As #fileNum
    Input #fileNum, P.nVerts
    ReDim P.verts(P.nVerts - 1) As m3Point
    For i = 0 To P.nVerts - 1
        Input #fileNum, P.verts(i).x, P.verts(i).y, P.verts(i).z
    Next i
    m3PolyInit = P
End Function
Public Function m3PolyToString(ByRef P As m3Poly) As String
    Dim S As String
    Dim i As Integer
    S = P.nVerts & vbCrLf
    For i = 0 To P.nVerts - 1
        S = S & m3PointToString(P.verts(i)) & vbCrLf
    Next i
    m3PolyToString = S
End Function

Public Sub m3PolyFill(ByRef obj As Object, ByRef P As m3Poly)
    Dim i As Integer
    Dim F As Double
    Dim dist As Double
    Dim Xorg As Integer
    Dim Yorg As Integer
    Xorg = -obj.ScaleLeft
    Yorg = obj.ScaleTop
    dist = Draw3D.m3GetDistance
    For i = 0 To P.nVerts - 1
        F = dist / (dist - P.verts(i).z)
        pt(i).x = Xorg + P.verts(i).x * F
        pt(i).y = Yorg - P.verts(i).y * F
    Next i
    Polygon obj.hdc, pt(0), P.nVerts
End Sub
Public Function m3polycenter(ByRef P As m3Poly) As m3Point
    Dim x As Double
    Dim y As Double
    Dim z As Double
    Dim i As Integer
    x = 0
    y = 0
    For i = 0 To P.nVerts - 1
        x = x + P.verts(i).x
        y = y + P.verts(i).y
        z = z + P.verts(i).z
    Next i
    m3polycenter = m3PointInit(x / P.nVerts, y / P.nVerts, z / P.nVerts)
End Function
Public Sub m3PolyApply(ByRef P As m3Poly, ByRef m As m3Matrix)
    Dim x As Double
    Dim y As Double
    Dim z As Double
    Dim i As Integer
    For i = 0 To P.nVerts - 1
        x = P.verts(i).x
        y = P.verts(i).y
        z = P.verts(i).z
        P.verts(i).x = x * m.m11 + y * m.m21 + z * m.m31 + m.m41
        P.verts(i).y = x * m.m12 + y * m.m22 + z * m.m32 + m.m42
        P.verts(i).z = x * m.m13 + y * m.m23 + z * m.m33 + m.m43
    Next i
End Sub

