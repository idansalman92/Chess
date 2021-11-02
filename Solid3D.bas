Attribute VB_Name = "Solid3D"
Option Explicit
Public Const MAX_VERTS = 20000
Private Type POINTAPI
        x As Long
        y As Long
        
End Type

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private pt(MAX_VERTS) As POINTAPI
Private pt1(MAX_VERTS) As m3Point

Public Type m3Face
    nLinks As Integer
    Links() As Integer
End Type
Public Type m3solid
    nVerts As Integer
    verts() As m3Point
    nFaces As Integer
    Faces() As m3Face
End Type
Public Sub m3SolidFill(ByRef Obj As Object, ByRef s As m3solid)
    Dim i As Integer
    Dim j As Integer
    Dim f As Double
    Dim dist As Double
    Dim xorg As Integer
    Dim yorg As Integer
    Dim v As m3Vector
    Dim u As m3Vector
    Dim n As m3Vector
    Dim p1 As Integer
    Dim P2 As Integer
    Dim p3 As Integer
    Dim nVerts As Integer
    xorg = -Obj.ScaleLeft
    yorg = Obj.ScaleTop
    dist = Draw3D.m3GetDistance
    For i = 0 To s.nVerts - 1
        f = dist / (dist - s.verts(i).z)
        pt1(i).x = s.verts(i).x * f
        pt1(i).y = s.verts(i).y * f
        pt1(i).z = s.verts(i).z
    Next i
    For i = 0 To s.nFaces - 1
        p1 = s.Faces(i).Links(0)
        P2 = s.Faces(i).Links(1)
        p3 = s.Faces(i).Links(2)
        v = m3VectorInit(pt1(p1), pt1(P2))
        u = m3VectorInit(pt1(P2), pt1(p3))
        n = m3VectorCross(v, u)
        If n.z >= 0 Then
            nVerts = s.Faces(i).nLinks
            For j = 0 To nVerts - 1
                p1 = s.Faces(i).Links(j)
                pt(j).x = xorg + pt1(p1).x
                pt(j).y = yorg - pt1(p1).y
            Next j
            Polygon Obj.hdc, pt(0), nVerts
        End If
    Next i
    
    
End Sub
Public Sub m3SolidFillShading(ByRef Obj As Object, ByRef s As m3solid)
    Dim i As Integer
    Dim j As Integer
    Dim ind As Integer
    Dim f As Double
    Dim v As m3Vector
    Dim u As m3Vector
    Dim n As m3Vector
    Dim nVerts As Integer
    Dim R As Byte
    Dim G As Byte
    Dim b As Byte
    Dim maxRGB As Byte
    Dim dot As Double
    Dim dotMax As Double
    Dim color As Long
    Dim colorTo As Long
    Dim borderColor As Long
    Dim xorg As Integer
    Dim yorg As Integer
    Dim dist As Double
    Dim lightVector As m3Vector
    Dim p1 As Integer
    Dim P2 As Integer
    Dim p3 As Integer
    xorg = -Obj.ScaleLeft
    yorg = Obj.ScaleTop
    dist = Draw3D.m3GetDistance
    lightVector = Draw3D.m3GetLightVector
    color = Obj.FillColor
    borderColor = Obj.ForeColor
    ' convert color to RGB components
    R = &HFF& And color
    G = (&HFF00& And color) \ &H100&
    b = (&HFF0000 And color) \ &H10000
    maxRGB = R
    If G > maxRGB Then maxRGB = G
    If b > maxRGB Then maxRGB = b
    dotMax = maxRGB / 255
    dotMax = dotMax + AmbFactor
    If dotMax * maxRGB > 255 Then dotMax = 255 / maxRGB
    For i = 0 To s.nVerts - 1
        f = dist / (dist - s.verts(i).z)
        pt1(i).x = s.verts(i).x * f
        pt1(i).y = s.verts(i).y * f
        pt1(i).z = s.verts(i).z

    Next i
    
    
    For i = 0 To s.nFaces - 1
        p1 = s.Faces(i).Links(0)
        P2 = s.Faces(i).Links(1)
        p3 = s.Faces(i).Links(2)
        v = m3VectorInit(pt1(p1), pt1(P2))
        u = m3VectorInit(pt1(P2), pt1(p3))
        n = m3VectorCross(v, u)
        If n.z >= 0 Then
            For j = 0 To s.Faces(i).nLinks - 1
                ind = s.Faces(i).Links(j)
                pt(j).x = xorg + pt1(ind).x
                pt(j).y = yorg - pt1(ind).y
            Next j
           
            m3VectorSetLen 1, n
            dot = lightVector.x * n.x + lightVector.y * n.y + lightVector.z * n.z
            If dot > 0.0000001 Then
                dot = dot + AmbFactor
                                             
            Else
                dot = AmbFactor
            End If
            If dot > dotMax Then dot = dotMax
            colorTo = RGB(dot * R, dot * G, dot * b)
            Obj.FillColor = colorTo
            Obj.ForeColor = colorTo
            Polygon Obj.hdc, pt(0), s.Faces(i).nLinks
        End If
    Next i
    
    Obj.FillColor = color
    Obj.ForeColor = borderColor
End Sub
Public Function m3GetPrizm(ByVal R As Double, ByVal h As Double, ByVal ppbase As Integer) As m3solid
    Dim s As m3solid
    Dim i As Integer
    Dim p As m3Point
    Dim M As m3Matrix
    s.nVerts = 2 * ppbase
    ReDim s.verts(0 To s.nVerts - 1) As m3Point
    s.nFaces = 2 + ppbase
    ReDim s.Faces(0 To s.nFaces - 1) As m3Face
    p = m3PointInit(R, 0, 0)
    M = Matrix3D.m3YRotate(2 * PI / ppbase)
    s.Faces(0).nLinks = ppbase
    ReDim s.Faces(0).Links(0 To ppbase - 1)
    s.Faces(1).nLinks = ppbase
    ReDim s.Faces(1).Links(0 To ppbase - 1)
    For i = 0 To ppbase - 1
        s.verts(i) = p
        s.verts(i + ppbase) = p
        s.verts(i + ppbase).y = h
        m3PointApply p, M
        s.Faces(0).Links(i) = ppbase - 1 - i
        s.Faces(1).Links(i) = ppbase + i
        s.Faces(i + 2).nLinks = 4
        ReDim s.Faces(i + 2).Links(0 To 3) As Integer
        s.Faces(i + 2).Links(0) = i
        s.Faces(i + 2).Links(1) = (i + 1) Mod ppbase
        s.Faces(i + 2).Links(2) = s.Faces(i + 2).Links(1) + ppbase
        s.Faces(i + 2).Links(3) = i + ppbase
    Next i
    m3GetPrizm = s
End Function
Public Function m3GetPyramid(ByVal R As Double, ByVal h As Double, ByVal ppbase As Integer) As m3solid
    Dim s As m3solid
    Dim i As Integer
    Dim p As m3Point
    Dim M As m3Matrix
    s.nVerts = ppbase + 1
    ReDim s.verts(0 To s.nVerts - 1) As m3Point
    s.nFaces = ppbase + 1
    ReDim s.Faces(0 To s.nFaces - 1) As m3Face
    p = m3PointInit(R, 0, 0)
    M = Matrix3D.m3YRotate(2 * PI / ppbase)
    s.Faces(0).nLinks = ppbase
    ReDim s.Faces(0).Links(0 To ppbase - 1) As Integer
    For i = 0 To ppbase - 1
        s.verts(i) = p
        m3PointApply p, M
        s.Faces(0).Links(i) = ppbase - 1 - i
        s.Faces(i + 1).nLinks = 3
        ReDim s.Faces(i + 1).Links(0 To 2) As Integer
        s.Faces(i + 1).Links(0) = i
        s.Faces(i + 1).Links(1) = (i + 1) Mod ppbase
        s.Faces(i + 1).Links(2) = ppbase
        
    Next i
    s.verts(ppbase) = m3PointInit(0, h, 0)
    
    m3GetPyramid = s
End Function
Public Function m3GetCube(ByVal w As Double) As m3solid
    Dim s As m3solid
    s.nVerts = 8
    ReDim s.verts(7) As m3Point
    s.verts(0) = m3PointInit(0, 0, w)
    s.verts(1) = m3PointInit(w, 0, w)
    s.verts(2) = m3PointInit(w, w, w)
    s.verts(3) = m3PointInit(0, w, w)
    s.verts(4) = m3PointInit(0, 0, 0)
    s.verts(5) = m3PointInit(w, 0, 0)
    s.verts(6) = m3PointInit(w, w, 0)
    s.verts(7) = m3PointInit(0, w, 0)
    s.nFaces = 6
    ReDim s.Faces(5) As m3Face
    
    'FronFace
    s.Faces(0).nLinks = 4
    ReDim s.Faces(0).Links(3) As Integer
    s.Faces(0).Links(0) = 0
    s.Faces(0).Links(1) = 1
    s.Faces(0).Links(2) = 2
    s.Faces(0).Links(3) = 3
    
    'Back Faces
    s.Faces(1).nLinks = 4
    ReDim s.Faces(1).Links(3) As Integer
    s.Faces(1).Links(0) = 4
    s.Faces(1).Links(1) = 7
    s.Faces(1).Links(2) = 6
    s.Faces(1).Links(3) = 5
    
    'UpFaces
    s.Faces(2).nLinks = 4
    ReDim s.Faces(2).Links(3) As Integer
    s.Faces(2).Links(0) = 3
    s.Faces(2).Links(1) = 2
    s.Faces(2).Links(2) = 6
    s.Faces(2).Links(3) = 7
    
    'DownFaces
    s.Faces(3).nLinks = 4
    ReDim s.Faces(3).Links(3) As Integer
    s.Faces(3).Links(0) = 0
    s.Faces(3).Links(1) = 4
    s.Faces(3).Links(2) = 5
    s.Faces(3).Links(3) = 1
    
    'LeftFace
    s.Faces(4).nLinks = 4
    ReDim s.Faces(4).Links(3) As Integer
    s.Faces(4).Links(0) = 0
    s.Faces(4).Links(1) = 3
    s.Faces(4).Links(2) = 7
    s.Faces(4).Links(3) = 4
    
    'RightFace
    s.Faces(5).nLinks = 4
    ReDim s.Faces(5).Links(3) As Integer
    s.Faces(5).Links(0) = 5
    s.Faces(5).Links(1) = 6
    s.Faces(5).Links(2) = 2
    s.Faces(5).Links(3) = 1
    
    m3GetCube = s
End Function
Public Sub m3solidApply(ByRef s As m3solid, ByRef M As m3Matrix)
    Dim x As Double
    Dim y As Double
    Dim z As Double
    Dim i As Integer
    For i = 0 To s.nVerts - 1
        x = s.verts(i).x
        y = s.verts(i).y
        z = s.verts(i).z
        s.verts(i).x = x * M.m11 + y * M.m21 + z * M.m31 + M.m41
        s.verts(i).y = x * M.m12 + y * M.m22 + z * M.m32 + M.m42
        s.verts(i).z = x * M.m13 + y * M.m23 + z * M.m33 + M.m43
    Next i
End Sub
Public Function m3solidcenter(ByRef s As m3solid) As m3Point
    Dim x As Double
    Dim y As Double
    Dim z As Double
    Dim i As Integer
    x = 0
    y = 0
    For i = 0 To s.nVerts - 1
        x = x + s.verts(i).x
        y = y + s.verts(i).y
        z = z + s.verts(i).z
    Next i
    m3solidcenter = m3PointInit(x / s.nVerts, y / s.nVerts, z / s.nVerts)
End Function
Public Function m3solidToString(ByRef s As m3solid) As String
   Dim str As String
    Dim i As Integer
    Dim j As Integer
    str = s.nVerts & vbTab & s.nFaces & vbCrLf
    For i = 0 To s.nVerts - 1
        str = str & m3PointToString(s.verts(i)) & vbCrLf
    Next i
    For i = 0 To s.nFaces - 1
        str = str & s.Faces(i).nLinks
        For j = 0 To s.Faces(i).nLinks - 1
            str = str & vbTab & s.Faces(i).Links(j)
        Next j
        str = str & vbCrLf
    Next i
    m3solidToString = str
End Function
Public Function SphareInit(ByVal R As Double, ByVal ppLayer As Integer) As m3solid
    Dim nLayers As Integer
    Dim s As m3solid
    Dim nVerts As Integer
    Dim nFaces As Integer
    Dim MZ As m3Matrix
    Dim MY As m3Matrix
    Dim p As m3Point
    Dim i As Integer
    Dim j As Integer
    Dim ind As Integer
    Dim ind1 As Integer
    nLayers = (ppLayer - 2) / 2
    nVerts = 2 + nLayers * ppLayer
    nFaces = 2 * ppLayer + (nLayers - 1) * ppLayer
    MZ = m3ZRotate(2 * PI / ppLayer)
    MY = m3YRotate(2 * PI / ppLayer)
    p = m3PointInit(0, R, 0)
    s.nVerts = nVerts
    ReDim s.verts(0 To nVerts - 1) As m3Point
    s.verts(0) = p
    s.verts(nVerts - 1) = m3PointInit(0, -R, 0)
    For i = 0 To nLayers - 1
        m3PointApply p, MZ
        For j = 0 To ppLayer - 1
            s.verts(1 + ppLayer * i + j) = p
            m3PointApply p, MY
        Next j
    Next i
    s.nFaces = nFaces
    ind = nVerts - 1 - ppLayer
    ReDim s.Faces(0 To nFaces - 1) As m3Face
    For i = 0 To ppLayer - 1
        s.Faces(i).nLinks = 3
        ReDim s.Faces(i).Links(0 To 2) As Integer
        s.Faces(i).Links(0) = 0
        s.Faces(i).Links(1) = i + 1
        s.Faces(i).Links(2) = (i + 1) Mod (ppLayer) + 1
        s.Faces(i + ppLayer).nLinks = 3
        ReDim s.Faces(i + ppLayer).Links(0 To 2) As Integer
        s.Faces(i + ppLayer).Links(0) = ind + i
        s.Faces(i + ppLayer).Links(1) = nVerts - 1
        s.Faces(i + ppLayer).Links(2) = ind + (i + 1) Mod ppLayer
    Next i
    
    For i = 0 To nLayers - 2
        For j = 0 To ppLayer - 1
            ind = 2 * ppLayer + i * ppLayer + j
            s.Faces(ind).nLinks = 4
            ReDim s.Faces(ind).Links(0 To 3) As Integer
            ind1 = 1 + i * ppLayer
            s.Faces(ind).Links(0) = ind1 + j
            s.Faces(ind).Links(1) = ind1 + ppLayer + j
            s.Faces(ind).Links(2) = ind1 + (j + 1) Mod ppLayer + ppLayer
            s.Faces(ind).Links(3) = ind1 + (j + 1) Mod ppLayer
        Next j
    Next i
        'S.nFaces = 2 * ppLayer
        SphareInit = s
End Function
