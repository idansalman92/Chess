Attribute VB_Name = "Figures3D"
Public Function m3SoldierGet(ByVal R As Double, ByVal ppLayer As Integer) As m3solid
    Dim nLayers As Integer
    Dim S As m3solid
    Dim nVerts As Integer
    Dim nFaces As Integer
    Dim MZ As m3Matrix
    Dim MY As m3Matrix
    Dim P As m3Point
    Dim i As Integer
    Dim j As Integer
    Dim ind As Integer
    Dim ind1 As Integer
    nLayers = (ppLayer - 2) / 2
    nVerts = 2 + nLayers * ppLayer
    nFaces = 2 * ppLayer + (nLayers - 1) * ppLayer
    MZ = m3ZRotate(2 * PI / ppLayer)
    MY = m3YRotate(2 * PI / ppLayer)
    P = m3PointInit(0, R, 0)
    S.nVerts = nVerts
    ReDim S.verts(0 To nVerts - 1) As m3Point
    S.verts(0) = P
    S.verts(nVerts - 1) = m3PointInit(0, -R, 0)
    For i = 0 To nLayers - 1
        m3PointApply P, MZ
        For j = 0 To ppLayer - 1
            S.verts(1 + ppLayer * i + j) = P
            m3PointApply P, MY
        Next j
    Next i
    S.nFaces = nFaces
    ind = nVerts - 1 - ppLayer
    ReDim S.Faces(0 To nFaces - 1) As m3Face
    For i = 0 To ppLayer - 1
        S.Faces(i).nLinks = 3
        ReDim S.Faces(i).Links(0 To 2) As Integer
        S.Faces(i).Links(0) = 0
        S.Faces(i).Links(1) = i + 1
        S.Faces(i).Links(2) = (i + 1) Mod (ppLayer) + 1
        S.Faces(i + ppLayer).nLinks = 3
        ReDim S.Faces(i + ppLayer).Links(0 To 2) As Integer
        S.Faces(i + ppLayer).Links(0) = ind + i
        S.Faces(i + ppLayer).Links(1) = nVerts - 1
        S.Faces(i + ppLayer).Links(2) = ind + (i + 1) Mod ppLayer
    Next i
    
    For i = 0 To nLayers - 2
        For j = 0 To ppLayer - 1
            ind = 2 * ppLayer + i * ppLayer + j
            S.Faces(ind).nLinks = 4
            ReDim S.Faces(ind).Links(0 To 3) As Integer
            ind1 = 1 + i * ppLayer
            S.Faces(ind).Links(0) = ind1 + j
            S.Faces(ind).Links(1) = ind1 + ppLayer + j
            S.Faces(ind).Links(2) = ind1 + (j + 1) Mod ppLayer + ppLayer
            S.Faces(ind).Links(3) = ind1 + (j + 1) Mod ppLayer
        Next j
    Next i
        'S.nFaces = 2 * ppLayer
        m3SoldierGet = S
End Function
