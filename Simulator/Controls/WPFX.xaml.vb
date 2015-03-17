Imports Microsoft.VisualBasic
Imports System
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Input
Imports System.Windows.Media.Media3D

Public Class WPFX
    Public Structure Triangle
        Dim Point1 As Point3D
        Dim Point2 As Point3D
        Dim Point3 As Point3D
        Dim Normal As Vector3D
    End Structure


    Public myViewport3D As Viewport3D
    Public myModel3DGroup As Model3DGroup
    Public myModelVisual3D As ModelVisual3D

    Public myBackground3DGroup As Model3DGroup
    Public myBackgroundVisual3D As ModelVisual3D

    Public CurrentCamera As PerspectiveCamera


    Public MouseIsDown As Boolean
    Private lastMouseRotate As Point
    Public DeltaXRotate As Double
    Public DeltaYRotate As Double
    Public MyHost As System.Windows.Forms.Integration.ElementHost
    Public Const PreRenderObjects As Integer = 3
    Public CenterModel As Point3D
    Sub LoadGraph(_myHost As System.Windows.Forms.Integration.ElementHost)
        myViewport3D = New Viewport3D()

        myModel3DGroup = New Model3DGroup()
        myModelVisual3D = New ModelVisual3D()
        CurrentCamera = New PerspectiveCamera()

        CastCamera(New Point3D(0, 0, 10), New Vector3D(0, 0, -10), 60)

        myModel3DGroup.Children.Add(CreateLights(New Vector3D(1, 1, 1), Colors.Black))
        myModel3DGroup.Children.Add(CreateLights(New Vector3D(1, -1, -1), Colors.Black))
        myModel3DGroup.Children.Add(CreateLights(New Vector3D(-1, 1, -1), Colors.Black))
        myModelVisual3D.Content = myModel3DGroup
        myViewport3D.Children.Add(myModelVisual3D)


        myBackground3DGroup = New Model3DGroup()
        myBackgroundVisual3D = New ModelVisual3D()
        myBackground3DGroup.Children.Add(CreateLights(New Vector3D(1, 1, 1), Colors.White))
        myBackground3DGroup.Children.Add(CreateLights(New Vector3D(1, -1, -1), Colors.White))
        myBackground3DGroup.Children.Add(CreateLights(New Vector3D(-1, 1, -1), Colors.White))
        myBackgroundVisual3D.Content = myBackground3DGroup
        myViewport3D.Children.Add(myBackgroundVisual3D)
        AddBackground()

        myViewport3D.Camera = CurrentCamera
        myViewport3D.ClipToBounds = True
        Me.Content = myViewport3D
        MyHost = _myHost
    End Sub
    Sub AddBackground()
        Dim myMeshGeometry3D As New MeshGeometry3D()
        Dim myNormalCollection As New Vector3DCollection()
        Dim myPositionCollection As New Point3DCollection()
        Dim myTriangleIndicesCollection As New Int32Collection()
        Dim CurrentGeometryModel3D As New GeometryModel3D


        myPositionCollection.Add(New Point3D(-10, -10, 0))
        myPositionCollection.Add(New Point3D(-10, 10, 0))
        myPositionCollection.Add(New Point3D(10, 10, 0))
        myPositionCollection.Add(New Point3D(10, -10, 0))

        myTriangleIndicesCollection.Add(0)
        myTriangleIndicesCollection.Add(2)
        myTriangleIndicesCollection.Add(1)

        myTriangleIndicesCollection.Add(0)
        myTriangleIndicesCollection.Add(3)
        myTriangleIndicesCollection.Add(2)

        myNormalCollection.Add(New Vector3D(0, 0, -10))
        myNormalCollection.Add(New Vector3D(0, 0, -10))
        myNormalCollection.Add(New Vector3D(0, 0, -10))

        myNormalCollection.Add(New Vector3D(0, 0, -10))
        myNormalCollection.Add(New Vector3D(0, 0, -10))
        myNormalCollection.Add(New Vector3D(0, 0, -10))



        myMeshGeometry3D.TriangleIndices = myTriangleIndicesCollection
        myMeshGeometry3D.Normals = myNormalCollection
        myMeshGeometry3D.Positions = myPositionCollection

        ' ===========================================================================================

        CurrentGeometryModel3D.Geometry = myMeshGeometry3D


        Dim myMaterial As New DiffuseMaterial(New SolidColorBrush(Colors.White))
        CurrentGeometryModel3D.Material = myMaterial

        myBackground3DGroup.Children.Add(CurrentGeometryModel3D)

    End Sub

    Sub CastCamera(_position As Point3D, _direction As Vector3D, FOV As Integer)
        CurrentCamera.Position = _position
        CurrentCamera.LookDirection = _direction
        CurrentCamera.FieldOfView = FOV
    End Sub
    Function CreateLights(_direction As Vector3D, _color As Color) As DirectionalLight
        Dim out As New DirectionalLight
        out.Color = _color
        out.Direction = _direction
        Return out
    End Function

    Function CreateSpehereInSpaceTriangles(Center As Point3D, Radius As Double) As ArrayList
        Dim Trlist As ArrayList = CreateUnitSphereTriangleList(25)
        Dim OutList As New ArrayList
        For i As Integer = 0 To Trlist.Count - 1
            Dim P1, P2, P3 As Point3D
            Dim Tr As Triangle

            P1.X = CType(Trlist(i), Triangle).Point1.X * Radius + Center.X
            P1.Y = CType(Trlist(i), Triangle).Point1.Y * Radius + Center.Y
            P1.Z = CType(Trlist(i), Triangle).Point1.Z * Radius + Center.Z

            P2.X = CType(Trlist(i), Triangle).Point2.X * Radius + Center.X
            P2.Y = CType(Trlist(i), Triangle).Point2.Y * Radius + Center.Y
            P2.Z = CType(Trlist(i), Triangle).Point2.Z * Radius + Center.Z

            P3.X = CType(Trlist(i), Triangle).Point3.X * Radius + Center.X
            P3.Y = CType(Trlist(i), Triangle).Point3.Y * Radius + Center.Y
            P3.Z = CType(Trlist(i), Triangle).Point3.Z * Radius + Center.Z

            Tr.Point1 = P1
            Tr.Point2 = P2
            Tr.Point3 = P3
            Tr.Normal = CType(Trlist(i), Triangle).Normal
            OutList.Add(Tr)
        Next
        Return OutList

    End Function

    Sub CreateCube(ByRef Normals As Vector3DCollection, ByRef Positions As Point3DCollection, ByRef TriangleIndices As Int32Collection, G As Grid)
        Dim Lx As Double = G.Location.X - G.Size.X_Length / 2
        Dim LxPlus As Double = G.Location.X + G.Size.X_Length / 2
        Dim Ly As Double = G.Location.Y - G.Size.Y_Length / 2
        Dim LyPlus As Double = G.Location.Y + G.Size.Y_Length / 2
        Dim Lz As Double = G.Location.Z - G.Size.Z_Length / 2
        Dim LzPlus As Double = G.Location.Z + G.Size.Z_Length / 2

        Positions.Add(New Point3D(Lx, Ly, Lz))
        Positions.Add(New Point3D(Lx, Ly, LzPlus))
        Positions.Add(New Point3D(Lx, LyPlus, Lz))
        Positions.Add(New Point3D(Lx, LyPlus, LzPlus))
        Positions.Add(New Point3D(LxPlus, Ly, Lz))
        Positions.Add(New Point3D(LxPlus, Ly, LzPlus))
        Positions.Add(New Point3D(LxPlus, LyPlus, Lz))
        Positions.Add(New Point3D(LxPlus, LyPlus, LzPlus))

        ' Plane Lx
        TriangleIndices.Add(0)
        TriangleIndices.Add(0 + 1)
        TriangleIndices.Add(0 + 3)

        TriangleIndices.Add(0)
        TriangleIndices.Add(0 + 3)
        TriangleIndices.Add(0 + 2)

        Normals.Add(New Vector3D(-1, 0, 0))
        Normals.Add(New Vector3D(-1, 0, 0))
        Normals.Add(New Vector3D(-1, 0, 0))

        Normals.Add(New Vector3D(-1, 0, 0))
        Normals.Add(New Vector3D(-1, 0, 0))
        Normals.Add(New Vector3D(-1, 0, 0))


        ' Plane Lx+
        TriangleIndices.Add(0 + 4)
        TriangleIndices.Add(0 + 7)
        TriangleIndices.Add(0 + 5)

        TriangleIndices.Add(0 + 4)
        TriangleIndices.Add(0 + 6)
        TriangleIndices.Add(0 + 7)


        Normals.Add(New Vector3D(1, 0, 0))
        Normals.Add(New Vector3D(1, 0, 0))
        Normals.Add(New Vector3D(1, 0, 0))

        Normals.Add(New Vector3D(1, 0, 0))
        Normals.Add(New Vector3D(1, 0, 0))
        Normals.Add(New Vector3D(1, 0, 0))

        ' Plane Ly
        TriangleIndices.Add(0 + 1)
        TriangleIndices.Add(0)
        TriangleIndices.Add(0 + 5)

        TriangleIndices.Add(0)
        TriangleIndices.Add(0 + 4)
        TriangleIndices.Add(0 + 5)

        Normals.Add(New Vector3D(0, 1, 0))
        Normals.Add(New Vector3D(0, 1, 0))
        Normals.Add(New Vector3D(0, 1, 0))

        Normals.Add(New Vector3D(0, 1, 0))
        Normals.Add(New Vector3D(0, 1, 0))
        Normals.Add(New Vector3D(0, 1, 0))


        ' Plane Ly+
        TriangleIndices.Add(0 + 2)
        TriangleIndices.Add(0 + 3)
        TriangleIndices.Add(0 + 7)

        TriangleIndices.Add(0 + 6)
        TriangleIndices.Add(0 + 2)
        TriangleIndices.Add(0 + 7)

        Normals.Add(New Vector3D(0, 1, 0))
        Normals.Add(New Vector3D(0, 1, 0))
        Normals.Add(New Vector3D(0, 1, 0))

        Normals.Add(New Vector3D(0, 1, 0))
        Normals.Add(New Vector3D(0, 1, 0))
        Normals.Add(New Vector3D(0, 1, 0))



        ' Plane Lz
        TriangleIndices.Add(0)
        TriangleIndices.Add(0 + 2)
        TriangleIndices.Add(0 + 6)

        TriangleIndices.Add(0 + 4)
        TriangleIndices.Add(0)
        TriangleIndices.Add(0 + 6)

        Normals.Add(New Vector3D(0, 0, 1))
        Normals.Add(New Vector3D(0, 0, 1))
        Normals.Add(New Vector3D(0, 0, 1))

        Normals.Add(New Vector3D(0, 0, 1))
        Normals.Add(New Vector3D(0, 0, 1))
        Normals.Add(New Vector3D(0, 0, 1))

        ' Plane Lz+
        TriangleIndices.Add(0 + 3)
        TriangleIndices.Add(0 + 1)
        TriangleIndices.Add(0 + 7)

        TriangleIndices.Add(0 + 1)
        TriangleIndices.Add(0 + 5)
        TriangleIndices.Add(0 + 7)

        Normals.Add(New Vector3D(0, 0, -1))
        Normals.Add(New Vector3D(0, 0, -1))
        Normals.Add(New Vector3D(0, 0, -1))

        Normals.Add(New Vector3D(0, 0, -1))
        Normals.Add(New Vector3D(0, 0, -1))
        Normals.Add(New Vector3D(0, 0, -1))

    End Sub
    Sub AddReplaceSingleGrid(G As Grid, Optional ix As Integer = -1)
        Dim myMeshGeometry3D As New MeshGeometry3D()
        Dim myNormalCollection As New Vector3DCollection()
        Dim myPositionCollection As New Point3DCollection()
        Dim myTriangleIndicesCollection As New Int32Collection()
        Dim CurrentGeometryModel3D As New GeometryModel3D


        CreateCube(myNormalCollection, myPositionCollection, myTriangleIndicesCollection, G)

        myMeshGeometry3D.TriangleIndices = myTriangleIndicesCollection
        myMeshGeometry3D.Normals = myNormalCollection
        myMeshGeometry3D.Positions = myPositionCollection
        'myMeshGeometry3D.Freeze()

        ' ===========================================================================================

        CurrentGeometryModel3D.Geometry = myMeshGeometry3D


        Dim myMaterial As New DiffuseMaterial(New SolidColorBrush(G.Color))  ' myHorizontalGradient)
        CurrentGeometryModel3D.Material = myMaterial

        If ix = -1 Then
            myModel3DGroup.Children.Add(CurrentGeometryModel3D)
            PointToAverage(G.Location, True)

        Else
            CType(myModel3DGroup.Children(ix + PreRenderObjects), GeometryModel3D).Geometry = myMeshGeometry3D
            PointToAverage(G.Location, False)
        End If
    End Sub
    Sub PointToAverage(NewPoint As Grid_Location, Added As Boolean)
        CenterModel.X = (CenterModel.X * (myModel3DGroup.Children.Count - PreRenderObjects) + NewPoint.X) / (myModel3DGroup.Children.Count - PreRenderObjects + 1)
        CenterModel.Y = (CenterModel.Y * (myModel3DGroup.Children.Count - PreRenderObjects) + NewPoint.Y) / (myModel3DGroup.Children.Count - PreRenderObjects + 1)
        CenterModel.Z = (CenterModel.Z * (myModel3DGroup.Children.Count - PreRenderObjects) + NewPoint.Z) / (myModel3DGroup.Children.Count - PreRenderObjects + 1)
    End Sub


    Sub CastColors(Grids() As Grid)
        For ix As Integer = 0 To SizeOf(Grids) - 1
            Dim myMaterial As New DiffuseMaterial(New SolidColorBrush(Grids(ix).Color))
            CType(myModel3DGroup.Children(ix + PreRenderObjects), GeometryModel3D).Material = myMaterial
        Next


    End Sub



    Sub LoadGridList(Grids() As Grid)
        For i As Integer = PreRenderObjects - 1 To myModel3DGroup.Children.Count - 2
            myModel3DGroup.Children.RemoveAt(PreRenderObjects)
        Next
        CenterModel.X = 0 : CenterModel.Y = 0 : CenterModel.Z = 0
        Dim sumX, sumY, sumZ As Double

        For i As Integer = 0 To SizeOf(Grids) - 1

            Dim myMeshGeometry3D As New MeshGeometry3D()
            Dim myNormalCollection As New Vector3DCollection()
            Dim myPositionCollection As New Point3DCollection()
            Dim myTriangleIndicesCollection As New Int32Collection()
            Dim CurrentGeometryModel3D As New GeometryModel3D

            CreateCube(myNormalCollection, myPositionCollection, myTriangleIndicesCollection, Grids(i))

            myMeshGeometry3D.TriangleIndices = myTriangleIndicesCollection
            myMeshGeometry3D.Normals = myNormalCollection
            myMeshGeometry3D.Positions = myPositionCollection
            myMeshGeometry3D.Freeze()

            ' ===========================================================================================

            CurrentGeometryModel3D.Geometry = myMeshGeometry3D

            Dim myMaterial As New DiffuseMaterial(New SolidColorBrush(IIf(Grids(i).Color.ColorContext Is Nothing, Colors.Green, Grids(i).Color)))  ' myHorizontalGradient)
            CurrentGeometryModel3D.Material = myMaterial

            myModel3DGroup.Children.Add(CurrentGeometryModel3D)

            sumX += Grids(i).Location.X
            sumY += Grids(i).Location.Y
            sumZ += Grids(i).Location.Z

        Next

        CenterModel.X = sumX / Grids.Length
        CenterModel.Y = sumY / Grids.Length
        CenterModel.Z = sumZ / Grids.Length
    End Sub



    Private Sub MainWindow_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseDown
        MouseIsDown = True
        lastMouseRotate.X = -1
        lastMouseRotate.Y = -1



    End Sub

    Private Sub WPFX_MouseEnter(sender As Object, e As MouseEventArgs) Handles Me.MouseEnter
        MyHost.Focus()
    End Sub

    'Sub Rotate(deltaX As Double, deltaY As Double)
    '    If MouseIsDown Then
    '        Dim m_viewMatrix As Matrix3D
    '        Dim wth As Double = myViewport3D.ActualWidth
    '        Dim hgt As Double = myViewport3D.ActualHeight
    '        If lastMouseRotate.X <> -1 Then

    '            Dim aY As Double = 180 * (e.GetPosition(Me).X - lastMouseRotate.X) / wth
    '            Dim aX As Double = 180 * (e.GetPosition(Me).Y - lastMouseRotate.Y) / hgt

    '            DeltaXRotate += aX
    '            DeltaYRotate += aY

    '            m_viewMatrix.RotateAt(New Quaternion(New Vector3D(0, 1, 0), DeltaYRotate), CenterModel)
    '            m_viewMatrix.RotateAt(New Quaternion(New Vector3D(1, 0, 0), DeltaXRotate), CenterModel)


    '            Dim TR As New Transform3DGroup
    '            TR.Children.Add(New MatrixTransform3D(m_viewMatrix))
    '            myModel3DGroup.Transform = TR

    '        End If

    '        lastMouseRotate = e.GetPosition(Me)



    '    End If
    'End Sub
    Private Sub MainWindow_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        If MouseIsDown Then
            Dim m_viewMatrix As Matrix3D
            Dim wth As Double = myViewport3D.ActualWidth
            Dim hgt As Double = myViewport3D.ActualHeight

            If lastMouseRotate.X <> -1 Then

                Dim aY As Double = 180 * (e.GetPosition(Me).X - lastMouseRotate.X) / wth
                Dim aX As Double = 180 * (e.GetPosition(Me).Y - lastMouseRotate.Y) / hgt

                DeltaXRotate += aX
                DeltaYRotate += aY

                m_viewMatrix.RotateAt(New Quaternion(New Vector3D(0, 1, 0), DeltaYRotate), CenterModel)
                m_viewMatrix.RotateAt(New Quaternion(New Vector3D(1, 0, 0), DeltaXRotate), CenterModel)


                Dim TR As New Transform3DGroup
                TR.Children.Add(New MatrixTransform3D(m_viewMatrix))
                myModel3DGroup.Transform = TR

            End If

            lastMouseRotate = e.GetPosition(Me)



        End If
        



    End Sub

    Private Sub MainWindow_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseUp
        MouseIsDown = False
        DeltaXRotate = DeltaXRotate Mod 360
        DeltaYRotate = DeltaYRotate Mod 360
    End Sub

    Function CreateUnitSphereTriangleList(Density As Integer) As ArrayList
        Dim Count As Integer = Density * 2
        Dim Steps As Double = Math.PI / Density


        Dim PointList(Density)() As Point3D
        For i As Integer = 0 To Density
            ReDim PointList(i)(Count)
        Next

        For tita As Integer = 0 To Density
            Dim vtita As Double = tita * (Math.PI / Density)
            For nphi As Integer = -Density To Density
                Dim vphi As Double = nphi * (Math.PI / Density)
                PointList(tita)(nphi + Density).X = Math.Sin(vtita) * Math.Cos(vphi)
                PointList(tita)(nphi + Density).Y = Math.Sin(vtita) * Math.Sin(vphi)
                PointList(tita)(nphi + Density).Z = Math.Cos(vtita)
            Next
        Next

        Dim TriangleList As New ArrayList

        For n_tita As Integer = 1 To PointList.GetLength(0) - 1
            For n_phi As Integer = 0 To PointList(n_tita).GetLength(0) - 2

                Dim Triangle1, Triangle2 As Triangle

                Triangle1.Point1 = PointList(n_tita)(n_phi)
                Triangle1.Point2 = PointList(n_tita)(n_phi + 1)
                Triangle1.Point3 = PointList(n_tita - 1)(n_phi)
                Triangle1.Normal = New Vector3D(Triangle1.Point1.X, Triangle1.Point1.Y, Triangle1.Point1.Z)

                Triangle2.Point1 = PointList(n_tita)(n_phi + 1)
                Triangle2.Point2 = PointList(n_tita - 1)(n_phi + 1)
                Triangle2.Point3 = PointList(n_tita - 1)(n_phi)
                Triangle2.Normal = New Vector3D(Triangle1.Point1.X, Triangle1.Point1.Y, Triangle1.Point1.Z)

                TriangleList.Add(Triangle1)
                TriangleList.Add(Triangle2)
            Next
        Next

        Return TriangleList
    End Function


    Public Sub Rotate()


        If lastMouseRotate.X <> -1 And lastMouseRotate.Y <> -1 Then
            Dim deltaX As Double = 1
            DeltaXRotate += deltaX

            Dim myRotateTransform3D As New RotateTransform3D()
            Dim myAxisAngleRotation3d As New AxisAngleRotation3D()

            myAxisAngleRotation3d.Axis = New Vector3D(3, 3, 0)

            myAxisAngleRotation3d.Angle = DeltaXRotate
            myRotateTransform3D.Rotation = myAxisAngleRotation3d
            myModel3DGroup.Transform = myRotateTransform3D

        End If



        lastMouseRotate.X += 2

    End Sub

    Private Sub WPFX_MouseWheel(sender As Object, e As MouseWheelEventArgs) Handles Me.MouseWheel

        Dim Px As New Point3D
        Px = CurrentCamera.Position
        Px.Z += e.Delta / 100
        CurrentCamera.Position = Px
    End Sub


End Class
