Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class TDC

    'core objects
    Dim MouseIsDown As Boolean
    Dim CameraZ As Double = -5

    Private D3DRoot As Manager
    Private D3DDev As Device
    Private D3DHelp As D3DX

    Private matCube1 As Matrix
    Private vbGrids As VertexBuffer
    Private vCube As CustomVertex.PositionNormalColored()

    Private vbGrids_coords As VertexBuffer
    Private vCube_coords As CustomVertex.PositionNormalColored()

    Private vbText As VertexBuffer
    Private Max, Min, Center As Grid_Location
    Private Started As Boolean

    Public Sub Start()
        Started = True
        Do While Started
            FrameRender()
            Application.DoEvents()
        Loop

    End Sub

    Dim d3dPP As PresentParameters


    Public Sub New()
        NewPeform()
    End Sub

    Sub NewPeform()
        d3dPP = New PresentParameters()
        With d3dPP
            .Windowed = True
            .SwapEffect = SwapEffect.Discard
            .BackBufferCount = 1
            .BackBufferFormat = D3DRoot.Adapters(0).CurrentDisplayMode.Format
            .BackBufferWidth = Me.ClientSize.Width()
            .BackBufferHeight = Me.ClientSize.Height()
            .PresentationInterval = PresentInterval.Immediate
        End With
        initialiseDevice(CType(Me, System.Windows.Forms.Control), d3dPP)
    End Sub

    Private Sub initialiseDevice(ByRef Target As System.Windows.Forms.Control, ByVal win As PresentParameters)
        Dim D3DCaps As Caps
        Dim DevCreate As Integer
        win.AutoDepthStencilFormat = DepthFormat.D16
        win.EnableAutoDepthStencil = True
        D3DCaps = D3DRoot.GetDeviceCaps(0, DeviceType.Hardware)

        If D3DCaps.DeviceCaps.SupportsHardwareTransformAndLight Then
            DevCreate = CreateFlags.HardwareVertexProcessing Or CreateFlags.FpuPreserve
        Else
            DevCreate = CreateFlags.SoftwareVertexProcessing Or CreateFlags.FpuPreserve
        End If

        Try
            D3DDev = New Device(0, DeviceType.Hardware, Target, DevCreate, win)
        Catch ex As Exception
            MsgBox("Error")
        End Try
        'D3DDev.Reset(win)

        With D3DDev
            .RenderState.ZBufferWriteEnable = True
            .RenderState.ZBufferEnable = True
            .RenderState.Lighting = True

            .RenderState.AlphaBlendEnable = True
            .RenderState.SourceBlend = Blend.SourceAlpha
            .RenderState.DestinationBlend = Blend.InvSourceAlpha

            .Transform.View = Matrix.LookAtLH(New Vector3(0, 0, -15), New Vector3(0, 0, 0), New Vector3(0, 1, 0))
            .Transform.Projection = Matrix.PerspectiveFovLH(Math.PI / 4, 4 / 3, 1, 100)
            .Transform.World = Matrix.Identity()


            D3DDev.Lights(0).Type = LightType.Directional
            D3DDev.Lights(0).Diffuse = Color.White
            D3DDev.Lights(0).Position = New Vector3(Max.X, Max.Y, Min.Z)
            D3DDev.Lights(0).Direction = New Vector3(-1, -1, 1)
            D3DDev.Lights(0).Update()
            D3DDev.Lights(0).Enabled = True

            D3DDev.Lights(1).Type = LightType.Directional
            D3DDev.Lights(1).Diffuse = Color.White
            D3DDev.Lights(1).Position = New Vector3(Min.X, Min.Y, Min.Z)
            D3DDev.Lights(1).Direction = New Vector3(0, 1, 0.5)
            D3DDev.Lights(1).Update()
            D3DDev.Lights(1).Enabled = True
        End With

    End Sub


    Sub DrawWells()
        Dim Coe As Double = ((Max.X - Min.X) * (Max.Y - Min.Y) * (Max.Z - Min.Z)) ^ (1 / 3)
        For i As Integer = 0 To SizeOf(Main.SimulationData.Model.Wells) - 1
            For j As Integer = 0 To SizeOf(Main.SimulationData.Model.Wells(i).Paths) - 1
                With Main.SimulationData.Model.Wells(i).Paths(j)
                    Dim radius As Double = .Diameter / (2 * Coe)
                    Dim Length As Double = Math.Sqrt((.X1 - .X2) ^ 2 + (.Y1 - .Y2) ^ 2 + (.Z1 - .Z2) ^ 2) / Coe
                    Dim PathMesh As Mesh = Mesh.Cylinder(D3DDev, radius, radius, Coe, 10, 3)

                    Dim matt As New Material
                    matt.Diffuse = Color.Black
                    D3DDev.Material = matt
                    'D3DDev.Transform.World = Matrix.Translation(New Vector3((.Location.X - Min.X) / (Max.X - Min.X) - .Size.X_Length / Coe, .Location.Y / Coe, .Location.Z / Coe)) * matCube
                    PathMesh.DrawSubset(0)
                End With

                

                With Main.SimulationData.Model.Grids(i)
                    'D3DDev.Transform.World = (Matrix.Translation(New Vector3(.Location.X * 5 - (.Size.X_Length * 5) / 3, .Location.Y * 5, .Location.Z * 5)) * matCube1) * Matrix.Scaling(0.2, 0.2, 0.2)

                End With



            Next
        Next


    End Sub
    Sub DrawTexts()
        'For i As Integer = 0 To SizeOf(Main.SimulationData.Model.Grids) - 1
        '    Main.SimulationData.Model.Grids(i).InsideText = "A"
        '    If Not String.IsNullOrEmpty(Main.SimulationData.Model.Grids(i).InsideText) Then
        '        Dim TextMesh As Mesh = Mesh.TextFromFont(D3DDev, New System.Drawing.Font("Arial", 8, FontStyle.Bold), Main.SimulationData.Model.Grids(i).InsideText, 0.001, 0.04)
        '        Dim matt As New Material
        '        matt.Diffuse = Color.Black
        '        D3DDev.Material = matt
        '        Dim Coe As Double = ((Max.X - Min.X) * (Max.Y - Min.Y) * (Max.Z - Min.Z)) ^ (1 / 3)
        '        With Main.SimulationData.Model.Grids(i)
        '            'D3DDev.Transform.World = (Matrix.Translation(New Vector3(.Location.X * 5 - (.Size.X_Length * 5) / 3, .Location.Y * 5, .Location.Z * 5)) * matCube1) * Matrix.Scaling(0.2, 0.2, 0.2)
        '            'D3DDev.Transform.World = Matrix.Translation(New Vector3((.Location.X - Min.X) / (Max.X - Min.X) - .Size.X_Length / Coe, .Location.Y / Coe, .Location.Z / Coe)) * matCube1
        '        End With

        '        TextMesh.DrawSubset(0)
        '    End If
        'Next

    End Sub

    Private Function MatrixTranslation(ByVal x As Single, ByVal y As Single, ByVal z As Single) As Matrix
        Dim mat As Matrix = New Matrix

        mat.M11 = 1.0F : mat.M12 = 0.0F : mat.M13 = 0.0F : mat.M14 = 0.0F
        mat.M21 = 0.0F : mat.M22 = 1.0F : mat.M23 = 0.0F : mat.M24 = 0.0F
        mat.M31 = 0.0F : mat.M32 = 0.0F : mat.M33 = 1.0F : mat.M34 = 0.0F
        mat.M41 = x : mat.M42 = y : mat.M43 = z : mat.M44 = 1.0F

        Return mat
    End Function

    Public Sub FrameRender()

        matCube1 = MatrixTranslation(-(Max.X + Min.X) / 2, -(Max.Y + Min.Y) / 2, -(Max.Z + Min.Z) / 2)
        matCube1 = Matrix.Multiply(matCube1, Matrix.Scaling(1, 1, 1))
        matCube1 = Matrix.Multiply(matCube1, Matrix.RotationX(Main.SimulationData.Graphics.X_Angle * (Math.PI / 180)))
        matCube1 = Matrix.Multiply(matCube1, Matrix.RotationY(Main.SimulationData.Graphics.Y_Angle * (Math.PI / 180)))

        If D3DDev Is Nothing Then Exit Sub

        D3DDev.Clear(ClearFlags.Target Or ClearFlags.ZBuffer, Color.White, 1.0F, 0)
        D3DDev.BeginScene()
        'DrawWells()


        D3DDev.Transform.World = matCube1


        D3DDev.VertexFormat = CustomVertex.PositionNormalColored.Format
        D3DDev.Transform.View = Matrix.LookAtLH(New Vector3((Max.X + Min.X) / 2, (Max.Y + Min.Y) / 2, CameraZ), New Vector3(0, 0, 0), New Vector3(0, 1, 0))

        If SizeOf(Main.SimulationData.Model.Grids) > 0 Then

            'D3DDev .DrawIndexedPrimitives (PrimitiveType.TriangleList , 0,0,SizeOf(Main.SimulationData.Model.Grids) * 24
            D3DDev.SetStreamSource(0, vbGrids, 0)
            Try
                D3DDev.DrawPrimitives(PrimitiveType.TriangleList, 0, SizeOf(Main.SimulationData.Model.Grids) * 24)
            Catch ex As Exception
                Dim h As Int16
                h = 0

            End Try


            

        End If


        'D3DDev.SetStreamSource(0, vbGrids_coords, 0)
        'D3DDev.VertexFormat = CustomVertex.PositionNormalColored.Format
        ''D3DDev.Transform.View = Matrix.LookAtLH(New Vector3((Max.X + Min.X) / 2, (Max.Y + Min.Y) / 2, CameraZ), New Vector3((Max.X + Min.X) / 2, (Max.Y + Min.Y) / 2, (Max.Z + Min.Z) / 2), New Vector3(0, 1, 0))
        'D3DDev.RenderState.Ambient = Color.White
        'D3DDev.DrawPrimitives(PrimitiveType.TriangleList, 0, 1)



        D3DDev.EndScene()
        D3DDev.Present()

    End Sub
    Sub Terminate()
        D3DDev.Dispose()
        System.GC.SuppressFinalize(D3DDev)
        System.GC.SuppressFinalize(Me)

    End Sub


    Sub CreateCoords()
        vbGrids_coords = New VertexBuffer(GetType(CustomVertex.PositionNormalColored), 3, D3DDev, 0, CustomVertex.PositionNormalColored.Format, Pool.Managed)
        vCube_coords = CType(vbGrids_coords.Lock(0, 0), CustomVertex.PositionNormalColored())

        Dim C As Color = Color.Black
        vCube_coords(0) = New CustomVertex.PositionNormalColored(-1, 0, 0, 0, 0, -1, C.ToArgb)
        vCube_coords(1) = New CustomVertex.PositionNormalColored(0, 1, 0, 0, 0, -1, C.ToArgb)
        vCube_coords(2) = New CustomVertex.PositionNormalColored(-1, 0, 0, 0, 0, -1, C.ToArgb)

        vbGrids_coords.Unlock()

    End Sub

    Public Sub LoadGeometry(ByRef SimulationData As Simulation, Optional ByVal Colored As Boolean = True, Optional ByVal SelectedIndex As Integer = -1)
        CreateCoords()
        If SizeOf(SimulationData.Model.Grids) = 0 Then Exit Sub

        Try
            Main.SimulationData = Main.SimulationData
            vbGrids = New VertexBuffer(GetType(CustomVertex.PositionNormalColored), SizeOf(SimulationData.Model.Grids) * 72, D3DDev, 0, CustomVertex.PositionNormalColored.Format, Pool.Managed)
            vCube = CType(vbGrids.Lock(0, 0), CustomVertex.PositionNormalColored())

            Max.X = SimulationData.Model.Grids(0).Location.X
            Max.Y = SimulationData.Model.Grids(0).Location.Y
            Max.Z = SimulationData.Model.Grids(0).Location.Z

            Min.X = Max.X
            Min.Y = Max.Y
            Min.Z = Max.Z

            For i As Integer = 0 To SizeOf(SimulationData.Model.Grids) - 1
                Dim x1 As Double = (SimulationData.Model.Grids(i).Location.X - SimulationData.Model.Grids(i).Size.X_Length / 2)
                Dim x2 As Double = (SimulationData.Model.Grids(i).Location.X + SimulationData.Model.Grids(i).Size.X_Length / 2)

                Dim y1 As Double = (SimulationData.Model.Grids(i).Location.Z - SimulationData.Model.Grids(i).Size.Z_Length / 2)
                Dim y2 As Double = (SimulationData.Model.Grids(i).Location.Z + SimulationData.Model.Grids(i).Size.Z_Length / 2)

                Dim z1 As Double = (SimulationData.Model.Grids(i).Location.Y - SimulationData.Model.Grids(i).Size.Y_Length / 2)
                Dim z2 As Double = (SimulationData.Model.Grids(i).Location.Y + SimulationData.Model.Grids(i).Size.Y_Length / 2)


                If x1 > Max.X Then Max.X = x1
                If x2 > Max.X Then Max.X = x2
                If x1 < Min.X Then Min.X = x1
                If x2 < Min.X Then Min.X = x2

                If z1 > Max.Z Then Max.Z = z1
                If z2 > Max.Z Then Max.Z = z2
                If z1 < Min.Z Then Min.Z = z1
                If z2 < Min.Z Then Min.Z = z2

                If y1 > Max.Y Then Max.Y = y1
                If y2 > Max.Y Then Max.Y = y2
                If y1 < Min.Y Then Min.Y = y1
                If y2 < Min.Y Then Min.Y = y2

            Next


            Dim Coe As Double = ((Max.X - Min.X) * (Max.Y - Min.Y) * (Max.Z - Min.Z)) ^ (1 / 3)
            Max.X = Max.X / Coe
            Min.X = Min.X / Coe

            Max.Y = Max.Y / Coe
            Min.Y = Min.Y / Coe

            Max.Z = Max.Z / Coe
            Min.Z = Min.Z / Coe


            For i As Integer = 0 To SizeOf(SimulationData.Model.Grids) - 1

                Dim x1 As Double = (SimulationData.Model.Grids(i).Location.X - SimulationData.Model.Grids(i).Size.X_Length / 2) / Coe
                Dim x2 As Double = (SimulationData.Model.Grids(i).Location.X + SimulationData.Model.Grids(i).Size.X_Length / 2) / Coe

                Dim y1 As Double = (SimulationData.Model.Grids(i).Location.Z - SimulationData.Model.Grids(i).Size.Z_Length / 2) / Coe
                Dim y2 As Double = (SimulationData.Model.Grids(i).Location.Z + SimulationData.Model.Grids(i).Size.Z_Length / 2) / Coe

                Dim z1 As Double = (SimulationData.Model.Grids(i).Location.Y - SimulationData.Model.Grids(i).Size.Y_Length / 2) / Coe
                Dim z2 As Double = (SimulationData.Model.Grids(i).Location.Y + SimulationData.Model.Grids(i).Size.Y_Length / 2) / Coe




                ' Z =========================
                ' Z Constant Forward \ Reverse
                Dim DefaultColor As Color = Color.FromArgb(70, 0, 0, 255)
                Dim SelectedColor As Color = Color.FromArgb(255, 255, 0, 0)

                vCube(i * 72) = New CustomVertex.PositionNormalColored(x1, y1, z1, 0, 0, -1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 1) = New CustomVertex.PositionNormalColored(x2, y2, z1, 0, 0, -1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 2) = New CustomVertex.PositionNormalColored(x2, y1, z1, 0, 0, -1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))

                vCube(i * 72 + 3) = New CustomVertex.PositionNormalColored(x1, y1, z1, 0, 0, -1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 4) = New CustomVertex.PositionNormalColored(x1, y2, z1, 0, 0, -1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 5) = New CustomVertex.PositionNormalColored(x2, y2, z1, 0, 0, -1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))


                vCube(i * 72 + 6) = New CustomVertex.PositionNormalColored(x1, y1, z1, 0, 0, 1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 7) = New CustomVertex.PositionNormalColored(x2, y1, z1, 0, 0, 1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 8) = New CustomVertex.PositionNormalColored(x2, y2, z1, 0, 0, 1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))

                vCube(i * 72 + 9) = New CustomVertex.PositionNormalColored(x1, y1, z1, 0, 0, 1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 10) = New CustomVertex.PositionNormalColored(x2, y2, z1, 0, 0, 1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 11) = New CustomVertex.PositionNormalColored(x1, y2, z1, 0, 0, 1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))

                vCube(i * 72 + 12) = New CustomVertex.PositionNormalColored(x1, y1, z2, 0, 0, -1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 13) = New CustomVertex.PositionNormalColored(x2, y2, z2, 0, 0, -1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 14) = New CustomVertex.PositionNormalColored(x2, y1, z2, 0, 0, -1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))

                vCube(i * 72 + 15) = New CustomVertex.PositionNormalColored(x1, y1, z2, 0, 0, -1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 16) = New CustomVertex.PositionNormalColored(x1, y2, z2, 0, 0, -1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 17) = New CustomVertex.PositionNormalColored(x2, y2, z2, 0, 0, -1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))


                vCube(i * 72 + 18) = New CustomVertex.PositionNormalColored(x1, y1, z2, 0, 0, 1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 19) = New CustomVertex.PositionNormalColored(x2, y1, z2, 0, 0, 1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 20) = New CustomVertex.PositionNormalColored(x2, y2, z2, 0, 0, 1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))

                vCube(i * 72 + 21) = New CustomVertex.PositionNormalColored(x1, y1, z2, 0, 0, 1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 22) = New CustomVertex.PositionNormalColored(x2, y2, z2, 0, 0, 1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 23) = New CustomVertex.PositionNormalColored(x1, y2, z2, 0, 0, 1, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))


                ' Y =========================
                ' Y Constant Forward \ Reverse
                vCube(i * 72 + 24) = New CustomVertex.PositionNormalColored(x1, y1, z1, 0, 1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 25) = New CustomVertex.PositionNormalColored(x2, y1, z2, 0, 1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 26) = New CustomVertex.PositionNormalColored(x2, y1, z1, 0, 1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))

                vCube(i * 72 + 27) = New CustomVertex.PositionNormalColored(x1, y1, z1, 0, 1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 28) = New CustomVertex.PositionNormalColored(x1, y1, z2, 0, 1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 29) = New CustomVertex.PositionNormalColored(x2, y1, z2, 0, 1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))

                vCube(i * 72 + 30) = New CustomVertex.PositionNormalColored(x1, y1, z1, 0, -1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 31) = New CustomVertex.PositionNormalColored(x2, y1, z1, 0, -1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 32) = New CustomVertex.PositionNormalColored(x2, y1, z2, 0, -1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))

                vCube(i * 72 + 33) = New CustomVertex.PositionNormalColored(x1, y1, z1, 0, -1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 34) = New CustomVertex.PositionNormalColored(x2, y1, z2, 0, -1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 35) = New CustomVertex.PositionNormalColored(x1, y1, z2, 0, -1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))

                vCube(i * 72 + 36) = New CustomVertex.PositionNormalColored(x1, y2, z1, 0, 1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 37) = New CustomVertex.PositionNormalColored(x2, y2, z2, 0, 1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 38) = New CustomVertex.PositionNormalColored(x2, y2, z1, 0, 1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))

                vCube(i * 72 + 39) = New CustomVertex.PositionNormalColored(x1, y2, z1, 0, 1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 40) = New CustomVertex.PositionNormalColored(x1, y2, z2, 0, 1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 41) = New CustomVertex.PositionNormalColored(x2, y2, z2, 0, 1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))

                vCube(i * 72 + 42) = New CustomVertex.PositionNormalColored(x1, y2, z1, 0, -1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 43) = New CustomVertex.PositionNormalColored(x2, y2, z1, 0, -1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 44) = New CustomVertex.PositionNormalColored(x2, y2, z2, 0, -1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))

                vCube(i * 72 + 45) = New CustomVertex.PositionNormalColored(x1, y2, z1, 0, -1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 46) = New CustomVertex.PositionNormalColored(x2, y2, z2, 0, -1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 47) = New CustomVertex.PositionNormalColored(x1, y2, z2, 0, -1, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))


                ' X =========================
                ' X Constant Forward \ Reverse
                vCube(i * 72 + 48) = New CustomVertex.PositionNormalColored(x1, y1, z1, -1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 49) = New CustomVertex.PositionNormalColored(x1, y2, z2, -1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 50) = New CustomVertex.PositionNormalColored(x1, y2, z1, -1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))

                vCube(i * 72 + 51) = New CustomVertex.PositionNormalColored(x1, y1, z1, -1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 52) = New CustomVertex.PositionNormalColored(x1, y1, z2, -1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 53) = New CustomVertex.PositionNormalColored(x1, y2, z2, -1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))

                vCube(i * 72 + 54) = New CustomVertex.PositionNormalColored(x1, y1, z1, 1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 55) = New CustomVertex.PositionNormalColored(x1, y2, z1, 1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 56) = New CustomVertex.PositionNormalColored(x1, y2, z2, 1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))

                vCube(i * 72 + 57) = New CustomVertex.PositionNormalColored(x1, y1, z1, 1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 58) = New CustomVertex.PositionNormalColored(x1, y2, z2, 1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 59) = New CustomVertex.PositionNormalColored(x1, y1, z2, 1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))

                vCube(i * 72 + 60) = New CustomVertex.PositionNormalColored(x2, y1, z1, -1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 61) = New CustomVertex.PositionNormalColored(x2, y2, z2, -1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 62) = New CustomVertex.PositionNormalColored(x2, y2, z1, -1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))

                vCube(i * 72 + 63) = New CustomVertex.PositionNormalColored(x2, y1, z1, -1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 64) = New CustomVertex.PositionNormalColored(x2, y1, z2, -1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 65) = New CustomVertex.PositionNormalColored(x2, y2, z2, -1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))

                vCube(i * 72 + 66) = New CustomVertex.PositionNormalColored(x2, y1, z1, 1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 67) = New CustomVertex.PositionNormalColored(x2, y2, z1, 1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 68) = New CustomVertex.PositionNormalColored(x2, y2, z2, 1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))

                vCube(i * 72 + 69) = New CustomVertex.PositionNormalColored(x2, y1, z1, 1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 70) = New CustomVertex.PositionNormalColored(x2, y2, z2, 1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))
                vCube(i * 72 + 71) = New CustomVertex.PositionNormalColored(x2, y1, z2, 1, 0, 0, IIf(Colored, SimulationData.Model.Grids(i).Color.ToArgb, IIf(i = SelectedIndex, SelectedColor.ToArgb, DefaultColor.ToArgb)))


            Next

            vbGrids.Unlock()


        Catch err As Exception
            MsgBox("loadGeometry(): " + Chr(13) + Chr(13) + err.ToString())
            Throw New Exception("could not complete loadGeometry()")
        End Try



        'End If
    End Sub

    Private Sub TDC_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        MouseIsDown = True
    End Sub

    Private Sub TDC_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
        Static L As Size
        If MouseIsDown Then
            If L.Height <> 0 And L.Width <> 0 Then

                Main.SimulationData.Graphics.X_Angle -= (e.Y - L.Height)
                Main.SimulationData.Graphics.Y_Angle -= (e.X - L.Width)
                FrameRender()
            End If
        End If
        L = e.Location

    End Sub

    Private Sub TDC_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        If Started Then
            Started = False
            Application.DoEvents()
            NewPeform()
            Application.DoEvents()
            Start()
        Else
            NewPeform()
        End If

    End Sub

    Private Sub TDC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TDC_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp
        MouseIsDown = False
    End Sub

    'Toggles
    Sub ToggleLighting()
        D3DDev.RenderState.Lighting = Not D3DDev.RenderState.Lighting
    End Sub
    Sub ToggleTransparency()
        D3DDev.RenderState.AlphaBlendEnable = Not D3DDev.RenderState.AlphaBlendEnable
    End Sub

    Private Sub TDC_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseWheel
        Dim Steps As Double = ((Max.X - Min.X) * (Max.Y - Min.Y) * (Max.Z - Min.Z)) ^ (1 / 3)
        CameraZ += (e.Delta) * (Math.Sqrt(Steps) / 1000)
        FrameRender()
    End Sub

    Private Sub TDC_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged

    End Sub


End Class
