Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class TDC

    'core objects
    Dim MouseIsDown As Boolean

    Private D3DRoot As Manager
    Private D3DDev As Device
    Private D3DHelp As D3DX

    Private matCube1 As Matrix
    Private vbGrids As VertexBuffer
    Private vbText As VertexBuffer


    Private Started As Boolean

    Public Sub Start(ByVal v As VU)
        Started = True
        Do While Started
            FrameRender(v)
            Application.DoEvents()
        Loop

    End Sub

    Dim d3dPP As PresentParameters
    Dim Data As VU


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
            .PresentationInterval = PresentInterval.One
        End With
        initialiseDevice(CType(Me, System.Windows.Forms.Control), d3dPP)
    End Sub

    Sub DrawTexts(ByVal _Data As VU)
        For i As Integer = 0 To SizeOf(_Data.Grids) - 1
            Dim TextMesh As Mesh = Mesh.TextFromFont(D3DDev, New System.Drawing.Font("Arial", 8, FontStyle.Bold), _Data.Grids(i).InsideText, 0.001, 0.04)
            Dim matt As New Material
            matt.Diffuse = Color.Black
            D3DDev.Material = matt

            With _Data.Grids(i)
                D3DDev.Transform.World = (Matrix.Translation(New Vector3(.X * 5 - (.Lx * 5) / 3, .Y * 5, .Z * 5)) * matCube1) * Matrix.Scaling(0.2, 0.2, 0.2)
            End With

            TextMesh.DrawSubset(0)
        Next

    End Sub

    Public Sub FrameRender(ByVal _Data As VU)

        matCube1 = Matrix.Identity()
        matCube1 = Matrix.Multiply(matCube1, Matrix.Scaling(1, 1, 1))
        matCube1 = Matrix.Multiply(matCube1, Matrix.RotationX(Data.X_Angel * (Math.PI / 180)))
        matCube1 = Matrix.Multiply(matCube1, Matrix.RotationY(Data.Y_Angel * (Math.PI / 180)))

        D3DDev.Clear(ClearFlags.Target Or ClearFlags.ZBuffer, Color.White, 1.0F, 0)
        D3DDev.BeginScene()

        DrawTexts(_Data)
        D3DDev.Transform.World = matCube1
        D3DDev.SetStreamSource(0, vbGrids, 0)
        D3DDev.VertexFormat = CustomVertex.PositionNormalColored.Format
        D3DDev.Transform.View = Matrix.LookAtLH(New Vector3(0, 0, -5), New Vector3(0, 0, 0), New Vector3(0, 1, 0))
        D3DDev.RenderState.Ambient = Color.White

        D3DDev.DrawPrimitives(PrimitiveType.TriangleList, 0, SizeOf(Data.Grids) * 24)
        D3DDev.EndScene()
        D3DDev.Present()

    End Sub

    Private Sub initialiseDevice(ByVal Target As System.Windows.Forms.Control, ByVal win As PresentParameters)
        Dim D3DCaps As Caps
        Dim DevCreate As Integer
        win.AutoDepthStencilFormat = DepthFormat.D16
        win.EnableAutoDepthStencil = True
        D3DCaps = D3DRoot.GetDeviceCaps(0, DeviceType.Hardware)

        If D3DCaps.DeviceCaps.SupportsHardwareTransformAndLight Then
            DevCreate = CreateFlags.HardwareVertexProcessing Or CreateFlags.MultiThreaded
        Else
            DevCreate = CreateFlags.SoftwareVertexProcessing Or CreateFlags.MultiThreaded
        End If

        D3DDev = New Device(0, DeviceType.Hardware, Target, DevCreate, win)

        With D3DDev
            .RenderState.ZBufferEnable = True
            .RenderState.Lighting = True

            .RenderState.AlphaBlendEnable = True
            .RenderState.SourceBlend = Blend.SourceAlpha
            .RenderState.DestinationBlend = Blend.InvSourceAlpha

            .Transform.View = Matrix.LookAtLH(New Vector3(0, 0, -15), New Vector3(0, 0, 0), New Vector3(0, 1, 0))
            .Transform.Projection = Matrix.PerspectiveFovLH(Math.PI / 4, 4 / 3, 1, 100)
            .Transform.World = Matrix.Identity()

            .Lights(0).Type = LightType.Directional
            .Lights(0).Diffuse = Color.White
            .Lights(0).Direction = New Vector3(-1, -1, 1)
            .Lights(0).Update()
            .Lights(0).Enabled = True

            .Lights(1).Type = LightType.Directional
            .Lights(1).Diffuse = Color.White
            .Lights(1).Direction = New Vector3(0, 1, 0.5)
            .Lights(1).Update()
            .Lights(1).Enabled = True




        End With

    End Sub

    


    Public Sub loadGeometry(ByVal _Data As VU)
        Try
            Data = _Data
            vbGrids = New VertexBuffer(GetType(CustomVertex.PositionNormalColored), SizeOf(Data.Grids) * 72, D3DDev, 0, CustomVertex.PositionNormalColored.Format, Pool.Managed)
            Dim vCube As CustomVertex.PositionNormalColored() = CType(vbGrids.Lock(0, 0), CustomVertex.PositionNormalColored())


            For i As Integer = 0 To SizeOf(_Data.Grids) - 1

                Dim x1 As Double = _Data.Grids(i).X - _Data.Grids(i).Lx / 2
                Dim x2 As Double = _Data.Grids(i).X + _Data.Grids(i).Lx / 2

                Dim y1 As Double = _Data.Grids(i).Y - _Data.Grids(i).Ly / 2
                Dim y2 As Double = _Data.Grids(i).Y + _Data.Grids(i).Ly / 2

                Dim z1 As Double = _Data.Grids(i).Z - _Data.Grids(i).Lz / 2
                Dim z2 As Double = _Data.Grids(i).Z + _Data.Grids(i).Lz / 2

                ' Z =========================
                ' Z Constant Forward \ Reverse
                vCube(i * 72) = New CustomVertex.PositionNormalColored(x1, y1, z1, 0, 0, -1, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 1) = New CustomVertex.PositionNormalColored(x2, y2, z1, 0, 0, -1, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 2) = New CustomVertex.PositionNormalColored(x2, y1, z1, 0, 0, -1, Data.Grids(i).Color.ToArgb)

                vCube(i * 72 + 3) = New CustomVertex.PositionNormalColored(x1, y1, z1, 0, 0, -1, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 4) = New CustomVertex.PositionNormalColored(x1, y2, z1, 0, 0, -1, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 5) = New CustomVertex.PositionNormalColored(x2, y2, z1, 0, 0, -1, Data.Grids(i).Color.ToArgb)


                vCube(i * 72 + 6) = New CustomVertex.PositionNormalColored(x1, y1, z1, 0, 0, 1, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 7) = New CustomVertex.PositionNormalColored(x2, y1, z1, 0, 0, 1, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 8) = New CustomVertex.PositionNormalColored(x2, y2, z1, 0, 0, 1, Data.Grids(i).Color.ToArgb)

                vCube(i * 72 + 9) = New CustomVertex.PositionNormalColored(x1, y1, z1, 0, 0, 1, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 10) = New CustomVertex.PositionNormalColored(x2, y2, z1, 0, 0, 1, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 11) = New CustomVertex.PositionNormalColored(x1, y2, z1, 0, 0, 1, Data.Grids(i).Color.ToArgb)

                vCube(i * 72 + 12) = New CustomVertex.PositionNormalColored(x1, y1, z2, 0, 0, -1, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 13) = New CustomVertex.PositionNormalColored(x2, y2, z2, 0, 0, -1, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 14) = New CustomVertex.PositionNormalColored(x2, y1, z2, 0, 0, -1, Data.Grids(i).Color.ToArgb)

                vCube(i * 72 + 15) = New CustomVertex.PositionNormalColored(x1, y1, z2, 0, 0, -1, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 16) = New CustomVertex.PositionNormalColored(x1, y2, z2, 0, 0, -1, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 17) = New CustomVertex.PositionNormalColored(x2, y2, z2, 0, 0, -1, Data.Grids(i).Color.ToArgb)


                vCube(i * 72 + 18) = New CustomVertex.PositionNormalColored(x1, y1, z2, 0, 0, 1, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 19) = New CustomVertex.PositionNormalColored(x2, y1, z2, 0, 0, 1, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 20) = New CustomVertex.PositionNormalColored(x2, y2, z2, 0, 0, 1, Data.Grids(i).Color.ToArgb)

                vCube(i * 72 + 21) = New CustomVertex.PositionNormalColored(x1, y1, z2, 0, 0, 1, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 22) = New CustomVertex.PositionNormalColored(x2, y2, z2, 0, 0, 1, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 23) = New CustomVertex.PositionNormalColored(x1, y2, z2, 0, 0, 1, Data.Grids(i).Color.ToArgb)


                ' Y =========================
                ' Y Constant Forward \ Reverse
                vCube(i * 72 + 24) = New CustomVertex.PositionNormalColored(x1, y1, z1, 0, 1, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 25) = New CustomVertex.PositionNormalColored(x2, y1, z2, 0, 1, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 26) = New CustomVertex.PositionNormalColored(x2, y1, z1, 0, 1, 0, Data.Grids(i).Color.ToArgb)

                vCube(i * 72 + 27) = New CustomVertex.PositionNormalColored(x1, y1, z1, 0, 1, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 28) = New CustomVertex.PositionNormalColored(x1, y1, z2, 0, 1, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 29) = New CustomVertex.PositionNormalColored(x2, y1, z2, 0, 1, 0, Data.Grids(i).Color.ToArgb)

                vCube(i * 72 + 30) = New CustomVertex.PositionNormalColored(x1, y1, z1, 0, -1, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 31) = New CustomVertex.PositionNormalColored(x2, y1, z1, 0, -1, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 32) = New CustomVertex.PositionNormalColored(x2, y1, z2, 0, -1, 0, Data.Grids(i).Color.ToArgb)

                vCube(i * 72 + 33) = New CustomVertex.PositionNormalColored(x1, y1, z1, 0, -1, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 34) = New CustomVertex.PositionNormalColored(x2, y1, z2, 0, -1, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 35) = New CustomVertex.PositionNormalColored(x1, y1, z2, 0, -1, 0, Data.Grids(i).Color.ToArgb)

                vCube(i * 72 + 36) = New CustomVertex.PositionNormalColored(x1, y2, z1, 0, 1, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 37) = New CustomVertex.PositionNormalColored(x2, y2, z2, 0, 1, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 38) = New CustomVertex.PositionNormalColored(x2, y2, z1, 0, 1, 0, Data.Grids(i).Color.ToArgb)

                vCube(i * 72 + 39) = New CustomVertex.PositionNormalColored(x1, y2, z1, 0, 1, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 40) = New CustomVertex.PositionNormalColored(x1, y2, z2, 0, 1, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 41) = New CustomVertex.PositionNormalColored(x2, y2, z2, 0, 1, 0, Data.Grids(i).Color.ToArgb)

                vCube(i * 72 + 42) = New CustomVertex.PositionNormalColored(x1, y2, z1, 0, -1, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 43) = New CustomVertex.PositionNormalColored(x2, y2, z1, 0, -1, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 44) = New CustomVertex.PositionNormalColored(x2, y2, z2, 0, -1, 0, Data.Grids(i).Color.ToArgb)

                vCube(i * 72 + 45) = New CustomVertex.PositionNormalColored(x1, y2, z1, 0, -1, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 46) = New CustomVertex.PositionNormalColored(x2, y2, z2, 0, -1, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 47) = New CustomVertex.PositionNormalColored(x1, y2, z2, 0, -1, 0, Data.Grids(i).Color.ToArgb)


                ' X =========================
                ' X Constant Forward \ Reverse
                vCube(i * 72 + 48) = New CustomVertex.PositionNormalColored(x1, y1, z1, -1, 0, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 49) = New CustomVertex.PositionNormalColored(x1, y2, z2, -1, 0, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 50) = New CustomVertex.PositionNormalColored(x1, y2, z1, -1, 0, 0, Data.Grids(i).Color.ToArgb)

                vCube(i * 72 + 51) = New CustomVertex.PositionNormalColored(x1, y1, z1, -1, 0, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 52) = New CustomVertex.PositionNormalColored(x1, y1, z2, -1, 0, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 53) = New CustomVertex.PositionNormalColored(x1, y2, z2, -1, 0, 0, Data.Grids(i).Color.ToArgb)

                vCube(i * 72 + 54) = New CustomVertex.PositionNormalColored(x1, y1, z1, 1, 0, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 55) = New CustomVertex.PositionNormalColored(x1, y2, z1, 1, 0, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 56) = New CustomVertex.PositionNormalColored(x1, y2, z2, 1, 0, 0, Data.Grids(i).Color.ToArgb)

                vCube(i * 72 + 57) = New CustomVertex.PositionNormalColored(x1, y1, z1, 1, 0, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 58) = New CustomVertex.PositionNormalColored(x1, y2, z2, 1, 0, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 59) = New CustomVertex.PositionNormalColored(x1, y1, z2, 1, 0, 0, Data.Grids(i).Color.ToArgb)

                vCube(i * 72 + 60) = New CustomVertex.PositionNormalColored(x2, y1, z1, -1, 0, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 61) = New CustomVertex.PositionNormalColored(x2, y2, z2, -1, 0, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 62) = New CustomVertex.PositionNormalColored(x2, y2, z1, -1, 0, 0, Data.Grids(i).Color.ToArgb)

                vCube(i * 72 + 63) = New CustomVertex.PositionNormalColored(x2, y1, z1, -1, 0, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 64) = New CustomVertex.PositionNormalColored(x2, y1, z2, -1, 0, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 65) = New CustomVertex.PositionNormalColored(x2, y2, z2, -1, 0, 0, Data.Grids(i).Color.ToArgb)

                vCube(i * 72 + 66) = New CustomVertex.PositionNormalColored(x2, y1, z1, 1, 0, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 67) = New CustomVertex.PositionNormalColored(x2, y2, z1, 1, 0, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 68) = New CustomVertex.PositionNormalColored(x2, y2, z2, 1, 0, 0, Data.Grids(i).Color.ToArgb)

                vCube(i * 72 + 69) = New CustomVertex.PositionNormalColored(x2, y1, z1, 1, 0, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 70) = New CustomVertex.PositionNormalColored(x2, y2, z2, 1, 0, 0, Data.Grids(i).Color.ToArgb)
                vCube(i * 72 + 71) = New CustomVertex.PositionNormalColored(x2, y1, z2, 1, 0, 0, Data.Grids(i).Color.ToArgb)

            Next



            vbGrids.Unlock()


        Catch err As Exception
            MsgBox("loadGeometry(): " + Chr(13) + Chr(13) + err.ToString())
            Throw New Exception("could not complete loadGeometry()")
        End Try
    End Sub

    Private Sub TDC_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        MouseIsDown = True
    End Sub

    Private Sub TDC_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
        Static L As Size
        If MouseIsDown Then
            If L.Height <> 0 And L.Width <> 0 Then

                Data.X_Angel -= (e.Y - L.Height)
                Data.Y_Angel -= (e.X - L.Width)
                FrameRender(Data)
            End If
        End If
        L = e.Location

    End Sub

    Private Sub TDC_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        NewPeform()
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
End Class
