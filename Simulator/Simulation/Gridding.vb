Public Module Gridding

    Sub CreateGrids(ByRef Grids() As Grid, ByVal Lx As Double, ByVal Ly As Double, ByVal Lz As Double, ByVal nx As Integer, ByVal ny As Integer, ByVal nz As Integer, ByVal Sw As Double, ByVal Sg As Double, ByVal Po As Double)
        Dim num As Integer = nx * ny * nz
        ReDim Grids(num - 1)
        Dim n As Integer = 0
        For i As Integer = 0 To nx - 1
            For j As Integer = 0 To ny - 1
                For k As Integer = 0 To nz - 1
                    Grids(n).Location.X = (Lx / nx) * i
                    Grids(n).Location.Y = (Ly / ny) * j
                    Grids(n).Location.Z = (Lz / nz) * k
                    Grids(n).Size.X_Length = Lx / nx
                    Grids(n).Size.Y_Length = Ly / ny
                    Grids(n).Size.Z_Length = Lz / nz
                    Grids(n).Saturation.Water = Sw
                    Grids(n).Saturation.Gas = Sg
                    Grids(n).Pressure.Oil = Po
                    Grids(n).Color = Windows.Media.Colors.Green
                    n += 1
                Next
            Next
        Next

    End Sub

End Module
