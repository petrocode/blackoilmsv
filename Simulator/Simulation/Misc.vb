Public Module Misc
    Sub Simulation_Copy(ByVal Data As Simulation, ByRef OnData As Simulation, Optional ByVal RedimIt As Boolean = True)

        With OnData
            .Exactness.IterationMSE = Data.Exactness.IterationMSE
            .Exactness.Equation = Data.Exactness.Equation
            .Exactness.OrthominMaxIteration = Data.Exactness.OrthominMaxIteration
            .RunType = Data.RunType

            .Methodology.Derivations.Gas_Saturation_Epsilon = Data.Methodology.Derivations.Gas_Saturation_Epsilon
            .Methodology.Derivations.Water_Saturation_Epsilon = Data.Methodology.Derivations.Water_Saturation_Epsilon
            .Methodology.Derivations.Oil_Pressure_Epsilon = Data.Methodology.Derivations.Oil_Pressure_Epsilon

            .Methodology.PVT.Gas.Compressibility = Data.Methodology.PVT.Gas.Compressibility
            .Methodology.PVT.Gas.Density = Data.Methodology.PVT.Gas.Density
            .Methodology.PVT.Gas.FVF = Data.Methodology.PVT.Gas.FVF
            .Methodology.PVT.Gas.Viscosity = Data.Methodology.PVT.Gas.Viscosity

            .Methodology.PVT.Oil.Compressibility = Data.Methodology.PVT.Oil.Compressibility
            .Methodology.PVT.Oil.Density = Data.Methodology.PVT.Oil.Density
            .Methodology.PVT.Oil.FVF = Data.Methodology.PVT.Oil.FVF
            .Methodology.PVT.Oil.Viscosity = Data.Methodology.PVT.Oil.Viscosity
            .Methodology.PVT.Rs = Data.Methodology.PVT.Rs

            .Methodology.PVT.Water.Compressibility = Data.Methodology.PVT.Water.Compressibility
            .Methodology.PVT.Water.Density = Data.Methodology.PVT.Water.Density
            .Methodology.PVT.Water.FVF = Data.Methodology.PVT.Water.FVF
            .Methodology.PVT.Water.Viscosity = Data.Methodology.PVT.Water.Viscosity


            .Methodology.Approach = Data.Methodology.Approach

            .Methodology.PVT = Data.Methodology.PVT

            .Options.Automatic_Grid_Adjacent_Recognition = Data.Options.Automatic_Grid_Adjacent_Recognition
            .Options.Grid_Transparency = Data.Options.Grid_Transparency
            .Options.PostPressureDistributionAcceleration = Data.Options.PostPressureDistributionAcceleration
            .Options.PPDA_i = Data.Options.PPDA_i
            .Options.PPDA_New = Data.Options.PPDA_New


            .System.Phase.Water = Data.System.Phase.Water
            .System.Phase.Oil = Data.System.Phase.Oil
            .System.Phase.Gas = Data.System.Phase.Gas
            .System.Units = Data.System.Units

            .Current_Step = Data.Current_Step

            .CurrentIterationCount = Data.CurrentIterationCount
            .CurrentIteration = Data.CurrentIteration

            If RedimIt Then
                ReDim .Time_Steps(SizeOf(Data.Time_Steps) - 1)
            End If

            For i As Integer = 0 To SizeOf(Data.Time_Steps) - 1
                .Time_Steps(i) = Data.Time_Steps(i)
            Next

            .Constants.Alpha = Data.Constants.Alpha
            .Constants.Beta = Data.Constants.Beta



            If RedimIt Then
                ReDim .Model.Grids(SizeOf(Data.Model.Grids) - 1)
            End If

            .Total_Time = Data.Total_Time

            If RedimIt Then
                ReDim .Schedules(SizeOf(Data.Schedules) - 1)
            End If

            For i As Integer = 0 To SizeOf(Data.Schedules) - 1
                .Schedules(i).Type = Data.Schedules(i).Type
                .Schedules(i).Orientation = Data.Schedules(i).Orientation
                .Schedules(i).Position = Data.Schedules(i).Position
                .Schedules(i).Target = Data.Schedules(i).Target
                If RedimIt Then
                    ReDim Preserve .Schedules(i).Data(Data.Schedules(i).Data.GetLength(0) - 1, Data.Schedules(i).Data.GetLength(1) - 1)
                End If
                For i1 As Integer = 0 To Data.Schedules(i).Data.GetLength(0) - 1
                    For j1 As Integer = 0 To Data.Schedules(i).Data.GetLength(1) - 1
                        .Schedules(i).Data(i1, j1) = Data.Schedules(i).Data(i1, j1)
                    Next
                Next
                .Schedules(i).ID = Data.Schedules(i).ID
            Next

            For i As Integer = 0 To SizeOf(.Model.Grids) - 1
                If RedimIt Then
                    ReDim .Model.Grids(i).Schedules(SizeOf(Data.Model.Grids(i).Schedules) - 1)
                End If
                For j As Integer = 0 To SizeOf(Data.Model.Grids(i).Schedules) - 1
                    .Model.Grids(i).Schedules(j) = Data.Model.Grids(i).Schedules(j)
                Next

                .Model.Grids(i).PVT.Temperature = Data.Model.Grids(i).PVT.Temperature
                .Model.Grids(i).Porosity = Data.Model.Grids(i).Porosity
                .Model.Grids(i).System = Data.Model.Grids(i).System

                .Model.Grids(i).Flow.Water = Data.Model.Grids(i).Flow.Water
                .Model.Grids(i).Flow.Gas = Data.Model.Grids(i).Flow.Gas
                .Model.Grids(i).Flow.Oil = Data.Model.Grids(i).Flow.Oil

                .Model.Grids(i).Saturation.Water = Data.Model.Grids(i).Saturation.Water
                .Model.Grids(i).Saturation.Oil = Data.Model.Grids(i).Saturation.Oil
                .Model.Grids(i).Saturation.Gas = Data.Model.Grids(i).Saturation.Gas

                .Model.Grids(i).SCAL.pcgo = Data.Model.Grids(i).SCAL.pcgo
                .Model.Grids(i).SCAL.pcgo_drg = Data.Model.Grids(i).SCAL.pcgo_drg
                .Model.Grids(i).SCAL.pcgo_imb = Data.Model.Grids(i).SCAL.pcgo_imb

                .Model.Grids(i).SCAL.pcow = Data.Model.Grids(i).SCAL.pcow
                .Model.Grids(i).SCAL.pcow_drg = Data.Model.Grids(i).SCAL.pcow_drg
                .Model.Grids(i).SCAL.pcgo_imb = Data.Model.Grids(i).SCAL.pcgo_imb


                .Model.Grids(i).SCAL.Krw.Kx = Data.Model.Grids(i).SCAL.Krw.Kx
                .Model.Grids(i).SCAL.Krw.Ky = Data.Model.Grids(i).SCAL.Krw.Ky
                .Model.Grids(i).SCAL.Krw.Kz = Data.Model.Grids(i).SCAL.Krw.Kz

                .Model.Grids(i).SCAL.Kro.Kx = Data.Model.Grids(i).SCAL.Kro.Kx
                .Model.Grids(i).SCAL.Kro.Ky = Data.Model.Grids(i).SCAL.Kro.Ky
                .Model.Grids(i).SCAL.Kro.Kz = Data.Model.Grids(i).SCAL.Kro.Kz

                .Model.Grids(i).SCAL.Krg.Kx = Data.Model.Grids(i).SCAL.Krg.Kx
                .Model.Grids(i).SCAL.Krg.Ky = Data.Model.Grids(i).SCAL.Krg.Ky
                .Model.Grids(i).SCAL.Krg.Kz = Data.Model.Grids(i).SCAL.Krg.Kz

                .Model.Grids(i).Size.X_Length = Data.Model.Grids(i).Size.X_Length
                .Model.Grids(i).Size.Y_Length = Data.Model.Grids(i).Size.Y_Length
                .Model.Grids(i).Size.Z_Length = Data.Model.Grids(i).Size.Z_Length

                .Model.Grids(i).RockType = Data.Model.Grids(i).RockType
                .Model.Grids(i).Volume = Data.Model.Grids(i).Volume

                .Model.Grids(i).PVT.Rs = Data.Model.Grids(i).PVT.Rs

                .Model.Grids(i).PVT.Water.Viscosity = Data.Model.Grids(i).PVT.Water.Viscosity
                .Model.Grids(i).PVT.Water.FVF = Data.Model.Grids(i).PVT.Water.FVF
                .Model.Grids(i).PVT.Water.Density = Data.Model.Grids(i).PVT.Water.Density

                .Model.Grids(i).PVT.Oil.Viscosity = Data.Model.Grids(i).PVT.Oil.Viscosity
                .Model.Grids(i).PVT.Oil.FVF = Data.Model.Grids(i).PVT.Oil.FVF
                .Model.Grids(i).PVT.Oil.Density = Data.Model.Grids(i).PVT.Oil.Density

                .Model.Grids(i).PVT.Gas.Viscosity = Data.Model.Grids(i).PVT.Gas.Viscosity
                .Model.Grids(i).PVT.Gas.FVF = Data.Model.Grids(i).PVT.Gas.FVF
                .Model.Grids(i).PVT.Gas.Density = Data.Model.Grids(i).PVT.Gas.Density

                .Model.Grids(i).Pressure.Water = Data.Model.Grids(i).Pressure.Water
                .Model.Grids(i).Pressure.Oil = Data.Model.Grids(i).Pressure.Oil
                .Model.Grids(i).Pressure.Gas = Data.Model.Grids(i).Pressure.Gas

                .Model.Grids(i).Location.X = Data.Model.Grids(i).Location.X
                .Model.Grids(i).Location.Y = Data.Model.Grids(i).Location.Y
                .Model.Grids(i).Location.Z = Data.Model.Grids(i).Location.Z

                .Model.Grids(i).Color = Data.Model.Grids(i).Color

                If SizeOf(Data.Model.Grids(i).Adjacents) > 0 Then
                    If RedimIt Then
                        ReDim .Model.Grids(i).Adjacents(SizeOf(Data.Model.Grids(i).Adjacents) - 1)
                    End If
                    For j As Integer = 0 To SizeOf(Data.Model.Grids(i).Adjacents) - 1
                        .Model.Grids(i).Adjacents(j).Area = Data.Model.Grids(i).Adjacents(j).Area
                        .Model.Grids(i).Adjacents(j).Target_Grid_Index = Data.Model.Grids(i).Adjacents(j).Target_Grid_Index
                        .Model.Grids(i).Adjacents(j).Type = Data.Model.Grids(i).Adjacents(j).Type
                    Next
                End If

            Next

            If SizeOf(Data.C_Bar_n) > 0 Then
                If RedimIt Then
                    ReDim .C_Bar_n(0, Data.C_Bar_n.GetLength(1) - 1)
                End If
                For i As Integer = 0 To Data.C_Bar_n.GetLength(1) - 1
                    .C_Bar_n(0, i) = Data.C_Bar_n(0, i)
                Next
            End If

            .Final_Step = Data.Final_Step
            .Finish_Time = Data.Finish_Time


            If SizeOf(Data.PVT.Table) > 0 Then
                If RedimIt Then
                    ReDim .PVT.Table(Data.PVT.Table.GetLength(0) - 1, Data.PVT.Table.GetLength(1) - 1)
                End If
                For i As Integer = 0 To Data.PVT.Table.GetLength(0) - 1
                    For j As Integer = 0 To Data.PVT.Table.GetLength(1) - 1
                        .PVT.Table(i, j) = Data.PVT.Table(i, j)
                    Next
                Next
            End If

            .PVT.Oil_At_Standard.Compressibility = Data.PVT.Oil_At_Standard.Compressibility
            .PVT.Water_At_Standard.Compressibility = Data.PVT.Water_At_Standard.Compressibility
            .PVT.Gas_At_Standard.Compressibility = Data.PVT.Gas_At_Standard.Compressibility
            .PVT.Condesate.Compressibility = Data.PVT.Condesate.Compressibility

            .PVT.Oil_At_Standard.Density = Data.PVT.Oil_At_Standard.Density
            .PVT.Water_At_Standard.Density = Data.PVT.Water_At_Standard.Density
            .PVT.Gas_At_Standard.Density = Data.PVT.Gas_At_Standard.Density

            .PVT.Oil_At_Standard.Viscosity = Data.PVT.Oil_At_Standard.Viscosity
            .PVT.Water_At_Standard.Viscosity = Data.PVT.Water_At_Standard.Viscosity
            .PVT.Gas_At_Standard.Viscosity = Data.PVT.Gas_At_Standard.Viscosity

            .PVT.Gas_Molecular_Weight = Data.PVT.Gas_Molecular_Weight
            .PVT.Gas_Critical_Pressure = Data.PVT.Gas_Critical_Pressure
            .PVT.Gas_Critical_Temperature = Data.PVT.Gas_Critical_Temperature

            .PVT.Gas_Bubble_Pressure = Data.PVT.Gas_Bubble_Pressure
            .PVT.StandingC = Data.PVT.StandingC


            If SizeOf(Data.Rocks) > 0 Then
                If RedimIt Then
                    ReDim .Rocks(SizeOf(Data.Rocks) - 1)
                End If
                For i As Integer = 0 To .Rocks.Length - 1
                    .Rocks(i).Initial_Porosity = Data.Rocks(i).Initial_Porosity
                    .Rocks(i).Cs = Data.Rocks(i).Cs
                    .Rocks(i).Permeability.Kx = Data.Rocks(i).Permeability.Kx
                    .Rocks(i).Permeability.Ky = Data.Rocks(i).Permeability.Ky
                    .Rocks(i).Permeability.Kz = Data.Rocks(i).Permeability.Kz
                    .Rocks(i).SCAL.Wettability.Water_Oil = Data.Rocks(i).SCAL.Wettability.Water_Oil
                    .Rocks(i).SCAL.Wettability.Water_Gas = Data.Rocks(i).SCAL.Wettability.Water_Gas
                    .Rocks(i).SCAL.Wettability.Oil_Gas = Data.Rocks(i).SCAL.Wettability.Oil_Gas

                    If SizeOf(Data.Rocks(i).SCAL.Tables.Water_Oil) > 0 Then
                        If RedimIt Then
                            ReDim .Rocks(i).SCAL.Tables.Water_Oil(Data.Rocks(i).SCAL.Tables.Water_Oil.GetLength(0) - 1, Data.Rocks(i).SCAL.Tables.Water_Oil.GetLength(1) - 1)
                        End If
                        For i1 As Integer = 0 To Data.Rocks(i).SCAL.Tables.Water_Oil.GetLength(0) - 1
                            For j1 As Integer = 0 To Data.Rocks(i).SCAL.Tables.Water_Oil.GetLength(1) - 1
                                .Rocks(i).SCAL.Tables.Water_Oil(i1, j1) = Data.Rocks(i).SCAL.Tables.Water_Oil(i1, j1)
                            Next
                        Next
                    End If


                    If SizeOf(Data.Rocks(i).SCAL.Tables.Oil_Gas) > 0 Then
                        If RedimIt Then
                            ReDim .Rocks(i).SCAL.Tables.Oil_Gas(Data.Rocks(i).SCAL.Tables.Oil_Gas.GetLength(0) - 1, Data.Rocks(i).SCAL.Tables.Oil_Gas.GetLength(1) - 1)
                        End If
                        For i1 As Integer = 0 To Data.Rocks(i).SCAL.Tables.Oil_Gas.GetLength(0) - 1
                            For j1 As Integer = 0 To Data.Rocks(i).SCAL.Tables.Oil_Gas.GetLength(1) - 1
                                .Rocks(i).SCAL.Tables.Oil_Gas(i1, j1) = Data.Rocks(i).SCAL.Tables.Oil_Gas(i1, j1)
                            Next
                        Next
                    End If


                    If SizeOf(Data.Rocks(i).SCAL.Tables.Gas_Water) > 0 Then
                        If RedimIt Then
                            ReDim .Rocks(i).SCAL.Tables.Gas_Water(Data.Rocks(i).SCAL.Tables.Gas_Water.GetLength(0) - 1, Data.Rocks(i).SCAL.Tables.Gas_Water.GetLength(1) - 1)
                        End If
                        For i1 As Integer = 0 To Data.Rocks(i).SCAL.Tables.Gas_Water.GetLength(0) - 1
                            For j1 As Integer = 0 To Data.Rocks(i).SCAL.Tables.Gas_Water.GetLength(1) - 1
                                .Rocks(i).SCAL.Tables.Gas_Water(i1, j1) = Data.Rocks(i).SCAL.Tables.Gas_Water(i1, j1)
                            Next
                        Next
                    End If


                    If SizeOf(Data.Rocks(i).SCAL.Tables.Gas_Water_Oil.Sw) > 0 Then
                        If RedimIt Then
                            ReDim .Rocks(i).SCAL.Tables.Gas_Water_Oil.Sw(SizeOf(Data.Rocks(i).SCAL.Tables.Gas_Water_Oil.Sw) - 1)
                            ReDim .Rocks(i).SCAL.Tables.Gas_Water_Oil.Values(SizeOf(Data.Rocks(i).SCAL.Tables.Gas_Water_Oil.Sw) - 1)
                        End If
                        For i1 As Integer = 0 To SizeOf(Data.Rocks(i).SCAL.Tables.Gas_Water_Oil.Sw) - 1
                            .Rocks(i).SCAL.Tables.Gas_Water_Oil.Sw(i1) = Data.Rocks(i).SCAL.Tables.Gas_Water_Oil.Sw(i1)
                            If RedimIt Then
                                ReDim .Rocks(i).SCAL.Tables.Gas_Water_Oil.Values(i1)(Data.Rocks(i).SCAL.Tables.Gas_Water_Oil.Values(i1).GetLength(0) - 1, Data.Rocks(i).SCAL.Tables.Gas_Water_Oil.Values(i1).GetLength(1) - 1)
                            End If
                            For j1 As Integer = 0 To Data.Rocks(i).SCAL.Tables.Gas_Water_Oil.Values(i1).GetLength(0) - 1
                                For j2 As Integer = 0 To Data.Rocks(i).SCAL.Tables.Gas_Water_Oil.Values(i1).GetLength(1) - 1
                                    .Rocks(i).SCAL.Tables.Gas_Water_Oil.Values(i1)(j1, j2) = Data.Rocks(i).SCAL.Tables.Gas_Water_Oil.Values(i1)(j1, j2)
                                Next
                            Next
                        Next
                    End If

                Next

            End If


        End With

    End Sub
    Function Simulation_Copy(ByVal Data As Simulation, Optional ByVal RedimIt As Boolean = True) As Simulation
        Dim out As New Simulation
        Simulation_Copy(Data, out, RedimIt)
        Return out
    End Function
    Function AddStringOfLength(ByVal str As String, ByVal length As Integer, ByVal S As String) As String
        For i As Integer = 0 To length - 1
            str += S
        Next
        Return str
    End Function
    Function SizeOf(ByVal arr As Array) As Integer
        If arr Is Nothing Then Return 0
        Return arr.Length
    End Function

    Function Grid_Parameter(ByVal nGrid As Grid, ByVal Parameter As String, ByVal System As System_of_Units) As Double

        Select Case Parameter
            Case "Po"
                Return nGrid.Pressure.Oil
            Case "Pw"
                Return nGrid.Pressure.Water
            Case "Pg"
                Return nGrid.Pressure.Gas
            Case "So"
                Return nGrid.Saturation.Oil
            Case "Sw"
                Return nGrid.Saturation.Water
            Case "Sg"
                Return nGrid.Saturation.Gas
            Case "Ko - X"
                Return nGrid.SCAL.Kro.Kx
            Case "Ko - Y"
                Return nGrid.SCAL.Kro.Ky
            Case "Ko - Z"
                Return nGrid.SCAL.Kro.Kz
            Case "Kw - X"
                Return nGrid.SCAL.Krw.Kx
            Case "Kw - Y"
                Return nGrid.SCAL.Krw.Ky
            Case "Kw - Z"
                Return nGrid.SCAL.Krw.Kz
            Case "Kg - X"
                Return nGrid.SCAL.Krg.Kx
            Case "Kg - Y"
                Return nGrid.SCAL.Krg.Ky
            Case "Kg - Z"
                Return nGrid.SCAL.Krg.Kz
            Case "Bo"
                Return nGrid.PVT.Oil.FVF
            Case "Bw"
                Return nGrid.PVT.Water.FVF
            Case "Bg"
                Return nGrid.PVT.Gas.FVF
            Case "Rs"
                Return nGrid.PVT.Rs
            Case "Viscosity - Oil"
                Return nGrid.PVT.Oil.Viscosity
            Case "Viscosity - Water"
                Return nGrid.PVT.Water.Viscosity
            Case "Viscosity - Gas"
                Return nGrid.PVT.Gas.Viscosity
            Case "Density - Oil"
                Return nGrid.PVT.Oil.Density
            Case "Density - Water"
                Return nGrid.PVT.Water.Density
            Case "Density - Gas"
                Return nGrid.PVT.Gas.Density
            Case "Gamma - Oil"
                Return nGrid.PVT.Oil.Pressure_Gradient(System)
            Case "Gamma - Water"
                Return nGrid.PVT.Water.Pressure_Gradient(System)
            Case "Gamma - Gas"
                Return nGrid.PVT.Gas.Pressure_Gradient(System)
            Case "X"
                Return nGrid.Location.X
            Case "Y"
                Return nGrid.Location.Y
            Case "Z"
                Return nGrid.Location.Z
            Case "Lx"
                Return nGrid.Size.X_Length
            Case "Ly"
                Return nGrid.Size.Y_Length
            Case "Lz"
                Return nGrid.Size.Z_Length
        End Select
    End Function

    Sub GetMaxMin(ByVal Parameter As String, ByRef Max As Double, ByRef Min As Double, ByRef Data As Simulation)
        For i As Integer = 0 To SizeOf(Data.Model.Grids) - 1
            If Grid_Parameter(Data.Model.Grids(i), Parameter, Data.System.Units) > Max Then Max = Grid_Parameter(Data.Model.Grids(i), Parameter, Data.System.Units)
            If Grid_Parameter(Data.Model.Grids(i), Parameter, Data.System.Units) < Min Then Min = Grid_Parameter(Data.Model.Grids(i), Parameter, Data.System.Units)
        Next
    End Sub
    Function RefreshColors(ByVal Parameter As String, ByVal Min As Double, ByVal Max As Double, ByVal MinColor As Color, ByVal MaxColor As Color, ByRef Data As Simulation) As Boolean
        Dim Rmin As Double = MinColor.R
        Dim Gmin As Double = MinColor.G
        Dim Bmin As Double = MinColor.B

        Dim Rmax As Double = MaxColor.R
        Dim Gmax As Double = MaxColor.G
        Dim Bmax As Double = MaxColor.B

        Dim RLen As Double = Rmax - Rmin
        Dim GLen As Double = Gmax - Gmin
        Dim BLen As Double = Bmax - Bmin

        For i As Integer = 0 To SizeOf(Data.Model.Grids) - 1
            Dim Value As Double = Grid_Parameter(Data.Model.Grids(i), Parameter, Data.System.Units)
            Dim portion As Double = 0
            Dim R, G, B, A As Integer
            portion = (Value - Min) / (Max - Min)

            If CSng(portion) < 0 Then
                Main.ColorBoundary1.ResetMin()
                Exit Function
            End If

            If CSng(portion) > 1 Then
                Main.ColorBoundary1.ResetMax()
                Exit Function
            End If

            If Not Double.IsNaN(portion) Then
                R = CInt(portion * RLen + Rmin)
                G = CInt(portion * GLen + Gmin)
                B = CInt(portion * BLen + Bmin)
                A = CInt(Data.Options.Grid_Transparency)
            End If
            If R >= 0 And B >= 0 And G >= 0 And R <= 255 And B <= 255 And G <= 255 Then
                Data.Model.Grids(i).Color = System.Windows.Media.Color.FromArgb(A, R, G, B)
            Else
                Main.Log("Value out of graph color boundary, Value = " + Value.ToString)
                Return False
            End If
        Next
        Return True

    End Function


End Module
