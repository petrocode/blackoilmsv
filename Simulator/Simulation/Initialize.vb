Public Module _Initialize

    Sub Initialize(ByRef Data As Simulation, Optional ByVal IsTimeStep As Boolean = True)
        Select Case Data.Methodology.Approach
            Case Simulation_Approach.Fully_Implicit
                For i As Integer = 0 To SizeOf(Data.Model.Grids) - 1
                    Initialize_Saturations(Data, Data.Model.Grids(i))
                    Initialize_Phase_System(Data.Model.Grids(i), Data.Rocks(Data.Model.Grids(i).RockType))
                    Initialize_SCAL(Data.Model.Grids(i), Data.Rocks(Data.Model.Grids(i).RockType), 1)
                    Initialize_Pressure(Data.Model.Grids(i))
                    Initialize_PVT(Data.Model.Grids(i), Data)
                    Initialize_Porosity(Data, i)
                    Initialize_Fully_Implicit_Schedule(Data, i, IsTimeStep)

                    Initialize_C_Bar_n(Data, i)
                Next
            Case Simulation_Approach.IMPES

                For i As Integer = 0 To SizeOf(Data.Model.Grids) - 1
                    If IsTimeStep Then Initialize_Saturations(Data, Data.Model.Grids(i))
                    If IsTimeStep Then Initialize_Phase_System(Data.Model.Grids(i), Data.Rocks(Data.Model.Grids(i).RockType))
                    If IsTimeStep Then Initialize_SCAL(Data.Model.Grids(i), Data.Rocks(Data.Model.Grids(i).RockType))
                    Initialize_Pressure(Data.Model.Grids(i))
                    Initialize_PVT(Data.Model.Grids(i), Data)
                    Initialize_Porosity(Data, i)
                    Initialize_IMPES_Schedule(Data, i, IsTimeStep)

                Next
        End Select
    End Sub

    Sub Capillary_Pressure_Hystersis(ByRef G As Grid, ByVal Wet As Wettability_Couple, ByVal NewSw As Decimal, ByVal NewSg As Decimal)
        Dim Delta_Water As Decimal = (NewSw - G.Saturation.Water)
        Dim Delta_Oil As Decimal = ((1 - NewSg - NewSw) - G.Saturation.Oil)
        Dim Delta_Gas As Decimal = (NewSg - G.Saturation.Gas)

        Dim kk As Rock


        Select Case Wet.Water_Oil
            Case Phase.Water
                If Delta_Water > Delta_Oil Then
                    G.SCAL.pcow = G.SCAL.pcow_imb
                Else
                    G.SCAL.pcow = G.SCAL.pcow_drg
                End If

            Case Phase.Oil
                If Delta_Water > Delta_Oil Then
                    G.SCAL.pcow = G.SCAL.pcow_drg
                Else
                    G.SCAL.pcow = G.SCAL.pcow_imb
                End If
        End Select


        Select Case Wet.Oil_Gas
            Case Phase.Oil
                If Delta_Oil > Delta_Gas Then
                    G.SCAL.pcgo = G.SCAL.pcgo_imb
                Else
                    G.SCAL.pcgo = G.SCAL.pcgo_drg
                End If
            Case Phase.Gas
                If Delta_Oil > Delta_Gas Then
                    G.SCAL.pcgo = G.SCAL.pcgo_drg
                Else
                    G.SCAL.pcgo = G.SCAL.pcgo_imb
                End If
        End Select



    End Sub



    Sub Initialize_Fully_Implicit_Schedule(ByRef Data As Simulation, ByVal iGrid As Integer, ByVal IsTimeStep As Boolean)



        For n As Integer = 0 To SizeOf(Data.Model.Grids(iGrid).Schedules) - 1
            Select Case Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Position
                Case Simulation_Schedule_Position.Surface

                    Select Case Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Type
                        Case Schedule_Type.Static_Oil_Pressure
                            Data.Model.Grids(iGrid).Flow.Oil = GetSurfaceFlow(Data, Data.Model.Grids(iGrid), Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Orientation, Phase.Oil, Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0))
                        Case Schedule_Type.Static_Gas_Pressure
                            Data.Model.Grids(iGrid).Flow.Gas = GetSurfaceFlow(Data, Data.Model.Grids(iGrid), Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Orientation, Phase.Gas, Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0))
                        Case Schedule_Type.Static_Water_Pressure
                            Data.Model.Grids(iGrid).Flow.Water = GetSurfaceFlow(Data, Data.Model.Grids(iGrid), Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Orientation, Phase.Water, Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0))
                        Case Schedule_Type.Dynamic_Oil_Pressure
                            Dim Pressure As Decimal = Interpolate(Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data, Data.Current_Time)
                            Data.Model.Grids(iGrid).Flow.Oil = GetSurfaceFlow(Data, Data.Model.Grids(iGrid), Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Orientation, Phase.Oil, Pressure)
                        Case Schedule_Type.Dynamic_Gas_Pressure
                            Dim Pressure As Decimal = Interpolate(Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data, Data.Current_Time)
                            Data.Model.Grids(iGrid).Flow.Gas = GetSurfaceFlow(Data, Data.Model.Grids(iGrid), Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Orientation, Phase.Gas, Pressure)
                        Case Schedule_Type.Dynamic_Water_Pressure
                            Dim Pressure As Decimal = Interpolate(Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data, Data.Current_Time)
                            Data.Model.Grids(iGrid).Flow.Water = GetSurfaceFlow(Data, Data.Model.Grids(iGrid), Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Orientation, Phase.Water, Pressure)
                    End Select

            End Select
        Next

        If IsTimeStep Then


            For n As Integer = 0 To SizeOf(Data.Model.Grids(iGrid).Schedules) - 1
                Select Case Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Position
                    Case Simulation_Schedule_Position.Internal
                        Select Case Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Type

                            Case Schedule_Type.Static_Oil_Flow
                                Data.Model.Grids(iGrid).Flow.Oil = Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0)
                            Case Schedule_Type.Static_Gas_Flow
                                Data.Model.Grids(iGrid).Flow.Gas = Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0)
                            Case Schedule_Type.Static_Water_Flow
                                Data.Model.Grids(iGrid).Flow.Water = Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0)
                            Case Schedule_Type.Static_Oil_Pressure
                                Data.Model.Grids(iGrid).Pressure.Oil = Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0)
                                Data.Model.Grids(iGrid).Pressure.Water = Data.Model.Grids(iGrid).Pressure.Oil - Data.Model.Grids(iGrid).SCAL.pcow
                                Data.Model.Grids(iGrid).Pressure.Gas = Data.Model.Grids(iGrid).SCAL.pcgo + Data.Model.Grids(iGrid).Pressure.Oil
                            Case Schedule_Type.Static_Gas_Pressure
                                Data.Model.Grids(iGrid).Pressure.Gas = Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0)
                                Data.Model.Grids(iGrid).Pressure.Oil = Data.Model.Grids(iGrid).Pressure.Gas - Data.Model.Grids(iGrid).SCAL.pcgo
                                Data.Model.Grids(iGrid).Pressure.Water = Data.Model.Grids(iGrid).Pressure.Oil - Data.Model.Grids(iGrid).SCAL.pcow
                            Case Schedule_Type.Static_Water_Pressure
                                Data.Model.Grids(iGrid).Pressure.Water = Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0)
                                Data.Model.Grids(iGrid).Pressure.Oil = Data.Model.Grids(iGrid).Pressure.Water + Data.Model.Grids(iGrid).SCAL.pcow
                                Data.Model.Grids(iGrid).Pressure.Gas = Data.Model.Grids(iGrid).SCAL.pcgo + Data.Model.Grids(iGrid).Pressure.Oil

                            Case Schedule_Type.Dynamic_Oil_Flow
                                Data.Model.Grids(iGrid).Flow.Oil = Interpolate(Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data, Data.Current_Time)
                            Case Schedule_Type.Dynamic_Gas_Flow
                                Data.Model.Grids(iGrid).Flow.Gas = Interpolate(Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data, Data.Current_Time)
                            Case Schedule_Type.Dynamic_Water_Flow
                                Data.Model.Grids(iGrid).Flow.Water = Interpolate(Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data, Data.Current_Time)
                            Case Schedule_Type.Dynamic_Oil_Pressure
                                Data.Model.Grids(iGrid).Pressure.Oil = Interpolate(Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data, Data.Current_Time)
                                Data.Model.Grids(iGrid).Pressure.Water = Data.Model.Grids(iGrid).Pressure.Oil - Data.Model.Grids(iGrid).SCAL.pcow
                                Data.Model.Grids(iGrid).Pressure.Gas = Data.Model.Grids(iGrid).SCAL.pcgo + Data.Model.Grids(iGrid).Pressure.Oil
                            Case Schedule_Type.Dynamic_Gas_Pressure
                                Data.Model.Grids(iGrid).Pressure.Gas = Interpolate(Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data, Data.Current_Time)
                                Data.Model.Grids(iGrid).Pressure.Oil = Data.Model.Grids(iGrid).Pressure.Gas - Data.Model.Grids(iGrid).SCAL.pcgo
                                Data.Model.Grids(iGrid).Pressure.Water = Data.Model.Grids(iGrid).Pressure.Oil - Data.Model.Grids(iGrid).SCAL.pcow
                            Case Schedule_Type.Dynamic_Water_Pressure
                                Data.Model.Grids(iGrid).Pressure.Water = Interpolate(Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data, Data.Current_Time)
                                Data.Model.Grids(iGrid).Pressure.Oil = Data.Model.Grids(iGrid).Pressure.Water + Data.Model.Grids(iGrid).SCAL.pcow
                                Data.Model.Grids(iGrid).Pressure.Gas = Data.Model.Grids(iGrid).SCAL.pcgo + Data.Model.Grids(iGrid).Pressure.Oil
                            Case Schedule_Type.Static_Water_Saturation
                                Data.Model.Grids(iGrid).Saturation.Water = Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0)


                        End Select

                End Select
            Next
        End If

    End Sub

    Sub Initialize_IMPES_Schedule(ByRef Data As Simulation, ByVal iGrid As Integer, ByVal IsTimeStep As Boolean)

        If Not IsTimeStep Then Exit Sub

        Static InWater As Decimal
        Static OutOil As Decimal


        For n As Integer = 0 To SizeOf(Data.Model.Grids(iGrid).Schedules) - 1
            Select Case Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Position
                Case Simulation_Schedule_Position.Surface

                    Select Case Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Type
                        Case Schedule_Type.Static_Oil_Pressure
                            Data.Model.Grids(iGrid).Flow.Oil = GetSurfaceFlow(Data, Data.Model.Grids(iGrid), Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Orientation, Phase.Oil, Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0))
                        Case Schedule_Type.Static_Gas_Pressure
                            Data.Model.Grids(iGrid).Flow.Gas = GetSurfaceFlow(Data, Data.Model.Grids(iGrid), Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Orientation, Phase.Gas, Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0))
                        Case Schedule_Type.Static_Water_Pressure
                            Data.Model.Grids(iGrid).Flow.Water = GetSurfaceFlow(Data, Data.Model.Grids(iGrid), Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Orientation, Phase.Water, Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0))
                        Case Schedule_Type.Dynamic_Oil_Pressure
                            Dim Pressure As Decimal = Interpolate(Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data, Data.Current_Time)
                            Data.Model.Grids(iGrid).Flow.Oil = GetSurfaceFlow(Data, Data.Model.Grids(iGrid), Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Orientation, Phase.Oil, Pressure)
                        Case Schedule_Type.Dynamic_Gas_Pressure
                            Dim Pressure As Decimal = Interpolate(Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data, Data.Current_Time)
                            Data.Model.Grids(iGrid).Flow.Gas = GetSurfaceFlow(Data, Data.Model.Grids(iGrid), Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Orientation, Phase.Gas, Pressure)
                        Case Schedule_Type.Dynamic_Water_Pressure
                            Dim Pressure As Decimal = Interpolate(Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data, Data.Current_Time)
                            Data.Model.Grids(iGrid).Flow.Water = GetSurfaceFlow(Data, Data.Model.Grids(iGrid), Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Orientation, Phase.Water, Pressure)
                    End Select

            End Select
        Next

        For n As Integer = 0 To SizeOf(Data.Model.Grids(iGrid).Schedules) - 1
            Select Case Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Position
                Case Simulation_Schedule_Position.Internal
                    Select Case Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Type

                        Case Schedule_Type.Static_Oil_Flow
                            Data.Model.Grids(iGrid).Flow.Oil = Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0)
                        Case Schedule_Type.Static_Gas_Flow
                            Data.Model.Grids(iGrid).Flow.Gas = Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0)
                        Case Schedule_Type.Static_Water_Flow
                            Data.Model.Grids(iGrid).Flow.Water = Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0)
                        Case Schedule_Type.Static_Oil_Pressure
                            Data.Model.Grids(iGrid).Pressure.Oil = Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0)
                            Data.Model.Grids(iGrid).Pressure.Water = Data.Model.Grids(iGrid).Pressure.Oil - Data.Model.Grids(iGrid).SCAL.pcow
                            Data.Model.Grids(iGrid).Pressure.Gas = Data.Model.Grids(iGrid).SCAL.pcgo + Data.Model.Grids(iGrid).Pressure.Oil
                        Case Schedule_Type.Static_Gas_Pressure
                            Data.Model.Grids(iGrid).Pressure.Gas = Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0)
                            Data.Model.Grids(iGrid).Pressure.Oil = Data.Model.Grids(iGrid).Pressure.Gas - Data.Model.Grids(iGrid).SCAL.pcgo
                            Data.Model.Grids(iGrid).Pressure.Water = Data.Model.Grids(iGrid).Pressure.Oil - Data.Model.Grids(iGrid).SCAL.pcow
                        Case Schedule_Type.Static_Water_Pressure
                            Data.Model.Grids(iGrid).Pressure.Water = Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0)
                            Data.Model.Grids(iGrid).Pressure.Oil = Data.Model.Grids(iGrid).Pressure.Water + Data.Model.Grids(iGrid).SCAL.pcow
                            Data.Model.Grids(iGrid).Pressure.Gas = Data.Model.Grids(iGrid).SCAL.pcgo + Data.Model.Grids(iGrid).Pressure.Oil

                        Case Schedule_Type.Dynamic_Oil_Flow
                            Data.Model.Grids(iGrid).Flow.Oil = Interpolate(Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data, Data.Current_Time)
                        Case Schedule_Type.Dynamic_Gas_Flow
                            Data.Model.Grids(iGrid).Flow.Gas = Interpolate(Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data, Data.Current_Time)
                        Case Schedule_Type.Dynamic_Water_Flow
                            Data.Model.Grids(iGrid).Flow.Water = Interpolate(Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data, Data.Current_Time)
                        Case Schedule_Type.Dynamic_Oil_Pressure
                            Data.Model.Grids(iGrid).Pressure.Oil = Interpolate(Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data, Data.Current_Time)
                            Data.Model.Grids(iGrid).Pressure.Water = Data.Model.Grids(iGrid).Pressure.Oil - Data.Model.Grids(iGrid).SCAL.pcow
                            Data.Model.Grids(iGrid).Pressure.Gas = Data.Model.Grids(iGrid).SCAL.pcgo + Data.Model.Grids(iGrid).Pressure.Oil
                        Case Schedule_Type.Dynamic_Gas_Pressure
                            Data.Model.Grids(iGrid).Pressure.Gas = Interpolate(Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data, Data.Current_Time)
                            Data.Model.Grids(iGrid).Pressure.Oil = Data.Model.Grids(iGrid).Pressure.Gas - Data.Model.Grids(iGrid).SCAL.pcgo
                            Data.Model.Grids(iGrid).Pressure.Water = Data.Model.Grids(iGrid).Pressure.Oil - Data.Model.Grids(iGrid).SCAL.pcow
                        Case Schedule_Type.Dynamic_Water_Pressure
                            Data.Model.Grids(iGrid).Pressure.Water = Interpolate(Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data, Data.Current_Time)
                            Data.Model.Grids(iGrid).Pressure.Oil = Data.Model.Grids(iGrid).Pressure.Water + Data.Model.Grids(iGrid).SCAL.pcow
                            Data.Model.Grids(iGrid).Pressure.Gas = Data.Model.Grids(iGrid).SCAL.pcgo + Data.Model.Grids(iGrid).Pressure.Oil
                        Case Schedule_Type.Static_Water_Saturation
                            Data.Model.Grids(iGrid).Saturation.Water = Data.Schedules(Data.Model.Grids(iGrid).Schedules(n)).Data(0, 0)


                    End Select

            End Select
        Next

    End Sub



    Sub Fp_Oil_Water_Gas(ByVal nGrid As Grid, ByVal mGrid As Grid, ByVal Method As Transmissibility_Fp_Averaging, ByRef Fp_Oil As Decimal, ByRef Fp_Water As Decimal, ByRef Fp_Gas As Decimal)
        Fp_Oil = Fp_Average(nGrid.PVT.Oil.Viscosity, mGrid.PVT.Oil.Viscosity, nGrid.PVT.Oil.FVF, mGrid.PVT.Oil.FVF, Method)
        Fp_Water = Fp_Average(nGrid.PVT.Water.Viscosity, mGrid.PVT.Water.Viscosity, nGrid.PVT.Water.FVF, mGrid.PVT.Water.FVF, Method)
        Fp_Gas = Fp_Average(nGrid.PVT.Gas.Viscosity, mGrid.PVT.Gas.Viscosity, nGrid.PVT.Gas.FVF, mGrid.PVT.Gas.FVF, Method)
    End Sub

    Function Fp_Average(ByVal mu1 As Decimal, ByVal mu2 As Decimal, ByVal b1 As Decimal, ByVal b2 As Decimal, ByVal method As Transmissibility_Fp_Averaging) As Decimal
        Select Case method
            Case Transmissibility_Fp_Averaging.Method1
                Return (1 / (((mu1 + mu2) / 2) * ((b1 + b2) / 2)))
            Case Transmissibility_Fp_Averaging.Method2
                Return ((1 / (mu1 * b1)) + (1 / (mu2 * b2))) / 2

        End Select
    End Function

    Function GetSurfaceFlow(ByVal Data As Simulation, ByVal G As Grid, ByVal Orientation As Adjacent_Type, ByVal P As Phase, ByVal outPressure As Decimal) As Decimal
        Dim Delta_P As Decimal
        Dim Area As Decimal
        Dim Delta_X As Decimal
        Dim K As Decimal
        Dim G_1 As New Grid
        With G_1
            .Pressure.Oil = G.Pressure.Oil
            .Pressure.Water = G.Pressure.Water
            .Pressure.Gas = G.Pressure.Gas
            Select Case P
                Case Phase.Oil
                    .Pressure.Oil = outPressure
                Case Phase.Water
                    .Pressure.Water = outPressure
                Case Phase.Gas
                    .Pressure.Gas = outPressure
            End Select
            .PVT.Temperature = G.PVT.Temperature
        End With

        Select Case P
            Case Phase.Oil
                Delta_P = G.Pressure.Oil - outPressure
                Select Case Orientation
                    Case Adjacent_Type.X
                        Area = G.Size.Z_Length * G.Size.Y_Length
                        If Delta_P > 0 Then
                            K = G.SCAL.Kro.Kx * Data.Rocks(G.RockType).Permeability.Kx
                        Else
                            K = Data.Rocks(G.RockType).SCAL.Tables.Water_Oil(1, 0) * Data.Rocks(G.RockType).Permeability.Kx
                        End If

                        Initialize_PVT(G_1, Data)
                        Dim Fp_Oil As Decimal = Fp_Average(G.PVT.Oil.Viscosity, G_1.PVT.Oil.Viscosity, G.PVT.Oil.FVF, G_1.PVT.Oil.FVF, Data.Methodology.Transmissibility_Fp_Averaging)
                        Return -(Data.Constants.Beta * K * Area * (Delta_P)) * (Fp_Oil / G.Size.X_Length)
                    Case Adjacent_Type.Y
                        Area = G.Size.Z_Length * G.Size.X_Length
                        If Delta_P > 0 Then
                            K = G.SCAL.Kro.Ky * Data.Rocks(G.RockType).Permeability.Ky
                        Else
                            K = Data.Rocks(G.RockType).SCAL.Tables.Water_Oil(1, 0) * Data.Rocks(G.RockType).Permeability.Ky
                        End If

                        Initialize_PVT(G_1, Data)
                        Dim Fp_Oil As Decimal = Fp_Average(G.PVT.Oil.Viscosity, G_1.PVT.Oil.Viscosity, G.PVT.Oil.FVF, G_1.PVT.Oil.FVF, Data.Methodology.Transmissibility_Fp_Averaging)
                        Return -(Data.Constants.Beta * K * Area * (Delta_P)) * (Fp_Oil / G.Size.Y_Length)
                    Case Adjacent_Type.Z
                        Area = G.Size.X_Length * G.Size.Y_Length
                        If Delta_P > 0 Then
                            K = G.SCAL.Kro.Kz * Data.Rocks(G.RockType).Permeability.Kz
                        Else
                            K = Data.Rocks(G.RockType).SCAL.Tables.Water_Oil(1, 0) * Data.Rocks(G.RockType).Permeability.Kz
                        End If

                        Initialize_PVT(G_1, Data)
                        Dim Fp_Oil As Decimal = Fp_Average(G.PVT.Oil.Viscosity, G_1.PVT.Oil.Viscosity, G.PVT.Oil.FVF, G_1.PVT.Oil.FVF, Data.Methodology.Transmissibility_Fp_Averaging)
                        Return -(Data.Constants.Beta * K * Area * (Delta_P)) * (Fp_Oil / G.Size.Z_Length)
                End Select
            Case Phase.Water
                Delta_P = G.Pressure.Water - outPressure
                Select Case Orientation
                    Case Adjacent_Type.X
                        Area = G.Size.Z_Length * G.Size.Y_Length
                        If Delta_P > 0 Then
                            K = G.SCAL.Krw.Kx * Data.Rocks(G.RockType).Permeability.Kx
                        Else
                            K = Data.Rocks(G.RockType).SCAL.Tables.Water_Oil(2, Data.Rocks(G.RockType).SCAL.Tables.Water_Oil.GetLength(1) - 1) * Data.Rocks(G.RockType).Permeability.Kx
                        End If

                        Initialize_PVT(G_1, Data)
                        Dim Fp_water As Decimal = Fp_Average(G.PVT.Water.Viscosity, G_1.PVT.Water.Viscosity, G.PVT.Water.FVF, G_1.PVT.Water.FVF, Data.Methodology.Transmissibility_Fp_Averaging)
                        Return -(Data.Constants.Beta * K * Area * (Delta_P)) * (Fp_water / G.Size.X_Length)
                    Case Adjacent_Type.Y
                        Area = G.Size.Z_Length * G.Size.X_Length
                        If Delta_P > 0 Then
                            K = G.SCAL.Krw.Ky * Data.Rocks(G.RockType).Permeability.Ky
                        Else
                            K = Data.Rocks(G.RockType).SCAL.Tables.Water_Oil(2, Data.Rocks(G.RockType).SCAL.Tables.Water_Oil.GetLength(1) - 1) * Data.Rocks(G.RockType).Permeability.Ky
                        End If

                        Initialize_PVT(G_1, Data)
                        Dim Fp_water As Decimal = Fp_Average(G.PVT.Water.Viscosity, G_1.PVT.Water.Viscosity, G.PVT.Water.FVF, G_1.PVT.Water.FVF, Data.Methodology.Transmissibility_Fp_Averaging)
                        Return -(Data.Constants.Beta * K * Area * (Delta_P)) * (Fp_water / G.Size.Y_Length)
                    Case Adjacent_Type.Z
                        Area = G.Size.X_Length * G.Size.Y_Length
                        If Delta_P > 0 Then
                            K = G.SCAL.Krw.Kz * Data.Rocks(G.RockType).Permeability.Kz
                        Else
                            K = Data.Rocks(G.RockType).SCAL.Tables.Water_Oil(2, Data.Rocks(G.RockType).SCAL.Tables.Water_Oil.GetLength(1) - 1) * Data.Rocks(G.RockType).Permeability.Kz
                        End If

                        Initialize_PVT(G_1, Data)
                        Dim Fp_water As Decimal = Fp_Average(G.PVT.Water.Viscosity, G_1.PVT.Water.Viscosity, G.PVT.Water.FVF, G_1.PVT.Water.FVF, Data.Methodology.Transmissibility_Fp_Averaging)
                        Return -(Data.Constants.Beta * K * Area * (Delta_P)) * (Fp_water / G.Size.Z_Length)
                End Select
            Case Phase.Gas
                Delta_P = G.Pressure.Gas - outPressure
                Select Case Orientation
                    Case Adjacent_Type.X
                        Area = G.Size.Z_Length * G.Size.Y_Length
                        If Delta_P > 0 Then
                            K = G.SCAL.Krg.Kx * Data.Rocks(G.RockType).Permeability.Kx
                        Else
                            K = Data.Rocks(G.RockType).SCAL.Tables.Oil_Gas(1, Data.Rocks(G.RockType).SCAL.Tables.Oil_Gas.GetLength(1) - 1) * Data.Rocks(G.RockType).Permeability.Kx
                        End If

                        Initialize_PVT(G_1, Data)
                        Dim Fp_gas As Decimal = Fp_Average(G.PVT.Gas.Viscosity, G_1.PVT.Gas.Viscosity, G.PVT.Gas.FVF, G_1.PVT.Gas.FVF, Data.Methodology.Transmissibility_Fp_Averaging)
                        Return -(Data.Constants.Beta * K * Area * (Delta_P)) * (Fp_gas / G.Size.X_Length)
                    Case Adjacent_Type.Y
                        Area = G.Size.Z_Length * G.Size.X_Length
                        If Delta_P > 0 Then
                            K = G.SCAL.Krg.Ky * Data.Rocks(G.RockType).Permeability.Ky
                        Else
                            K = Data.Rocks(G.RockType).SCAL.Tables.Oil_Gas(1, Data.Rocks(G.RockType).SCAL.Tables.Oil_Gas.GetLength(1) - 1) * Data.Rocks(G.RockType).Permeability.Ky
                        End If

                        Initialize_PVT(G_1, Data)
                        Dim Fp_gas As Decimal = Fp_Average(G.PVT.Gas.Viscosity, G_1.PVT.Gas.Viscosity, G.PVT.Gas.FVF, G_1.PVT.Gas.FVF, Data.Methodology.Transmissibility_Fp_Averaging)
                        Return -(Data.Constants.Beta * K * Area * (Delta_P)) * (Fp_gas / G.Size.Y_Length)
                    Case Adjacent_Type.Z
                        Area = G.Size.X_Length * G.Size.Y_Length
                        If Delta_P > 0 Then
                            K = G.SCAL.Krg.Kz * Data.Rocks(G.RockType).Permeability.Kz
                        Else
                            K = Data.Rocks(G.RockType).SCAL.Tables.Oil_Gas(1, Data.Rocks(G.RockType).SCAL.Tables.Oil_Gas.GetLength(1) - 1) * Data.Rocks(G.RockType).Permeability.Kz
                        End If

                        Initialize_PVT(G_1, Data)
                        Dim Fp_gas As Decimal = Fp_Average(G.PVT.Gas.Viscosity, G_1.PVT.Gas.Viscosity, G.PVT.Gas.FVF, G_1.PVT.Gas.FVF, Data.Methodology.Transmissibility_Fp_Averaging)
                        Return -(Data.Constants.Beta * K * Area * (Delta_P)) * (Fp_gas / G.Size.Z_Length)
                End Select
        End Select

    End Function

    Sub Initialize_Saturations(ByVal Data As Simulation, ByRef nGrid As Grid)
        nGrid.Saturation.Oil = 1 - (nGrid.Saturation.Water + nGrid.Saturation.Gas)
    End Sub

    Sub Initialize_Phase_System(ByRef nGrid As Grid, ByVal Rock As Rock)

        Dim swmin As Decimal
        Dim swmax As Decimal
        Dim sgmin As Decimal
        Dim sgmax As Decimal

        If SizeOf(Rock.SCAL.Tables.Water_Oil) > 0 Then
            swmin = Rock.SCAL.Tables.Water_Oil(0, 0)
            swmax = Rock.SCAL.Tables.Water_Oil(0, Rock.SCAL.Tables.Water_Oil.GetLength(1) - 1)
        End If
        If SizeOf(Rock.SCAL.Tables.Oil_Gas) > 0 Then
            sgmin = Rock.SCAL.Tables.Oil_Gas(0, 0)
            sgmax = Rock.SCAL.Tables.Oil_Gas(0, Rock.SCAL.Tables.Oil_Gas.GetLength(1) - 1)
        End If

        If ((nGrid.Saturation.Water <= swmin) And (nGrid.Saturation.Gas <= sgmin)) Then
            nGrid.System.Phase = Grid_Phase_System.Oil_Only
        ElseIf (nGrid.Saturation.Water >= swmax) Then
            nGrid.System.Phase = Grid_Phase_System.Water_Only
        ElseIf ((nGrid.Saturation.Water <= swmin) And (nGrid.Saturation.Gas >= sgmax)) Then
            nGrid.System.Phase = Grid_Phase_System.Gas_Only
        ElseIf ((nGrid.Saturation.Water > swmin) And (nGrid.Saturation.Water < swmax) And (nGrid.Saturation.Gas <= sgmin)) Then
            nGrid.System.Phase = Grid_Phase_System.Oil_Water
        ElseIf ((nGrid.Saturation.Water < swmin) And (nGrid.Saturation.Gas > sgmin) And (nGrid.Saturation.Gas <= sgmax)) Then
            nGrid.System.Phase = Grid_Phase_System.Oil_Gas
        ElseIf (((1 - nGrid.Saturation.Water - nGrid.Saturation.Gas) < (1 - sgmax)) And ((1 - nGrid.Saturation.Water - nGrid.Saturation.Gas) < (1 - swmax))) Then
            nGrid.System.Phase = Grid_Phase_System.Gas_Water
        Else
            nGrid.System.Phase = Grid_Phase_System.Oil_Gas_Water
        End If

    End Sub

    Sub Initialize_SCAL(ByRef nGrid As Grid, ByVal Rock As Rock, Optional ByVal Source As Integer = 0)

        Dim Columns() As Integer = {1, 2, 3, 4}

        nGrid.SCAL.Kro.Kx = 0
        nGrid.SCAL.Kro.Ky = 0
        nGrid.SCAL.Kro.Kz = 0

        nGrid.SCAL.Krw.Kx = 0
        nGrid.SCAL.Krw.Ky = 0
        nGrid.SCAL.Krw.Kz = 0

        nGrid.SCAL.Krg.Kx = 0
        nGrid.SCAL.Krg.Ky = 0
        nGrid.SCAL.Krg.Kz = 0

        With Rock

            Select Case nGrid.System.Phase
                Case Grid_Phase_System.Oil_Only

                    nGrid.SCAL.Kro.Kx = .SCAL.Tables.Water_Oil(1, 0)
                    nGrid.SCAL.Kro.Ky = .SCAL.Tables.Water_Oil(1, 0)
                    nGrid.SCAL.Kro.Kz = .SCAL.Tables.Water_Oil(1, 0)

                    nGrid.SCAL.pcow_drg = .SCAL.Tables.Water_Oil(3, 0)
                    nGrid.SCAL.pcow_imb = .SCAL.Tables.Water_Oil(4, 0)

                Case Grid_Phase_System.Gas_Only

                    nGrid.SCAL.Krg.Kx = .SCAL.Tables.Water_Oil(2, 0)
                    nGrid.SCAL.Krg.Ky = .SCAL.Tables.Water_Oil(2, 0)
                    nGrid.SCAL.Krg.Kz = .SCAL.Tables.Water_Oil(2, 0)

                Case Grid_Phase_System.Water_Only

                    nGrid.SCAL.Krw.Kx = .SCAL.Tables.Water_Oil(2, .SCAL.Tables.Water_Oil.GetLength(1) - 1)
                    nGrid.SCAL.Krw.Ky = .SCAL.Tables.Water_Oil(2, .SCAL.Tables.Water_Oil.GetLength(1) - 1)
                    nGrid.SCAL.Krw.Kz = .SCAL.Tables.Water_Oil(2, .SCAL.Tables.Water_Oil.GetLength(1) - 1)

                    nGrid.SCAL.pcow_drg = .SCAL.Tables.Water_Oil(3, .SCAL.Tables.Water_Oil.GetLength(1) - 1)
                    nGrid.SCAL.pcow_imb = .SCAL.Tables.Water_Oil(4, .SCAL.Tables.Water_Oil.GetLength(1) - 1)



                Case Grid_Phase_System.Oil_Water
                    Dim Water_Oil_Values() As Decimal
                    Water_Oil_Values = Linear_Interpolation_2D(nGrid.Saturation.Water, .SCAL.Tables.Water_Oil, 0, Columns)

                    nGrid.SCAL.Kro.Kx = Water_Oil_Values(0)
                    nGrid.SCAL.Kro.Ky = Water_Oil_Values(0)
                    nGrid.SCAL.Kro.Kz = Water_Oil_Values(0)

                    nGrid.SCAL.Krw.Kx = Water_Oil_Values(1)
                    nGrid.SCAL.Krw.Ky = Water_Oil_Values(1)
                    nGrid.SCAL.Krw.Kz = Water_Oil_Values(1)

                    nGrid.SCAL.pcow_drg = Water_Oil_Values(2)
                    nGrid.SCAL.pcow_imb = Water_Oil_Values(3)



                Case Grid_Phase_System.Oil_Gas
                    Dim Oil_Gas_Values() As Decimal
                    Oil_Gas_Values = Linear_Interpolation_2D(nGrid.Saturation.Gas, .SCAL.Tables.Oil_Gas, 0, Columns)

                    nGrid.SCAL.Kro.Kx = Oil_Gas_Values(0)
                    nGrid.SCAL.Kro.Ky = Oil_Gas_Values(0)
                    nGrid.SCAL.Kro.Kz = Oil_Gas_Values(0)

                    nGrid.SCAL.Krg.Kx = Oil_Gas_Values(1)
                    nGrid.SCAL.Krg.Ky = Oil_Gas_Values(1)
                    nGrid.SCAL.Krg.Kz = Oil_Gas_Values(1)

                    nGrid.SCAL.pcgo_drg = Oil_Gas_Values(2)
                    nGrid.SCAL.pcgo_imb = Oil_Gas_Values(3)


                Case Grid_Phase_System.Gas_Water
                    Dim Water_Gas_Values() As Decimal
                    'If SizeOf(.SCAL.Tables.Gas_Water) > 0 Then
                    '    Water_Gas_Values = Linear_Interpolation_2D(nGrid.Saturation.Water, .SCAL.Tables.Gas_Water, 0, Columns)

                    '    nGrid.SCAL.Krg.Kx = Water_Gas_Values(0)
                    '    nGrid.SCAL.Krg.Ky = Water_Gas_Values(0)
                    '    nGrid.SCAL.Krg.Kz = Water_Gas_Values(0)

                    '    nGrid.SCAL.Krw.Kx = Water_Gas_Values(1)
                    '    nGrid.SCAL.Krw.Ky = Water_Gas_Values(1)
                    '    nGrid.SCAL.Krw.Kz = Water_Gas_Values(1)
                    'End If

                Case Grid_Phase_System.Oil_Gas_Water


            End Select



            Select Case Rock.SCAL.Wettability.Water_Oil
                Case Phase.Water
                    If nGrid.Delta.Water > nGrid.Delta.Oil Then
                        nGrid.SCAL.pcow = nGrid.SCAL.pcow_imb
                    Else
                        nGrid.SCAL.pcow = nGrid.SCAL.pcow_drg
                    End If

                Case Phase.Oil
                    If nGrid.Delta.Water > nGrid.Delta.Oil Then
                        nGrid.SCAL.pcow = nGrid.SCAL.pcow_drg
                    Else
                        nGrid.SCAL.pcow = nGrid.SCAL.pcow_imb
                    End If
            End Select


            Select Case Rock.SCAL.Wettability.Oil_Gas
                Case Phase.Oil
                    If nGrid.Delta.Oil > nGrid.Delta.Gas Then
                        nGrid.SCAL.pcgo = nGrid.SCAL.pcgo_imb
                    Else
                        nGrid.SCAL.pcgo = nGrid.SCAL.pcgo_drg
                    End If
                Case Phase.Gas
                    If nGrid.Delta.Oil > nGrid.Delta.Gas Then
                        nGrid.SCAL.pcgo = nGrid.SCAL.pcgo_drg
                    Else
                        nGrid.SCAL.pcgo = nGrid.SCAL.pcgo_imb
                    End If
            End Select

            If nGrid.SCAL.pcow = 0 Then
                nGrid.SCAL.pcow = (nGrid.SCAL.pcow_drg + nGrid.SCAL.pcow_imb) / 2
            End If
            If nGrid.SCAL.pcgo = 0 Then
                nGrid.SCAL.pcgo = (nGrid.SCAL.pcgo_drg + nGrid.SCAL.pcgo_imb) / 2
            End If

        End With
    End Sub

    Sub Initialize_Pressure(ByRef nGrid As Grid)
        nGrid.Pressure.Gas = nGrid.SCAL.pcgo + nGrid.Pressure.Oil
        nGrid.Pressure.Water = nGrid.Pressure.Oil - nGrid.SCAL.pcow
    End Sub

    Sub Initialize_Porosity(ByRef Data As Simulation, ByVal iGrid As Integer)
        ' Assumed Constatnt
    End Sub


    Sub Initialize_PVT(ByRef nGrid As Grid, ByVal Data As Simulation)
        With nGrid

            Dim vals() As Decimal

            If Data.Methodology.PVT.Oil.FVF = PVT_Methods.Table Or Data.Methodology.PVT.Oil.Viscosity = PVT_Methods.Table Or Data.Methodology.PVT.Oil.Density = PVT_Methods.Table Or Data.Methodology.PVT.Oil.Compressibility = PVT_Methods.Table Or Data.Methodology.PVT.Rs = PVT_Methods.Table Then
                Dim oil_columns() As Integer = {1, 2, 5, 8, 11}
                vals = Linear_Interpolation_2D(.Pressure.Oil, Data.PVT.Table, 0, oil_columns)
            End If

            If Data.Methodology.PVT.Rs = PVT_Methods.Table Then
                .PVT.Rs = vals(0)
            Else
                .PVT.Rs = PVT_Correlation(Data.Methodology.PVT.Rs, nGrid, Data)
            End If
            If Data.Methodology.PVT.Oil.FVF = PVT_Methods.Table Then
                .PVT.Oil.FVF = vals(1)
            Else
                .PVT.Oil.FVF = PVT_Correlation(Data.Methodology.PVT.Oil.FVF, nGrid, Data)
            End If
            If Data.Methodology.PVT.Oil.Density = PVT_Methods.Table Then
                .PVT.Oil.Density = vals(2)
            Else
                .PVT.Oil.Density = PVT_Correlation(Data.Methodology.PVT.Oil.Density, nGrid, Data)
            End If
            If Data.Methodology.PVT.Oil.Viscosity = PVT_Methods.Table Then
                .PVT.Oil.Viscosity = vals(3)
            Else
                .PVT.Oil.Viscosity = PVT_Correlation(Data.Methodology.PVT.Oil.Viscosity, nGrid, Data)
            End If
            If Data.Methodology.PVT.Oil.Compressibility = PVT_Methods.Table Then
                .PVT.Oil.Compressibility = vals(4)
            Else
                .PVT.Oil.Compressibility = PVT_Correlation(Data.Methodology.PVT.Oil.Compressibility, nGrid, Data)
            End If



            If Data.Methodology.PVT.Gas.FVF = PVT_Methods.Table Or Data.Methodology.PVT.Gas.Viscosity = PVT_Methods.Table Or Data.Methodology.PVT.Gas.Density = PVT_Methods.Table Or Data.Methodology.PVT.Gas.Compressibility = PVT_Methods.Table Then
                Dim gas_columns() As Integer = {3, 6, 9, 12}
                vals = Linear_Interpolation_2D(.Pressure.Gas, Data.PVT.Table, 0, gas_columns)
            End If

            If Data.Methodology.PVT.Gas.FVF = PVT_Methods.Table Then
                .PVT.Gas.FVF = vals(0)
            Else
                .PVT.Gas.FVF = PVT_Correlation(Data.Methodology.PVT.Gas.FVF, nGrid, Data)
            End If
            If Data.Methodology.PVT.Gas.Density = PVT_Methods.Table Then
                .PVT.Gas.Density = vals(1)
            Else
                .PVT.Gas.Density = PVT_Correlation(Data.Methodology.PVT.Gas.Density, nGrid, Data)
            End If
            If Data.Methodology.PVT.Gas.Viscosity = PVT_Methods.Table Then
                .PVT.Gas.Viscosity = vals(2)
            Else
                .PVT.Gas.Viscosity = PVT_Correlation(Data.Methodology.PVT.Gas.Viscosity, nGrid, Data)
            End If
            If Data.Methodology.PVT.Gas.Compressibility = PVT_Methods.Table Then
                .PVT.Gas.Compressibility = vals(3)
            Else
                .PVT.Gas.Compressibility = PVT_Correlation(Data.Methodology.PVT.Gas.Compressibility, nGrid, Data)
            End If


            If Data.Methodology.PVT.Water.FVF = PVT_Methods.Table Or Data.Methodology.PVT.Water.Viscosity = PVT_Methods.Table Or Data.Methodology.PVT.Water.Density = PVT_Methods.Table Or Data.Methodology.PVT.Water.Compressibility = PVT_Methods.Table Then
                Dim water_columns() As Integer = {4, 7, 10, 13}
                vals = Linear_Interpolation_2D(.Pressure.Water, Data.PVT.Table, 0, water_columns)
            End If

            If Data.Methodology.PVT.Water.FVF = PVT_Methods.Table Then
                .PVT.Water.FVF = vals(0)
            Else
                .PVT.Water.FVF = PVT_Correlation(Data.Methodology.PVT.Water.FVF, nGrid, Data)
            End If
            If Data.Methodology.PVT.Water.Density = PVT_Methods.Table Then
                .PVT.Water.Density = vals(1)
            Else
                .PVT.Water.Density = PVT_Correlation(Data.Methodology.PVT.Water.Density, nGrid, Data)
            End If
            If Data.Methodology.PVT.Water.Viscosity = PVT_Methods.Table Then
                .PVT.Water.Viscosity = vals(2)
            Else
                .PVT.Water.Viscosity = PVT_Correlation(Data.Methodology.PVT.Water.Viscosity, nGrid, Data)
            End If
            If Data.Methodology.PVT.Water.Compressibility = PVT_Methods.Table Then
                .PVT.Water.Compressibility = vals(3)
            Else
                .PVT.Water.Compressibility = PVT_Correlation(Data.Methodology.PVT.Water.Compressibility, nGrid, Data)
            End If





        End With


    End Sub

    Function PVT_Correlation(ByVal Method As PVT_Methods, ByVal G As Grid, ByVal Data As Simulation) As Decimal
        Dim Pstr As Decimal
        Select Case Data.System.Units
            Case System_of_Units.Field
                Pstr = 14.65
            Case System_of_Units.Metric
                Pstr = 100
        End Select

        Select Case Method

            ' FVF Correlations ============================================================================================
            Case PVT_Methods.Oil_FVF_1
                Return (Math.E ^ (-Data.PVT.Oil_At_Standard.Compressibility * (G.Pressure.Oil - Pstr)))
            Case PVT_Methods.Water_FVF_1
                Return (Math.E ^ (-Data.PVT.Water_At_Standard.Compressibility * (G.Pressure.Water - Pstr)))
            Case PVT_Methods.Gas_FVF_1
                If G.Pressure.Gas < Data.PVT.Gas_Bubble_Pressure Then
                    Dim Z As Decimal = z_factor_Shell(G.Pressure.Gas, G.PVT.Temperature, Data.PVT.Gas_Critical_Pressure, Data.PVT.Gas_Critical_Temperature)

                    Select Case Data.System.Units
                        Case System_of_Units.Field
                            Return 0.00502 * ((Z * G.PVT.Temperature) / G.Pressure.Gas)
                        Case System_of_Units.Metric
                            Return ((0.00502 * (((Z * G.PVT.Temperature) / Simulation_Constants.Temperature) / (G.Pressure.Gas / Simulation_Constants.Pressure)))) / Simulation_Constants.BgConversion
                    End Select



                Else
                    Dim Z As Decimal = z_factor_Shell(Data.PVT.Gas_Bubble_Pressure, G.PVT.Temperature, Data.PVT.Gas_Critical_Pressure, Data.PVT.Gas_Critical_Temperature)
                    Dim Bg_Bubble As Decimal
                    Dim Bg As Decimal
                    Select Case Data.System.Units
                        Case System_of_Units.Field
                            Bg_Bubble = 0.00502 * ((Z * G.PVT.Temperature) / Data.PVT.Gas_Bubble_Pressure)
                        Case System_of_Units.Metric
                            Bg_Bubble = (0.00502 * (((Z * G.PVT.Temperature) / Simulation_Constants.Temperature) / (Data.PVT.Gas_Bubble_Pressure / Simulation_Constants.Pressure))) / Simulation_Constants.BgConversion
                    End Select

                    Bg = Bg_Bubble * (Math.E ^ (-Data.PVT.Condesate.Compressibility * (G.Pressure.Gas - Data.PVT.Gas_Bubble_Pressure)))
                    Return Bg

                End If
                '==========================================================================================================


                ' Density Correlations ============================================================================================
            Case PVT_Methods.Oil_Density_1
                Return Data.PVT.Oil_At_Standard.Density * (Math.E ^ (Data.PVT.Oil_At_Standard.Compressibility * (G.Pressure.Oil - Pstr)))
            Case PVT_Methods.Water_Density_1
                Return Data.PVT.Water_At_Standard.Density * (Math.E ^ (Data.PVT.Water_At_Standard.Compressibility * (G.Pressure.Water - Pstr)))
            Case PVT_Methods.Gas_Density_1
                Dim Bg As Decimal = PVT_Correlation(Data.Methodology.PVT.Gas.FVF, G, Data) * Simulation_Constants.Volume_Conversion_Factor
                Select Case Data.System.Units
                    Case System_of_Units.Field
                        Return Data.PVT.Gas_At_Standard.Density / Bg
                    Case System_of_Units.Metric
                        Return (Data.PVT.Gas_At_Standard.Density / Simulation_Constants.Phase_Density) / Bg
                End Select
                '==========================================================================================================


                ' Viscosity Correlations ============================================================================================
            Case PVT_Methods.Oil_Viscosity_1
                Dim T As Decimal = G.PVT.Temperature
                Select Case Data.System.Units
                    Case System_of_Units.Field
                        T = T - 460
                    Case System_of_Units.Metric
                        T = ((T - 273) * 18 + 32)
                End Select

                Dim y As Decimal = 10 ^ (3.0324 - 0.02023 * API(G.PVT.Oil.Gamma_Water(Data.System.Units), Data.System.Units))
                Select Case Data.System.Units
                    Case System_of_Units.Field
                        Return (10 ^ (y * (T ^ (-1.163)))) - 1
                    Case System_of_Units.Metric
                        Return ((10 ^ (y * (T ^ (-1.163)))) - 1) / Simulation_Constants.Viscosity
                End Select


            Case PVT_Methods.Water_Viscosity_1
                Dim T As Decimal = G.PVT.Temperature
                Select Case Data.System.Units
                    Case System_of_Units.Field
                        T = T - 460
                    Case System_of_Units.Metric
                        T = ((T - 273) * 18 + 32)
                End Select

                Select Case Data.System.Units
                    Case System_of_Units.Field
                        Return Math.E ^ (1.003 - 0.01479 * T + 0.00001982 * (T ^ 2))
                    Case System_of_Units.Metric
                        Return (Math.E ^ (1.003 - 0.01479 * T + 0.00001982 * (T ^ 2))) / Simulation_Constants.Viscosity
                End Select


            Case PVT_Methods.Gas_Viscosity_1
                Dim T As Decimal = G.PVT.Temperature
                Select Case Data.System.Units
                    Case System_of_Units.Metric
                        T = (((T - 273) * 18 + 32) + 460)
                End Select

                Dim M As Decimal = Data.PVT.Gas_Molecular_Weight
                Dim x As Decimal = 3.448 + 986.4 / T + (0.01009 * M)
                Dim y As Decimal = 2.447 - 0.2224 * x
                Dim K As Decimal = ((9.379 + 0.01607 * M) * (T ^ 1.5)) / (209.2 + 19.26 * M + T)
                Dim rho_g As Decimal = PVT_Correlation(Data.Methodology.PVT.Gas.Density, G, Data)

                Select Case Data.System.Units
                    Case System_of_Units.Field
                        rho_g = (rho_g * Simulation_Constants.Phase_Density) / 1000
                    Case System_of_Units.Metric
                        rho_g = rho_g / 1000
                End Select

                Dim mu As Decimal = K * (0.0001) * (Math.E ^ (x * (rho_g ^ y)))



                Select Case Data.System.Units
                    Case System_of_Units.Field
                        Return mu
                    Case System_of_Units.Metric
                        Return mu / Simulation_Constants.Viscosity
                End Select


                '==========================================================================================================

            Case PVT_Methods.Oil_Viscosity_Static

                Return Data.PVT.Oil_At_Standard.Viscosity
            Case PVT_Methods.Water_Viscosity_Static
                Return Data.PVT.Water_At_Standard.Viscosity
            Case PVT_Methods.Gas_Viscosity_Static
                Return Data.PVT.Gas_At_Standard.Viscosity

            Case PVT_Methods.Oil_Compressibility_Static
                Return Data.PVT.Oil_At_Standard.Compressibility
            Case PVT_Methods.Water_Compressibility_Static
                Return Data.PVT.Water_At_Standard.Compressibility
            Case PVT_Methods.Gas_Compressibility_Static
                Return Data.PVT.Gas_At_Standard.Compressibility

            Case PVT_Methods.RS_1
                Return 0

        End Select


    End Function
    Function z_factor_Shell(ByVal Pressure As Decimal, ByVal Temperature As Decimal, ByVal Cp As Decimal, ByVal Ct As Decimal) As Decimal
        Dim Pr As Decimal = Pressure / Cp
        Dim Tr As Decimal = Temperature / Ct

        Dim Z, ZA, ZB, ZC, ZD, ZE, ZG, ZF As Decimal

        ZA = -0.101 - 0.36 * Tr + 1.3868 * Math.Sqrt(Tr - 0.919)
        ZB = 0.021 + (0.04275 / (Tr - 0.65))
        ZE = 0.32 * Math.E ^ (-19.53 * (Tr - 1))
        ZF = 0.122 * Math.E ^ (-11.3 * (Tr - 1))
        ZD = (0.0657 / (Tr - 0.86)) - 0.037
        ZC = 0.6222 - 0.224 * Tr
        ZG = Pr * (ZC + ZD * Pr + ZE * (Pr ^ 4))


        Z = ZA + ZB * Pr + (1 - ZA) * (Math.E ^ (-ZG)) - ZF * ((Pr / 10) ^ 4)
        Return Z
    End Function

    Function API(ByVal Gamma As Decimal, ByVal System As System_of_Units) As Decimal
        Select Case System
            Case System_of_Units.Field
                Return (141.5 / Gamma) - 131.5
            Case System_of_Units.Metric
                Return (141.5 / (Gamma * Simulation_Constants.Phase_Gravity)) - 131.5
        End Select

    End Function
    Sub Initialize_C_Bar_n(ByRef Data As Simulation, ByVal iGrid As Integer)
        If SizeOf(Data.C_Bar_n) = 0 Then ReDim Data.C_Bar_n(0, SizeOf(Data.Model.Grids) * 3 - 1)
        With Data.Model.Grids(iGrid)
            Data.C_Bar_n(0, iGrid * 3) = (.Volume / (Data.Constants.Alpha * Data.Time_Step)) * _
            ((.Porosity * .Saturation.Water) / .PVT.Water.FVF)

            Data.C_Bar_n(0, iGrid * 3 + 1) = (.Volume / (Data.Constants.Alpha * Data.Time_Step)) * _
            ((.Porosity * .Saturation.Gas / .PVT.Gas.FVF) + (.Porosity * .PVT.Rs * _
            (1 - .Saturation.Water - .Saturation.Gas) / .PVT.Oil.FVF))

            Data.C_Bar_n(0, iGrid * 3 + 2) = (.Volume / (Data.Constants.Alpha * Data.Time_Step)) _
            * ((.Porosity * (1 - .Saturation.Water - .Saturation.Gas)) / .PVT.Oil.FVF)
        End With
    End Sub


End Module
