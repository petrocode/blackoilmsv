
'                            Black Oil Reservoir Simulator
'                         Optimized for Core Scale Simulation
'                                 By: Ameer Emamjomeh
'                              amir.emamjomeh@gmail.com

' 
' 3D-3Phase Black-Oil Model
' Multi-Size Multi-Neighbor Gridding
' Fully-Implicit Method
' Implementation of RockTypes 
' Capillary Pressure Hysteresis
' 3D-3Phase Permeability
' Dynamic Porosity
' Time-Step Schedule
' Integrated Modification\Simulation\Visualization Environment
' Simultanious Simulation and Visualization
' Grid Internal/Surface Schedule
' 3D Visualization Using DirectXs

' ======================================== Project Schedule ==========================================
' Date                           Description                                              Module
' September 15 2010              Project Start
' September 16 2010              Data Structures and Initialization Functions           - Preperation 
' September 17 2010              SCAL & PVT Functions                                   - Initialize 
' September 18 2010              Orthomin Solution - Matrix Computations                - Mathematics 
' September 19 2010              Writing Simulator's Core                               - Simulator and Others
' September 23 2010              First Run                                              
' September 24 2010              Test and Debug                                         
' September 28 2010              First Exact Output



Public Class Core


    Public Event Iteration(ByVal Data_n As Simulation, ByVal Data_v As Simulation, ByVal IM As Implicit_Matrix, ByVal MSE As Iteration_MSE)
    Public Event Time_Step(ByVal History As Stack(Of Simulation), ByVal Data_n As Simulation, ByVal Total_Time As Double, ByVal Iterations As Integer)
    Public Event Simulation_End(ByVal History As Stack(Of Simulation), ByVal Total_Time As Double, ByVal Iterations As Integer)


    Private Paused As Boolean
    Public History As New Stack(Of Simulation)

    '======================================= Main Subroutine of Simulator =================================='
    '-------------------------------------------------------------------------------------------------------'

    Sub Run(ByRef Data As Simulation)

        Paused = False

        Dim Total_Time As Double
        Dim IM As Implicit_Matrix
        Dim Implicit_MSE As Iteration_MSE


        ' Preparing Simulator  
        Set_System_of_Units(Data)                                                'Setting Alpha and Beta for Unit Conversions
        Initialize_Grids_Porosities(Data)                                        'Initializing Grid Porosities Based on Their Rock Types
        Calculate_Grids_Volume(Data)                                             'Setting Grid Volumes Using Their Size Parameters
        Prepare_Adjacents(Data)                                                  'Locating Adjacent Grids and Calculating Their Common Surface Area
        Prepare_Implicit_Matrix(Data, IM)                                        'Setting Matrix Dimensions Based on Grid Numbers


        'Time Step Level 
        Do While (Total_Time < Data.Finish_Time And Data.Current_Step < Data.Final_Step) And Not Paused

            Data.Current_Step += 1                                               'Time Step Counter
            IM.Iteration = 0                                                     'Iteration Counter Initialize   
            Initialize(Data)                                                     'Calculating Model Parameters for Current Sw, Sg and Po
            History.Push(Simulation_Copy(Data))                                  'Saving a Copy of Model at Current Step
            Implicit_MSE = Data.Exactness.Iteration                              'Initializing Implicit MSE Goal

            'Iteration Level
            Do While (Implicit_MSE.Sw >= Data.Exactness.Iteration.Sw Or Implicit_MSE.Sg >= Data.Exactness.Iteration.Sg Or Implicit_MSE.Po >= Data.Exactness.Iteration.Po) And Not Paused

                Create_Implicit_Matrix(History.First, Data, IM, History)         'Preparing Equations for Calculating Current Deflection from n+1 Time Step
                Solve(IM.J, IM.Delta_X, IM.Mines_R, 10, Data.Exactness.Equation) 'Solving Equations Matrix
                Implicit_MSE = Calculate_Implicit_Iteration_MSE(IM)              'Calculating Current Iterative Error
                Update_Simulatin_Data(Data, IM)                                  'Setting New Sw, Sg and Po Values for k+1 Iteration
                Initialize(Data)                                                 'Updating Model for k+1 Iteration

                RaiseEvent Iteration(History.First, Data, IM, Implicit_MSE)      'Iteration Progress Event   

            Loop

            Total_Time += Data.Time_Step                                         'Cummulative Passed Time
            RaiseEvent Time_Step(History, Data, Total_Time, IM.Iteration)        'Time Step Progress Event
        Loop

        RaiseEvent Simulation_End(History, Total_Time, IM.Iteration)             'End of Simulation Event

    End Sub


    Function Simulation_Copy(ByVal Data As Simulation) As Simulation
        Dim out As New Simulation
        With out
            .Exactness.Iteration = Data.Exactness.Iteration
            .Exactness.Equation = Data.Exactness.Equation

            .Methodology.Derivations.Gas_Saturation_Epsilon = Data.Methodology.Derivations.Gas_Saturation_Epsilon
            .Methodology.Derivations.Water_Saturation_Epsilon = Data.Methodology.Derivations.Water_Saturation_Epsilon
            .Methodology.Derivations.Oil_Pressure_Epsilon = Data.Methodology.Derivations.Oil_Pressure_Epsilon

            .System.Phase.Water = Data.System.Phase.Water
            .System.Phase.Oil = Data.System.Phase.Oil
            .System.Phase.Gas = Data.System.Phase.Gas

            .System.Units = Data.System.Units

            .Current_Step = Data.Current_Step
            ReDim .Time_Steps(SizeOf(Data.Time_Steps) - 1)
            For i As Integer = 0 To SizeOf(Data.Time_Steps) - 1
                .Time_Steps(i) = Data.Time_Steps(i)
            Next

            .Alpha = Data.Alpha
            .Beta = Data.Beta

            .gravity_acceleration_field = Data.gravity_acceleration_field
            .gravity_acceleration_metric = Data.gravity_acceleration_metric

            ReDim .Grids(SizeOf(Data.Grids) - 1)
            For i As Integer = 0 To SizeOf(.Grids) - 1
                .Grids(i).Porosity = Data.Grids(i).Porosity
                .Grids(i).System = Data.Grids(i).System

                .Grids(i).Saturation.Water = Data.Grids(i).Saturation.Water
                .Grids(i).Saturation.Oil = Data.Grids(i).Saturation.Oil
                .Grids(i).Saturation.Gas = Data.Grids(i).Saturation.Gas

                .Grids(i).SCAL.pcgo = Data.Grids(i).SCAL.pcgo
                .Grids(i).SCAL.pcow = Data.Grids(i).SCAL.pcow

                .Grids(i).SCAL.Water.Kx = Data.Grids(i).SCAL.Water.Kx
                .Grids(i).SCAL.Water.Ky = Data.Grids(i).SCAL.Water.Ky
                .Grids(i).SCAL.Water.Kz = Data.Grids(i).SCAL.Water.Kz

                .Grids(i).SCAL.Oil.Kx = Data.Grids(i).SCAL.Oil.Kx
                .Grids(i).SCAL.Oil.Ky = Data.Grids(i).SCAL.Oil.Ky
                .Grids(i).SCAL.Oil.Kz = Data.Grids(i).SCAL.Oil.Kz

                .Grids(i).SCAL.Gas.Kx = Data.Grids(i).SCAL.Gas.Kx
                .Grids(i).SCAL.Gas.Ky = Data.Grids(i).SCAL.Gas.Ky
                .Grids(i).SCAL.Gas.Kz = Data.Grids(i).SCAL.Gas.Kz

                .Grids(i).Size.X_Length = Data.Grids(i).Size.X_Length
                .Grids(i).Size.Y_Length = Data.Grids(i).Size.Y_Length
                .Grids(i).Size.Z_Length = Data.Grids(i).Size.Z_Length

                .Grids(i).RockType = Data.Grids(i).RockType
                .Grids(i).Volume = Data.Grids(i).Volume


                .Grids(i).PVT.Rs = Data.Grids(i).PVT.Rs

                .Grids(i).PVT.Water.Viscosity = Data.Grids(i).PVT.Water.Viscosity
                .Grids(i).PVT.Water.Gamma = Data.Grids(i).PVT.Water.Gamma
                .Grids(i).PVT.Water.FVF = Data.Grids(i).PVT.Water.FVF
                .Grids(i).PVT.Water.Density = Data.Grids(i).PVT.Water.Density

                .Grids(i).PVT.Oil.Viscosity = Data.Grids(i).PVT.Oil.Viscosity
                .Grids(i).PVT.Oil.Gamma = Data.Grids(i).PVT.Oil.Gamma
                .Grids(i).PVT.Oil.FVF = Data.Grids(i).PVT.Oil.FVF
                .Grids(i).PVT.Oil.Density = Data.Grids(i).PVT.Oil.Density

                .Grids(i).PVT.Gas.Viscosity = Data.Grids(i).PVT.Gas.Viscosity
                .Grids(i).PVT.Gas.Gamma = Data.Grids(i).PVT.Gas.Gamma
                .Grids(i).PVT.Gas.FVF = Data.Grids(i).PVT.Gas.FVF
                .Grids(i).PVT.Gas.Density = Data.Grids(i).PVT.Gas.Density

                .Grids(i).Pressure.Water = Data.Grids(i).Pressure.Water
                .Grids(i).Pressure.Oil = Data.Grids(i).Pressure.Oil
                .Grids(i).Pressure.Gas = Data.Grids(i).Pressure.Gas

                .Grids(i).Location.X = Data.Grids(i).Location.X
                .Grids(i).Location.Y = Data.Grids(i).Location.Y
                .Grids(i).Location.Z = Data.Grids(i).Location.Z
            Next

            ReDim .C_Bar_n(0, Data.C_Bar_n.GetLength(1) - 1)
            For i As Integer = 0 To Data.C_Bar_n.GetLength(1) - 1
                .C_Bar_n(0, i) = Data.C_Bar_n(0, i)
            Next


            .Finish_Time = Data.Finish_Time

        End With
        Return out
    End Function

    Sub Update_Simulatin_Data(ByRef Data As Simulation, ByVal IM As Implicit_Matrix)
        For i As Integer = 0 To Data.Grids.Length - 1
            Data.Grids(i).Saturation.Water += IM.Delta_X(0, i * 3)
            Data.Grids(i).Saturation.Gas += IM.Delta_X(0, i * 3 + 1)
            Data.Grids(i).Pressure.Oil += IM.Delta_X(0, i * 3 + 2)
        Next
    End Sub

    Function Calculate_Implicit_Iteration_MSE(ByVal IM As Implicit_Matrix) As Iteration_MSE
        Static Last_IM As Implicit_Matrix

        Dim Sw_MSE, Sg_MSE, Po_MSE As Double
        Dim out As New Iteration_MSE

        For i As Integer = 0 To IM.Delta_X.GetLength(1) / 3 - 1
            Sw_MSE += Math.Abs(IM.Delta_X(0, i * 3))
            Sg_MSE += Math.Abs(IM.Delta_X(0, i * 3 + 1))
            Po_MSE += Math.Abs(IM.Delta_X(0, i * 3 + 2))
        Next
        out.Sw = Sw_MSE / (IM.Delta_X.GetLength(1) / 3)
        out.Sg = Sg_MSE / (IM.Delta_X.GetLength(1) / 3)
        out.Po = Po_MSE / (IM.Delta_X.GetLength(1) / 3)
        Return out

    End Function
    Sub Copy_Deflection(ByVal IM As Implicit_Matrix)
        Dim str As String
        For i As Integer = 0 To IM.Delta_X.Length - 1
            str += IM.Delta_X(0, i).ToString + Chr(Keys.Tab)
        Next
        Clipboard.SetText(str)

    End Sub
    Sub Copy_Matrix(ByVal IM As Implicit_Matrix)
        Dim str As String
        For j As Integer = 0 To IM.J.GetLength(1) - 1
            For i As Integer = 0 To IM.J.GetLength(0) - 1
                str += IM.J(i, j).ToString '+ Chr(Keys.Tab)
            Next
            'str += vbCrLf
        Next

        Clipboard.SetText(str)

    End Sub
    Sub Create_Implicit_Matrix(ByVal Data_n0 As Simulation, ByRef Data_n1 As Simulation, ByRef IM As Implicit_Matrix, ByVal History As Stack(Of Simulation))

        IM_Zero(IM)
        IM.Iteration += 1
        ' Grid Loop =====================================================================
        For nGrid As Integer = 0 To Data_n1.Grids.Length - 1

            Dim Grid_Params As New Grid_Parameters

            For iAdjacent As Integer = 0 To Data_n1.Grids(nGrid).Adjacents.Length - 1

                Calculate_F_Bracet_Parameters(Data_n0, Data_n1, nGrid, iAdjacent, Grid_Params)
                Add_F_Bracet_Parameters_m(Data_n1, nGrid, iAdjacent, Grid_Params, IM)
                Add_F_Bracet_Parameters_n(Data_n1, nGrid, iAdjacent, Grid_Params)
                Add_Mines_R_Bar_Parameters(Data_n1, nGrid, iAdjacent, Grid_Params, IM)

            Next

            Calculate_C_Bracet(Data_n0, Data_n1, nGrid, Grid_Params)
            Add_Bracet_Parameters_n(Data_n1, nGrid, Grid_Params, IM)
            Add_C_Bar(Data_n0, Data_n1, nGrid, Grid_Params, IM)
            Add_Q_Bar(Data_n1, nGrid, IM)

        Next


    End Sub
    Sub Add_Q_Bar(ByVal Data_v As Simulation, ByVal iGrid As Integer, ByRef IM As Implicit_Matrix)
        With IM
            .Mines_R(0, iGrid * 3) -= Data_v.Grids(iGrid).Flow.Water
            .Mines_R(0, iGrid * 3 + 1) -= Data_v.Grids(iGrid).Flow.Gas
            .Mines_R(0, iGrid * 3 + 2) -= Data_v.Grids(iGrid).Flow.Oil
        End With
    End Sub
    Sub Add_C_Bar(ByVal Data_n As Simulation, ByVal Data_v As Simulation, ByVal iGrid As Integer, ByRef Grid_Params As Grid_Parameters, ByRef IM As Implicit_Matrix)
        With IM
            .Mines_R(0, iGrid * 3) += (((Data_v.Grids(iGrid).Volume / (Data_v.Alpha * Data_n.Time_Step)) * _
            ((Data_v.Grids(iGrid).Porosity * Data_v.Grids(iGrid).Saturation.Water) / Data_v.Grids(iGrid).PVT.Water.FVF)) - _
            Data_n.C_Bar_n(0, iGrid * 3))

            .Mines_R(0, iGrid * 3 + 1) += (((Data_v.Grids(iGrid).Volume / (Data_v.Alpha * Data_n.Time_Step)) * _
            ((Data_v.Grids(iGrid).Porosity * Data_v.Grids(iGrid).Saturation.Gas / Data_v.Grids(iGrid).PVT.Gas.FVF) + _
            (Data_v.Grids(iGrid).Porosity * Data_v.Grids(iGrid).PVT.Rs * (1 - Data_v.Grids(iGrid).Saturation.Water - _
            Data_v.Grids(iGrid).Saturation.Gas) / _
            Data_v.Grids(iGrid).PVT.Oil.FVF))) - _
            Data_n.C_Bar_n(0, iGrid * 3 + 1))

            .Mines_R(0, iGrid * 3 + 2) += (((Data_v.Grids(iGrid).Volume / (Data_n.Alpha * Data_n.Time_Step)) * _
            (Data_v.Grids(iGrid).Porosity * (1 - Data_v.Grids(iGrid).Saturation.Water - Data_v.Grids(iGrid).Saturation.Gas) _
            / Data_v.Grids(iGrid).PVT.Oil.FVF)) - _
            Data_n.C_Bar_n(0, iGrid * 3 + 2))
        End With

    End Sub

    Sub Add_Mines_R_Bar_Parameters(ByVal Data As Simulation, ByVal iGrid As Integer, ByVal iAdjacent As Integer, ByVal Grid_Params As Grid_Parameters, ByRef IM As Implicit_Matrix)
        Dim adjID As Integer = Data.Grids(iGrid).Adjacents(iAdjacent).Target_Grid_Index
        With IM
            .Mines_R(0, iGrid * 3) -= (Grid_Params.R.Tw_nm * ((Data.Grids(adjID).Pressure.Oil - Data.Grids(iGrid).Pressure.Oil) - _
            (Data.Grids(adjID).SCAL.pcow - Data.Grids(iGrid).SCAL.pcow) - Grid_Params.R.average_Gamma_Water * _
            (Data.Grids(adjID).Location.Z - Data.Grids(iGrid).Location.Z)))

            .Mines_R(0, iGrid * 3 + 1) -= ((Grid_Params.R.Tg_nm * ((Data.Grids(adjID).Pressure.Oil - Data.Grids(iGrid).Pressure.Oil) + _
            (Data.Grids(adjID).SCAL.pcgo - Data.Grids(iGrid).SCAL.pcgo) - Grid_Params.R.average_Gamma_Gas * _
            (Data.Grids(adjID).Location.Z - Data.Grids(iGrid).Location.Z))) + _
            (Grid_Params.R.To_nm * (Data.Grids(iGrid).PVT.Rs + Data.Grids(adjID).PVT.Rs) / 2) * _
            ((Data.Grids(adjID).Pressure.Oil - Data.Grids(iGrid).Pressure.Oil) - _
             Grid_Params.R.average_Gamma_Oil * (Data.Grids(adjID).Location.Z - Data.Grids(iGrid).Location.Z)))

            .Mines_R(0, iGrid * 3 + 2) -= (Grid_Params.R.To_nm * ((Data.Grids(adjID).Pressure.Oil - Data.Grids(iGrid).Pressure.Oil) - _
             Grid_Params.R.average_Gamma_Oil * (Data.Grids(adjID).Location.Z - Data.Grids(iGrid).Location.Z)))

        End With
    End Sub

    Sub Add_Bracet_Parameters_n(ByVal Data As Simulation, ByVal iGrid As Integer, ByRef Grid_Params As Grid_Parameters, ByRef IM As Implicit_Matrix)
        With IM
            .J(iGrid * 3, iGrid * 3) = Grid_Params.R_nn.Rw_Sw
            .J(iGrid * 3 + 1, iGrid * 3) = Grid_Params.R_nn.Rw_Sg
            .J(iGrid * 3 + 2, iGrid * 3) = Grid_Params.R_nn.Rw_Po

            .J(iGrid * 3, iGrid * 3 + 1) = Grid_Params.R_nn.Rg_Sw
            .J(iGrid * 3 + 1, iGrid * 3 + 1) = Grid_Params.R_nn.Rg_Sg
            .J(iGrid * 3 + 2, iGrid * 3 + 1) = Grid_Params.R_nn.Rg_Po

            .J(iGrid * 3, iGrid * 3 + 2) = Grid_Params.R_nn.Ro_Sw
            .J(iGrid * 3 + 1, iGrid * 3 + 2) = Grid_Params.R_nn.Ro_Sg
            .J(iGrid * 3 + 2, iGrid * 3 + 2) = Grid_Params.R_nn.Ro_Po
        End With
    End Sub

    Sub Calculate_C_Bracet(ByVal Data_0 As Simulation, ByVal Data_1 As Simulation, ByVal iGrid As Integer, ByRef Grid_Params As Grid_Parameters)

        Dim Phi_prime As Double = IIf(Data_1.Grids(iGrid).Pressure.Oil = Data_0.Grids(iGrid).Pressure.Oil, 0, _
        (Data_1.Grids(iGrid).Porosity - Data_0.Grids(iGrid).Porosity) / _
        (Data_1.Grids(iGrid).Pressure.Oil - Data_0.Grids(iGrid).Pressure.Oil))

        Dim Rs_prime As Double = IIf(Data_1.Grids(iGrid).Pressure.Oil = Data_0.Grids(iGrid).Pressure.Oil, 0, _
        (Data_1.Grids(iGrid).PVT.Rs - Data_0.Grids(iGrid).PVT.Rs) / _
        (Data_1.Grids(iGrid).Pressure.Oil - Data_0.Grids(iGrid).Pressure.Oil))

        Dim one_over_bo_prime As Double = IIf(Data_1.Grids(iGrid).Pressure.Oil = Data_0.Grids(iGrid).Pressure.Oil, 0, _
        ((1 / Data_1.Grids(iGrid).PVT.Oil.FVF) - (1 / Data_0.Grids(iGrid).PVT.Oil.FVF)) / _
        (Data_1.Grids(iGrid).Pressure.Oil - Data_0.Grids(iGrid).Pressure.Oil))

        Dim one_over_bg_prime As Double = IIf(Data_1.Grids(iGrid).Pressure.Oil = Data_0.Grids(iGrid).Pressure.Oil, 0, _
        ((1 / Data_1.Grids(iGrid).PVT.Gas.FVF) - (1 / Data_0.Grids(iGrid).PVT.Gas.FVF)) / _
        (Data_1.Grids(iGrid).Pressure.Oil - Data_0.Grids(iGrid).Pressure.Oil))

        Dim one_over_bw_prime As Double = IIf(Data_1.Grids(iGrid).Pressure.Oil = Data_0.Grids(iGrid).Pressure.Oil, 0, _
        ((1 / Data_1.Grids(iGrid).PVT.Water.FVF) - (1 / Data_0.Grids(iGrid).PVT.Water.FVF)) / _
        (Data_1.Grids(iGrid).Pressure.Oil - Data_0.Grids(iGrid).Pressure.Oil))

        ' Cww ---
        Grid_Params.R_nn.Rw_Sw -= (Data_0.Grids(iGrid).Volume / (Data_0.Alpha * Data_0.Time_Step)) * (Data_1.Grids(iGrid).Porosity / Data_1.Grids(iGrid).PVT.Water.FVF)

        ' Cwg ---
        Grid_Params.R_nn.Rw_Sg -= 0

        ' Cwp ---
        Grid_Params.R_nn.Rw_Po -= (Data_0.Grids(iGrid).Volume / (Data_0.Alpha * Data_0.Time_Step)) * _
        ((Phi_prime / Data_0.Grids(iGrid).PVT.Water.FVF) + _
        Data_1.Grids(iGrid).Porosity * one_over_bw_prime) * Data_0.Grids(iGrid).Saturation.Water

        ' Cow ---
        Grid_Params.R_nn.Ro_Sw -= -(Data_0.Grids(iGrid).Volume / (Data_0.Alpha * Data_0.Time_Step)) * _
        (Data_1.Grids(iGrid).Porosity / Data_1.Grids(iGrid).PVT.Oil.FVF)

        ' Cog ---
        Grid_Params.R_nn.Ro_Sg -= -(Data_0.Grids(iGrid).Volume / (Data_0.Alpha * Data_0.Time_Step)) * _
        (Data_1.Grids(iGrid).Porosity / Data_1.Grids(iGrid).PVT.Oil.FVF)

        ' Cop ---
        Grid_Params.R_nn.Ro_Po -= (Data_0.Grids(iGrid).Volume / (Data_0.Alpha * Data_0.Time_Step)) * _
        ((Phi_prime / Data_0.Grids(iGrid).PVT.Oil.FVF) + Data_1.Grids(iGrid).Porosity * one_over_bo_prime) * _
        (1 - Data_0.Grids(iGrid).Saturation.Water - Data_0.Grids(iGrid).Saturation.Gas)

        ' Cgw ---
        Grid_Params.R_nn.Rg_Sw -= -(Data_0.Grids(iGrid).Volume / (Data_0.Alpha * Data_0.Time_Step)) * _
        (Data_1.Grids(iGrid).Porosity / Data_1.Grids(iGrid).PVT.Oil.FVF) * Data_1.Grids(iGrid).PVT.Rs

        ' Cgg ---
        Grid_Params.R_nn.Rg_Sg -= (Data_0.Grids(iGrid).Volume / (Data_0.Alpha * Data_0.Time_Step)) * _
        ((Data_1.Grids(iGrid).Porosity / Data_1.Grids(iGrid).PVT.Gas.FVF) - _
        (Data_1.Grids(iGrid).Porosity / Data_1.Grids(iGrid).PVT.Oil.FVF) * Data_1.Grids(iGrid).PVT.Rs)

        ' Cgp ---
        Grid_Params.R_nn.Rg_Po -= (Data_0.Grids(iGrid).Volume / (Data_0.Alpha * Data_0.Time_Step)) * ( _
        (((Phi_prime / Data_0.Grids(iGrid).PVT.Oil.FVF + Data_1.Grids(iGrid).Porosity * one_over_bo_prime) * _
         Data_0.Grids(iGrid).PVT.Rs) + (Data_1.Grids(iGrid).Porosity / Data_1.Grids(iGrid).PVT.Oil.FVF) * Rs_prime) * _
         (1 - Data_0.Grids(iGrid).Saturation.Water - Data_0.Grids(iGrid).Saturation.Gas) + _
         ((Phi_prime / Data_0.Grids(iGrid).PVT.Gas.FVF) + Data_1.Grids(iGrid).Porosity * one_over_bg_prime) * _
         Data_0.Grids(iGrid).Saturation.Gas _
         )
    End Sub

    Sub Add_F_Bracet_Parameters_n(ByVal Data As Simulation, ByVal iGrid As Integer, ByVal iAdjacent As Integer, ByRef Grid_Params As Grid_Parameters)
        Dim adjID As Integer = Data.Grids(iGrid).Adjacents(iAdjacent).Target_Grid_Index

        Dim nGrid As Grid = Data.Grids(iGrid)
        Dim mGrid As Grid = Data.Grids(adjID)


        ' ∂Rwn /∂Swn (0,0)
        Grid_Params.R_nn.Rw_Sw += ((mGrid.Pressure.Oil - nGrid.Pressure.Oil) _
            - (mGrid.SCAL.pcow - nGrid.SCAL.pcow) _
            - Grid_Params.R.average_Gamma_Water * (mGrid.Location.Z - nGrid.Location.Z)) * Grid_Params.R.round_Tw_Sw_n _
            + Grid_Params.R.Tw_nm * Grid_Params.R.Pcow_Prime_n

        ' ∂Rwn /∂Sgn (1,0)
        Grid_Params.R_nn.Rw_Sg += 0

        ' ∂Rwn /∂Pon (2,0)
        Grid_Params.R_nn.Rw_Po += (-Grid_Params.R.Tw_nm + ((mGrid.Pressure.Oil - nGrid.Pressure.Oil) - _
             (mGrid.SCAL.pcow - nGrid.SCAL.pcow) - (Grid_Params.R.average_Gamma_Water * _
            (mGrid.Location.Z - nGrid.Location.Z))) * Grid_Params.R.round_Tw_Po_n)

        ' ∂Ron /∂Swn (0,2)
        Grid_Params.R_nn.Ro_Sw += ((mGrid.Pressure.Oil - nGrid.Pressure.Oil) - _
            (Grid_Params.R.average_Gamma_Oil * (mGrid.Location.Z - nGrid.Location.Z))) * Grid_Params.R.round_To_Sw_n

        ' ∂Ron /∂Sgn (1,2)
        Grid_Params.R_nn.Ro_Sg += ((mGrid.Pressure.Oil - nGrid.Pressure.Oil) - _
            (Grid_Params.R.average_Gamma_Oil * (mGrid.Location.Z - nGrid.Location.Z))) * Grid_Params.R.round_To_Sg_n

        ' ∂Ron /∂Pon (2,2)
        Grid_Params.R_nn.Ro_Po += (-Grid_Params.R.To_nm + ((mGrid.Pressure.Oil - nGrid.Pressure.Oil) - _
            (Grid_Params.R.average_Gamma_Oil * (mGrid.Location.Z - nGrid.Location.Z))) * Grid_Params.R.round_To_Po_n)

        ' ∂Rgn /∂Swn (0,1)
        Grid_Params.R_nn.Rg_Sw += (((mGrid.Pressure.Oil - nGrid.Pressure.Oil) _
            - Grid_Params.R.average_Gamma_Oil * (mGrid.Location.Z - nGrid.Location.Z)) * (Grid_Params.R.round_Tw_Sw_n _
            * (mGrid.PVT.Rs + nGrid.PVT.Rs) / 2))


        ' ∂Rgn /∂Sgn (1,1)
        Grid_Params.R_nn.Rg_Sg += (((mGrid.Pressure.Oil - nGrid.Pressure.Oil) + _
            (mGrid.SCAL.pcgo - nGrid.SCAL.pcgo) - (Grid_Params.R.average_Gamma_Gas * _
            (mGrid.Location.Z - nGrid.Location.Z))) * Grid_Params.R.round_Tg_Sg_n _
            - Grid_Params.R.Tg_nm * Grid_Params.R.Pcgo_Prime_n + _
            ((mGrid.Pressure.Oil - nGrid.Pressure.Oil) - (Grid_Params.R.average_Gamma_Oil * (mGrid.Location.Z - nGrid.Location.Z))) * _
            Grid_Params.R.round_To_Sg_n * ((mGrid.PVT.Rs + nGrid.PVT.Rs) / 2))

        ' ∂Rgn /∂Pon (2,1)
        Grid_Params.R_nn.Rg_Po += (-Grid_Params.R.Tg_nm + _
            ((mGrid.Pressure.Oil - nGrid.Pressure.Oil) + _
            (mGrid.SCAL.pcgo - nGrid.SCAL.pcgo) - (Grid_Params.R.average_Gamma_Gas * _
            (mGrid.Location.Z - nGrid.Location.Z))) * Grid_Params.R.round_Tg_Po_n - _
            Grid_Params.R.To_nm * (mGrid.PVT.Rs + nGrid.PVT.Rs) / 2 + _
            ((mGrid.Pressure.Oil - nGrid.Pressure.Oil) - (Grid_Params.R.average_Gamma_Oil * (mGrid.Location.Z - nGrid.Location.Z))) * _
            Grid_Params.R.round_To_Po_n * ((mGrid.PVT.Rs + nGrid.PVT.Rs) / 2))



    End Sub

    Sub Add_F_Bracet_Parameters_m(ByVal Data As Simulation, ByVal iGrid As Integer, ByVal iAdjacent As Integer, ByVal Grid_Params As Grid_Parameters, ByRef IM As Implicit_Matrix)
        Dim adjID As Integer = Data.Grids(iGrid).Adjacents(iAdjacent).Target_Grid_Index

        Dim nGrid As Grid = Data.Grids(iGrid)
        Dim mGrid As Grid = Data.Grids(adjID)

        With IM
            ' ∂Rwn /∂Swm (0,0)
            .J(adjID * 3, iGrid * 3) = ((mGrid.Pressure.Oil - nGrid.Pressure.Oil) _
            - (mGrid.SCAL.pcow - nGrid.SCAL.pcow) _
            - Grid_Params.R.average_Gamma_Water * (mGrid.Location.Z - nGrid.Location.Z)) * Grid_Params.R.round_Tw_Sw_m _
            - Grid_Params.R.Tw_nm * Grid_Params.R.Pcow_Prime_m

            ' ∂Rwn /∂Sgm (1,0)
            .J(adjID * 3 + 1, iGrid * 3) = 0

            ' ∂Rwn /∂Pom (2,0)
            .J(adjID * 3 + 2, iGrid * 3) = Grid_Params.R.Tw_nm + ((mGrid.Pressure.Oil - nGrid.Pressure.Oil) - _
             (mGrid.SCAL.pcow - nGrid.SCAL.pcow) - (Grid_Params.R.average_Gamma_Water * _
            (mGrid.Location.Z - nGrid.Location.Z))) * Grid_Params.R.round_Tw_Po_m

            ' ∂Ron /∂Swm (0,2)
            .J(adjID * 3, iGrid * 3 + 2) = ((mGrid.Pressure.Oil - nGrid.Pressure.Oil) - _
            (Grid_Params.R.average_Gamma_Oil * (mGrid.Location.Z - nGrid.Location.Z))) * Grid_Params.R.round_To_Sw_m

            ' ∂Ron /∂Sgm (1,2)
            .J(adjID * 3 + 1, iGrid * 3 + 2) = ((mGrid.Pressure.Oil - nGrid.Pressure.Oil) - _
            (Grid_Params.R.average_Gamma_Oil * (mGrid.Location.Z - nGrid.Location.Z))) * Grid_Params.R.round_To_Sg_m

            ' ∂Ron /∂Pom (2,2)
            .J(adjID * 3 + 2, iGrid * 3 + 2) = Grid_Params.R.To_nm + ((mGrid.Pressure.Oil - nGrid.Pressure.Oil) - _
            (Grid_Params.R.average_Gamma_Oil * (mGrid.Location.Z - nGrid.Location.Z))) * Grid_Params.R.round_To_Po_m

            ' ∂Rgn /∂Swm (0,1)
            .J(adjID * 3, iGrid * 3 + 1) = ((mGrid.Pressure.Oil - nGrid.Pressure.Oil) - _
            (Grid_Params.R.average_Gamma_Oil * (mGrid.Location.Z - nGrid.Location.Z))) * _
            (Grid_Params.R.round_To_Sw_m * (mGrid.PVT.Rs + nGrid.PVT.Rs) / 2)

            ' ∂Rgn /∂Sgm (1,1)
            .J(adjID * 3 + 1, iGrid * 3 + 1) = ((mGrid.Pressure.Oil - nGrid.Pressure.Oil) + _
            (mGrid.SCAL.pcgo - nGrid.SCAL.pcgo) - (Grid_Params.R.average_Gamma_Gas * _
            (mGrid.Location.Z - nGrid.Location.Z))) * Grid_Params.R.round_Tg_Sg_m _
            + Grid_Params.R.Tg_nm * Grid_Params.R.Pcgo_Prime_m + _
            ((mGrid.Pressure.Oil - nGrid.Pressure.Oil) - (Grid_Params.R.average_Gamma_Oil * (mGrid.Location.Z - nGrid.Location.Z))) * _
            Grid_Params.R.round_To_Sg_m * ((mGrid.PVT.Rs + nGrid.PVT.Rs) / 2)

            ' ∂Rgn /∂Pom (2,1)
            .J(adjID * 3 + 2, iGrid * 3 + 1) = Grid_Params.R.Tg_nm + _
            ((mGrid.Pressure.Oil - nGrid.Pressure.Oil) + _
            (mGrid.SCAL.pcgo - nGrid.SCAL.pcgo) - (Grid_Params.R.average_Gamma_Gas * _
            (mGrid.Location.Z - nGrid.Location.Z))) * Grid_Params.R.round_Tg_Po_m + _
            Grid_Params.R.To_nm * (mGrid.PVT.Rs + nGrid.PVT.Rs) / 2 + _
            ((mGrid.Pressure.Oil - nGrid.Pressure.Oil) - (Grid_Params.R.average_Gamma_Oil * (mGrid.Location.Z - nGrid.Location.Z))) * _
            Grid_Params.R.round_To_Po_m * ((mGrid.PVT.Rs + nGrid.PVT.Rs) / 2)

        End With



    End Sub

    Sub Calculate_F_Bracet_Parameters(ByVal data_n As Simulation, ByVal Data_v As Simulation, ByVal nGrid As Integer, ByVal iAdjacent As Integer, ByRef Grid_Params As Grid_Parameters)
        Dim _area As Double = Data_v.Grids(nGrid).Adjacents(iAdjacent).Area
        Dim adjID As Integer = Data_v.Grids(nGrid).Adjacents(iAdjacent).Target_Grid_Index

        Dim For_Water_Saturation_Round_Grid_m As New Grid
        Dim For_Gas_Saturation_Round_Grid_m As New Grid
        Dim For_Pressure_Round_Grid_m As New Grid

        Dim For_Water_Saturation_Round_Grid_n As New Grid
        Dim For_Gas_Saturation_Round_Grid_n As New Grid
        Dim For_Pressure_Round_Grid_n As New Grid

        Dim l As Double
        Dim Fp_Oil As Double
        Dim Fp_Water As Double
        Dim Fp_Gas As Double
        Dim ko, kw, kg As Double


        Grid_Params.R.average_Gamma_Water = (data_n.Grids(nGrid).PVT.Water.Gamma + data_n.Grids(adjID).PVT.Water.Gamma) / 2
        Grid_Params.R.average_Gamma_Oil = (data_n.Grids(nGrid).PVT.Oil.Gamma + data_n.Grids(adjID).PVT.Oil.Gamma) / 2
        Grid_Params.R.average_Gamma_Gas = (data_n.Grids(nGrid).PVT.Gas.Gamma + data_n.Grids(adjID).PVT.Gas.Gamma) / 2

        Initialize_Forward_Grids(Data_v, nGrid, adjID, For_Gas_Saturation_Round_Grid_n, For_Water_Saturation_Round_Grid_n, For_Pressure_Round_Grid_n, For_Gas_Saturation_Round_Grid_m, For_Water_Saturation_Round_Grid_m, For_Pressure_Round_Grid_m)


        ' Non Round Region ==============================================================================================
        Fp_Oil_Water_Gas(Data_v.Grids(nGrid), Data_v.Grids(adjID), Data_v.Methodology.Transmissibility_Fp_Averaging, Fp_Oil, Fp_Water, Fp_Gas)
        Ko_Kw_Kg(Data_v.Grids(nGrid), Data_v.Grids(adjID), Data_v.Grids(nGrid).Adjacents(iAdjacent).Type, ko, kw, kg, l)

        Grid_Params.R.To_nm = IIf(data_n.System.Phase.Oil, ((ko * _area * Data_v.Beta) / l) * Fp_Oil, 0)
        Grid_Params.R.Tw_nm = IIf(data_n.System.Phase.Water, ((kw * _area * Data_v.Beta) / l) * Fp_Water, 0)
        Grid_Params.R.Tg_nm = IIf(data_n.System.Phase.Gas, ((kg * _area * Data_v.Beta) / l) * Fp_Gas, 0)

        'm Region ====================================================================================================================================================================================
        'Sw-m Round Region ====================================================================================================================================================================
        Ko_Kw_Kg(Data_v.Grids(nGrid), For_Water_Saturation_Round_Grid_m, Data_v.Grids(nGrid).Adjacents(iAdjacent).Type, ko, kw, kg, l)

        Grid_Params.R.round_To_Sw_m = IIf(data_n.System.Phase.Water, (((ko * _area * Data_v.Beta) / l) * Fp_Oil - Grid_Params.R.To_nm) / Data_v.Methodology.Derivations.Water_Saturation_Epsilon, 0)
        Grid_Params.R.round_Tw_Sw_m = IIf(data_n.System.Phase.Water, (((kw * _area * Data_v.Beta) / l) * Fp_Water - Grid_Params.R.Tw_nm) / Data_v.Methodology.Derivations.Water_Saturation_Epsilon, 0)
        Grid_Params.R.Pcow_Prime_m = (For_Water_Saturation_Round_Grid_m.SCAL.pcow - Data_v.Grids(adjID).SCAL.pcow) / Data_v.Methodology.Derivations.Water_Saturation_Epsilon

        'Sg-m Round Region ====================================================================================================================================================================
        Ko_Kw_Kg(Data_v.Grids(nGrid), For_Gas_Saturation_Round_Grid_m, Data_v.Grids(nGrid).Adjacents(iAdjacent).Type, ko, kw, kg, l)

        Grid_Params.R.round_To_Sg_m = IIf(data_n.System.Phase.Gas, (((ko * _area * Data_v.Beta) / l) * Fp_Oil - Grid_Params.R.To_nm) / Data_v.Methodology.Derivations.Gas_Saturation_Epsilon, 0)
        Grid_Params.R.round_Tg_Sg_m = IIf(data_n.System.Phase.Gas, (((kg * _area * Data_v.Beta) / l) * Fp_Gas - Grid_Params.R.Tg_nm) / Data_v.Methodology.Derivations.Gas_Saturation_Epsilon, 0)
        Grid_Params.R.Pcgo_Prime_m = (For_Gas_Saturation_Round_Grid_m.SCAL.pcgo - Data_v.Grids(adjID).SCAL.pcgo) / Data_v.Methodology.Derivations.Gas_Saturation_Epsilon

        'Po-m Round Region ====================================================================================================================================================================
        Fp_Oil_Water_Gas(Data_v.Grids(nGrid), For_Pressure_Round_Grid_m, Data_v.Methodology.Transmissibility_Fp_Averaging, Fp_Oil, Fp_Water, Fp_Gas)
        Ko_Kw_Kg(Data_v.Grids(nGrid), Data_v.Grids(adjID), Data_v.Grids(nGrid).Adjacents(iAdjacent).Type, ko, kw, kg, l)
        Dim Rs1 As Double = (Data_v.Grids(nGrid).PVT.Rs + Data_v.Grids(adjID).PVT.Rs) / 2
        Dim Rs2 As Double = (Data_v.Grids(nGrid).PVT.Rs + For_Pressure_Round_Grid_m.PVT.Rs) / 2

        Grid_Params.R.round_To_Po_m = (((ko * _area * Data_v.Beta) / l) * Fp_Oil - Grid_Params.R.To_nm) / Data_v.Methodology.Derivations.Oil_Pressure_Epsilon
        Grid_Params.R.round_To_Rs_Po_m = (Rs2 * (((ko * _area * Data_v.Beta) / l) * Fp_Oil) - (Grid_Params.R.To_nm * Rs1)) / Data_v.Methodology.Derivations.Oil_Pressure_Epsilon
        Grid_Params.R.round_Tw_Po_m = (((kw * _area * Data_v.Beta) / l) * Fp_Water - Grid_Params.R.Tw_nm) / Data_v.Methodology.Derivations.Oil_Pressure_Epsilon
        Grid_Params.R.round_Tg_Po_m = (((kg * _area * Data_v.Beta) / l) * Fp_Gas - Grid_Params.R.Tg_nm) / Data_v.Methodology.Derivations.Oil_Pressure_Epsilon

        'n Region ====================================================================================================================================================================================
        'Sw-n Round Region ====================================================================================================================================================================
        Fp_Oil_Water_Gas(Data_v.Grids(nGrid), Data_v.Grids(adjID), Data_v.Methodology.Transmissibility_Fp_Averaging, Fp_Oil, Fp_Water, Fp_Gas)
        Ko_Kw_Kg(For_Water_Saturation_Round_Grid_n, Data_v.Grids(adjID), Data_v.Grids(nGrid).Adjacents(iAdjacent).Type, ko, kw, kg, l)

        Grid_Params.R.round_To_Sw_n = IIf(data_n.System.Phase.Water, (((ko * _area * Data_v.Beta) / l) * Fp_Oil - Grid_Params.R.To_nm) / Data_v.Methodology.Derivations.Water_Saturation_Epsilon, 0)
        Grid_Params.R.round_Tw_Sw_n = IIf(data_n.System.Phase.Water, (((kw * _area * Data_v.Beta) / l) * Fp_Water - Grid_Params.R.Tw_nm) / Data_v.Methodology.Derivations.Water_Saturation_Epsilon, 0)
        Grid_Params.R.Pcow_Prime_n = (For_Water_Saturation_Round_Grid_n.SCAL.pcow - Data_v.Grids(nGrid).SCAL.pcow) / Data_v.Methodology.Derivations.Water_Saturation_Epsilon

        'Sg-n Round Region ====================================================================================================================================================================
        Ko_Kw_Kg(For_Gas_Saturation_Round_Grid_n, Data_v.Grids(adjID), Data_v.Grids(nGrid).Adjacents(iAdjacent).Type, ko, kw, kg, l)

        Grid_Params.R.round_To_Sg_n = IIf(data_n.System.Phase.Gas, (((ko * _area * Data_v.Beta) / l) * Fp_Oil - Grid_Params.R.To_nm) / Data_v.Methodology.Derivations.Gas_Saturation_Epsilon, 0)
        Grid_Params.R.round_Tg_Sg_n = IIf(data_n.System.Phase.Gas, (((kg * _area * Data_v.Beta) / l) * Fp_Gas - Grid_Params.R.Tg_nm) / Data_v.Methodology.Derivations.Gas_Saturation_Epsilon, 0)
        Grid_Params.R.Pcgo_Prime_n = (For_Gas_Saturation_Round_Grid_n.SCAL.pcgo - Data_v.Grids(nGrid).SCAL.pcgo) / Data_v.Methodology.Derivations.Gas_Saturation_Epsilon

        'Po-n Round Region ====================================================================================================================================================================
        Fp_Oil_Water_Gas(For_Pressure_Round_Grid_n, Data_v.Grids(adjID), Data_v.Methodology.Transmissibility_Fp_Averaging, Fp_Oil, Fp_Water, Fp_Gas)
        Ko_Kw_Kg(Data_v.Grids(nGrid), Data_v.Grids(adjID), Data_v.Grids(nGrid).Adjacents(iAdjacent).Type, ko, kw, kg, l)
        Rs1 = (Data_v.Grids(nGrid).PVT.Rs + Data_v.Grids(adjID).PVT.Rs) / 2
        Rs2 = (For_Water_Saturation_Round_Grid_n.PVT.Rs + Data_v.Grids(adjID).PVT.Rs) / 2

        Grid_Params.R.round_To_Po_n = (((ko * _area * Data_v.Beta) / l) * Fp_Oil - Grid_Params.R.To_nm) / Data_v.Methodology.Derivations.Oil_Pressure_Epsilon
        Grid_Params.R.round_To_Rs_Po_n = (Rs2 * (((ko * _area * Data_v.Beta) / l) * Fp_Oil) - (Grid_Params.R.To_nm * Rs1)) / Data_v.Methodology.Derivations.Oil_Pressure_Epsilon
        Grid_Params.R.round_Tw_Po_n = (((kw * _area * Data_v.Beta) / l) * Fp_Water - Grid_Params.R.Tw_nm) / Data_v.Methodology.Derivations.Oil_Pressure_Epsilon
        Grid_Params.R.round_Tg_Po_n = (((kg * _area * Data_v.Beta) / l) * Fp_Gas - Grid_Params.R.Tg_nm) / Data_v.Methodology.Derivations.Oil_Pressure_Epsilon

    End Sub

    Sub Initialize_Forward_Grids(ByVal Data As Simulation, ByVal nGrid As Integer, ByVal adjID As Integer, ByRef For_Gas_Saturation_Round_Grid_n As Grid, ByRef For_Water_Saturation_Round_Grid_n As Grid, ByRef For_Pressure_Round_Grid_n As Grid, ByRef For_Gas_Saturation_Round_Grid_m As Grid, ByRef For_Water_Saturation_Round_Grid_m As Grid, ByRef For_Pressure_Round_Grid_m As Grid)
        For_Water_Saturation_Round_Grid_n.Saturation.Water = (Data.Grids(nGrid).Saturation.Water + Data.Methodology.Derivations.Water_Saturation_Epsilon)
        For_Gas_Saturation_Round_Grid_n.Saturation.Gas = (Data.Grids(nGrid).Saturation.Gas + Data.Methodology.Derivations.Gas_Saturation_Epsilon)
        For_Pressure_Round_Grid_n.Pressure.Oil = (Data.Grids(nGrid).Pressure.Oil + Data.Methodology.Derivations.Oil_Pressure_Epsilon)

        For_Water_Saturation_Round_Grid_m.Saturation.Water = (Data.Grids(adjID).Saturation.Water + Data.Methodology.Derivations.Water_Saturation_Epsilon)
        For_Gas_Saturation_Round_Grid_m.Saturation.Gas = (Data.Grids(adjID).Saturation.Gas + Data.Methodology.Derivations.Gas_Saturation_Epsilon)
        For_Pressure_Round_Grid_m.Pressure.Oil = (Data.Grids(adjID).Pressure.Oil + Data.Methodology.Derivations.Oil_Pressure_Epsilon)


        For_Water_Saturation_Round_Grid_n.Pressure.Oil = Data.Grids(nGrid).Pressure.Oil
        For_Gas_Saturation_Round_Grid_n.Pressure.Oil = Data.Grids(nGrid).Pressure.Oil
        For_Pressure_Round_Grid_n.Saturation.Water = Data.Grids(nGrid).Saturation.Water
        For_Pressure_Round_Grid_n.Saturation.Oil = Data.Grids(nGrid).Saturation.Oil

        For_Water_Saturation_Round_Grid_m.Pressure.Oil = Data.Grids(adjID).Pressure.Oil
        For_Gas_Saturation_Round_Grid_m.Pressure.Oil = Data.Grids(adjID).Pressure.Oil
        For_Pressure_Round_Grid_m.Saturation.Water = Data.Grids(adjID).Saturation.Water
        For_Pressure_Round_Grid_m.Saturation.Oil = Data.Grids(adjID).Saturation.Oil


        Initialize_Saturations(For_Water_Saturation_Round_Grid_n)
        Initialize_Saturations(For_Gas_Saturation_Round_Grid_n)
        Initialize_Pressure(For_Pressure_Round_Grid_n)

        Initialize_Saturations(For_Water_Saturation_Round_Grid_m)
        Initialize_Saturations(For_Gas_Saturation_Round_Grid_m)
        Initialize_Pressure(For_Pressure_Round_Grid_m)


        Initialize_Phase_System(For_Water_Saturation_Round_Grid_n, Data.Rocks(Data.Grids(nGrid).RockType))
        Initialize_Phase_System(For_Gas_Saturation_Round_Grid_n, Data.Rocks(Data.Grids(nGrid).RockType))
        Initialize_Phase_System(For_Pressure_Round_Grid_n, Data.Rocks(Data.Grids(nGrid).RockType))

        Initialize_Phase_System(For_Water_Saturation_Round_Grid_m, Data.Rocks(Data.Grids(adjID).RockType))
        Initialize_Phase_System(For_Gas_Saturation_Round_Grid_m, Data.Rocks(Data.Grids(adjID).RockType))
        Initialize_Phase_System(For_Pressure_Round_Grid_m, Data.Rocks(Data.Grids(adjID).RockType))

        Initialize_SCAL(For_Water_Saturation_Round_Grid_n, Data.Rocks(Data.Grids(nGrid).RockType))
        Initialize_SCAL(For_Gas_Saturation_Round_Grid_n, Data.Rocks(Data.Grids(nGrid).RockType))
        Initialize_PVT(For_Pressure_Round_Grid_n, Data)

        Initialize_SCAL(For_Water_Saturation_Round_Grid_m, Data.Rocks(Data.Grids(adjID).RockType))
        Initialize_SCAL(For_Gas_Saturation_Round_Grid_m, Data.Rocks(Data.Grids(adjID).RockType))
        Initialize_PVT(For_Pressure_Round_Grid_m, Data)
    End Sub

    Sub Fp_Oil_Water_Gas(ByVal nGrid As Grid, ByVal mGrid As Grid, ByVal Method As Transmissibility_Fp_Averaging, ByRef Fp_Oil As Double, ByRef Fp_Water As Double, ByRef Fp_Gas As Double)
        Fp_Oil = Fp_Average(nGrid.PVT.Oil.Viscosity, mGrid.PVT.Oil.Viscosity, nGrid.PVT.Oil.FVF, mGrid.PVT.Oil.FVF, Method)
        Fp_Water = Fp_Average(nGrid.PVT.Water.Viscosity, mGrid.PVT.Water.Viscosity, nGrid.PVT.Water.FVF, mGrid.PVT.Water.FVF, Method)
        Fp_Gas = Fp_Average(nGrid.PVT.Gas.Viscosity, mGrid.PVT.Gas.Viscosity, nGrid.PVT.Gas.FVF, mGrid.PVT.Gas.FVF, Method)
    End Sub

    Sub Ko_Kw_Kg(ByVal nGrid As Grid, ByVal mGrid As Grid, ByVal AdjacentType As Adjacent_Type, ByRef Ko As Double, ByRef Kw As Double, ByRef Kg As Double, ByRef L As Double)
        Select Case AdjacentType
            Case Adjacent_Type.X
                L = (nGrid.Size.X_Length + mGrid.Size.X_Length) / 2
                Ko = Average_Permeability(nGrid.SCAL.Oil.Kx, mGrid.SCAL.Oil.Kx, nGrid.Size.X_Length / 2, mGrid.Size.X_Length / 2)
                Kw = Average_Permeability(nGrid.SCAL.Water.Kx, mGrid.SCAL.Water.Kx, nGrid.Size.X_Length / 2, mGrid.Size.X_Length / 2)
                Kg = Average_Permeability(nGrid.SCAL.Gas.Kx, mGrid.SCAL.Gas.Kx, nGrid.Size.X_Length / 2, mGrid.Size.X_Length / 2)
            Case Adjacent_Type.Y
                L = (nGrid.Size.Y_Length + mGrid.Size.Y_Length) / 2
                Kg = Average_Permeability(nGrid.SCAL.Gas.Ky, mGrid.SCAL.Gas.Ky, nGrid.Size.Y_Length / 2, mGrid.Size.Y_Length / 2)
                Kw = Average_Permeability(nGrid.SCAL.Water.Ky, mGrid.SCAL.Water.Ky, nGrid.Size.Y_Length / 2, mGrid.Size.Y_Length / 2)
                Ko = Average_Permeability(nGrid.SCAL.Oil.Ky, mGrid.SCAL.Oil.Ky, nGrid.Size.Y_Length / 2, mGrid.Size.Y_Length / 2)

            Case Adjacent_Type.Z
                L = (nGrid.Size.Z_Length + mGrid.Size.Z_Length) / 2
                Ko = Average_Permeability(nGrid.SCAL.Oil.Kz, mGrid.SCAL.Oil.Kz, nGrid.Size.Z_Length / 2, mGrid.Size.Z_Length / 2)
                Kw = Average_Permeability(nGrid.SCAL.Water.Kz, mGrid.SCAL.Water.Kz, nGrid.Size.Z_Length / 2, mGrid.Size.Z_Length / 2)
                Kg = Average_Permeability(nGrid.SCAL.Gas.Kz, mGrid.SCAL.Gas.Kz, nGrid.Size.Z_Length / 2, mGrid.Size.Z_Length / 2)
        End Select

    End Sub

    Function Fp_Average(ByVal mu1 As Double, ByVal mu2 As Double, ByVal b1 As Double, ByVal b2 As Double, ByVal method As Transmissibility_Fp_Averaging) As Double
        Select Case method
            Case Transmissibility_Fp_Averaging.Method1
                Return (1 / (((mu1 + mu2) / 2) * ((b1 + b2) / 2)))
            Case Transmissibility_Fp_Averaging.Method2
                Return ((1 / (mu1 * b1)) + (1 / (mu2 * b2))) / 2

        End Select
    End Function

    Sub Prepare_Adjacents(ByRef Data As Simulation)
        If Data.Options.Automatic_Grid_Adjacent_Recognition Then

            For i As Integer = 0 To SizeOf(Data.Grids) - 1
                Grid_Adjacent_Recognition(Data, i)
            Next
            Calculate_Adjacents_Area(Data)

            For i As Integer = 0 To SizeOf(Data.Grids) - 1
                Dim n As Integer = 0
                Dim Adjacents(SizeOf(Data.Grids(i).Adjacents) - 1) As Grid_Adjacent
                For j As Integer = 0 To SizeOf(Data.Grids(i).Adjacents) - 1
                    Adjacents(j).Area = Data.Grids(i).Adjacents(j).Area
                    Adjacents(j).Type = Data.Grids(i).Adjacents(j).Type
                    Adjacents(j).Target_Grid_Index = Data.Grids(i).Adjacents(j).Target_Grid_Index
                    If Data.Grids(i).Adjacents(j).Area > 0 Then
                        n += 1
                    End If
                Next

                ReDim Data.Grids(i).Adjacents(n - 1)
                n = 0

                For j As Integer = 0 To SizeOf(Adjacents) - 1
                    If Adjacents(j).Area > 0 Then
                        Data.Grids(i).Adjacents(n).Area = Adjacents(j).Area
                        Data.Grids(i).Adjacents(n).Type = Adjacents(j).Type
                        Data.Grids(i).Adjacents(n).Target_Grid_Index = Adjacents(j).Target_Grid_Index
                        n += 1
                    End If
                Next
            Next
        Else
            Calculate_Adjacents_Area(Data)
        End If
    End Sub

    Sub Grid_Adjacent_Recognition(ByRef Data As Simulation, ByVal i As Integer)
        Data.Grids(i).Adjacents = Nothing
        Dim n As Integer = 0
        For j As Integer = 0 To SizeOf(Data.Grids) - 1
            If (Data.Grids(j).Location.X + Data.Grids(j).Size.X_Length / 2 = Data.Grids(i).Location.X - Data.Grids(i).Size.X_Length / 2) _
            Or (Data.Grids(j).Location.X - Data.Grids(j).Size.X_Length / 2 = Data.Grids(i).Location.X + Data.Grids(i).Size.X_Length / 2) Then
                n += 1
            End If

            If (Data.Grids(j).Location.Y + Data.Grids(j).Size.Y_Length / 2 = Data.Grids(i).Location.Y - Data.Grids(i).Size.Y_Length / 2) _
            Or (Data.Grids(j).Location.Y - Data.Grids(j).Size.Y_Length / 2 = Data.Grids(i).Location.Y + Data.Grids(i).Size.Y_Length / 2) Then
                n += 1
            End If

            If (Data.Grids(j).Location.Z + Data.Grids(j).Size.Z_Length / 2 = Data.Grids(i).Location.Z - Data.Grids(i).Size.Z_Length / 2) _
            Or (Data.Grids(j).Location.Z - Data.Grids(j).Size.Z_Length / 2 = Data.Grids(i).Location.Z + Data.Grids(i).Size.Z_Length / 2) Then
                n += 1
            End If
        Next

        ReDim Data.Grids(i).Adjacents(n - 1)
        n = 0
        For j As Integer = 0 To SizeOf(Data.Grids) - 1
            If (Data.Grids(j).Location.X + Data.Grids(j).Size.X_Length / 2 = Data.Grids(i).Location.X - Data.Grids(i).Size.X_Length / 2) _
            Or (Data.Grids(j).Location.X - Data.Grids(j).Size.X_Length / 2 = Data.Grids(i).Location.X + Data.Grids(i).Size.X_Length / 2) Then
                Data.Grids(i).Adjacents(n).Target_Grid_Index = j
                Data.Grids(i).Adjacents(n).Type = Adjacent_Type.X
                n += 1
            End If

            If (Data.Grids(j).Location.Y + Data.Grids(j).Size.Y_Length / 2 = Data.Grids(i).Location.Y - Data.Grids(i).Size.Y_Length / 2) _
            Or (Data.Grids(j).Location.Y - Data.Grids(j).Size.Y_Length / 2 = Data.Grids(i).Location.Y + Data.Grids(i).Size.Y_Length / 2) Then
                Data.Grids(i).Adjacents(n).Target_Grid_Index = j
                Data.Grids(i).Adjacents(n).Type = Adjacent_Type.Y
                n += 1
            End If

            If (Data.Grids(j).Location.Z + Data.Grids(j).Size.Z_Length / 2 = Data.Grids(i).Location.Z - Data.Grids(i).Size.Z_Length / 2) _
            Or (Data.Grids(j).Location.Z - Data.Grids(j).Size.Z_Length / 2 = Data.Grids(i).Location.Z + Data.Grids(i).Size.Z_Length / 2) Then
                Data.Grids(i).Adjacents(n).Target_Grid_Index = j
                Data.Grids(i).Adjacents(n).Type = Adjacent_Type.Z
                n += 1
            End If
        Next


    End Sub

    Sub IM_Zero(ByRef IM As Implicit_Matrix)
        For i As Integer = 0 To IM.J.GetLength(0) - 1
            For j As Integer = 0 To IM.J.GetLength(1) - 1
                IM.J(i, j) = 0
            Next
            IM.Mines_R(0, i) = 0
        Next
    End Sub

    Sub StopSimulation()
        Paused = True
    End Sub


End Class
