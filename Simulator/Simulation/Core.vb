

' ============================================== Project Schedule ==============================================
' Date                           Description                                              Module
' September 15 2010              Project Start
' September 16 2010              Data Structures and Initialization Functions           - Preperation 
' September 17 2010              SCAL & PVT Functions                                   - Initialize 
' September 18 2010              Orthomin Solution - Matrix Computations                - Mathematics 
' September 19 2010              Writing Simulator's Core                               - Simulator and Others
' September 23 2010              First Run                                              
' September 24 2010              Test and Debug                                         
' September 28 2010              First Accurate Output
'---------------------------------------------------------------------------------------------------------------
' March     28 2013              Revision for WPF Implementation                        - WPFX Control
'===============================================================================================================


''' <summary>
'''                                                         BLACK-OIL MSV 1.0
'''  ______________________________Open Source Simulator for Black-Oil Model of Hydrocarbon Reservoirs _____________________________
'''                                                 Copyright (C) 2014  Amir Emamjomeh                                              
'''                                                      amir.emamjomeh@gmail.com
''' 
''' This program is free software: you can redistribute it and/or  modify it  under the terms  of the  GNU  General  Public  License 
''' as published by the Free Software Foundation, either version 3 of the License, or any later version. This program is distributed 
''' in the hope that it will be useful,but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR 
''' A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
''' 
''' You should have received a copy of the GNU General Public Licensealong with this program. If not,see http://www.gnu.org/licenses
'''_________________________________________________________________________________________________________________________________
''' </summary>
Public Class Core


    Public Event Iteration(ByVal Data_n As Simulation, ByVal Data_v As Simulation, ByVal FI_IM As Fully_Implicit_Matrix, ByVal FI_MSE As Iteration_MSE, ByVal IMPES_IM As IMPES_Matrix, ByVal IMPES_MSE As Iteration_MSE)
    Public Event SolutionFailed(ByVal Data_n As Simulation, ByVal Data_v As Simulation, ByVal IM As Fully_Implicit_Matrix, ByVal MSE As Iteration_MSE)
    Public Event Time_Step(ByVal History As Simulation_History, ByVal Data_n As Simulation)
    Public Event Simulation_End(ByVal History As Simulation_History)


    Public Paused As Boolean
    Public Finished As Boolean


    '============================================== Main Subroutine of Simulator ==============================================='
    '---------------------------------------------------------------------------------------------------------------------------'


    Sub Run(ByVal Data As Simulation, ByRef History As Simulation_History)
        Dim strx As String = ""


        Paused = False                                                                         ' Simulation is not Paused
        Finished = False                                                                       ' Simulation is not Finished

        Dim FI_Input_Matrix As Fully_Implicit_Matrix
        Dim FI_Output_Matrix As Fully_Implicit_Matrix
        Dim FullyImplicit_MSE As Iteration_MSE

        Dim IMPES_Matrix As IMPES_Matrix
        Dim IMPES_MSE As Iteration_MSE

        ' Preparing Simulator   
        History.Clear()                                                                       'Clear History Array
        History.Length = Data.Final_Step                                                      'Initialize History Array   

        Set_System_of_Units(Data)                                                             'Sets Coefficients for Each System of Units
        Prepare_Adjacents(Data)                                                               'Determines Grids With Surface in Common
        Initialize_Grids_Porosities(Data)                                                     'Copies Porosity Value to All Grids
        Calculate_Grids_Volume(Data)                                                          'Calculates Grid Volume Using Dimension
        Initialize_Gas_Critical_Pressure_Temperature(Data)                                    'Calculates Gas Critical Pressure and Temperature at Current State

        Prepare_Implicit_Matrix(Data, FI_Input_Matrix)                                        'Allocates Memory to Input Implicit Matrix
        Prepare_Implicit_Matrix(Data, FI_Output_Matrix)                                       'Allocates Memory to Output Implicit Matrix  
        Prepare_IMPES_Matrix(Data, IMPES_Matrix)                                              'Allocates Memory to IMPES matrix

        History.Cast(Data)                                                                    'Copies Current Simulation Data to All History Simulation Elements
        Application.DoEvents()                                                                'Enforces the Previous Methods to Accomplish Before Simulation is Started

        'Time Step Level 
        Do While (Data.Total_Time < Data.Finish_Time And Data.Current_Step < Data.Final_Step) And Not Paused

            Data.Current_Step += 1                                                            'One Time Step Forward
            History.Current = Data.Current_Step                                               'Sets Index of Current Time Step in Progress in History

            FI_Input_Matrix.Iteration = 0                                                     'Initialize Fully Implicit Matrix
            IMPES_Matrix.Iteration = 0                                                        'Initialize IMPES Matrix  

            Initialize(Data)                                                                  'Initialize Model Parameters for New Time Step
            Simulation_Copy(Data, History.Simulations(Data.Current_Step - 1), False)          'Copies Current Simulation State to History Element

            Select Case Data.Methodology.Approach

                Case Simulation_Approach.Fully_Implicit

                    '====================================================================================================='
                    '--------------Fully----------------------------------------------------------------------------------'
                    '              *****  *         *  ******   *       *****     ***     *****  ***********              '
                    '                *    **       **  *     *  *         *     *     *     *    *    *    *              '
                    '                *    * *     * *  *     *  *         *    *            *         *                   '
                    '                *    *  *   *  *  ******   *         *    *            *         *                   '
                    '                *    *   * *   *  *        *         *    *            *         *                   '
                    '                *    *    *    *  *        *         *     *     *     *         *                   '
                    '              ***** **         ** **       ******  *****     ***     *****      ***                  '
                    '-----------------------------------------------------------------------------------------------------'
                    '====================================================================================================='


                    FullyImplicit_MSE = Data.Exactness.IterationMSE

                    Do While (FullyImplicit_MSE.Sw >= Data.Exactness.IterationMSE.Sw Or FullyImplicit_MSE.Sg >= Data.Exactness.IterationMSE.Sg Or FullyImplicit_MSE.Po >= Data.Exactness.IterationMSE.Po) And Not Paused And FI_Input_Matrix.Iteration < Data.Exactness.IterationMSE.Maximum

                        Create_Implicit_Matrix(History.Simulations(History.Current - 1), Data, FI_Input_Matrix)
                        Solve_Orthomin(FI_Input_Matrix.J, FI_Input_Matrix.Delta_X, FI_Input_Matrix.Mines_R, Data.Exactness.OrthominMaxIteration, Data.Exactness.Equation)
                        FullyImplicit_MSE = Calculate_Fully_Implicit_Iteration_MSE(FI_Input_Matrix)
                        FullyImplicit_Update_Simulatin_Data(Data, FI_Input_Matrix)
                        Initialize(Data, False)
                        RaiseEvent Iteration(History.Simulations(History.Current - 1), Data, FI_Input_Matrix, FullyImplicit_MSE, IMPES_Matrix, IMPES_MSE)
                    Loop

                    Data.CurrentIterationCount = FI_Input_Matrix.Iteration

                Case Simulation_Approach.IMPES

                    '====================================================================================================='
                    '-----------------------------------------------------------------------------------------------------'
                    '                           *****  *         *  ******   ******   *****                               '
                    '                             *    **       **  *     *  *       *     *                              '
                    '                             *    * *     * *  *     *  *       *                                    '
                    '                             *    *  *   *  *  ******   ****      ***                                '
                    '                             *    *   * *   *  *        *             *                              '
                    '                             *    *    *    *  *        *       *     *                              '
                    '                           ***** **         ** **       ******   *****                               '
                    '-----------------------------------------------------------------------------------------------------'
                    '====================================================================================================='

                    Dim Last_IM As New IMPES_Matrix
                    Do While (IMPES_MSE.Po >= Data.Exactness.IterationMSE.Po Or IMPES_Matrix.Iteration <= 1) And Not Paused And IMPES_Matrix.Iteration < Data.Exactness.IterationMSE.Maximum

                        Create_IMPES_Matrix(History.Simulations(History.Current - 1), Data, IMPES_Matrix)
                        Initialize_IMPES_Po_Matrix(Data, IMPES_Matrix.Po)
                        Solve(IMPES_Matrix.LeftSide, IMPES_Matrix.Po, IMPES_Matrix.Target, 100, Data.Exactness.Equation, Equation_Solution_Method.Orhtomin)
                        Update_IMPES_Pressure(Data, IMPES_Matrix)
                        If Last_IM.Po IsNot Nothing Then IMPES_MSE = Calculate_IMPES_Iteration_MSE(IMPES_Matrix, Last_IM)
                        copy_IMPES_Matrix(IMPES_Matrix, Last_IM)
                        Initialize(Data, False)
                        RaiseEvent Iteration(History.First, Data, FI_Input_Matrix, FullyImplicit_MSE, IMPES_Matrix, IMPES_MSE)

                    Loop

                    Update_IMPES_Saturation(History.First, Data, IMPES_Matrix)
                    Data.CurrentIterationCount = IMPES_Matrix.Iteration

            End Select

            Data.Total_Time += Data.Time_Step
            RaiseEvent Time_Step(History, Data)                                          'Time Step Progress Event

        Loop

        Finished = True                                                                  'End of Simulation Indicator
        RaiseEvent Simulation_End(History)                                               'End of Simulation Event

    End Sub

    Sub Copy_FullyImplicit_Matrix(ByVal Data_n As Simulation, ByVal Data_v As Simulation, ByVal _To As Fully_Implicit_Matrix)
        If SizeOf(_To.Delta_X) = 0 Then ReDim _To.Delta_X(0, SizeOf(Data_v.Model.Grids) * 3 - 1)
        For i As Integer = 0 To SizeOf(Data_v.Model.Grids) - 1
            _To.Delta_X(0, i * 3) = Data_v.Model.Grids(i).Saturation.Water - Data_n.Model.Grids(i).Saturation.Water
            _To.Delta_X(0, i * 3 + 1) = Data_v.Model.Grids(i).Saturation.Gas - Data_n.Model.Grids(i).Saturation.Gas
            _To.Delta_X(0, i * 3 + 2) = Data_v.Model.Grids(i).Pressure.Oil - Data_n.Model.Grids(i).Pressure.Oil
        Next

    End Sub
    Sub Initialize_IMPES_Po_Matrix(ByVal Data As Simulation, ByRef Po(,) As Decimal)
        For i As Integer = 0 To SizeOf(Data.Model.Grids) - 1
            Po(0, i) = Data.Model.Grids(i).Pressure.Oil
        Next
    End Sub
    Sub copy_IMPES_Matrix(ByVal Imp_from As IMPES_Matrix, ByRef Imp_to As IMPES_Matrix)
        ReDim Imp_to.Po(0, Imp_from.Po.GetLength(1) - 1)
        For i As Integer = 0 To Imp_from.Po.GetLength(1) - 1
            Imp_to.Po(0, i) = Imp_from.Po(0, i)
        Next

    End Sub
    Sub FullyImplicit_Update_Simulatin_Data(ByRef Data As Simulation, ByVal IM As Fully_Implicit_Matrix)
        For i As Integer = 0 To Data.Model.Grids.Length - 1
            Dim nGrid As Grid = Data.Model.Grids(i)
            Dim Rock As Rock = Data.Rocks(nGrid.RockType)

            Dim NewSw As Decimal
            Dim NewSg As Decimal
            Dim NewPo As Decimal

            CorrectFullyImplicitResult(Data, i, IM.Delta_X(0, i * 3), IM.Delta_X(0, i * 3 + 1), IM.Delta_X(0, i * 3 + 2), NewSw, NewSg, NewPo)
            Capillary_Pressure_Hystersis(Data.Model.Grids(i), Data.Rocks(Data.Model.Grids(i).RockType).SCAL.Wettability, NewSw, NewSg)

            nGrid.Delta.Water = (Data.Model.Grids(i).Saturation.Water - NewSw)
            nGrid.Delta.Oil = (Data.Model.Grids(i).Saturation.Oil - (1 - NewSw - NewSg))
            nGrid.Delta.Gas = (Data.Model.Grids(i).Saturation.Gas - NewSg)

            Data.Model.Grids(i).Saturation.Water = IIf(NewSw < 0, 0, NewSw)
            Data.Model.Grids(i).Saturation.Gas = IIf(NewSg < 0, 0, NewSg)
            Data.Model.Grids(i).Pressure.Oil = NewPo
        Next
    End Sub
    Sub CorrectFullyImplicitResult(ByRef Data As Simulation, ByVal i As Integer, ByVal DeltaSw As Decimal, ByVal DeltaSg As Decimal, ByVal DeltaPo As Decimal, ByRef newSw As Decimal, ByRef newSg As Decimal, ByRef newPo As Decimal)
        Dim oil, water, gas As Boolean
        Dim Section As Decimal = 1
        Simulation_Needs_PVT_Table(Data, oil, water, gas)

        If oil Then
            If Data.Model.Grids(i).Pressure.Oil + DeltaPo > Data.PVT.Table(0, Data.PVT.Table.GetLength(1) - 1) Then
                newPo = Data.PVT.Table(0, Data.PVT.Table.GetLength(1) - 1) - Data.Methodology.Derivations.Oil_Pressure_Epsilon
                Section = (newPo - Data.Model.Grids(i).Pressure.Oil) / DeltaPo
            ElseIf Data.Model.Grids(i).Pressure.Oil + DeltaPo < Data.PVT.Table(0, 0) Then
                newPo = Data.PVT.Table(0, 0) + Data.Methodology.Derivations.Oil_Pressure_Epsilon
                Section = (newPo - Data.Model.Grids(i).Pressure.Oil) / DeltaPo
            Else
                newPo = Data.Model.Grids(i).Pressure.Oil + DeltaPo
            End If
        Else
            newPo = Data.Model.Grids(i).Pressure.Oil + DeltaPo
        End If

        newSw = Data.Model.Grids(i).Saturation.Water + Section * DeltaSw
        newSg = Data.Model.Grids(i).Saturation.Gas + Section * DeltaSg
    End Sub

    Sub Update_IMPES_Pressure(ByRef Data As Simulation, ByVal IM As IMPES_Matrix)
        For i As Integer = 0 To Data.Model.Grids.Length - 1
            Data.Model.Grids(i).Pressure.Oil = IM.Po(0, i)
        Next
    End Sub

    Function Calculate_IMPES_Iteration_MSE(ByVal New_IM As IMPES_Matrix, ByVal Last_IM As IMPES_Matrix) As Iteration_MSE
        Dim Po_MSE As Decimal
        Dim out As New Iteration_MSE

        For i As Integer = 0 To New_IM.Po.GetLength(1) - 1
            Po_MSE += Math.Abs(New_IM.Po(0, i) - Last_IM.Po(0, i))
        Next
        out.Po = Po_MSE / (New_IM.Po.GetLength(1))
        Return out
    End Function

    Function Calculate_Fully_Implicit_Iteration_MSE(ByVal IM As Fully_Implicit_Matrix) As Iteration_MSE

        Dim Sw_MSE, Sg_MSE, Po_MSE As Decimal
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
    Function Simulation_Needs_PVT_Table(ByVal Data As Simulation, ByRef ForOil As Boolean, ByRef ForGas As Boolean, ByRef ForWater As Boolean) As Boolean
        ForOil = Data.Methodology.PVT.Oil.FVF = PVT_Methods.Table Or Data.Methodology.PVT.Oil.Viscosity = PVT_Methods.Table Or Data.Methodology.PVT.Oil.Density = PVT_Methods.Table Or Data.Methodology.PVT.Oil.Compressibility = PVT_Methods.Table Or Data.Methodology.PVT.Rs = PVT_Methods.Table
        ForWater = Data.Methodology.PVT.Water.FVF = PVT_Methods.Table Or Data.Methodology.PVT.Water.Viscosity = PVT_Methods.Table Or Data.Methodology.PVT.Water.Density = PVT_Methods.Table Or Data.Methodology.PVT.Water.Compressibility = PVT_Methods.Table Or Data.Methodology.PVT.Rs = PVT_Methods.Table
        ForGas = Data.Methodology.PVT.Gas.FVF = PVT_Methods.Table Or Data.Methodology.PVT.Gas.Viscosity = PVT_Methods.Table Or Data.Methodology.PVT.Gas.Density = PVT_Methods.Table Or Data.Methodology.PVT.Gas.Compressibility = PVT_Methods.Table Or Data.Methodology.PVT.Rs = PVT_Methods.Table
        Return ForGas Or ForOil Or ForWater
    End Function

    Sub Update_IMPES_Saturation(ByVal Data_n0 As Simulation, ByRef Data_n1 As Simulation, ByRef IM As IMPES_Matrix)
        For nGrid As Integer = 0 To Data_n1.Model.Grids.Length - 1

            Dim Grid_Params As New Grid_Parameters
            Calculate_C_Bracet(Data_n0, Data_n1, nGrid, Grid_Params)

            Dim Sigma_Sw As Decimal = 0
            Dim Sigma_Sg As Decimal = 0
            Dim Sigma3 As Decimal = 0
            Dim L_Param1 As Decimal = 0
            Dim R_Param1 As Decimal = 0
            Dim R_Param2 As Decimal = 0

            Dim Cww As Decimal = -Grid_Params.R_nn.Rw_Sw
            Dim Cwg As Decimal = -Grid_Params.R_nn.Rw_Sg
            Dim Cwp As Decimal = -Grid_Params.R_nn.Rw_Po
            Dim Cow As Decimal = -Grid_Params.R_nn.Ro_Sw
            Dim Cog As Decimal = -Grid_Params.R_nn.Ro_Sg
            Dim Cop As Decimal = -Grid_Params.R_nn.Ro_Po
            Dim Cgw As Decimal = -Grid_Params.R_nn.Rg_Sw
            Dim Cgg As Decimal = -Grid_Params.R_nn.Rg_Sg
            Dim Cgp As Decimal = -Grid_Params.R_nn.Rg_Po

            Dim qo As Decimal = Data_n0.Model.Grids(nGrid).Flow.Oil
            Dim qw As Decimal = Data_n0.Model.Grids(nGrid).Flow.Water
            Dim qg As Decimal = Data_n0.Model.Grids(nGrid).Flow.Gas

            Dim Po_n0 As Decimal = Data_n0.Model.Grids(nGrid).Pressure.Oil
            Dim Po_n1 As Decimal = Data_n1.Model.Grids(nGrid).Pressure.Oil
            Dim Delta_Swn1_Swn0 As Decimal

            For iAdjacent As Integer = 0 To Data_n1.Model.Grids(nGrid).Adjacents.Length - 1
                Dim iTarget As Integer = Data_n1.Model.Grids(nGrid).Adjacents(iAdjacent).Target_Grid_Index
                Calculate_Transmissibilities(Data_n0, nGrid, iAdjacent, Grid_Params)

                Dim _To As Decimal = Grid_Params.R.To_nm
                Dim _Tw As Decimal = Grid_Params.R.Tw_nm
                Dim _Tg As Decimal = Grid_Params.R.Tg_nm
                Dim Gamma_Water As Decimal = Grid_Params.R.average_Gamma_Water
                Dim Gamma_Oil As Decimal = Grid_Params.R.average_Gamma_Oil
                Dim Gamma_Gas As Decimal = Grid_Params.R.average_Gamma_Gas

                Dim Delta_P_n1 As Decimal = Data_n1.Model.Grids(iTarget).Pressure.Oil - Data_n1.Model.Grids(nGrid).Pressure.Oil
                Dim Delta_Pc_n0 As Decimal = Data_n0.Model.Grids(iTarget).SCAL.pcow - Data_n0.Model.Grids(nGrid).SCAL.pcow
                Dim Delta_Z As Decimal = Data_n0.Model.Grids(iTarget).Location.Z - Data_n0.Model.Grids(nGrid).Location.Z


                Sigma_Sw += (_Tw * (Delta_P_n1 - Delta_Pc_n0 - Gamma_Water * Delta_Z))
                Sigma_Sg += (_To * (Delta_P_n1 - Gamma_Oil * Delta_Z))

            Next

            Delta_Swn1_Swn0 = (1 / Cww) * (Sigma_Sw - Cwp * (Po_n1 - Po_n0) + qw)

            If Data_n0.System.Phase.Water Then
                Data_n1.Model.Grids(nGrid).Saturation.Water += Delta_Swn1_Swn0
            Else
                Data_n1.Model.Grids(nGrid).Saturation.Water = 0
            End If

            If Data_n0.System.Phase.Gas Then
                Data_n1.Model.Grids(nGrid).Saturation.Gas += (1 / Cog) * (Sigma_Sg - Cop * (Po_n1 - Po_n0) - Cow * Delta_Swn1_Swn0 + qo)
            Else
                Data_n1.Model.Grids(nGrid).Saturation.Gas = 0
            End If

        Next

    End Sub

    Sub Create_IMPES_Matrix(ByVal Data_n0 As Simulation, ByRef Data_n1 As Simulation, ByRef IM As IMPES_Matrix)
        IMPES_Zero(IM)
        IM.Iteration += 1

        ' Grid Loop =====================================================================
        For nGrid As Integer = 0 To Data_n1.Model.Grids.Length - 1

            Dim Grid_Params As New Grid_Parameters
            Calculate_C_Bracet(Data_n0, Data_n1, nGrid, Grid_Params)

            Dim Po_n1m_Coe(Data_n1.Model.Grids.Length - 1) As Decimal
            Dim Po_n1n_Coe As Decimal = 0
            Dim Sigma1 As Decimal = 0
            Dim Sigma2 As Decimal = 0
            Dim Sigma3 As Decimal = 0
            Dim L_Param1 As Decimal = 0
            Dim R_Param1 As Decimal = 0
            Dim R_Param2 As Decimal = 0

            Dim iTarget As Integer
            For iAdjacent As Integer = 0 To Data_n1.Model.Grids(nGrid).Adjacents.Length - 1

                iTarget = Data_n1.Model.Grids(nGrid).Adjacents(iAdjacent).Target_Grid_Index
                Calculate_Transmissibilities(Data_n0, nGrid, iAdjacent, Grid_Params)
                Add_Sigma(Data_n0, Data_n1, Po_n1m_Coe(iTarget), Po_n1n_Coe, Sigma1, Sigma2, Sigma3, Grid_Params, nGrid, iAdjacent)
            Next
            Calculate_Static_Params(Data_n0, Data_n1, L_Param1, R_Param1, R_Param2, Grid_Params, nGrid, IM.Iteration)
            Combine_IMPES_Parameters(IM, nGrid, Po_n1m_Coe, Po_n1n_Coe, Sigma1, Sigma2, Sigma3, L_Param1, R_Param1, R_Param2)

        Next

    End Sub

    Sub Combine_IMPES_Parameters(ByRef IM As IMPES_Matrix, ByVal nGrid As Integer, ByVal Po_n1m_Coe() As Decimal, ByVal Po_n1n_Coe As Decimal, ByVal Sigma1 As Decimal, ByVal Sigma2 As Decimal, ByVal Sigma3 As Decimal, ByVal L_Param1 As Decimal, ByVal R_Param1 As Decimal, ByVal R_Param2 As Decimal)
        Dim Right_Hand_Side As Decimal = -R_Param1 - R_Param2 + Sigma1 - Sigma2 + Sigma3
        With IM
            .Target(0, nGrid) = Right_Hand_Side
            For i As Integer = 0 To SizeOf(Po_n1m_Coe) - 1
                If i <> nGrid Then
                    .LeftSide(i, nGrid) = Po_n1m_Coe(i)
                End If
            Next
            .LeftSide(nGrid, nGrid) = -(L_Param1 + Po_n1n_Coe)
        End With
    End Sub

    Sub Calculate_Static_Params(ByVal Data_0 As Simulation, ByVal Data_v As Simulation, ByRef L_Param1 As Decimal, ByRef R_Param1 As Decimal, ByRef R_Param2 As Decimal, ByVal Grid_Params As Grid_Parameters, ByVal nGrid As Integer, ByVal Iteration As Integer)
        Dim Bo As Decimal = Data_v.Model.Grids(nGrid).PVT.Oil.FVF
        Dim Bg As Decimal = Data_v.Model.Grids(nGrid).PVT.Gas.FVF
        Dim Bw As Decimal = Data_v.Model.Grids(nGrid).PVT.Water.FVF
        Dim Rs_n0 As Decimal = Data_0.Model.Grids(nGrid).PVT.Rs
        Dim Rs_n1 As Decimal = Data_v.Model.Grids(nGrid).PVT.Rs

        Dim Cww As Decimal = -Grid_Params.R_nn.Rw_Sw
        Dim Cwg As Decimal = -Grid_Params.R_nn.Rw_Sg
        Dim Cwp As Decimal = -Grid_Params.R_nn.Rw_Po
        Dim Cow As Decimal = -Grid_Params.R_nn.Ro_Sw
        Dim Cog As Decimal = -Grid_Params.R_nn.Ro_Sg
        Dim Cop As Decimal = -Grid_Params.R_nn.Ro_Po
        Dim Cgw As Decimal = -Grid_Params.R_nn.Rg_Sw
        Dim Cgg As Decimal = -Grid_Params.R_nn.Rg_Sg
        Dim Cgp As Decimal = -Grid_Params.R_nn.Rg_Po

        Dim Po_n As Decimal = Data_0.Model.Grids(nGrid).Pressure.Oil
        Dim qo As Decimal = Data_0.Model.Grids(nGrid).Flow.Oil
        Dim qw As Decimal = Data_0.Model.Grids(nGrid).Flow.Water
        Dim qg As Decimal = Data_0.Model.Grids(nGrid).Flow.Gas

        If Iteration <= 1 Then
            qo = 0
            qw = 0
            qg = 0
        End If

        L_Param1 = (Bo - Rs_n1 * Bg) * Cop + Bw * Cwp + Bg * Cgp
        R_Param1 = ((Bo - Rs_n1 * Bg) * Cop + Bw * Cwp + Bg * Cgp) * Po_n
        R_Param2 = ((Bo - Rs_n1 * Bg) * qo + Bw * qw + Bg * qg)

    End Sub

    Sub Add_Sigma(ByVal Data_0 As Simulation, ByVal Data_v As Simulation, ByRef Param_Po_n1m As Decimal, ByRef Param_Po_n1n As Decimal, ByRef Sigma1 As Decimal, ByRef Sigma2 As Decimal, ByRef Sigma3 As Decimal, ByVal Grid_Params As Grid_Parameters, ByVal nGrid As Integer, ByVal iAdjacent As Integer)
        Dim iTarget As Integer = Data_0.Model.Grids(nGrid).Adjacents(iAdjacent).Target_Grid_Index

        Dim Bo As Decimal = Data_v.Model.Grids(nGrid).PVT.Oil.FVF
        Dim Bg As Decimal = Data_v.Model.Grids(nGrid).PVT.Gas.FVF
        Dim Bw As Decimal = Data_v.Model.Grids(nGrid).PVT.Water.FVF
        Dim Rs_n0 As Decimal = Data_0.Model.Grids(nGrid).PVT.Rs
        Dim Rs_n1 As Decimal = Data_v.Model.Grids(nGrid).PVT.Rs
        Dim _To As Decimal = Grid_Params.R.To_nm
        Dim _Tw As Decimal = Grid_Params.R.Tw_nm
        Dim _Tg As Decimal = Grid_Params.R.Tg_nm
        Dim Gamma_Water As Decimal = Grid_Params.R.average_Gamma_Water
        Dim Gamma_Oil As Decimal = Grid_Params.R.average_Gamma_Oil
        Dim Gamma_Gas As Decimal = Grid_Params.R.average_Gamma_Gas

        Param_Po_n1m = ((Bo - Rs_n1 * Bg) * _To + Bw * _Tw + Bg * (_Tg + _To * Rs_n0))
        Param_Po_n1n += ((Bo - Rs_n1 * Bg) * _To + Bw * _Tw + Bg * (_Tg + Rs_n0 * _To))

        Sigma1 += (Bw * _Tw * (Data_0.Model.Grids(iTarget).SCAL.pcow - Data_0.Model.Grids(nGrid).SCAL.pcow))
        Sigma2 += (Bg * _Tg * (Data_0.Model.Grids(iTarget).SCAL.pcgo - Data_0.Model.Grids(nGrid).SCAL.pcgo))
        Sigma3 += ((Bo - Rs_n1 * Bg) * _To * Gamma_Oil + Bw * _Tw * Gamma_Water + Bg * (_Tg * Gamma_Gas + (_To * Rs_n0) * Gamma_Water)) * (Data_0.Model.Grids(iTarget).Location.Z - Data_0.Model.Grids(nGrid).Location.Z)

    End Sub


    Sub Create_Implicit_Matrix(ByVal Data_n0 As Simulation, ByRef Data_n1 As Simulation, ByRef IM As Fully_Implicit_Matrix)

        Fully_Implicit_Zero(IM)
        IM.Iteration += 1

        ' Grid Loop =====================================================================
        For nGrid As Integer = 0 To Data_n1.Model.Grids.Length - 1

            Dim Grid_Params As New Grid_Parameters

            For iAdjacent As Integer = 0 To Data_n1.Model.Grids(nGrid).Adjacents.Length - 1

                Calculate_F_Bracet_Parameters(Data_n0, Data_n1, nGrid, iAdjacent, Grid_Params)
                Add_F_Bracet_Parameters_m(Data_n1, nGrid, iAdjacent, Grid_Params, IM)
                Add_F_Bracet_Parameters_n(Data_n1, nGrid, iAdjacent, Grid_Params)
                Add_Mines_R_Bar_F_Parameters(Data_n1, nGrid, iAdjacent, Grid_Params, IM)

            Next

            Calculate_C_Bracet(Data_n0, Data_n1, nGrid, Grid_Params)
            Add_Round_R_Round_Q(Data_n0, Data_n1, nGrid, Grid_Params)
            Add_Bracet_Parameters_n(Data_n1, nGrid, Grid_Params, IM)
            Add_C_Bar(Data_n0, Data_n1, nGrid, Grid_Params, IM)
            Add_Q_Bar(Data_n1, nGrid, IM)

        Next

    End Sub
    Sub Add_Round_R_Round_Q(ByVal Data_0 As Simulation, ByVal Data_1 As Simulation, ByVal iGrid As Integer, ByRef Grid_Params As Grid_Parameters)
        Dim Qo_Po As Decimal
        Dim Qw_Po As Decimal
        Dim Qg_Po As Decimal
        Dim Qo_Sw As Decimal
        Dim Qw_Sw As Decimal
        Dim Qg_Sw As Decimal
        Dim Qo_Sg As Decimal
        Dim Qw_Sg As Decimal
        Dim Qg_Sg As Decimal

        If SizeOf(Data_0.Model.Grids(iGrid).Schedules) = 0 Then Exit Sub

        For i As Integer = 0 To SizeOf(Data_0.Model.Grids(iGrid).Schedules) - 1
            If Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Position = Simulation_Schedule_Position.Surface Then
                Dim _g_Po As New Grid
                Dim _g_Sw As New Grid
                Dim _g_Sg As New Grid

                Dim Q1, Q2 As Decimal
                With _g_Po
                    .Pressure.Oil = Data_1.Model.Grids(iGrid).Pressure.Oil + Data_1.Methodology.Derivations.Oil_Pressure_Epsilon
                    .Pressure.Water = Data_1.Model.Grids(iGrid).Pressure.Water
                    .Pressure.Gas = Data_1.Model.Grids(iGrid).Pressure.Gas
                    .Saturation.Water = Data_1.Model.Grids(iGrid).Saturation.Water
                    .Saturation.Gas = Data_1.Model.Grids(iGrid).Saturation.Gas
                    .PVT.Temperature = Data_1.Model.Grids(iGrid).PVT.Temperature
                    Copy_Size(Data_1.Model.Grids(iGrid), _g_Po)
                    Copy_Location(Data_1.Model.Grids(iGrid), _g_Po)
                End With

                Initialize_Saturations(Data_0, _g_Po)
                Initialize_Phase_System(_g_Po, Data_0.Rocks(Data_0.Model.Grids(iGrid).RockType))
                Initialize_SCAL(_g_Po, Data_0.Rocks(Data_0.Model.Grids(iGrid).RockType), 2)
                Initialize_Pressure(_g_Po)
                Initialize_PVT(_g_Po, Data_0)

                With _g_Sw
                    .Pressure.Oil = Data_1.Model.Grids(iGrid).Pressure.Oil
                    .Pressure.Water = Data_1.Model.Grids(iGrid).Pressure.Water
                    .Pressure.Gas = Data_1.Model.Grids(iGrid).Pressure.Gas
                    .Saturation.Water = Data_1.Model.Grids(iGrid).Saturation.Water + Data_1.Methodology.Derivations.Water_Saturation_Epsilon
                    .Saturation.Gas = Data_1.Model.Grids(iGrid).Saturation.Gas
                    .PVT.Temperature = Data_1.Model.Grids(iGrid).PVT.Temperature
                    Copy_Size(Data_1.Model.Grids(iGrid), _g_Sw)
                    Copy_Location(Data_1.Model.Grids(iGrid), _g_Sw)
                End With

                Initialize_Saturations(Data_0, _g_Sw)
                Initialize_Phase_System(_g_Sw, Data_0.Rocks(Data_0.Model.Grids(iGrid).RockType))
                Initialize_SCAL(_g_Sw, Data_0.Rocks(Data_0.Model.Grids(iGrid).RockType))
                Initialize_Pressure(_g_Sw)
                Initialize_PVT(_g_Sw, Data_0)

                With _g_Sg
                    .Pressure.Oil = Data_1.Model.Grids(iGrid).Pressure.Oil
                    .Pressure.Water = Data_1.Model.Grids(iGrid).Pressure.Water
                    .Pressure.Gas = Data_1.Model.Grids(iGrid).Pressure.Gas
                    .Saturation.Water = Data_1.Model.Grids(iGrid).Saturation.Water
                    .Saturation.Gas = Data_1.Model.Grids(iGrid).Saturation.Gas + Data_1.Methodology.Derivations.Gas_Saturation_Epsilon
                    .PVT.Temperature = Data_1.Model.Grids(iGrid).PVT.Temperature
                    Copy_Size(Data_1.Model.Grids(iGrid), _g_Sg)
                    Copy_Location(Data_1.Model.Grids(iGrid), _g_Sg)
                End With

                Initialize_Saturations(Data_0, _g_Sg)
                Initialize_Phase_System(_g_Sg, Data_0.Rocks(Data_0.Model.Grids(iGrid).RockType))
                Initialize_SCAL(_g_Sg, Data_0.Rocks(Data_0.Model.Grids(iGrid).RockType))
                Initialize_Pressure(_g_Sg)
                Initialize_PVT(_g_Sg, Data_0)




                '==============================================================================================================
                Select Case Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Type
                    Case Schedule_Type.Static_Oil_Pressure
                        Q1 = Data_0.Model.Grids(iGrid).Flow.Oil

                        Q2 = GetSurfaceFlow(Data_1, _g_Po, Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Orientation, Phase.Oil, Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Data(0, 0))
                        Qo_Po += (Q2 - Q1) / Data_1.Methodology.Derivations.Oil_Pressure_Epsilon
                        Q2 = GetSurfaceFlow(Data_1, _g_Sw, Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Orientation, Phase.Oil, Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Data(0, 0))
                        Qo_Sw += (Q2 - Q1) / Data_1.Methodology.Derivations.Water_Saturation_Epsilon
                        Q2 = GetSurfaceFlow(Data_1, _g_Sg, Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Orientation, Phase.Oil, Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Data(0, 0))
                        Qo_Sg += (Q2 - Q1) / Data_1.Methodology.Derivations.Gas_Saturation_Epsilon

                    Case Schedule_Type.Static_Water_Pressure
                        Q1 = Data_0.Model.Grids(iGrid).Flow.Water

                        Q2 = GetSurfaceFlow(Data_1, _g_Po, Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Orientation, Phase.Water, Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Data(0, 0))
                        Qw_Po += (Q2 - Q1) / Data_1.Methodology.Derivations.Oil_Pressure_Epsilon
                        Q2 = GetSurfaceFlow(Data_1, _g_Sw, Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Orientation, Phase.Water, Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Data(0, 0))
                        Qw_Sw += (Q2 - Q1) / Data_1.Methodology.Derivations.Water_Saturation_Epsilon
                        Q2 = GetSurfaceFlow(Data_1, _g_Sg, Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Orientation, Phase.Water, Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Data(0, 0))
                        Qw_Sg += (Q2 - Q1) / Data_1.Methodology.Derivations.Gas_Saturation_Epsilon
                    Case Schedule_Type.Static_Gas_Pressure
                        Q1 = Data_0.Model.Grids(iGrid).Flow.Gas

                        Q2 = GetSurfaceFlow(Data_1, _g_Po, Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Orientation, Phase.Gas, Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Data(0, 0))
                        Qg_Po += (Q2 - Q1) / Data_1.Methodology.Derivations.Oil_Pressure_Epsilon
                        Q2 = GetSurfaceFlow(Data_1, _g_Sw, Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Orientation, Phase.Gas, Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Data(0, 0))
                        Qg_Sw += (Q2 - Q1) / Data_1.Methodology.Derivations.Water_Saturation_Epsilon
                        Q2 = GetSurfaceFlow(Data_1, _g_Sg, Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Orientation, Phase.Gas, Data_0.Schedules(Data_0.Model.Grids(iGrid).Schedules(i)).Data(0, 0))
                        Qg_Sg += (Q2 - Q1) / Data_1.Methodology.Derivations.Gas_Saturation_Epsilon
                End Select
            End If
        Next


        With Grid_Params
            .R_nn.Ro_Po += Qo_Po
            .R_nn.Rw_Po += Qw_Po
            .R_nn.Rg_Po += Qg_Po

            .R_nn.Ro_Sw += Qo_Sw
            .R_nn.Rw_Sw += Qw_Sw
            .R_nn.Rg_Sw += Qg_Sw

            .R_nn.Ro_Sg += Qo_Sg
            .R_nn.Rw_Sg += Qw_Sg
            .R_nn.Rg_Sg += Qg_Sg
        End With

    End Sub
    Sub Add_Q_Bar(ByVal Data_v As Simulation, ByVal iGrid As Integer, ByRef IM As Fully_Implicit_Matrix)
        With IM
            .Mines_R(0, iGrid * 3) -= Data_v.Model.Grids(iGrid).Flow.Water
            .Mines_R(0, iGrid * 3 + 1) -= Data_v.Model.Grids(iGrid).Flow.Gas
            .Mines_R(0, iGrid * 3 + 2) -= Data_v.Model.Grids(iGrid).Flow.Oil
        End With
    End Sub
    Sub Add_C_Bar(ByVal Data_n As Simulation, ByVal Data_v As Simulation, ByVal iGrid As Integer, ByRef Grid_Params As Grid_Parameters, ByRef IM As Fully_Implicit_Matrix)
        Dim Vb As Decimal = Data_v.Model.Grids(iGrid).Volume
        Dim Phi As Decimal = Data_v.Model.Grids(iGrid).Porosity
        Dim t As Decimal = Data_n.Time_Step
        Dim a As Decimal = Data_v.Constants.Alpha
        Dim Sw As Decimal = Data_v.Model.Grids(iGrid).Saturation.Water
        Dim Sg As Decimal = Data_v.Model.Grids(iGrid).Saturation.Gas
        Dim Bw As Decimal = Data_v.Model.Grids(iGrid).PVT.Water.FVF
        Dim Bg As Decimal = Data_v.Model.Grids(iGrid).PVT.Gas.FVF
        Dim Bo As Decimal = Data_v.Model.Grids(iGrid).PVT.Oil.FVF
        Dim Rs As Decimal = Data_v.Model.Grids(iGrid).PVT.Rs

        With IM
            .Mines_R(0, iGrid * 3) += (((Vb / (a * t)) * ((Phi * Sw) / Bw)) - Data_n.C_Bar_n(0, iGrid * 3))
            .Mines_R(0, iGrid * 3 + 1) += ((((Vb / (a * t)) * ((Phi * Sg) / Bg) + (Phi * Rs * (1 - Sw - Sg)) / Bo)) - Data_n.C_Bar_n(0, iGrid * 3 + 1))
            .Mines_R(0, iGrid * 3 + 2) += (((Vb / (a * t)) * (Phi * (1 - Sw - Sg)) / Bo) - Data_n.C_Bar_n(0, iGrid * 3 + 2))
        End With

    End Sub

    Sub Add_Mines_R_Bar_F_Parameters(ByVal Data As Simulation, ByVal iGrid As Integer, ByVal iAdjacent As Integer, ByVal Grid_Params As Grid_Parameters, ByRef IM As Fully_Implicit_Matrix)
        Dim adjID As Integer = Data.Model.Grids(iGrid).Adjacents(iAdjacent).Target_Grid_Index
        Dim Pm As Decimal = Data.Model.Grids(adjID).Pressure.Oil
        Dim Pn As Decimal = Data.Model.Grids(iGrid).Pressure.Oil
        Dim _Tw As Decimal = Grid_Params.R.Tw_nm
        Dim _To As Decimal = Grid_Params.R.To_nm
        Dim _Tg As Decimal = Grid_Params.R.Tg_nm
        Dim pcgo_m As Decimal = Data.Model.Grids(adjID).SCAL.pcgo
        Dim pcow_m As Decimal = Data.Model.Grids(adjID).SCAL.pcow
        Dim pcgo_n As Decimal = Data.Model.Grids(iGrid).SCAL.pcgo
        Dim pcow_n As Decimal = Data.Model.Grids(iGrid).SCAL.pcow

        Dim Gw As Decimal = Grid_Params.R.average_Gamma_Water
        Dim Go As Decimal = Grid_Params.R.average_Gamma_Oil
        Dim Gg As Decimal = Grid_Params.R.average_Gamma_Gas
        Dim Zm As Decimal = Data.Model.Grids(adjID).Location.Z
        Dim Zn As Decimal = Data.Model.Grids(iGrid).Location.Z

        Dim Rsm As Decimal = Data.Model.Grids(adjID).PVT.Rs
        Dim Rsn As Decimal = Data.Model.Grids(iGrid).PVT.Rs

        With IM
            .Mines_R(0, iGrid * 3) -= (_Tw * ((Pm - Pn) - (pcow_m - pcow_n) - Gw * (Zm - Zn)))
            .Mines_R(0, iGrid * 3 + 1) -= ((_Tg * ((Pm - Pn) + (pcgo_m - pcgo_n) - Gg * (Zm - Zn))) + (_To * (Rsn + Rsm) / 2) * ((Pm - Pn) - Go * (Zm - Zn)))
            .Mines_R(0, iGrid * 3 + 2) -= (_To * (((Pm - Pn) - Go * (Zm - Zn))))
        End With

    End Sub

    Sub Add_Bracet_Parameters_n(ByVal Data As Simulation, ByVal iGrid As Integer, ByRef Grid_Params As Grid_Parameters, ByRef IM As Fully_Implicit_Matrix)
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

        Dim Phi_prime As Decimal
        Dim Rs_prime As Decimal
        Dim one_over_bo_prime As Decimal
        Dim one_over_bg_prime As Decimal
        Dim one_over_bw_prime As Decimal
        Dim DeltaPressure As Double = (Data_1.Model.Grids(iGrid).Pressure.Oil - Data_0.Model.Grids(iGrid).Pressure.Oil)
        If DeltaPressure <> 0 Then

            Phi_prime = (Data_1.Model.Grids(iGrid).Porosity - Data_0.Model.Grids(iGrid).Porosity) / DeltaPressure
            Rs_prime = (Data_1.Model.Grids(iGrid).PVT.Rs - Data_0.Model.Grids(iGrid).PVT.Rs) / DeltaPressure
            one_over_bo_prime = ((1 / Data_1.Model.Grids(iGrid).PVT.Oil.FVF) - (1 / Data_0.Model.Grids(iGrid).PVT.Oil.FVF)) / DeltaPressure
            one_over_bg_prime = ((1 / Data_1.Model.Grids(iGrid).PVT.Gas.FVF) - (1 / Data_0.Model.Grids(iGrid).PVT.Gas.FVF)) / DeltaPressure
            one_over_bw_prime = ((1 / Data_1.Model.Grids(iGrid).PVT.Water.FVF) - (1 / Data_0.Model.Grids(iGrid).PVT.Water.FVF)) / DeltaPressure

        End If

        Dim Vb As Decimal = Data_0.Model.Grids(iGrid).Volume
        Dim a As Decimal = Data_0.Constants.Alpha
        Dim t As Decimal = Data_0.Time_Step
        Dim Phi As Decimal = Data_1.Model.Grids(iGrid).Porosity
        Dim Bo_1 As Decimal = Data_1.Model.Grids(iGrid).PVT.Oil.FVF
        Dim Bw_1 As Decimal = Data_1.Model.Grids(iGrid).PVT.Water.FVF
        Dim Bg_1 As Decimal = Data_1.Model.Grids(iGrid).PVT.Gas.FVF
        Dim Bo_0 As Decimal = Data_0.Model.Grids(iGrid).PVT.Oil.FVF
        Dim Bw_0 As Decimal = Data_0.Model.Grids(iGrid).PVT.Water.FVF
        Dim Bg_0 As Decimal = Data_0.Model.Grids(iGrid).PVT.Gas.FVF

        Dim Sg As Decimal = Data_0.Model.Grids(iGrid).Saturation.Gas
        Dim Sw As Decimal = Data_0.Model.Grids(iGrid).Saturation.Water

        Dim Rs_1 As Decimal = Data_1.Model.Grids(iGrid).PVT.Rs
        Dim Rs_0 As Decimal = Data_0.Model.Grids(iGrid).PVT.Rs


        ' Cww ---
        Grid_Params.R_nn.Rw_Sw -= (Vb / (a * t)) * (Phi / Bw_1)

        ' Cwg ---
        Grid_Params.R_nn.Rw_Sg -= 0

        ' Cwp ---
        Grid_Params.R_nn.Rw_Po -= (Vb / (a * t)) * ((Phi_prime / Bw_0) + Phi * one_over_bw_prime) * Sw

        ' Cow ---
        Grid_Params.R_nn.Ro_Sw -= -(Vb / (a * t)) * (Phi / Bo_1)

        ' Cog ---
        Grid_Params.R_nn.Ro_Sg -= -(Vb / (a * t)) * (Phi / Bo_1)

        ' Cop ---
        Grid_Params.R_nn.Ro_Po -= (Vb / (a * t)) * ((Phi_prime / Bo_0) + Phi * one_over_bo_prime) * (1 - Sw - Sg)

        ' Cgw ---
        Grid_Params.R_nn.Rg_Sw -= -(Vb / (a * t)) * (Phi / Bo_1) * Rs_1

        ' Cgg ---
        Grid_Params.R_nn.Rg_Sg -= (Vb / (a * t)) * ((Phi / Bg_1) - (Phi / Bo_1) * Rs_1)

        ' Cgp ---
        Grid_Params.R_nn.Rg_Po -= (Vb / (a * t)) * ((((Phi_prime / Bo_0 + Phi * one_over_bo_prime) * Rs_0) + (Phi / Bo_1) * Rs_prime) * (1 - Sw - Sg) + ((Phi_prime / Bg_0) + Phi * one_over_bg_prime) * Sg)
    End Sub


    Sub Add_F_Bracet_Parameters_n(ByVal Data As Simulation, ByVal iGrid As Integer, ByVal iAdjacent As Integer, ByRef Grid_Params As Grid_Parameters)
        Dim adjID As Integer = Data.Model.Grids(iGrid).Adjacents(iAdjacent).Target_Grid_Index

        Dim nGrid As Grid = Data.Model.Grids(iGrid)
        Dim mGrid As Grid = Data.Model.Grids(adjID)

        Dim Pm As Decimal = mGrid.Pressure.Oil
        Dim Pn As Decimal = nGrid.Pressure.Oil
        Dim Zm As Decimal = mGrid.Location.Z
        Dim Zn As Decimal = nGrid.Location.Z
        Dim Rsm As Decimal = mGrid.PVT.Rs
        Dim Rsn As Decimal = nGrid.PVT.Rs

        Dim _Tw As Decimal = Grid_Params.R.Tw_nm
        Dim _To As Decimal = Grid_Params.R.To_nm
        Dim _Tg As Decimal = Grid_Params.R.Tg_nm

        Dim Tw_Sw_n As Decimal = Grid_Params.R.round_Tw_Sw_n
        Dim Tw_Po_n As Decimal = Grid_Params.R.round_Tw_Po_n

        Dim To_Sw_n As Decimal = Grid_Params.R.round_To_Sw_n
        Dim To_Sg_n As Decimal = Grid_Params.R.round_To_Sg_n
        Dim To_Po_n As Decimal = Grid_Params.R.round_To_Po_n

        Dim Tg_Sg_n As Decimal = Grid_Params.R.round_Tg_Sg_n
        Dim Tg_Po_n As Decimal = Grid_Params.R.round_Tg_Po_n

        Dim Gw As Decimal = Grid_Params.R.average_Gamma_Water
        Dim Go As Decimal = Grid_Params.R.average_Gamma_Oil
        Dim Gg As Decimal = Grid_Params.R.average_Gamma_Gas

        Dim d_pcow_n As Decimal = Grid_Params.R.Pcow_Prime_n
        Dim d_pcgo_n As Decimal = Grid_Params.R.Pcgo_Prime_n

        Dim pcgo_m As Decimal = mGrid.SCAL.pcgo
        Dim pcgo_n As Decimal = nGrid.SCAL.pcgo

        Dim pcow_m As Decimal = mGrid.SCAL.pcow
        Dim pcow_n As Decimal = nGrid.SCAL.pcow

        Dim dZ As Decimal = (Zm - Zn)
        Dim dP As Decimal = (Pm - Pn)
        Dim dPcow As Decimal = (pcow_m - pcow_n)
        Dim dPcgo As Decimal = (pcgo_m - pcgo_n)
        Dim Rs As Decimal = (Rsm + Rsn) / 2


        ' ∂Rwn /∂Swn (0,0)
        Grid_Params.R_nn.Rw_Sw += ((dP - dPcow - Gw * dZ) * Tw_Sw_n + _Tw * d_pcow_n)

        ' ∂Rwn /∂Sgn (1,0)
        Grid_Params.R_nn.Rw_Sg += 0

        ' ∂Rwn /∂Pon (2,0)
        Grid_Params.R_nn.Rw_Po += (-_Tw + (dP - dPcow - (Gw * dZ)) * Tw_Po_n)

        ' ∂Rgn /∂Swn (0,1)
        Grid_Params.R_nn.Rg_Sw += ((dP - Go * dZ) * (To_Sw_n * Rs))

        ' ∂Rgn /∂Sgn (1,1)
        Grid_Params.R_nn.Rg_Sg += ((dP + dPcgo - Gg * dZ) * Tg_Sg_n - _Tg * d_pcgo_n + (dP - Go * dZ) * To_Sg_n * Rs)

        ' ∂Rgn /∂Pon (2,1)
        Grid_Params.R_nn.Rg_Po += (-_Tg + (dP + dPcgo - Gg * dZ) * Tg_Po_n - _To * Rs + (dP - Go * dZ) * To_Po_n * Rs)

        ' ∂Ron /∂Swn (0,2)
        Grid_Params.R_nn.Ro_Sw += (dP - Go * dZ) * To_Sw_n

        ' ∂Ron /∂Sgn (1,2)
        Grid_Params.R_nn.Ro_Sg += (dP - Go * dZ) * To_Sg_n

        ' ∂Ron /∂Pon (2,2)
        Grid_Params.R_nn.Ro_Po += (-_To + (dP - Go * dZ) * To_Po_n)

    End Sub

    Sub Add_F_Bracet_Parameters_m(ByVal Data As Simulation, ByVal iGrid As Integer, ByVal iAdjacent As Integer, ByVal Grid_Params As Grid_Parameters, ByRef IM As Fully_Implicit_Matrix)
        Dim adjID As Integer = Data.Model.Grids(iGrid).Adjacents(iAdjacent).Target_Grid_Index

        Dim nGrid As Grid = Data.Model.Grids(iGrid)
        Dim mGrid As Grid = Data.Model.Grids(adjID)

        Dim Pm As Decimal = mGrid.Pressure.Oil
        Dim Pn As Decimal = nGrid.Pressure.Oil
        Dim Zm As Decimal = mGrid.Location.Z
        Dim Zn As Decimal = nGrid.Location.Z
        Dim Rsm As Decimal = mGrid.PVT.Rs
        Dim Rsn As Decimal = nGrid.PVT.Rs

        Dim _Tw As Decimal = Grid_Params.R.Tw_nm
        Dim _To As Decimal = Grid_Params.R.To_nm
        Dim _Tg As Decimal = Grid_Params.R.Tg_nm

        Dim Tw_Sw_m As Decimal = Grid_Params.R.round_Tw_Sw_m
        Dim Tw_Po_m As Decimal = Grid_Params.R.round_Tw_Po_m

        Dim To_Sw_m As Decimal = Grid_Params.R.round_To_Sw_m
        Dim To_Sg_m As Decimal = Grid_Params.R.round_To_Sg_m
        Dim To_Po_m As Decimal = Grid_Params.R.round_To_Po_m

        Dim Tg_Sg_m As Decimal = Grid_Params.R.round_Tg_Sg_m
        Dim Tg_Po_m As Decimal = Grid_Params.R.round_Tg_Po_m


        Dim Gw As Decimal = Grid_Params.R.average_Gamma_Water
        Dim Go As Decimal = Grid_Params.R.average_Gamma_Oil
        Dim Gg As Decimal = Grid_Params.R.average_Gamma_Gas

        Dim d_pcow_m As Decimal = Grid_Params.R.Pcow_Prime_m
        Dim d_pcgo_m As Decimal = Grid_Params.R.Pcgo_Prime_m


        Dim dZ As Decimal = (Zm - Zn)
        Dim dP As Decimal = (Pm - Pn)
        Dim dPcow As Decimal = (mGrid.SCAL.pcow - nGrid.SCAL.pcow)
        Dim dPcgo As Decimal = (mGrid.SCAL.pcgo - nGrid.SCAL.pcgo)
        Dim Rs As Decimal = (Rsm + Rsn) / 2

        With IM
            ' ∂Rwn /∂Swm (0,0)
            .J(adjID * 3, iGrid * 3) += ((dP - dPcow - Gw * dZ) * Tw_Sw_m - _Tw * d_pcow_m)

            ' ∂Rwn /∂Sgm (1,0)
            .J(adjID * 3 + 1, iGrid * 3) += 0

            ' ∂Rwn /∂Pom (2,0)
            .J(adjID * 3 + 2, iGrid * 3) += (_Tw + (dP - dPcow - Gw * dZ) * Tw_Po_m)

            ' ∂Rgn /∂Swm (0,1)
            .J(adjID * 3, iGrid * 3 + 1) += ((dP - Go * dZ) * (To_Sw_m * Rs))

            ' ∂Rgn /∂Sgm (1,1)
            .J(adjID * 3 + 1, iGrid * 3 + 1) += ((dP + dPcgo - Gg * dZ) * Tg_Sg_m + _Tg * d_pcgo_m + (dP - Go * dZ) * To_Sg_m * Rs)

            ' ∂Rgn /∂Pom (2,1)
            .J(adjID * 3 + 2, iGrid * 3 + 1) += (_Tg + (dP + dPcgo - Gg * dZ) * Tg_Po_m + _To * Rs + (dP - Go * dZ) * To_Po_m * Rs)

            ' ∂Ron /∂Swm (0,2)
            .J(adjID * 3, iGrid * 3 + 2) += (dP - Go * dZ) * To_Sw_m

            ' ∂Ron /∂Sgm (1,2)
            .J(adjID * 3 + 1, iGrid * 3 + 2) += (dP - Go * dZ) * To_Sg_m

            ' ∂Ron /∂Pom (2,2)
            .J(adjID * 3 + 2, iGrid * 3 + 2) += (_To + (dP - Go * dZ) * To_Po_m)

        End With

    End Sub

    Sub Calculate_Average_Gamma(ByVal Data_n As Simulation, ByVal nGrid As Integer, ByVal iAdjacent As Integer, ByRef Grid_Params As Grid_Parameters)
        Dim _area As Decimal = Data_n.Model.Grids(nGrid).Adjacents(iAdjacent).Area
        Dim adjID As Integer = Data_n.Model.Grids(nGrid).Adjacents(iAdjacent).Target_Grid_Index
        Grid_Params.R.average_Gamma_Water = (Data_n.Model.Grids(nGrid).PVT.Water.Pressure_Gradient(Data_n.System.Units) + Data_n.Model.Grids(adjID).PVT.Water.Pressure_Gradient(Data_n.System.Units)) / 2
        Grid_Params.R.average_Gamma_Oil = (Data_n.Model.Grids(nGrid).PVT.Oil.Pressure_Gradient(Data_n.System.Units) + Data_n.Model.Grids(adjID).PVT.Oil.Pressure_Gradient(Data_n.System.Units)) / 2
        Grid_Params.R.average_Gamma_Gas = (Data_n.Model.Grids(nGrid).PVT.Gas.Pressure_Gradient(Data_n.System.Units) + Data_n.Model.Grids(adjID).PVT.Gas.Pressure_Gradient(Data_n.System.Units)) / 2

    End Sub
    Sub Calculate_Transmissibilities(ByVal Data_n As Simulation, ByVal nGrid As Integer, ByVal iAdjacent As Integer, ByRef Grid_Params As Grid_Parameters)
        Dim _area As Decimal = Data_n.Model.Grids(nGrid).Adjacents(iAdjacent).Area
        Dim adjID As Integer = Data_n.Model.Grids(nGrid).Adjacents(iAdjacent).Target_Grid_Index

        Dim l As Decimal
        Dim Fp_Oil As Decimal
        Dim Fp_Water As Decimal
        Dim Fp_Gas As Decimal
        Dim ko, kw, kg As Decimal

        Calculate_Average_Gamma(Data_n, nGrid, iAdjacent, Grid_Params)

        Fp_Oil_Water_Gas(Data_n.Model.Grids(nGrid), Data_n.Model.Grids(adjID), Data_n.Methodology.Transmissibility_Fp_Averaging, Fp_Oil, Fp_Water, Fp_Gas)
        Ko_Kw_Kg(Data_n, Data_n.Model.Grids(nGrid), Data_n.Model.Grids(adjID), Data_n.Model.Grids(nGrid).Adjacents(iAdjacent).Type, ko, kw, kg, l, Grid_Params.R.average_Gamma_Water, Grid_Params.R.average_Gamma_Oil, Grid_Params.R.average_Gamma_Gas)

        Grid_Params.R.To_nm = IIf(Data_n.System.Phase.Oil, ((ko * _area * Data_n.Constants.Beta) / l) * Fp_Oil, 0)
        Grid_Params.R.Tw_nm = IIf(Data_n.System.Phase.Water, ((kw * _area * Data_n.Constants.Beta) / l) * Fp_Water, 0)
        Grid_Params.R.Tg_nm = IIf(Data_n.System.Phase.Gas, ((kg * _area * Data_n.Constants.Beta) / l) * Fp_Gas, 0)

    End Sub

    Sub GetStream(ByVal nGrid As Grid, ByVal mGrid As Grid, ByVal Type As Adjacent_Type, ByVal Gamma_Oil As Decimal, ByVal Gamma_Gas As Decimal, ByVal Gamma_Water As Decimal, ByRef Oil As Boolean, ByRef Gas As Boolean, ByRef Water As Boolean)

        Select Case Type
            Case Adjacent_Type.X, Adjacent_Type.Y
                Oil = nGrid.Pressure.Oil > mGrid.Pressure.Oil
                Water = nGrid.Pressure.Water > mGrid.Pressure.Water
                Gas = nGrid.Pressure.Gas > mGrid.Pressure.Gas
            Case Adjacent_Type.Z
                Oil = nGrid.Pressure.Oil + Gamma_Oil * nGrid.Location.Z > mGrid.Pressure.Oil + Gamma_Oil * mGrid.Location.Z
                Water = nGrid.Pressure.Water + Gamma_Water * nGrid.Location.Z > mGrid.Pressure.Water + Gamma_Water * mGrid.Location.Z
                Gas = nGrid.Pressure.Gas + Gamma_Gas * nGrid.Location.Z > mGrid.Pressure.Gas + Gamma_Gas * mGrid.Location.Z
        End Select
    End Sub
    Sub Calculate_F_Bracet_Parameters(ByVal data_n As Simulation, ByVal Data_v As Simulation, ByVal nGrid As Integer, ByVal iAdjacent As Integer, ByRef Grid_Params As Grid_Parameters)
        Dim _area As Decimal = Data_v.Model.Grids(nGrid).Adjacents(iAdjacent).Area
        Dim adjID As Integer = Data_v.Model.Grids(nGrid).Adjacents(iAdjacent).Target_Grid_Index

        Dim n_Grid As Grid = Data_v.Model.Grids(nGrid)
        Dim m_Grid As Grid = Data_v.Model.Grids(adjID)
        Dim Adjacent_T As Adjacent_Type = n_Grid.Adjacents(iAdjacent).Type

        Dim For_Sw_Round_m As New Grid
        Dim For_Sg_Round_m As New Grid
        Dim For_Po_Round_m As New Grid

        Dim For_Sw_Round_n As New Grid
        Dim For_Sg_Round_n As New Grid
        Dim For_Po_Round_n As New Grid

        Dim l As Decimal
        Dim Fp_Oil As Decimal
        Dim Fp_Water As Decimal
        Dim Fp_Gas As Decimal
        Dim ko, kw, kg As Decimal

        Dim UpStreamOil, UpStreamWater, UpStreamGas As Boolean

        Calculate_Average_Gamma(data_n, nGrid, iAdjacent, Grid_Params)
        Initialize_Forward_Grids(Data_v, nGrid, adjID, For_Sg_Round_n, For_Sw_Round_n, For_Po_Round_n, For_Sg_Round_m, For_Sw_Round_m, For_Po_Round_m)

        GetStream(n_Grid, m_Grid, Adjacent_T, Grid_Params.R.average_Gamma_Oil, Grid_Params.R.average_Gamma_Gas, Grid_Params.R.average_Gamma_Water, UpStreamOil, UpStreamGas, UpStreamWater)

        ' Non Round Region ==============================================================================================
        Fp_Oil_Water_Gas(n_Grid, m_Grid, Data_v.Methodology.Transmissibility_Fp_Averaging, Fp_Oil, Fp_Water, Fp_Gas)
        Ko_Kw_Kg(data_n, n_Grid, m_Grid, Adjacent_T, ko, kw, kg, l, Grid_Params.R.average_Gamma_Water, Grid_Params.R.average_Gamma_Oil, Grid_Params.R.average_Gamma_Gas)

        Grid_Params.R.To_nm = IIf(data_n.System.Phase.Oil, ((ko * _area * Data_v.Constants.Beta) / l) * Fp_Oil, 0)
        Grid_Params.R.Tw_nm = IIf(data_n.System.Phase.Water, ((kw * _area * Data_v.Constants.Beta) / l) * Fp_Water, 0)
        Grid_Params.R.Tg_nm = IIf(data_n.System.Phase.Gas, ((kg * _area * Data_v.Constants.Beta) / l) * Fp_Gas, 0)

        'm Region ====================================================================================================================================================================================
        'Sw-m Round Region ====================================================================================================================================================================
        Ko_Kw_Kg(data_n, n_Grid, For_Sw_Round_m, Adjacent_T, ko, kw, kg, l, Grid_Params.R.average_Gamma_Water, Grid_Params.R.average_Gamma_Oil, Grid_Params.R.average_Gamma_Gas)

        Grid_Params.R.round_To_Sw_m = IIf(data_n.System.Phase.Water, (((ko * _area * Data_v.Constants.Beta) / l) * Fp_Oil - Grid_Params.R.To_nm) / Data_v.Methodology.Derivations.Water_Saturation_Epsilon, 0)
        Grid_Params.R.round_Tw_Sw_m = IIf(data_n.System.Phase.Water, (((kw * _area * Data_v.Constants.Beta) / l) * Fp_Water - Grid_Params.R.Tw_nm) / Data_v.Methodology.Derivations.Water_Saturation_Epsilon, 0)
        Grid_Params.R.Pcow_Prime_m = (For_Sw_Round_m.SCAL.pcow - m_Grid.SCAL.pcow) / Data_v.Methodology.Derivations.Water_Saturation_Epsilon

        'Sg-m Round Region ====================================================================================================================================================================
        Ko_Kw_Kg(data_n, n_Grid, For_Sg_Round_m, Adjacent_T, ko, kw, kg, l, Grid_Params.R.average_Gamma_Water, Grid_Params.R.average_Gamma_Oil, Grid_Params.R.average_Gamma_Gas)

        Grid_Params.R.round_To_Sg_m = IIf(data_n.System.Phase.Gas, (((ko * _area * Data_v.Constants.Beta) / l) * Fp_Oil - Grid_Params.R.To_nm) / Data_v.Methodology.Derivations.Gas_Saturation_Epsilon, 0)
        Grid_Params.R.round_Tg_Sg_m = IIf(data_n.System.Phase.Gas, (((kg * _area * Data_v.Constants.Beta) / l) * Fp_Gas - Grid_Params.R.Tg_nm) / Data_v.Methodology.Derivations.Gas_Saturation_Epsilon, 0)
        Grid_Params.R.Pcgo_Prime_m = (For_Sg_Round_m.SCAL.pcgo - m_Grid.SCAL.pcgo) / Data_v.Methodology.Derivations.Gas_Saturation_Epsilon

        'Po-m Round Region ====================================================================================================================================================================
        Fp_Oil_Water_Gas(n_Grid, For_Po_Round_m, Data_v.Methodology.Transmissibility_Fp_Averaging, Fp_Oil, Fp_Water, Fp_Gas)
        Ko_Kw_Kg(data_n, n_Grid, m_Grid, Adjacent_T, ko, kw, kg, l, Grid_Params.R.average_Gamma_Water, Grid_Params.R.average_Gamma_Oil, Grid_Params.R.average_Gamma_Gas)
        Dim Rs1 As Decimal = (n_Grid.PVT.Rs + m_Grid.PVT.Rs) / 2
        Dim Rs2 As Decimal = (n_Grid.PVT.Rs + For_Po_Round_m.PVT.Rs) / 2

        Grid_Params.R.round_To_Po_m = (((ko * _area * Data_v.Constants.Beta) / l) * Fp_Oil - Grid_Params.R.To_nm) / Data_v.Methodology.Derivations.Oil_Pressure_Epsilon
        Grid_Params.R.round_To_Rs_Po_m = (Rs2 * (((ko * _area * Data_v.Constants.Beta) / l) * Fp_Oil) - (Grid_Params.R.To_nm * Rs1)) / Data_v.Methodology.Derivations.Oil_Pressure_Epsilon
        Grid_Params.R.round_Tw_Po_m = (((kw * _area * Data_v.Constants.Beta) / l) * Fp_Water - Grid_Params.R.Tw_nm) / Data_v.Methodology.Derivations.Oil_Pressure_Epsilon
        Grid_Params.R.round_Tg_Po_m = (((kg * _area * Data_v.Constants.Beta) / l) * Fp_Gas - Grid_Params.R.Tg_nm) / Data_v.Methodology.Derivations.Oil_Pressure_Epsilon

        'n Region ====================================================================================================================================================================================
        'Sw-n Round Region ====================================================================================================================================================================
        Fp_Oil_Water_Gas(n_Grid, m_Grid, Data_v.Methodology.Transmissibility_Fp_Averaging, Fp_Oil, Fp_Water, Fp_Gas)
        Ko_Kw_Kg(data_n, For_Sw_Round_n, m_Grid, Adjacent_T, ko, kw, kg, l, Grid_Params.R.average_Gamma_Water, Grid_Params.R.average_Gamma_Oil, Grid_Params.R.average_Gamma_Gas)

        Grid_Params.R.round_To_Sw_n = IIf(data_n.System.Phase.Water, (((ko * _area * Data_v.Constants.Beta) / l) * Fp_Oil - Grid_Params.R.To_nm) / Data_v.Methodology.Derivations.Water_Saturation_Epsilon, 0)
        Grid_Params.R.round_Tw_Sw_n = IIf(data_n.System.Phase.Water, (((kw * _area * Data_v.Constants.Beta) / l) * Fp_Water - Grid_Params.R.Tw_nm) / Data_v.Methodology.Derivations.Water_Saturation_Epsilon, 0)
        Grid_Params.R.Pcow_Prime_n = (For_Sw_Round_n.SCAL.pcow - n_Grid.SCAL.pcow) / Data_v.Methodology.Derivations.Water_Saturation_Epsilon

        'Sg-n Round Region ====================================================================================================================================================================
        Ko_Kw_Kg(data_n, For_Sg_Round_n, m_Grid, Adjacent_T, ko, kw, kg, l, Grid_Params.R.average_Gamma_Water, Grid_Params.R.average_Gamma_Oil, Grid_Params.R.average_Gamma_Gas)

        Grid_Params.R.round_To_Sg_n = IIf(data_n.System.Phase.Gas, (((ko * _area * Data_v.Constants.Beta) / l) * Fp_Oil - Grid_Params.R.To_nm) / Data_v.Methodology.Derivations.Gas_Saturation_Epsilon, 0)
        Grid_Params.R.round_Tg_Sg_n = IIf(data_n.System.Phase.Gas, (((kg * _area * Data_v.Constants.Beta) / l) * Fp_Gas - Grid_Params.R.Tg_nm) / Data_v.Methodology.Derivations.Gas_Saturation_Epsilon, 0)
        Grid_Params.R.Pcgo_Prime_n = (For_Sg_Round_n.SCAL.pcgo - n_Grid.SCAL.pcgo) / Data_v.Methodology.Derivations.Gas_Saturation_Epsilon

        'Po-n Round Region ====================================================================================================================================================================
        Fp_Oil_Water_Gas(For_Po_Round_n, m_Grid, Data_v.Methodology.Transmissibility_Fp_Averaging, Fp_Oil, Fp_Water, Fp_Gas)
        Ko_Kw_Kg(data_n, n_Grid, m_Grid, Adjacent_T, ko, kw, kg, l, Grid_Params.R.average_Gamma_Water, Grid_Params.R.average_Gamma_Oil, Grid_Params.R.average_Gamma_Gas)
        Rs1 = (n_Grid.PVT.Rs + m_Grid.PVT.Rs) / 2
        Rs2 = (For_Sw_Round_n.PVT.Rs + m_Grid.PVT.Rs) / 2

        Grid_Params.R.round_To_Po_n = (((ko * _area * Data_v.Constants.Beta) / l) * Fp_Oil - Grid_Params.R.To_nm) / Data_v.Methodology.Derivations.Oil_Pressure_Epsilon
        Grid_Params.R.round_To_Rs_Po_n = (Rs2 * (((ko * _area * Data_v.Constants.Beta) / l) * Fp_Oil) - (Grid_Params.R.To_nm * Rs1)) / Data_v.Methodology.Derivations.Oil_Pressure_Epsilon
        Grid_Params.R.round_Tw_Po_n = (((kw * _area * Data_v.Constants.Beta) / l) * Fp_Water - Grid_Params.R.Tw_nm) / Data_v.Methodology.Derivations.Oil_Pressure_Epsilon
        Grid_Params.R.round_Tg_Po_n = (((kg * _area * Data_v.Constants.Beta) / l) * Fp_Gas - Grid_Params.R.Tg_nm) / Data_v.Methodology.Derivations.Oil_Pressure_Epsilon

    End Sub
    Dim strx As String
    Sub Copy_SCAL(ByVal From_Grid As Grid, ByRef TO_Grid As Grid)
        With From_Grid.SCAL
            TO_Grid.SCAL.pcgo_imb = From_Grid.SCAL.pcgo_imb
            TO_Grid.SCAL.pcgo_drg = From_Grid.SCAL.pcgo_drg

            TO_Grid.SCAL.pcow_drg = From_Grid.SCAL.pcow_drg
            TO_Grid.SCAL.pcow_imb = From_Grid.SCAL.pcow_imb

            TO_Grid.SCAL.pcow = From_Grid.SCAL.pcow
            TO_Grid.SCAL.pcgo = From_Grid.SCAL.pcgo
            TO_Grid.SCAL.Kro.Kx = From_Grid.SCAL.Kro.Kx
            TO_Grid.SCAL.Kro.Ky = From_Grid.SCAL.Kro.Ky
            TO_Grid.SCAL.Kro.Kz = From_Grid.SCAL.Kro.Kz
            TO_Grid.SCAL.Krw.Kx = From_Grid.SCAL.Krw.Kx
            TO_Grid.SCAL.Krw.Ky = From_Grid.SCAL.Krw.Ky
            TO_Grid.SCAL.Krw.Kz = From_Grid.SCAL.Krw.Kz
            TO_Grid.SCAL.Krg.Kx = From_Grid.SCAL.Krg.Kx
            TO_Grid.SCAL.Krg.Ky = From_Grid.SCAL.Krg.Ky
            TO_Grid.SCAL.Krg.Kz = From_Grid.SCAL.Krg.Kz
        End With

    End Sub
    Sub Copy_Size(ByVal From_Grid As Grid, ByRef To_Grid As Grid)
        With To_Grid.Size
            .X_Length = From_Grid.Size.X_Length
            .Y_Length = From_Grid.Size.Y_Length
            .Z_Length = From_Grid.Size.Z_Length
        End With
    End Sub
    Sub Copy_Location(ByVal From_Grid As Grid, ByRef To_Grid As Grid)
        With To_Grid.Location
            .X = From_Grid.Location.X
            .Y = From_Grid.Location.Y
            .Z = From_Grid.Location.Z
        End With
    End Sub
    Sub Copy_Pressure(ByVal From_Grid As Grid, ByRef To_Grid As Grid)
        With To_Grid.Pressure
            .Oil = From_Grid.Pressure.Oil
            .Water = From_Grid.Pressure.Water
            .Gas = From_Grid.Pressure.Gas
        End With
    End Sub

    Sub Initialize_Forward_Grids(ByVal Data As Simulation, ByVal nGrid As Integer, ByVal adjID As Integer, ByRef For_Sg_Round_n As Grid, ByRef For_Sw_Round_n As Grid, ByRef For_Po_Round_n As Grid, ByRef For_Sg_Round_m As Grid, ByRef For_Sw_Round_m As Grid, ByRef For_Po_Round_m As Grid)

        'Round n Region ================================================================
        Copy_Size(Data.Model.Grids(nGrid), For_Sw_Round_n)
        Copy_Location(Data.Model.Grids(nGrid), For_Sw_Round_n)
        Copy_Pressure(Data.Model.Grids(nGrid), For_Sw_Round_n)
        For_Sw_Round_n.Saturation.Water = (Data.Model.Grids(nGrid).Saturation.Water + Data.Methodology.Derivations.Water_Saturation_Epsilon)
        For_Sw_Round_n.Saturation.Gas = Data.Model.Grids(nGrid).Saturation.Gas
        For_Sw_Round_n.PVT.Temperature = Data.Model.Grids(nGrid).PVT.Temperature
        Initialize_Saturations(Data, For_Sw_Round_n)
        Initialize_Phase_System(For_Sw_Round_n, Data.Rocks(Data.Model.Grids(nGrid).RockType))
        If Data.System.Phase.Water Then
            Initialize_SCAL(For_Sw_Round_n, Data.Rocks(Data.Model.Grids(nGrid).RockType))
        Else
            Copy_SCAL(Data.Model.Grids(nGrid), For_Sw_Round_n)
        End If

        Copy_Size(Data.Model.Grids(nGrid), For_Sg_Round_n)
        Copy_Location(Data.Model.Grids(nGrid), For_Sg_Round_n)
        Copy_Pressure(Data.Model.Grids(nGrid), For_Sg_Round_n)
        For_Sg_Round_n.Saturation.Gas = (Data.Model.Grids(nGrid).Saturation.Gas + Data.Methodology.Derivations.Gas_Saturation_Epsilon)
        For_Sg_Round_n.Saturation.Water = Data.Model.Grids(nGrid).Saturation.Water
        For_Sg_Round_n.PVT.Temperature = Data.Model.Grids(nGrid).PVT.Temperature
        Initialize_Saturations(Data, For_Sg_Round_n)
        Initialize_Phase_System(For_Sg_Round_n, Data.Rocks(Data.Model.Grids(nGrid).RockType))
        If Data.System.Phase.Gas Then
            Initialize_SCAL(For_Sg_Round_n, Data.Rocks(Data.Model.Grids(nGrid).RockType))
        Else
            Copy_SCAL(Data.Model.Grids(nGrid), For_Sg_Round_n)
        End If
        Copy_Size(Data.Model.Grids(nGrid), For_Po_Round_n)
        Copy_Location(Data.Model.Grids(nGrid), For_Po_Round_n)
        For_Po_Round_n.Pressure.Oil = (Data.Model.Grids(nGrid).Pressure.Oil + Data.Methodology.Derivations.Oil_Pressure_Epsilon)
        For_Po_Round_n.Pressure.Water = (Data.Model.Grids(nGrid).Pressure.Water + Data.Methodology.Derivations.Oil_Pressure_Epsilon)
        For_Po_Round_n.Pressure.Gas = (Data.Model.Grids(nGrid).Pressure.Gas + Data.Methodology.Derivations.Oil_Pressure_Epsilon)

        For_Po_Round_n.Saturation.Water = Data.Model.Grids(nGrid).Saturation.Water
        For_Po_Round_n.Saturation.Oil = Data.Model.Grids(nGrid).Saturation.Oil
        For_Po_Round_n.PVT.Temperature = Data.Model.Grids(nGrid).PVT.Temperature
        Initialize_Phase_System(For_Po_Round_n, Data.Rocks(Data.Model.Grids(nGrid).RockType))
        Initialize_PVT(For_Po_Round_n, Data)
        ' ================================================================================


        'Round m Region ==================================================================
        Copy_Size(Data.Model.Grids(adjID), For_Sw_Round_m)
        Copy_Location(Data.Model.Grids(adjID), For_Sw_Round_m)
        Copy_Pressure(Data.Model.Grids(adjID), For_Sw_Round_m)
        For_Sw_Round_m.Saturation.Water = (Data.Model.Grids(adjID).Saturation.Water + Data.Methodology.Derivations.Water_Saturation_Epsilon)
        For_Sw_Round_m.Saturation.Gas = Data.Model.Grids(adjID).Saturation.Gas
        For_Sw_Round_m.PVT.Temperature = Data.Model.Grids(adjID).PVT.Temperature
        Initialize_Saturations(Data, For_Sw_Round_m)
        Initialize_Phase_System(For_Sw_Round_m, Data.Rocks(Data.Model.Grids(adjID).RockType))
        If Data.System.Phase.Water Then
            Initialize_SCAL(For_Sw_Round_m, Data.Rocks(Data.Model.Grids(adjID).RockType))
        Else
            Copy_SCAL(Data.Model.Grids(adjID), For_Sw_Round_m)
        End If


        Copy_Size(Data.Model.Grids(adjID), For_Sg_Round_m)
        Copy_Location(Data.Model.Grids(adjID), For_Sg_Round_m)
        Copy_Pressure(Data.Model.Grids(adjID), For_Sg_Round_m)
        For_Sg_Round_m.Saturation.Gas = (Data.Model.Grids(adjID).Saturation.Gas + Data.Methodology.Derivations.Gas_Saturation_Epsilon)
        For_Sg_Round_m.Saturation.Water = Data.Model.Grids(adjID).Saturation.Water
        For_Sg_Round_m.PVT.Temperature = Data.Model.Grids(adjID).PVT.Temperature
        Initialize_Saturations(Data, For_Sg_Round_m)
        Initialize_Phase_System(For_Sg_Round_m, Data.Rocks(Data.Model.Grids(adjID).RockType))
        If Data.System.Phase.Gas Then
            Initialize_SCAL(For_Sg_Round_m, Data.Rocks(Data.Model.Grids(adjID).RockType))
        Else
            Copy_SCAL(Data.Model.Grids(adjID), For_Sg_Round_m)
        End If


        Copy_Size(Data.Model.Grids(adjID), For_Po_Round_m)
        Copy_Location(Data.Model.Grids(adjID), For_Po_Round_m)
        For_Po_Round_m.Pressure.Oil = (Data.Model.Grids(adjID).Pressure.Oil + Data.Methodology.Derivations.Oil_Pressure_Epsilon)
        For_Po_Round_m.Pressure.Water = (Data.Model.Grids(adjID).Pressure.Water + Data.Methodology.Derivations.Oil_Pressure_Epsilon)
        For_Po_Round_m.Pressure.Gas = (Data.Model.Grids(adjID).Pressure.Gas + Data.Methodology.Derivations.Oil_Pressure_Epsilon)
        For_Po_Round_m.Saturation.Water = Data.Model.Grids(adjID).Saturation.Water
        For_Po_Round_m.Saturation.Oil = Data.Model.Grids(adjID).Saturation.Oil
        For_Po_Round_m.PVT.Temperature = Data.Model.Grids(adjID).PVT.Temperature
        Initialize_Phase_System(For_Po_Round_m, Data.Rocks(Data.Model.Grids(adjID).RockType))
        Initialize_PVT(For_Po_Round_m, Data)
        ' ================================================================================

    End Sub

    Sub Ko_Kw_Kg(ByVal Data As Simulation, ByVal nGrid As Grid, ByVal mGrid As Grid, ByVal AdjacentType As Adjacent_Type, ByRef Ko As Decimal, ByRef Kw As Decimal, ByRef Kg As Decimal, ByRef L As Decimal, ByVal Gamma_Water As Decimal, ByVal Gamma_Oil As Decimal, ByVal Gamma_Gas As Decimal)
        Dim K_Absolute_Average As Decimal
        Select Case AdjacentType
            Case Adjacent_Type.X
                L = (nGrid.Size.X_Length + mGrid.Size.X_Length) / 2
                K_Absolute_Average = Average_Permeability(Data.Rocks(nGrid.RockType).Permeability.Kx, Data.Rocks(mGrid.RockType).Permeability.Kx, nGrid.Size.X_Length / 2, mGrid.Size.X_Length / 2)
                If nGrid.Pressure.Oil > mGrid.Pressure.Oil Then
                    Ko = K_Absolute_Average * nGrid.SCAL.Kro.Kx
                Else
                    Ko = K_Absolute_Average * mGrid.SCAL.Kro.Kx
                End If
                If nGrid.Pressure.Water > mGrid.Pressure.Water Then
                    Kw = K_Absolute_Average * nGrid.SCAL.Krw.Kx
                Else
                    Kw = K_Absolute_Average * mGrid.SCAL.Krw.Kx
                End If
                If nGrid.Pressure.Gas > mGrid.Pressure.Gas Then
                    Kg = K_Absolute_Average * nGrid.SCAL.Krg.Kx
                Else
                    Kg = K_Absolute_Average * mGrid.SCAL.Krg.Kx
                End If
            Case Adjacent_Type.Y
                L = (nGrid.Size.Y_Length + mGrid.Size.Y_Length) / 2
                K_Absolute_Average = Average_Permeability(Data.Rocks(nGrid.RockType).Permeability.Ky, Data.Rocks(mGrid.RockType).Permeability.Ky, nGrid.Size.Y_Length / 2, mGrid.Size.Y_Length / 2)
                If nGrid.Pressure.Oil > mGrid.Pressure.Oil Then
                    Ko = K_Absolute_Average * nGrid.SCAL.Kro.Ky
                Else
                    Ko = K_Absolute_Average * mGrid.SCAL.Kro.Ky
                End If
                If nGrid.Pressure.Water > mGrid.Pressure.Water Then
                    Kw = K_Absolute_Average * nGrid.SCAL.Krw.Ky
                Else
                    Kw = K_Absolute_Average * mGrid.SCAL.Krw.Ky
                End If
                If nGrid.Pressure.Gas > mGrid.Pressure.Gas Then
                    Kg = K_Absolute_Average * nGrid.SCAL.Krg.Ky
                Else
                    Kg = K_Absolute_Average * mGrid.SCAL.Krg.Ky
                End If
            Case Adjacent_Type.Z
                L = (nGrid.Size.Z_Length + mGrid.Size.Z_Length) / 2
                K_Absolute_Average = Average_Permeability(Data.Rocks(nGrid.RockType).Permeability.Kz, Data.Rocks(mGrid.RockType).Permeability.Kz, nGrid.Size.Z_Length / 2, mGrid.Size.Z_Length / 2)
                If nGrid.Pressure.Oil + Gamma_Oil * nGrid.Location.Z > mGrid.Pressure.Oil + Gamma_Oil * mGrid.Location.Z Then
                    Ko = K_Absolute_Average * nGrid.SCAL.Kro.Kz
                Else
                    Ko = K_Absolute_Average * mGrid.SCAL.Kro.Kz
                End If
                If nGrid.Pressure.Water + Gamma_Water * nGrid.Location.Z > mGrid.Pressure.Water + Gamma_Water * mGrid.Location.Z Then
                    Kw = K_Absolute_Average * nGrid.SCAL.Krw.Kz
                Else
                    Kw = K_Absolute_Average * mGrid.SCAL.Krw.Kz
                End If
                If nGrid.Pressure.Gas + Gamma_Gas * nGrid.Location.Z > mGrid.Pressure.Gas + Gamma_Gas * mGrid.Location.Z Then
                    Kg = K_Absolute_Average * nGrid.SCAL.Krg.Kz
                Else
                    Kg = K_Absolute_Average * mGrid.SCAL.Krg.Kz
                End If
        End Select

    End Sub


    Sub Prepare_Adjacents(ByRef Data As Simulation)
        If Data.Options.Automatic_Grid_Adjacent_Recognition Then

            For i As Integer = 0 To SizeOf(Data.Model.Grids) - 1
                Grid_Adjacent_Recognition(Data, i)
            Next
            Calculate_Adjacents_Area(Data)

            For i As Integer = 0 To SizeOf(Data.Model.Grids) - 1
                Dim n As Integer = 0
                Dim Adjacents(SizeOf(Data.Model.Grids(i).Adjacents) - 1) As Grid_Adjacent
                For j As Integer = 0 To SizeOf(Data.Model.Grids(i).Adjacents) - 1
                    Adjacents(j).Area = Data.Model.Grids(i).Adjacents(j).Area
                    Adjacents(j).Type = Data.Model.Grids(i).Adjacents(j).Type
                    Adjacents(j).Target_Grid_Index = Data.Model.Grids(i).Adjacents(j).Target_Grid_Index
                    If Data.Model.Grids(i).Adjacents(j).Area > 0 Then
                        n += 1
                    End If
                Next

                ReDim Data.Model.Grids(i).Adjacents(n - 1)
                n = 0

                For j As Integer = 0 To SizeOf(Adjacents) - 1
                    If Adjacents(j).Area > 0 Then
                        Data.Model.Grids(i).Adjacents(n).Area = Adjacents(j).Area
                        Data.Model.Grids(i).Adjacents(n).Type = Adjacents(j).Type
                        Data.Model.Grids(i).Adjacents(n).Target_Grid_Index = Adjacents(j).Target_Grid_Index
                        n += 1
                    End If
                Next
            Next
        Else
            Calculate_Adjacents_Area(Data)
        End If
    End Sub

    Sub Grid_Adjacent_Recognition(ByRef Data As Simulation, ByVal i As Integer)
        Data.Model.Grids(i).Adjacents = Nothing
        Dim n As Integer = 0
        For j As Integer = 0 To SizeOf(Data.Model.Grids) - 1
            If (Data.Model.Grids(j).Location.X + Data.Model.Grids(j).Size.X_Length / 2 = Data.Model.Grids(i).Location.X - Data.Model.Grids(i).Size.X_Length / 2) _
            Or (Data.Model.Grids(j).Location.X - Data.Model.Grids(j).Size.X_Length / 2 = Data.Model.Grids(i).Location.X + Data.Model.Grids(i).Size.X_Length / 2) Then
                n += 1
            End If

            If (Data.Model.Grids(j).Location.Y + Data.Model.Grids(j).Size.Y_Length / 2 = Data.Model.Grids(i).Location.Y - Data.Model.Grids(i).Size.Y_Length / 2) _
            Or (Data.Model.Grids(j).Location.Y - Data.Model.Grids(j).Size.Y_Length / 2 = Data.Model.Grids(i).Location.Y + Data.Model.Grids(i).Size.Y_Length / 2) Then
                n += 1
            End If

            If Math.Abs((Data.Model.Grids(j).Location.Z + Data.Model.Grids(j).Size.Z_Length / 2) - (Data.Model.Grids(i).Location.Z - Data.Model.Grids(i).Size.Z_Length / 2)) < 0.0000000001 _
            Or Math.Abs((Data.Model.Grids(j).Location.Z - Data.Model.Grids(j).Size.Z_Length / 2) - (Data.Model.Grids(i).Location.Z + Data.Model.Grids(i).Size.Z_Length / 2)) < 0.0000000001 Then
                n += 1
            End If
        Next

        ReDim Data.Model.Grids(i).Adjacents(n - 1)
        n = 0
        For j As Integer = 0 To SizeOf(Data.Model.Grids) - 1
            If (Data.Model.Grids(j).Location.X + Data.Model.Grids(j).Size.X_Length / 2 = Data.Model.Grids(i).Location.X - Data.Model.Grids(i).Size.X_Length / 2) _
            Or (Data.Model.Grids(j).Location.X - Data.Model.Grids(j).Size.X_Length / 2 = Data.Model.Grids(i).Location.X + Data.Model.Grids(i).Size.X_Length / 2) Then
                Data.Model.Grids(i).Adjacents(n).Target_Grid_Index = j
                Data.Model.Grids(i).Adjacents(n).Type = Adjacent_Type.X
                n += 1
            End If

            If (Data.Model.Grids(j).Location.Y + Data.Model.Grids(j).Size.Y_Length / 2 = Data.Model.Grids(i).Location.Y - Data.Model.Grids(i).Size.Y_Length / 2) _
            Or (Data.Model.Grids(j).Location.Y - Data.Model.Grids(j).Size.Y_Length / 2 = Data.Model.Grids(i).Location.Y + Data.Model.Grids(i).Size.Y_Length / 2) Then
                Data.Model.Grids(i).Adjacents(n).Target_Grid_Index = j
                Data.Model.Grids(i).Adjacents(n).Type = Adjacent_Type.Y
                n += 1
            End If

            If Math.Abs((Data.Model.Grids(j).Location.Z + Data.Model.Grids(j).Size.Z_Length / 2) - (Data.Model.Grids(i).Location.Z - Data.Model.Grids(i).Size.Z_Length / 2)) < 0.0000000001 _
            Or Math.Abs((Data.Model.Grids(j).Location.Z - Data.Model.Grids(j).Size.Z_Length / 2) - (Data.Model.Grids(i).Location.Z + Data.Model.Grids(i).Size.Z_Length / 2)) < 0.0000000001 Then
                Data.Model.Grids(i).Adjacents(n).Target_Grid_Index = j
                Data.Model.Grids(i).Adjacents(n).Type = Adjacent_Type.Z
                n += 1
            End If
        Next

    End Sub

    Sub IMPES_Zero(ByRef IM As IMPES_Matrix)
        For i As Integer = 0 To IM.LeftSide.GetLength(0) - 1
            For j As Integer = 0 To IM.LeftSide.GetLength(1) - 1
                IM.LeftSide(i, j) = 0
            Next
            IM.Po(0, i) = 0
            IM.Target(0, i) = 0
        Next
    End Sub

    Sub Fully_Implicit_Zero(ByRef IM As Fully_Implicit_Matrix)
        For i As Integer = 0 To IM.J.GetLength(0) - 1
            For j As Integer = 0 To IM.J.GetLength(1) - 1
                IM.J(i, j) = 0
            Next
            IM.Mines_R(0, i) = 0
            IM.Delta_X(0, i) = 0
        Next
    End Sub

    Sub StopSimulation()
        Paused = True
    End Sub

End Class
