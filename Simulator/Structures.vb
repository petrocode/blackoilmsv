Public Module Structures

    ' ---------------------- Grid Displacement System Enumeration --------------------------- 
    Enum Grid_Couple_Displacement_System
        Unknown
        Drainage
        Imbibition
    End Enum
    ' --------------------------------------------------------------------------------------- 


    ' --------------------------- System of Units Enumeration ------------------------------- 
    Enum System_of_Units
        Field
        Metric
    End Enum
    ' --------------------------------------------------------------------------------------- 


    ' --------------------------------- Phase Enumeration ----------------------------------- 
    Enum Phase
        Unknown
        Water
        Gas
        Oil
    End Enum
    ' --------------------------------------------------------------------------------------- 


    ' ------------------------ Grid Boundary Condition Enumeration -------------------------- 
    Enum Boundary_Condition_Type
        Normal
        Parameter
    End Enum
    ' --------------------------------------------------------------------------------------- 


    ' -------------------------- Simulation Run Type Enumeration ----------------------------
    Enum Simulation_RunType
        PureSimulation
        HistoryMatching_Corey
        HistoryMatching_Burdine
        HistoryMatching_SM
        HistoryMatching_LET

    End Enum
    ' --------------------------------------------------------------------------------------- 


    ' ------------------------- Grid Flow-Rate Switch Enumeration --------------------------- 
    Enum Grid_Flow_Rate_Switch_Type
        Interpolate
        StepByStep
    End Enum
    ' --------------------------------------------------------------------------------------- 


    ' -------------------------------- PVT Methods Enumeration ------------------------------
    Enum PVT_Methods
        Table
        Oil_FVF_1
        Gas_FVF_1
        Water_FVF_1
        Oil_Density_1
        Gas_Density_1
        Water_Density_1
        Oil_Viscosity_1
        Gas_Viscosity_1
        Water_Viscosity_1
        Oil_Viscosity_Static
        Gas_Viscosity_Static
        Water_Viscosity_Static
        Oil_Compressibility_Static
        Gas_Compressibility_Static
        Water_Compressibility_Static
        Gas_Critical_Pressure_1
        Gas_Critical_Temperature_1
        RS_1
        Correlation2
        Correlation3
    End Enum
    ' --------------------------------------------------------------------------------------- 


    ' ------------------------------------ Grid Phase System -------------------------------- 
    Enum Grid_Phase_System
        Oil_Only
        Gas_Only
        Water_Only
        Oil_Gas
        Oil_Water
        Gas_Water
        Oil_Gas_Water
    End Enum
    ' --------------------------------------------------------------------------------------- 


    ' ---------------------------------- Simulation Approach -------------------------------- 
    Enum Simulation_Approach
        Fully_Implicit
        IMPES
        SEQ
    End Enum
    ' --------------------------------------------------------------------------------------- 


    ' ------------------------ Transmisiibility Fp Averaging Method ------------------------- 
    Enum Transmissibility_Fp_Averaging
        Method1
        Method2
        Method3
    End Enum
    ' --------------------------------------------------------------------------------------- 


    ' ----------------------------- Simulation Schedule Location ---------------------------- 
    Enum Simulation_Schedule_Position
        Internal
        Surface
    End Enum
    ' --------------------------------------------------------------------------------------- 


    ' --------------------- Simulation Schedule Target Enumeration -------------------------- 
    Enum Simulation_Schedule_Target
        Grid
        Well
    End Enum
    ' --------------------------------------------------------------------------------------- 


    ' ------------------------------- Schedule Type Enumeration ----------------------------- 
    Enum Schedule_Type
        Equations
        Static_Oil_Pressure
        Static_Gas_Pressure
        Static_Water_Pressure
        Static_Oil_Flow
        Static_Gas_Flow
        Static_Water_Flow
        Static_Water_Saturation
        Static_Gas_Saturation
        Dynamic_Oil_Pressure
        Dynamic_Gas_Pressure
        Dynamic_Water_Pressure
        Dynamic_Oil_Flow
        Dynamic_Gas_Flow
        Dynamic_Water_Flow
        Dynamic_Water_Saturation
        Dynamic_Gas_Saturation
    End Enum
    ' --------------------------------------------------------------------------------------- 


    ' -------------------------Equation Solution Methods Enumeration-------------------------
    Enum Equation_Solution_Method
        LSOR
        Orhtomin
    End Enum
    ' --------------------------------------------------------------------------------------- 


    ' --------------------------------Adjacent Type Enumeration------------------------------ 
    Enum Adjacent_Type
        X
        Y
        Z
    End Enum
    '--------------------------------------------------------------------------------------- 


    ' ------------------------------Wettability Couple Structure---------------------------- 
    Structure Wettability_Couple
        Dim Water_Oil As Phase
        Dim Oil_Gas As Phase
        Dim Water_Gas As Phase
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' ------------------------- Three Phase Permeability Structure --------------------------
    Structure Three_Phase_Permeability
        Dim Sw() As Decimal
        Dim Values()(,) As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' ------------------------------Wettability Couple Structure---------------------------- 
    Structure RockType_SCAL_Tables
        Dim Water_Oil(,) As Decimal
        Dim Oil_Gas(,) As Decimal
        Dim Gas_Water(,) As Decimal
        Dim Gas_Water_Oil As Three_Phase_Permeability
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' ------------------------ Grid Displacement System Structure --------------------------- 
    Structure Grid_Displacement_System
        Dim Water_Oil As Grid_Couple_Displacement_System
        Dim Water_Gas As Grid_Couple_Displacement_System
        Dim Oil_Gas As Grid_Couple_Displacement_System
    End Structure
    ' --------------------------------------------------------------------------------------- 


    '' -------------------------------- LET Parameter Structure ------------------------------ 
    'Structure PC_Set
    '    Dim Sat As Double
    '    Dim Pc As Double
    'End Structure
    '' --------------------------------------------------------------------------------------- 


    '' -------------------------------- LET Parameter Structure ------------------------------ 
    'Structure PC_LET
    '    Dim PCs() As PC_Set
    '    Dim MinPc As Double
    '    Dim MaxPc As Double
    'End Structure
    '' --------------------------------------------------------------------------------------- 



    ' -------------------------------- LET Parameter Structure ------------------------------ 
    Structure LET_Parameter
        Dim MSE As Double

        Dim L_WO_W As Double
        Dim E_WO_W As Double
        Dim T_WO_W As Double

        Dim L_WO_O As Double
        Dim E_WO_O As Double
        Dim T_WO_O As Double

        Dim L_GO_G As Double
        Dim E_GO_G As Double
        Dim T_GO_G As Double

        Dim L_GO_O As Double
        Dim E_GO_O As Double
        Dim T_GO_O As Double

        Dim L_WG_W As Double
        Dim E_WG_W As Double
        Dim T_WG_W As Double

        Dim L_WG_G As Double
        Dim E_WG_G As Double
        Dim T_WG_G As Double

        Dim Sinj_Min As Double
        Dim Sinj_Max As Double
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' ---------------------------------In-Time Flow Structure-------------------------------- 
    Structure Fluid_Flow
        Dim Oil As Decimal
        Dim Gas As Decimal
        Dim Water As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' ---------------------------------Grid System Structure--------------------------------- 
    Structure Grid_System
        Dim Phase As Grid_Phase_System
        Dim Displacement As Grid_Displacement_System
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' --------------------------- PVT Phase Methods Enumeration -----------------------------
    Structure PVT_Phase_Method
        Dim Viscosity As PVT_Methods
        Dim FVF As PVT_Methods
        Dim Density As PVT_Methods
        Dim Compressibility As PVT_Methods
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' --------------------------- PVT Phases Methods Enumeration ----------------------------
    Structure PVT_Phases_Methods
        Dim Oil As PVT_Phase_Method
        Dim Gas As PVT_Phase_Method
        Dim Water As PVT_Phase_Method
        Dim Rs As PVT_Methods
        Dim Gas_Critical_Pressure As PVT_Methods
        Dim Gas_Critical_Temperature As PVT_Methods
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' --------------------------------Grid Location Structure-------------------------------- 
    Structure Grid_Location
        Dim X As Decimal
        Dim Y As Decimal
        Dim Z As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' -----------------------------------Grid Size Structure--------------------------------- 
    Structure Grid_Size
        Dim X_Length As Decimal
        Dim Y_Length As Decimal
        Dim Z_Length As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' --------------------------------Grid Saturation Structure------------------------------ 
    Structure Saturation_Profile
        Dim Oil As Decimal
        Dim Water As Decimal
        Dim Gas As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' --------------------------------Grid Saturation Structure------------------------------ 
    Structure Pressure_Profile
        Dim Oil As Decimal
        Dim Water As Decimal
        Dim Gas As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' ----------------------------- Simulation Schedule Structure----------------------------
    Structure Simulation_Schedule
        Dim ID As Integer
        Dim Target As Simulation_Schedule_Target
        Dim Type As Schedule_Type
        Dim Position As Simulation_Schedule_Position
        Dim Orientation As Adjacent_Type
        Dim Data(,) As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' --------------------- Simulation Derivation Parameters Structure ---------------------- 
    Structure Simulation_Derivation_Parameters
        Dim Water_Saturation_Epsilon As Decimal
        Dim Gas_Saturation_Epsilon As Decimal
        Dim Oil_Pressure_Epsilon As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' --------------------------- Simulation Methodology Structure -------------------------- 
    Structure Simulation_Methodology
        Dim Transmissibility_Fp_Averaging As Transmissibility_Fp_Averaging
        Dim Derivations As Simulation_Derivation_Parameters
        Dim Approach As Simulation_Approach
        Dim PVT As PVT_Phases_Methods
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' ----------------------------- Simulation Options Structure ---------------------------- 
    Structure Simulation_Options
        Dim Automatic_Grid_Adjacent_Recognition As Boolean
        Dim PostPressureDistributionAcceleration As Boolean
        Dim PPDA_New As Double
        Dim PPDA_i As Integer
        Dim Grid_Transparency As Byte
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' -----------------------------------Grid Side Structure--------------------------------- 
    Structure Side
        Dim xCoord, yCoord As Decimal
        Dim xLen, yLen As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' ------------------------------Permeability Data Structure------------------------------ 
    Structure Permeability
        Dim Kx As Decimal
        Dim Ky As Decimal
        Dim Kz As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' -------------------------Grid Adjacent Connection Structure---------------------------- 
    Structure Grid_Adjacent
        Dim Target_Grid_Index As Integer
        Dim Area As Decimal
        Dim Type As Adjacent_Type
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' -------------------------- Fully Implicit Matrix Structure ---------------------------- 
    Structure Fully_Implicit_Matrix
        Dim Iteration As Integer
        Dim J(,) As Decimal
        Dim Delta_X(,) As Decimal
        Dim Mines_R(,) As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' ------------------------------- IMPES Matrix Structure -------------------------------- 
    Structure IMPES_Matrix
        Dim Iteration As Integer
        Dim LeftSide(,) As Decimal
        Dim Po(,) As Decimal
        Dim Target(,) As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 

    ' ------------------------------- IMPES Matrix Structure -------------------------------- 
    Structure SEQ_Matrix
        Dim Iteration As Integer
        Dim LeftSide(,) As Decimal
        Dim Po(,) As Decimal
        Dim Target(,) As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' -------------------------------- Fluid Property Structure------------------------------ 
    Structure Fluid_Property
        Dim FVF As Decimal
        Dim Density As Decimal
        Dim Viscosity As Decimal
        Dim Compressibility As Decimal

        ReadOnly Property Gamma_Air(ByVal System As System_of_Units) As Decimal
            Get
                Select Case System
                    Case System_of_Units.Field
                        Return Density
                    Case System_of_Units.Metric
                        Return Density / 29
                End Select

            End Get
        End Property
        ReadOnly Property Gamma_Water(ByVal System As System_of_Units) As Decimal
            Get
                Select Case System
                    Case System_of_Units.Field
                        Return Density / 62.4
                    Case System_of_Units.Metric
                        Return (Density / Simulation_Constants.Phase_Density) / 62.4
                End Select
            End Get
        End Property

        ReadOnly Property Pressure_Gradient(ByVal System As System_of_Units) As Decimal
            Get
                Select Case System
                    Case System_of_Units.Field
                        Return Density * 0.00021584 * 32.174
                    Case System_of_Units.Metric
                        Return Density * 0.001 * 9.8066352
                End Select
            End Get
        End Property
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' -----------------------------  Grid Parameters Structure ------------------------------
    Structure Grid_Parameters
        Dim R As R_Parameters_Structure
        Dim R_nn As R_nn_Parameters_Structure
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' -------------------------- R nn Bracket Parameters Structure --------------------------
    Structure R_nn_Parameters_Structure
        Dim Ro_Po As Decimal
        Dim Ro_Sw As Decimal
        Dim Ro_Sg As Decimal

        Dim Rw_Po As Decimal
        Dim Rw_Sw As Decimal
        Dim Rw_Sg As Decimal

        Dim Rg_Po As Decimal
        Dim Rg_Sw As Decimal
        Dim Rg_Sg As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' -------------------------- R Bracket Parameters Structure ----------------------------- 
    Structure R_Parameters_Structure

        Dim average_Gamma_Oil As Decimal
        Dim average_Gamma_Water As Decimal
        Dim average_Gamma_Gas As Decimal


        Dim round_To_Po_m As Decimal
        Dim round_To_Sw_m As Decimal
        Dim round_To_Sg_m As Decimal
        Dim round_To_Rs_Po_m As Decimal

        Dim round_Tw_Po_m As Decimal
        Dim round_Tw_Sw_m As Decimal

        Dim round_Tg_Po_m As Decimal
        Dim round_Tg_Sg_m As Decimal

        Dim round_To_Po_n As Decimal
        Dim round_To_Sw_n As Decimal
        Dim round_To_Sg_n As Decimal
        Dim round_To_Rs_Po_n As Decimal

        Dim round_Tw_Po_n As Decimal
        Dim round_Tw_Sw_n As Decimal

        Dim round_Tg_Po_n As Decimal
        Dim round_Tg_Sw_n As Decimal
        Dim round_Tg_Sg_n As Decimal

        Dim Pcow_Prime_m As Decimal
        Dim Pcgo_Prime_m As Decimal

        Dim Pcow_Prime_n As Decimal
        Dim Pcgo_Prime_n As Decimal

        Dim To_nm As Decimal
        Dim Tg_nm As Decimal
        Dim Tw_nm As Decimal

    End Structure




    ' -------------------------------Grid PVT Data Structure -------------------------------- 
    Structure Grid_PVT_Data
        Dim Oil As Fluid_Property
        Dim Water As Fluid_Property
        Dim Gas As Fluid_Property
        Dim Rs As Decimal
        Dim Temperature As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' -------------------------------Grid SCAL Data Structure -------------------------------- 
    Structure Grid_Scal_Data
        Dim Kro As Permeability
        Dim Krw As Permeability
        Dim Krg As Permeability
        Dim pcow As Decimal
        Dim pcgo As Decimal
        Dim pcow_drg, pcow_imb As Decimal
        Dim pcgo_drg, pcgo_imb As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' --------------------------------- Well Path Structure --------------------------------- 
    Structure Well_Path
        Dim X1, Y1, Z1 As Decimal
        Dim X2, Y2, Z2 As Decimal
        Dim Diameter As Decimal
        Dim Skin As Decimal
        Dim Flow As Boolean
    End Structure

    ' --------------------------------------------------------------------------------------- 


    ' ----------------------------------- Well Structure ------------------------------------ 
    Structure Well
        Dim Name As String
        Dim DefaultDiameter, DefaultSkin As Decimal
        Dim Paths() As Well_Path
        Dim Schedules() As Integer
    End Structure
    ' ---------------------------------------------------------------------------------------


    ' ----------------------------------Grid Data Structure---------------------------------- 
    Structure Grid
        Dim ID As Integer

        Dim RockType As Integer
        Dim Adjacents() As Grid_Adjacent

        Dim Location As Grid_Location
        Dim Size As Grid_Size
        Dim Saturation As Saturation_Profile
        Dim Delta As Saturation_Profile
        Dim Pressure As Pressure_Profile
        Dim Flow As Fluid_Flow
        Dim Schedules() As Integer

        Dim PVT As Grid_PVT_Data
        Dim SCAL As Grid_Scal_Data

        Dim Volume As Decimal
        Dim Porosity As Decimal

        Dim System As Grid_System
        Dim Color As System.Windows.Media.Color
        Dim InsideText As String



    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' -------------------------------- RockType Data Structure------------------------------- 
    Structure Rock
        Dim ID As Integer
        Dim Initial_Porosity As Decimal
        Dim Cs As Decimal
        Dim Permeability As Permeability
        Dim SCAL As RockType_SCAL_Data
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' -------------------------------- Grid Depletion Structure------------------------------ 
    Structure Grid_Depletion
        Dim Target_Adjacent_Index As Integer
        Dim Phase As Phase
    End Structure
    ' --------------------------------------------------------------------------------------- 



    ' ----------------------------------- ANN Feed Structure---------------------------------
    Structure ANN_Feed
        Dim Input() As Decimal
        Dim Target() As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 



    ' ------------------------------ RockType SCAL Data Structure---------------------------- 
    Structure RockType_SCAL_Data
        Dim Wettability As Wettability_Couple
        Dim Tables As RockType_SCAL_Tables
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' ------------------------- Simulation History Data Structure---------------------------- 
    Structure Simulation_History
        Private _Simulations() As Simulation
        Private _length As Integer
        Public Property Length() As Integer
            Get
                Return _length
            End Get
            Set(ByVal value As Integer)

                If value > SizeOf(_Simulations) Then
                    ReDim Preserve _Simulations(value - 1)
                    For i As Integer = SizeOf(_Simulations) - 1 To value - 1
                        _Simulations(i) = New Simulation
                    Next
                End If
                _length = value
            End Set
        End Property
        Public Property Simulations(ByVal i As Integer) As Simulation
            Get
                Return _Simulations(i)
            End Get
            Set(ByVal value As Simulation)
                _Simulations(i) = value
            End Set
        End Property
        Public Sub Cast(ByVal Data As Simulation)
            For i As Integer = 0 To SizeOf(_Simulations) - 1
                _Simulations(i) = Simulation_Copy(Data)
                _Simulations(i).Current_Step = i + 1
            Next
        End Sub


        Sub Clear()
            _length = 0
            _Simulations = Nothing
            Current = 0
        End Sub
        Public ReadOnly Property First() As Simulation
            Get
                Return _Simulations(Current - 1)
            End Get
        End Property
        Dim Current As Integer

        'Public ReadOnly Property Count() As Integer
        '    Get
        '        Return Current
        '    End Get
        'End Property
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' ---------------------- Iteration MSE Structure -------------------------
    Structure Iteration_MSE
        Dim Maximum As Integer
        Dim Sw As Decimal
        Dim Sg As Decimal
        Dim Po As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' ------------------------------- Simulation MSE Structure ------------------------------ 
    Structure Simulation_Exactness
        Dim IterationMSE As Iteration_MSE
        Dim OrthominMaxIteration As Integer
        Dim Equation As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' ------------------------------- Simulation MSE Structure ------------------------------ 
    Structure Simulation_Graphics
        Dim X_Angle As Decimal
        Dim Y_Angle As Decimal
        Dim Depth As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' ----------------------------- Simulation System Structure ----------------------------- 
    Structure Simulation_System
        Dim Units As System_of_Units
        Dim Phase As Simulation_Phase_Sysyem
    End Structure

    ' --------------------------------------------------------------------------------------- 


    ' --------------------------- Simulation Constants Structure ---------------------------- 
    Structure Simulation_Constants
        Dim Beta, Alpha As Decimal
        Const Length As Decimal = 0.3048
        Const Area As Decimal = 0.09290304
        Const Permeability As Decimal = 0.9869233
        Const Viscosity As Decimal = 0.001
        Const Gas_FVF As Decimal = 5.5519314
        Const Liquid_FVF As Decimal = 1
        Const Rs As Decimal = 0.18101175
        Const Pressure As Decimal = 6.894757
        Const Pressure_Gradient As Decimal = 22.62059
        Const Phase_Gravity As Decimal = 22.62059
        Const Gas_Flow_Rate As Decimal = 0.0286364
        Const Liquid_Flow_Rate As Decimal = 0.1589873
        Const Volumetric_Velocity As Decimal = 1.7103717
        Const Vb As Decimal = 0.02831685
        Const Phase_Density As Decimal = 16.01846
        Const gc As Decimal = 0.3048
        Const Compressibility As Decimal = 0.1450377
        Const Temperature As Decimal = 0.555555555555
        Const BgConversion As Decimal = 5.5519314
        Const Volume_Conversion_Factor As Decimal = 5.614583
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' ------------------------- Simulation Phase System Structure --------------------------- 
    Structure Simulation_Phase_Sysyem
        Dim Oil As Boolean
        Dim Water As Boolean
        Dim Gas As Boolean
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' -------------------------------- Simulation Structure --------------------------------- 
    Structure Simulation_Model
        Dim Grids() As Grid
        Dim Wells() As Well
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' -------------------------------- Simulation Structure --------------------------------- 
    Structure Simulation_PVT_Data
        Dim Table(,) As Decimal
        Dim Oil_At_Standard As Fluid_Property
        Dim Water_At_Standard As Fluid_Property
        Dim Gas_At_Standard As Fluid_Property
        Dim Condesate As Fluid_Property

        Dim Gas_Molecular_Weight As Decimal
        Dim Gas_Critical_Pressure As Decimal
        Dim Gas_Critical_Temperature As Decimal
        Dim Gas_Bubble_Pressure As Decimal
        Dim StandingC As Decimal
    End Structure
    ' --------------------------------------------------------------------------------------- 


    ' -------------------------------- Simulation Structure --------------------------------- 
    Structure Simulation
        Dim RunType As Simulation_RunType
        Dim Schedules() As Simulation_Schedule
        Dim Model As Simulation_Model
        Dim Rocks() As Rock
        Dim PVT As Simulation_PVT_Data
        Dim Constants As Simulation_Constants
        Dim System As Simulation_System
        Dim Current_Step As Integer
        Dim Time_Steps() As Decimal
        Dim Finish_Time As Decimal
        Dim Total_Time As Decimal
        Dim Final_Step As Integer
        Dim Current_Time As Decimal
        Dim Exactness As Simulation_Exactness
        Dim C_Bar_n(,) As Decimal
        Dim Methodology As Simulation_Methodology
        Dim Options As Simulation_Options
        Dim Graphics As Simulation_Graphics
        Dim CurrentIterationCount As Integer
        Dim CurrentIteration As Integer


        Function Time_Step() As Decimal
            'If CurrentIteration <= 1 Then Return StarterTimeStep
            If Current_Step = 0 Then Return Time_Steps(0)

            If Current_Step <= SizeOf(Time_Steps) Then
                Return Time_Steps(Current_Step - 1)
            Else
                Return Time_Steps(SizeOf(Time_Steps) - 1)
            End If
        End Function
    End Structure
    ' --------------------------------------------------------------------------------------- 

End Module
