Public Module Preparation

    ' ======================  Setting System of Units Coefficients ====================== '
    Sub Set_System_of_Units(ByRef Data As Simulation)
        Select Case Data.System.Units
            Case System_of_Units.Field
                Data.Constants.Beta = 1.127
                Data.Constants.Alpha = 5.614583
            Case System_of_Units.Metric
                Data.Constants.Beta = 0.0000864
                Data.Constants.Alpha = 1
        End Select
    End Sub
    ' ----------------------------------------------------------------------------------- '



    ' =================  Initializing Grid Porosities from Rock Type ==================== '
    Sub Initialize_Grids_Porosities(ByRef Data As Simulation)
        For i As Integer = 0 To SizeOf(Data.Model.Grids) - 1
            Data.Model.Grids(i).Porosity = Data.Rocks(Data.Model.Grids(i).RockType).Initial_Porosity
        Next
    End Sub
    ' ----------------------------------------------------------------------------------- '


    ' ===============  Initializing Gas Critical Pressure and Temperature =============== '
    Sub Initialize_Gas_Critical_Pressure_Temperature(ByRef Data As Simulation)

        Dim Gamma_Gas As Double = Data.PVT.Gas_At_Standard.Gamma_Air(Data.System.Units)
        Select Case Data.Methodology.PVT.Gas_Critical_Pressure
            Case PVT_Methods.Gas_Critical_Pressure_1
                Data.PVT.Gas_Critical_Pressure = 756.8 - 131 * Gamma_Gas - 3.6 * (Gamma_Gas ^ 2)
        End Select
        Select Case Data.Methodology.PVT.Gas_Critical_Temperature
            Case PVT_Methods.Gas_Critical_Temperature_1
                Data.PVT.Gas_Critical_Temperature = 169.2 + 349.5 * Gamma_Gas - 74 * (Gamma_Gas ^ 2)
        End Select

    End Sub
    ' ----------------------------------------------------------------------------------- '


    ' ============================  Initializing Grid Volumes =========================== '
    Sub Calculate_Grids_Volume(ByRef Data As Simulation)
        For i As Integer = 0 To SizeOf(Data.Model.Grids) - 1
            With Data.Model.Grids(i).Size
                Data.Model.Grids(i).Volume = .X_Length * .Y_Length * .Z_Length
            End With
        Next
    End Sub
    ' ----------------------------------------------------------------------------------- '



    ' ==========================  Calculate Adjacent Grids Area ========================= '
    Sub Calculate_Adjacents_Area(ByRef Data As Simulation)
        For i As Integer = 0 To Data.Model.Grids.Length - 1
            For j As Integer = 0 To SizeOf(Data.Model.Grids(i).Adjacents) - 1
                Dim s1, s2 As Side
                Dim iGrid As Integer = Data.Model.Grids(i).Adjacents(j).Target_Grid_Index

                Select Case Data.Model.Grids(i).Adjacents(j).Type
                    Case Adjacent_Type.X
                        s1.xCoord = Data.Model.Grids(iGrid).Location.Y
                        s1.yCoord = Data.Model.Grids(iGrid).Location.Z
                        s1.xLen = Data.Model.Grids(iGrid).Size.Y_Length
                        s1.yLen = Data.Model.Grids(iGrid).Size.Z_Length

                        s2.xCoord = Data.Model.Grids(i).Location.Y
                        s2.yCoord = Data.Model.Grids(i).Location.Z
                        s2.xLen = Data.Model.Grids(i).Size.Y_Length
                        s2.yLen = Data.Model.Grids(i).Size.Z_Length

                    Case Adjacent_Type.Y
                        s1.xCoord = Data.Model.Grids(iGrid).Location.X
                        s1.yCoord = Data.Model.Grids(iGrid).Location.Z
                        s1.xLen = Data.Model.Grids(iGrid).Size.X_Length
                        s1.yLen = Data.Model.Grids(iGrid).Size.Z_Length

                        s2.xCoord = Data.Model.Grids(i).Location.X
                        s2.yCoord = Data.Model.Grids(i).Location.Z
                        s2.xLen = Data.Model.Grids(i).Size.X_Length
                        s2.yLen = Data.Model.Grids(i).Size.Z_Length

                    Case Adjacent_Type.Z
                        s1.xCoord = Data.Model.Grids(iGrid).Location.X
                        s1.yCoord = Data.Model.Grids(iGrid).Location.Y
                        s1.xLen = Data.Model.Grids(iGrid).Size.X_Length
                        s1.yLen = Data.Model.Grids(iGrid).Size.Y_Length

                        s2.xCoord = Data.Model.Grids(i).Location.X
                        s2.yCoord = Data.Model.Grids(i).Location.Y
                        s2.xLen = Data.Model.Grids(i).Size.X_Length
                        s2.yLen = Data.Model.Grids(i).Size.Y_Length

                End Select

                Data.Model.Grids(i).Adjacents(j).Area = Common_Area(s1, s2)

            Next
        Next
    End Sub
    ' ----------------------------------------------------------------------------------- '



    ' ===========================  Prepareing Implicit Matrix =========================== '
    Sub Prepare_Implicit_Matrix(ByRef Data As Simulation, ByRef IM As Fully_Implicit_Matrix)
        ReDim IM.J((Data.Model.Grids.Length * 3) - 1, (Data.Model.Grids.Length * 3) - 1)
        ReDim IM.Delta_X(0, (Data.Model.Grids.Length * 3) - 1)
        ReDim IM.Mines_R(0, (Data.Model.Grids.Length * 3) - 1)
        ReDim Data.C_Bar_n(0, (Data.Model.Grids.Length * 3) - 1)
        IM.Iteration = 0
    End Sub
    ' ----------------------------------------------------------------------------------- '


    ' =============================  Prepareing IMPES Matrix ===--======================== '
    Sub Prepare_IMPES_Matrix(ByRef Data As Simulation, ByRef IM As IMPES_Matrix)
        ReDim IM.LeftSide(Data.Model.Grids.Length - 1, Data.Model.Grids.Length - 1)
        ReDim IM.Po(0, Data.Model.Grids.Length - 1)
        ReDim IM.Target(0, Data.Model.Grids.Length - 1)
        IM.Iteration = 0
    End Sub
    ' ----------------------------------------------------------------------------------- '






End Module
