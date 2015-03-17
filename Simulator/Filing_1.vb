Module Filing
    Sub SaveFile(ByVal Name As String, ByVal Data As Simulation)
        '>> Instantiation
        Dim XmlDoc As New Xml.XmlDocument
        XmlDoc.LoadXml("<Data></Data>")

        '>> Definitions ----------------------------------------------------------------------

        Dim myGlobals As Xml.XmlElement = XmlDoc.CreateElement("Global")
        Dim myGlobals_System As Xml.XmlElement = XmlDoc.CreateElement("System")
        Dim myGlobals_System_Units As Xml.XmlElement = XmlDoc.CreateElement("Units")
        Dim myGlobals_System_Phases As Xml.XmlElement = XmlDoc.CreateElement("Phases")
        Dim myGlobals_Exactness As Xml.XmlElement = XmlDoc.CreateElement("MSE")
        Dim myGlobals_Methodology As Xml.XmlElement = XmlDoc.CreateElement("Methodology")
        Dim myGlobals_Exactness_Iteration As Xml.XmlElement = XmlDoc.CreateElement("Iteration")
        Dim myGlobals_TimeSteps As Xml.XmlElement = XmlDoc.CreateElement("TimeSteps")
        Dim myGlobals_Constants As Xml.XmlElement = XmlDoc.CreateElement("Constants")

        Dim myModel As Xml.XmlElement = XmlDoc.CreateElement("Model")
        Dim myModel_Grids As Xml.XmlElement = XmlDoc.CreateElement("Grids")
        Dim myModel_Wells As Xml.XmlElement = XmlDoc.CreateElement("Wells")

        Dim myPVT As Xml.XmlElement = XmlDoc.CreateElement("PVT")
        Dim myPVT_Data As Xml.XmlElement = XmlDoc.CreateElement("Data")
        Dim myPVT_Options As Xml.XmlElement = XmlDoc.CreateElement("Options")

        Dim myRockTypes As Xml.XmlElement = XmlDoc.CreateElement("RockTypes")
        Dim myRockTypes_Absolutes As Xml.XmlElement
        Dim myRockTypes_Relatives As Xml.XmlElement
        Dim myRockTypes_Relatives_KROW As Xml.XmlElement
        Dim myRockTypes_Relatives_KRGO As Xml.XmlElement
        Dim myRockTypes_Relatives_KRWG As Xml.XmlElement
        Dim myRockTypes_Relatives_KROWG As Xml.XmlElement

        Dim mySchedules As Xml.XmlElement = XmlDoc.CreateElement("Schedules")
        Dim mySchedules_Grids As Xml.XmlElement = XmlDoc.CreateElement("Grids")
        Dim mySchedules_Wells As Xml.XmlElement = XmlDoc.CreateElement("Wells")


        '>> Settings -------------------------------------------------------------------------

        myGlobals_System_Units.SetAttribute("System", Data.System.Units.ToString)
        myGlobals_System_Phases.SetAttribute("Water", Data.System.Phase.Water)
        myGlobals_System_Phases.SetAttribute("Oil", Data.System.Phase.Oil)
        myGlobals_System_Phases.SetAttribute("Gas", Data.System.Phase.Gas)
        myGlobals_Exactness.SetAttribute("Equation", Data.Exactness.Equation)
        myGlobals_Exactness_Iteration.SetAttribute("Maximum", Data.Exactness.IterationMSE.Maximum)
        myGlobals_Exactness_Iteration.SetAttribute("OrthominIterations", Data.Exactness.OrthominMaxIteration)
        myGlobals_Exactness_Iteration.SetAttribute("Sw", Data.Exactness.IterationMSE.Sw)
        myGlobals_Exactness_Iteration.SetAttribute("Sg", Data.Exactness.IterationMSE.Sg)
        myGlobals_Exactness_Iteration.SetAttribute("Po", Data.Exactness.IterationMSE.Po)
        myGlobals_Methodology.SetAttribute("Fp_Method", Data.Methodology.Transmissibility_Fp_Averaging)
        myGlobals_Methodology.SetAttribute("Derivative_Sw", Data.Methodology.Derivations.Water_Saturation_Epsilon)
        myGlobals_Methodology.SetAttribute("Derivative_Sg", Data.Methodology.Derivations.Gas_Saturation_Epsilon)
        myGlobals_Methodology.SetAttribute("Derivative_Po", Data.Methodology.Derivations.Oil_Pressure_Epsilon)

        Dim TimeSteps As String
        For i As Integer = 0 To SizeOf(Data.Time_Steps) - 1
            TimeSteps += Data.Time_Steps(i).ToString
            If i < SizeOf(Data.Time_Steps) - 1 Then TimeSteps += ","
        Next

        myGlobals_TimeSteps.SetAttribute("TimeSteps", TimeSteps)
        myGlobals_TimeSteps.SetAttribute("FinalStep", Data.Final_Step)
        myGlobals_TimeSteps.SetAttribute("FinishTime", Data.Finish_Time)

        For i As Integer = 0 To SizeOf(Data.Model.Grids) - 1
            Dim myGrid As Xml.XmlElement = XmlDoc.CreateElement("Grid")
            myGrid.SetAttribute("X", Data.Model.Grids(i).Location.X)
            myGrid.SetAttribute("Y", Data.Model.Grids(i).Location.Y)
            myGrid.SetAttribute("Z", Data.Model.Grids(i).Location.Z)

            myGrid.SetAttribute("Lx", Data.Model.Grids(i).Size.X_Length)
            myGrid.SetAttribute("Ly", Data.Model.Grids(i).Size.Y_Length)
            myGrid.SetAttribute("Lz", Data.Model.Grids(i).Size.Z_Length)

            myGrid.SetAttribute("Sw", Data.Model.Grids(i).Saturation.Water)
            myGrid.SetAttribute("Sg", Data.Model.Grids(i).Saturation.Gas)
            myGrid.SetAttribute("Po", Data.Model.Grids(i).Pressure.Oil)

            Dim strSchedules As String = ""
            For n As Integer = 0 To SizeOf(Data.Model.Grids(i).Schedules) - 1
                strSchedules += Data.Model.Grids(i).Schedules(n).ToString
                If n < SizeOf(Data.Model.Grids(i).Schedules) - 1 Then strSchedules += ","
            Next

            myGrid.SetAttribute("RockType", Data.Model.Grids(i).RockType)
            myGrid.SetAttribute("Schedules", strSchedules)
            myModel_Grids.AppendChild(myGrid)
        Next

        For i As Integer = 0 To SizeOf(Data.Model.Wells) - 1
            Dim myWell As Xml.XmlElement = XmlDoc.CreateElement("Well")
            myWell.SetAttribute("Name", Data.Model.Wells(i).Name)
            myWell.SetAttribute("DefaultDiameter", Data.Model.Wells(i).DefaultDiameter.ToString)
            myWell.SetAttribute("DefaultSkin", Data.Model.Wells(i).DefaultSkin.ToString)
            For j As Integer = 0 To SizeOf(Data.Model.Wells(i).Paths) - 1
                Dim Path As Xml.XmlElement = XmlDoc.CreateElement("Path")
                Path.SetAttribute("Start", Data.Model.Wells(i).Paths(j).X1.ToString + "," + Data.Model.Wells(i).Paths(j).Y1.ToString + "," + Data.Model.Wells(i).Paths(j).Z1.ToString)
                Path.SetAttribute("Finish", Data.Model.Wells(i).Paths(j).X2.ToString + "," + Data.Model.Wells(i).Paths(j).Y2.ToString + "," + Data.Model.Wells(i).Paths(j).Z2.ToString)
                Path.SetAttribute("Flow", Data.Model.Wells(i).Paths(j).Flow)
                Path.SetAttribute("Diameter", Data.Model.Wells(i).Paths(j).Diameter.ToString)
                Path.SetAttribute("Skin", Data.Model.Wells(i).Paths(j).Skin.Tostring)
                myWell.AppendChild(Path)
            Next


            Dim strSchedules As String = ""
            For n As Integer = 0 To SizeOf(Data.Model.Wells(i).Schedules) - 1
                strSchedules += Data.Model.Wells(i).Schedules(n).ToString
                If n < SizeOf(Data.Model.Wells(i).Schedules) - 1 Then strSchedules += ","
            Next
            myWell.SetAttribute("Schedules", strSchedules)
            myModel_Wells.AppendChild(myWell)

        Next

        If SizeOf(Data.PVT.Table) > 0 Then
            For i As Integer = 0 To Data.PVT.Table.GetLength(1) - 1
                Dim myRow As Xml.XmlElement = XmlDoc.CreateElement("Row")
                myRow.SetAttribute("Pressure", Data.PVT.Table(0, i))
                myRow.SetAttribute("Rs", Data.PVT.Table(1, i))
                myRow.SetAttribute("Bo", Data.PVT.Table(2, i))
                myRow.SetAttribute("Bg", Data.PVT.Table(3, i))
                myRow.SetAttribute("Bw", Data.PVT.Table(4, i))
                myRow.SetAttribute("Oil_Density", Data.PVT.Table(5, i))
                myRow.SetAttribute("Gas_Density", Data.PVT.Table(6, i))
                myRow.SetAttribute("Water_Density", Data.PVT.Table(7, i))
                myRow.SetAttribute("Oil_Viscosity", Data.PVT.Table(8, i))
                myRow.SetAttribute("Gas_Viscosity", Data.PVT.Table(9, i))
                myRow.SetAttribute("Water_Viscosity", Data.PVT.Table(10, i))
                myRow.SetAttribute("Oil_Compressibility", Data.PVT.Table(11, i))
                myRow.SetAttribute("Gas_Compressibility", Data.PVT.Table(12, i))
                myRow.SetAttribute("Water_Compressibility", Data.PVT.Table(13, i))
                myPVT_Data.AppendChild(myRow)
            Next
        End If


        For i As Integer = 0 To SizeOf(Data.Rocks) - 1
            Dim myRT As Xml.XmlElement = XmlDoc.CreateElement("Rock")
            myRT.SetAttribute("ID", Data.Rocks(i).ID)
            myRockTypes_Absolutes = XmlDoc.CreateElement("Absoultes")
            myRockTypes_Relatives = XmlDoc.CreateElement("Relatives")
            myRockTypes_Relatives_KROW = XmlDoc.CreateElement("KROW")
            myRockTypes_Relatives_KRGO = XmlDoc.CreateElement("KRGO")
            myRockTypes_Relatives_KRWG = XmlDoc.CreateElement("KRWG")
            myRockTypes_Relatives_KROWG = XmlDoc.CreateElement("KROWG")

            myRockTypes_Absolutes.SetAttribute("Initial_Porosity", Data.Rocks(i).Initial_Porosity)
            myRockTypes_Absolutes.SetAttribute("Cs", Data.Rocks(i).Cs)
            myRockTypes_Absolutes.SetAttribute("Kx", Data.Rocks(i).Permeability.Kx)
            myRockTypes_Absolutes.SetAttribute("Ky", Data.Rocks(i).Permeability.Ky)
            myRockTypes_Absolutes.SetAttribute("Kz", Data.Rocks(i).Permeability.Kz)

            If SizeOf(Data.Rocks(i).SCAL.Tables.Water_Oil) > 0 Then
                For j As Integer = 0 To Data.Rocks(i).SCAL.Tables.Water_Oil.GetLength(1) - 1
                    Dim myKROWset As Xml.XmlElement = XmlDoc.CreateElement("Set")
                    myKROWset.SetAttribute("Sw", Data.Rocks(i).SCAL.Tables.Water_Oil(0, j))
                    myKROWset.SetAttribute("Kro", Data.Rocks(i).SCAL.Tables.Water_Oil(1, j))
                    myKROWset.SetAttribute("Krw", Data.Rocks(i).SCAL.Tables.Water_Oil(2, j))
                    myKROWset.SetAttribute("Pc_Drainage", Data.Rocks(i).SCAL.Tables.Water_Oil(3, j))
                    myKROWset.SetAttribute("Pc_Imbibition", Data.Rocks(i).SCAL.Tables.Water_Oil(4, j))
                    myRockTypes_Relatives_KROW.AppendChild(myKROWset)
                Next
            End If

            If SizeOf(Data.Rocks(i).SCAL.Tables.Oil_Gas) > 0 Then
                For j As Integer = 0 To Data.Rocks(i).SCAL.Tables.Oil_Gas.GetLength(1) - 1
                    Dim myKRGOset As Xml.XmlElement = XmlDoc.CreateElement("Set")
                    myKRGOset.SetAttribute("Sg", Data.Rocks(i).SCAL.Tables.Oil_Gas(0, j))
                    myKRGOset.SetAttribute("Kro", Data.Rocks(i).SCAL.Tables.Oil_Gas(1, j))
                    myKRGOset.SetAttribute("Krg", Data.Rocks(i).SCAL.Tables.Oil_Gas(2, j))
                    myKRGOset.SetAttribute("Pc_Drainage", Data.Rocks(i).SCAL.Tables.Oil_Gas(3, j))
                    myKRGOset.SetAttribute("Pc_Imbibition", Data.Rocks(i).SCAL.Tables.Oil_Gas(4, j))
                    myRockTypes_Relatives_KRGO.AppendChild(myKRGOset)
                Next
            End If

            If SizeOf(Data.Rocks(i).SCAL.Tables.Gas_Water) > 0 Then
                For j As Integer = 0 To Data.Rocks(i).SCAL.Tables.Gas_Water.GetLength(1) - 1
                    Dim myKRWGset As Xml.XmlElement = XmlDoc.CreateElement("Set")
                    myKRWGset.SetAttribute("Sw", Data.Rocks(i).SCAL.Tables.Gas_Water(0, j))
                    myKRWGset.SetAttribute("Krw", Data.Rocks(i).SCAL.Tables.Gas_Water(1, j))
                    myKRWGset.SetAttribute("Krg", Data.Rocks(i).SCAL.Tables.Gas_Water(2, j))
                    myKRWGset.SetAttribute("Pc_Drainage", Data.Rocks(i).SCAL.Tables.Gas_Water(3, j))
                    myKRWGset.SetAttribute("Pc_Imbibition", Data.Rocks(i).SCAL.Tables.Gas_Water(4, j))
                    myRockTypes_Relatives_KRWG.AppendChild(myKRWGset)
                Next
            End If

            If SizeOf(Data.Rocks(i).SCAL.Tables.Gas_Water_Oil.Sw) > 0 Then
                For j As Integer = 0 To SizeOf(Data.Rocks(i).SCAL.Tables.Gas_Water_Oil.Sw) - 1
                    Dim NewSw As Xml.XmlElement = XmlDoc.CreateElement("Sw")
                    NewSw.SetAttribute("Sw", Data.Rocks(i).SCAL.Tables.Gas_Water_Oil.Sw(j))
                    For k As Integer = 0 To Data.Rocks(i).SCAL.Tables.Gas_Water_Oil.Values(j).GetLength(1) - 1
                        Dim NewSet As Xml.XmlElement = XmlDoc.CreateElement("Set")
                        'NewSet.SetAttribute()
                    Next



                    Dim myKRWGset As Xml.XmlElement = XmlDoc.CreateElement("Set")
                    myKRWGset.SetAttribute("Sg", Data.Rocks(i).SCAL.Tables.Gas_Water(0, j))
                    myKRWGset.SetAttribute("Krw", Data.Rocks(i).SCAL.Tables.Gas_Water(1, j))
                    myKRWGset.SetAttribute("Krg", Data.Rocks(i).SCAL.Tables.Gas_Water(2, j))
                    myKRWGset.SetAttribute("Pc_Drainage", Data.Rocks(i).SCAL.Tables.Gas_Water(3, j))
                    myKRWGset.SetAttribute("Pc_Imbibition", Data.Rocks(i).SCAL.Tables.Gas_Water(4, j))
                    myRockTypes_Relatives_KRWG.AppendChild(myKRWGset)
                Next
            End If

            myRockTypes_Relatives.AppendChild(myRockTypes_Relatives_KROW)
            myRockTypes_Relatives.AppendChild(myRockTypes_Relatives_KRGO)
            myRockTypes_Relatives.AppendChild(myRockTypes_Relatives_KRWG)
            myRockTypes_Relatives.AppendChild(myRockTypes_Relatives_KROWG)

            myRT.AppendChild(myRockTypes_Absolutes)
            myRT.AppendChild(myRockTypes_Relatives)

            myRockTypes.AppendChild(myRT)
        Next

        For i As Integer = 0 To SizeOf(Data.Schedules) - 1
            Dim mySC As Xml.XmlElement = XmlDoc.CreateElement("Schedule")
            Dim strData As String = ""

            mySC.SetAttribute("ID", Data.Schedules(i).ID)
            mySC.SetAttribute("Target", Data.Schedules(i).Target)
            mySC.SetAttribute("Type", Data.Schedules(i).Type)
            mySC.SetAttribute("Position", Data.Schedules(i).Position)
            mySC.SetAttribute("Orientation", Data.Schedules(i).Orientation)

            If SizeOf(Data.Schedules(i).Data) > 0 Then
                For i1 As Integer = 0 To Data.Schedules(i).Data.GetLength(0) - 1
                    For i2 As Integer = 0 To Data.Schedules(i).Data.GetLength(1) - 1
                        strData += Data.Schedules(i).Data(i1, i2).ToString + ","
                    Next
                Next
                strData = strData.Substring(0, strData.Length - 1)
            End If

            mySC.SetAttribute("Data", strData)
            mySchedules.AppendChild(mySC)
        Next

        '>> Connections ---------------------------------------------------------------------

        myGlobals_System.AppendChild(myGlobals_System_Units)
        myGlobals_System.AppendChild(myGlobals_System_Phases)
        myGlobals.AppendChild(myGlobals_System)
        myGlobals_Exactness.AppendChild(myGlobals_Exactness_Iteration)
        myGlobals.AppendChild(myGlobals_Exactness)
        myGlobals.AppendChild(myGlobals_TimeSteps)
        myGlobals.AppendChild(myGlobals_Constants)
        myGlobals.AppendChild(myGlobals_Methodology)

        myModel.AppendChild(myModel_Grids)
        myModel.AppendChild(myModel_Wells)

        myPVT.AppendChild(myPVT_Data)
        myPVT.AppendChild(myPVT_Options)



        '>> Finialization
        XmlDoc.DocumentElement.AppendChild(myGlobals)
        XmlDoc.DocumentElement.AppendChild(myModel)
        XmlDoc.DocumentElement.AppendChild(myPVT)
        XmlDoc.DocumentElement.AppendChild(myRockTypes)
        XmlDoc.DocumentElement.AppendChild(mySchedules)

        XmlDoc.Save(Name)


    End Sub

    Sub OpenFile(ByVal strFile As String, ByRef Data As Simulation)

        Dim XMLdoc As New Xml.XmlDocument
        XMLdoc.Load(strFile)

        ' Globals 
        Data.System.Units = IIf(XMLdoc.ChildNodes(0).ChildNodes(0).ChildNodes(0).ChildNodes(0).Attributes("System").Value = "Field", 0, 1)
        Data.System.Phase.Water = XMLdoc.ChildNodes(0).ChildNodes(0).ChildNodes(0).ChildNodes(1).Attributes("Water").Value
        Data.System.Phase.Oil = XMLdoc.ChildNodes(0).ChildNodes(0).ChildNodes(0).ChildNodes(1).Attributes("Oil").Value
        Data.System.Phase.Gas = XMLdoc.ChildNodes(0).ChildNodes(0).ChildNodes(0).ChildNodes(1).Attributes("Gas").Value
        Try
            Data.Exactness.IterationMSE.Maximum = XMLdoc.ChildNodes(0).ChildNodes(0).ChildNodes(1).ChildNodes(0).Attributes("Maximum").Value
        Catch ex As Exception
            Data.Exactness.IterationMSE.Maximum = 100
        End Try
        Try
            Data.Exactness.OrthominMaxIteration = XMLdoc.ChildNodes(0).ChildNodes(0).ChildNodes(1).ChildNodes(0).Attributes("OrthominIterations").Value
        Catch ex As Exception
            Data.Exactness.OrthominMaxIteration = 20
        End Try

        Data.Exactness.Equation = XMLdoc.ChildNodes(0).ChildNodes(0).ChildNodes(1).Attributes("Equation").Value
        Data.Exactness.IterationMSE.Sw = XMLdoc.ChildNodes(0).ChildNodes(0).ChildNodes(1).ChildNodes(0).Attributes("Sw").Value
        Data.Exactness.IterationMSE.Sg = XMLdoc.ChildNodes(0).ChildNodes(0).ChildNodes(1).ChildNodes(0).Attributes("Sg").Value
        Data.Exactness.IterationMSE.Po = XMLdoc.ChildNodes(0).ChildNodes(0).ChildNodes(1).ChildNodes(0).Attributes("Po").Value

        Data.Final_Step = XMLdoc.ChildNodes(0).ChildNodes(0).ChildNodes(2).Attributes("FinalStep").Value
        Data.Finish_Time = XMLdoc.ChildNodes(0).ChildNodes(0).ChildNodes(2).Attributes("FinishTime").Value
        If Val(XMLdoc.ChildNodes(0).ChildNodes(0).ChildNodes(2).Attributes("TimeSteps").Value) > 0 Then
            Dim vals_time_step() As String = Split(XMLdoc.ChildNodes(0).ChildNodes(0).ChildNodes(2).Attributes("TimeSteps").Value, ",")
            ReDim Data.Time_Steps(SizeOf(vals_time_step) - 1)
            For i As Integer = 0 To SizeOf(vals_time_step) - 1
                Data.Time_Steps(i) = Val(vals_time_step(i))
            Next
        End If


        Data.Methodology.Transmissibility_Fp_Averaging = XMLdoc.ChildNodes(0).ChildNodes(0).ChildNodes(4).Attributes("Fp_Method").Value
        Data.Methodology.Derivations.Water_Saturation_Epsilon = XMLdoc.ChildNodes(0).ChildNodes(0).ChildNodes(4).Attributes("Derivative_Sw").Value
        Data.Methodology.Derivations.Gas_Saturation_Epsilon = XMLdoc.ChildNodes(0).ChildNodes(0).ChildNodes(4).Attributes("Derivative_Sg").Value
        Data.Methodology.Derivations.Oil_Pressure_Epsilon = XMLdoc.ChildNodes(0).ChildNodes(0).ChildNodes(4).Attributes("Derivative_Po").Value

        ' Model  
        ' Grids
        ReDim Data.Model.Grids(XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(0).ChildNodes.Count - 1)
        For i As Integer = 0 To SizeOf(Data.Model.Grids) - 1
            Data.Model.Grids(i).Location.X = XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(0).ChildNodes(i).Attributes("X").Value
            Data.Model.Grids(i).Location.Y = XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(0).ChildNodes(i).Attributes("Y").Value
            Data.Model.Grids(i).Location.Z = XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(0).ChildNodes(i).Attributes("Z").Value
            Data.Model.Grids(i).Size.X_Length = XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(0).ChildNodes(i).Attributes("Lx").Value
            Data.Model.Grids(i).Size.Y_Length = XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(0).ChildNodes(i).Attributes("Ly").Value
            Data.Model.Grids(i).Size.Z_Length = XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(0).ChildNodes(i).Attributes("Lz").Value
            Data.Model.Grids(i).Saturation.Water = XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(0).ChildNodes(i).Attributes("Sw").Value
            Data.Model.Grids(i).Saturation.Gas = XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(0).ChildNodes(i).Attributes("Sg").Value
            Data.Model.Grids(i).Pressure.Oil = XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(0).ChildNodes(i).Attributes("Po").Value
            Data.Model.Grids(i).RockType = XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(0).ChildNodes(i).Attributes("RockType").Value
            Dim strSchedule As String = XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(0).ChildNodes(i).Attributes("Schedules").Value
            If strSchedule <> "" Then
                Dim sc() As String = Split(strSchedule, ",")
                ReDim Preserve Data.Model.Grids(i).Schedules(SizeOf(sc) - 1)
                For j As Integer = 0 To SizeOf(sc) - 1
                    Data.Model.Grids(i).Schedules(j) = sc(j)
                Next
            End If
        Next

        ' Wells
        ReDim Data.Model.Wells(XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(1).ChildNodes.Count - 1)
        For i As Integer = 0 To SizeOf(Data.Model.Wells) - 1
            Data.Model.Wells(i).Name = XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(1).ChildNodes(i).Attributes("Name").Value
            Data.Model.Wells(i).DefaultDiameter = XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(1).ChildNodes(i).Attributes("DefaultDiameter").Value
            Data.Model.Wells(i).DefaultSkin = XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(1).ChildNodes(i).Attributes("DefaultSkin").Value

            Dim strSchedule As String = XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(1).ChildNodes(i).Attributes("Schedules").Value
            If strSchedule <> "" Then
                Dim sc() As String = Split(strSchedule, ",")
                ReDim Preserve Data.Model.Wells(i).Schedules(SizeOf(sc) - 1)
                For j As Integer = 0 To SizeOf(sc) - 1
                    Data.Model.Wells(i).Schedules(j) = sc(j)
                Next
            End If

            ReDim Data.Model.Wells(i).Paths(XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(1).ChildNodes(i).ChildNodes.Count - 1)
            For j As Integer = 0 To SizeOf(Data.Model.Wells(i).Paths) - 1
                Dim strStart() As String = Split(XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(1).ChildNodes(i).ChildNodes(j).Attributes("Start").Value, ",")
                Dim strFinish() As String = Split(XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(1).ChildNodes(i).ChildNodes(j).Attributes("Finish").Value, ",")
                Data.Model.Wells(i).Paths(j).X1 = strStart(0)
                Data.Model.Wells(i).Paths(j).Y1 = strStart(1)
                Data.Model.Wells(i).Paths(j).Z1 = strStart(2)

                Data.Model.Wells(i).Paths(j).X2 = strFinish(0)
                Data.Model.Wells(i).Paths(j).Y2 = strFinish(1)
                Data.Model.Wells(i).Paths(j).Z2 = strFinish(2)

                Data.Model.Wells(i).Paths(j).Diameter = XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(1).ChildNodes(i).ChildNodes(j).Attributes("Diameter").Value
                Data.Model.Wells(i).Paths(j).Flow = (XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(1).ChildNodes(i).ChildNodes(j).Attributes("Flow").Value = "True")
                Data.Model.Wells(i).Paths(j).Skin = XMLdoc.ChildNodes(0).ChildNodes(1).ChildNodes(1).ChildNodes(i).ChildNodes(j).Attributes("Skin").Value
            Next

        Next

        ' PVT
        ReDim Data.PVT.Table(13, XMLdoc.ChildNodes(0).ChildNodes(2).ChildNodes(0).ChildNodes.Count - 1)
        For i As Integer = 0 To XMLdoc.ChildNodes(0).ChildNodes(2).ChildNodes(0).ChildNodes.Count - 1
            Data.PVT.Table(0, i) = XMLdoc.ChildNodes(0).ChildNodes(2).ChildNodes(0).ChildNodes(i).Attributes("Pressure").Value
            Data.PVT.Table(1, i) = XMLdoc.ChildNodes(0).ChildNodes(2).ChildNodes(0).ChildNodes(i).Attributes("Rs").Value
            Data.PVT.Table(2, i) = XMLdoc.ChildNodes(0).ChildNodes(2).ChildNodes(0).ChildNodes(i).Attributes("Bo").Value
            Data.PVT.Table(3, i) = XMLdoc.ChildNodes(0).ChildNodes(2).ChildNodes(0).ChildNodes(i).Attributes("Bg").Value
            Data.PVT.Table(4, i) = XMLdoc.ChildNodes(0).ChildNodes(2).ChildNodes(0).ChildNodes(i).Attributes("Bw").Value
            Data.PVT.Table(5, i) = XMLdoc.ChildNodes(0).ChildNodes(2).ChildNodes(0).ChildNodes(i).Attributes("Oil_Density").Value
            Data.PVT.Table(6, i) = XMLdoc.ChildNodes(0).ChildNodes(2).ChildNodes(0).ChildNodes(i).Attributes("Gas_Density").Value
            Data.PVT.Table(7, i) = XMLdoc.ChildNodes(0).ChildNodes(2).ChildNodes(0).ChildNodes(i).Attributes("Water_Density").Value
            Data.PVT.Table(8, i) = XMLdoc.ChildNodes(0).ChildNodes(2).ChildNodes(0).ChildNodes(i).Attributes("Oil_Viscosity").Value
            Data.PVT.Table(9, i) = XMLdoc.ChildNodes(0).ChildNodes(2).ChildNodes(0).ChildNodes(i).Attributes("Gas_Viscosity").Value
            Data.PVT.Table(10, i) = XMLdoc.ChildNodes(0).ChildNodes(2).ChildNodes(0).ChildNodes(i).Attributes("Water_Viscosity").Value
            Data.PVT.Table(11, i) = XMLdoc.ChildNodes(0).ChildNodes(2).ChildNodes(0).ChildNodes(i).Attributes("Oil_Compressibility").Value
            Data.PVT.Table(12, i) = XMLdoc.ChildNodes(0).ChildNodes(2).ChildNodes(0).ChildNodes(i).Attributes("Gas_Compressibility").Value
            Data.PVT.Table(13, i) = XMLdoc.ChildNodes(0).ChildNodes(2).ChildNodes(0).ChildNodes(i).Attributes("Water_Compressibility").Value
        Next

        ' Rock Types
        ReDim Data.Rocks(XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes.Count - 1)
        For i As Integer = 0 To XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes.Count - 1
            Data.Rocks(i).ID = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).Attributes("ID").Value
            Data.Rocks(i).Initial_Porosity = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(0).Attributes("Initial_Porosity").Value
            Data.Rocks(i).Cs = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(0).Attributes("Cs").Value
            Data.Rocks(i).Permeability.Kx = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(0).Attributes("Kx").Value
            Data.Rocks(i).Permeability.Ky = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(0).Attributes("Ky").Value
            Data.Rocks(i).Permeability.Kz = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(0).Attributes("Kz").Value

            If XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(0).ChildNodes.Count > 0 Then
                ReDim Data.Rocks(i).SCAL.Tables.Water_Oil(4, XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(0).ChildNodes.Count - 1)
                For j As Integer = 0 To XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(0).ChildNodes.Count - 1
                    Data.Rocks(i).SCAL.Tables.Water_Oil(0, j) = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(0).ChildNodes(j).Attributes("Sw").Value
                    Data.Rocks(i).SCAL.Tables.Water_Oil(1, j) = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(0).ChildNodes(j).Attributes("Kro").Value
                    Data.Rocks(i).SCAL.Tables.Water_Oil(2, j) = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(0).ChildNodes(j).Attributes("Krw").Value
                    Data.Rocks(i).SCAL.Tables.Water_Oil(3, j) = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(0).ChildNodes(j).Attributes("Pc_Drainage").Value
                    Data.Rocks(i).SCAL.Tables.Water_Oil(4, j) = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(0).ChildNodes(j).Attributes("Pc_Imbibition").Value
                Next
            End If

            If XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(1).ChildNodes.Count > 0 Then
                ReDim Data.Rocks(i).SCAL.Tables.Oil_Gas(4, XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(1).ChildNodes.Count - 1)
                For j As Integer = 0 To XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(1).ChildNodes.Count - 1
                    Data.Rocks(i).SCAL.Tables.Oil_Gas(0, j) = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(1).ChildNodes(j).Attributes("Sg").Value
                    Data.Rocks(i).SCAL.Tables.Oil_Gas(1, j) = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(1).ChildNodes(j).Attributes("Kro").Value
                    Data.Rocks(i).SCAL.Tables.Oil_Gas(2, j) = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(1).ChildNodes(j).Attributes("Krg").Value
                    Data.Rocks(i).SCAL.Tables.Oil_Gas(3, j) = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(1).ChildNodes(j).Attributes("Pc_Drainage").Value
                    Data.Rocks(i).SCAL.Tables.Oil_Gas(4, j) = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(1).ChildNodes(j).Attributes("Pc_Imbibition").Value
                Next
            End If

            If XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(2).ChildNodes.Count > 0 Then
                ReDim Data.Rocks(i).SCAL.Tables.Gas_Water(4, XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(2).ChildNodes.Count - 1)
                For j As Integer = 0 To XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(2).ChildNodes.Count - 1
                    Data.Rocks(i).SCAL.Tables.Gas_Water(0, j) = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(2).ChildNodes(j).Attributes("Sw").Value
                    Data.Rocks(i).SCAL.Tables.Gas_Water(1, j) = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(2).ChildNodes(j).Attributes("Krw").Value
                    Data.Rocks(i).SCAL.Tables.Gas_Water(2, j) = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(2).ChildNodes(j).Attributes("Krg").Value
                    Data.Rocks(i).SCAL.Tables.Gas_Water(3, j) = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(2).ChildNodes(j).Attributes("Pc_Drainage").Value
                    Data.Rocks(i).SCAL.Tables.Gas_Water(4, j) = XMLdoc.ChildNodes(0).ChildNodes(3).ChildNodes(i).ChildNodes(1).ChildNodes(2).ChildNodes(j).Attributes("Pc_Imbibition").Value
                Next
            End If
        Next

        ' Schedules
        If XMLdoc.ChildNodes(0).ChildNodes(4).ChildNodes.Count > 0 Then
            ReDim Data.Schedules(XMLdoc.ChildNodes(0).ChildNodes(4).ChildNodes.Count - 1)
            For i As Integer = 0 To XMLdoc.ChildNodes(0).ChildNodes(4).ChildNodes.Count - 1
                Data.Schedules(i).ID = XMLdoc.ChildNodes(0).ChildNodes(4).ChildNodes(i).Attributes("ID").Value
                Data.Schedules(i).Target = XMLdoc.ChildNodes(0).ChildNodes(4).ChildNodes(i).Attributes("Target").Value
                Data.Schedules(i).Type = XMLdoc.ChildNodes(0).ChildNodes(4).ChildNodes(i).Attributes("Type").Value
                Data.Schedules(i).Position = XMLdoc.ChildNodes(0).ChildNodes(4).ChildNodes(i).Attributes("Position").Value
                Data.Schedules(i).Orientation = XMLdoc.ChildNodes(0).ChildNodes(4).ChildNodes(i).Attributes("Orientation").Value
                If XMLdoc.ChildNodes(0).ChildNodes(4).ChildNodes(i).Attributes("Data").Value <> "" Then
                    Dim strData() As String = Split(XMLdoc.ChildNodes(0).ChildNodes(4).ChildNodes(i).Attributes("Data").Value, ",")
                    If SizeOf(strData) = 1 Then
                        ReDim Preserve Data.Schedules(i).Data(0, 0)
                        Data.Schedules(i).Data(0, 0) = Val(strData(0))
                    Else
                        ReDim Preserve Data.Schedules(i).Data(1, SizeOf(strData) / 2 - 1)
                        For n As Integer = 0 To SizeOf(strData) / 2 - 1
                            Data.Schedules(i).Data(0, n) = Val(strData(n * 2))
                            Data.Schedules(i).Data(1, n) = Val(strData(n * 2 + 1))
                        Next
                    End If
                End If
            Next
        End If

        Data.Options.Grid_Transparency = 140
        Data.Options.Automatic_Grid_Adjacent_Recognition = True

    End Sub
End Module
