''' <summary>
'''                                                         BLACK-OIL MSV 1.0
'''  ______________________________Open Source Simulator for Black-Oil Model of Hydrocarbon Reservoirs _____________________________
'''                                                Copyright (C) 2014  Amir Emamjomeh                                              
'''                                                     amir.emamjomeh@gmail.com
''' 
''' This program is free software: you can redistribute it and/or  modify it  under the terms  of the  GNU  General  Public  License 
''' as published by the Free Software Foundation, either version 3 of the License, or any later version. This program is distributed 
''' in the hope that it will be useful,but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR 
''' A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
''' 
''' You should have received a copy of the GNU General Public Licensealong with this program. If not,see http://www.gnu.org/licenses
'''_________________________________________________________________________________________________________________________________
''' </summary>
Public Class Main

    Public WithEvents ucEdit As WPFX
    Public WithEvents ucSimulation As WPFX

    Friend WithEvents mySimulator As New Core

    Public History As New Simulation_History
    Public SimulationData As New Simulation
    Public ErrorMessage As String

    Private SimulationGraphLoaded As Boolean
    Private isBulkGrid As Boolean
    Private sum As Double

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If Button1.Tag = "Run" Then
            Button1.BackgroundImage = My.Resources.play_button3
            Button1.Tag = "Pause"
            TC.Maximum = SimulationData.Final_Step
            RunCheck(True)
            If History.Current = 0 Then
                mySimulator.Run(Simulation_Copy(SimulationData), History)
            Else
                mySimulator.Run(History.Simulations(History.Current), History)
            End If
            RunCheck(False)
        Else
            PauseSimulation()
        End If
    End Sub

    Sub RunCheck(ByVal IsRunning As Boolean)
        TV.Enabled = Not IsRunning
    End Sub

    Sub PauseSimulation()
        Button1.Tag = "Run"
        Button1.BackgroundImage = My.Resources.play_button1
        RunCheck(False)
        mySimulator.StopSimulation()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim dlgSave As New SaveFileDialog
        With dlgSave
            .ShowDialog()
            .Filter = "*.xml|*.xml"
            If .FileName <> "" Then
                SaveFile(.FileName + ".xml", SimulationData)
            End If
        End With
    End Sub
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        InitializeTreeView(TV)
        DockGroupBoxes()

        ColorBoundary1.ComboBox1.SelectedIndex = 3
        ColorBoundary1.ValueBox1.Text = 1000
        ColorBoundary1.ValueBox2.Text = 4000
        RoundButtons()
        PrepareDGVCurrent()
        ReDim SimulationData.Rocks(0)
        SimulationData.Rocks(0).ID = 1

        Wpfx1.LoadGraph(ElementHost2)
        Wpfx2.LoadGraph(ElementHost3)

    End Sub

    Sub PrepareDGVCurrent()
        With dgvCurrent
            .RowCount = 5
            .RowHeadersWidth = 80
            .Rows(0).HeaderCell.Value = "Step"
            .Rows(1).HeaderCell.Value = "Length"
            .Rows(2).HeaderCell.Value = "Total"
            .Rows(3).HeaderCell.Value = "Iterations"
        End With
    End Sub

    Sub RoundButtons()
        Dim GP As New Drawing2D.GraphicsPath
        GP.AddEllipse(New Rectangle(4, 4, 41, 41))
        Button1.Region = New Region(GP)
        Button9.Region = New Region(GP)
    End Sub

    Sub DockGroupBoxes()
        gbxSimulation.Dock = DockStyle.Fill
        gbxPVT.Dock = DockStyle.Fill
        gbxModel.Dock = DockStyle.Fill
        gbxSRockType.Dock = DockStyle.Fill
        gbxSchedule.Dock = DockStyle.Fill
    End Sub

    Private Sub TV_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TV.AfterSelect
        gbxSimulation.Visible = (TV.SelectedNode.Level = 1) And TV.SelectedNode.Text = "Simulation"
        gbxModel.Visible = (TV.SelectedNode.Level = 2 And TV.SelectedNode.Text = "Model")
        gbxPVT.Visible = (TV.SelectedNode.Level = 2 And TV.SelectedNode.Text = "PVT")
        gbxSRockType.Visible = (TV.SelectedNode.Level = 2 And TV.SelectedNode.Text = "Rock Types")
        gbxSchedule.Visible = (TV.SelectedNode.Level = 2 And TV.SelectedNode.Text = "Schedules")

        Dim adress() As String = Split(TV.SelectedNode.FullPath, "\")
        Label14.Text = adress(SizeOf(adress) - 1)
        Select Case TV.SelectedNode.Level
            Case 1
                If TV.SelectedNode.Text = "Simulation" Then
                    castSimulation()
                    If CheckData() Then
                        SplitContainer3.Visible = True
                    Else
                        SplitContainer3.Visible = False
                    End If
                End If
            Case 2
                Select Case TV.SelectedNode.Text
                    Case "Model"
                        castModel()
                    Case "PVT"
                        castPVT()
                    Case "Rock Types"
                        castRockTypes()
                    Case "Schedules"
                        castSchedules()
                End Select
        End Select
    End Sub

    Function CheckData() As Boolean
        Return True
    End Function
    Sub Log(ByVal str As String)

        LogBox.Text += str + vbCrLf
        LogBox.SelectionStart = LogBox.TextLength - 1
        LogBox.ScrollToCaret()
    End Sub

    Private Sub mySimulator_Simulation_End(ByVal History As Simulation_History) Handles mySimulator.Simulation_End
        Button1.BackgroundImage = My.Resources.play_button1
        Button1.Tag = "Run"
        RunCheck(False)
    End Sub
    Private Sub SolutionFailed(ByVal Data_n As Simulation, ByVal Data_v As Simulation, ByVal IM As Fully_Implicit_Matrix, ByVal MSE As Iteration_MSE) Handles mySimulator.SolutionFailed
        Log("Solution Failed at Iteration:" + IM.Iteration.ToString)

    End Sub
    Private Sub mySimulator_Time_Step(ByVal History As Simulation_History, ByVal Data As Structures.Simulation) Handles mySimulator.Time_Step

        Select Case Data.RunType
            Case Simulation_RunType.HistoryMatching_LET

            Case Simulation_RunType.PureSimulation
                Dim n1 As Integer = 17 - Data.Current_Step.ToString.Length
                Dim n2 As Integer = 17 - Data.Total_Time.ToString.Length
                Dim n3 As Integer = 17 - Data.CurrentIterationCount.ToString.Length
                Dim nsw1 As Integer = 17 - Data.Model.Grids(0).Saturation.Water.ToString
                Dim nsw2 As Integer = 17 - Data.Model.Grids(3).Saturation.Water.ToString
                Dim no1 As Integer = 17 - Data.Model.Grids(0).Pressure.Oil.ToString
                Dim nw2 As Integer = 17 - Data.Model.Grids(3).Pressure.Water.ToString

                Log("Time Step: " & AddStringOfLength(Data.Current_Step.ToString, n1, " ") _
                & " Time: " & AddStringOfLength(Data.Total_Time.ToString, n2, " ") & _
                " Iterations: " & AddStringOfLength(Data.CurrentIterationCount.ToString, n3, " "))


                CastCurrentTrackOnDGV()
                TC.Value = Data.Current_Step

                RefreshColors(ColorBoundary1.ComboBox1.Text, ColorBoundary1.ValueBox1.Text, ColorBoundary1.ValueBox2.Text, ColorBoundary1.Button1.BackColor, ColorBoundary1.Button2.BackColor, Data)


                Wpfx1.CastColors(Data.Model.Grids)
                DrawProgressIndicator()
                TC.Value = Data.Current_Step
                Application.DoEvents()

        End Select

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        StopSimulation()
    End Sub
    Sub StopSimulation()
        mySimulator.StopSimulation()
        Button1.BackgroundImage = My.Resources.play_button1
        TC.Value = 0
        LogBox.Clear()
        Button1.Tag = "Run"
        SimulationData.Current_Step = 0
        SimulationData.Current_Time = 0
        RunCheck(False)
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged

        If TabControl1.SelectedTab.Text = "Run" Then
            If SplitContainer3.Visible Then
                Application.DoEvents()
                If History.Current = 0 Then
                    Initialize_Grids_Porosities(SimulationData)
                    Calculate_Grids_Volume(SimulationData)
                    Set_System_of_Units(SimulationData)
                    Initialize(SimulationData)
                    Wpfx1.LoadGridList(SimulationData.Model.Grids)
                End If
            End If
        End If

    End Sub

    Sub ReloadGraph()
        Try
            If History.Current = 0 Then
                RefreshColors(ColorBoundary1.ComboBox1.Text, ColorBoundary1.ValueBox1.Text, ColorBoundary1.ValueBox2.Text, ColorBoundary1.Button1.BackColor, ColorBoundary1.Button2.BackColor, SimulationData)
                Wpfx1.CastColors(SimulationData.Model.Grids)
            Else
                If History.Current - TC.Value <= 0 Then TC.Value = History.Current - 1
                RefreshColors(ColorBoundary1.ComboBox1.Text, ColorBoundary1.ValueBox1.Text, ColorBoundary1.ValueBox2.Text, ColorBoundary1.Button1.BackColor, ColorBoundary1.Button2.BackColor, History.Simulations(TC.Value))
                Wpfx1.CastColors(History.Simulations(TC.Value).Model.Grids)
            End If
        Catch ex As Exception
        End Try
        Application.DoEvents()

    End Sub

    Private Sub tmrLoadSimulationGraph_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrLoadSimulationGraph.Tick
        tmrLoadSimulationGraph.Enabled = False
        SimulationGraphLoaded = True
        Application.DoEvents()
    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint
        If History.Current > 0 Then
            DrawProgressIndicator()
        Else
            e.Graphics.Clear(Color.White)
        End If
    End Sub


#Region "Tracker"

    Sub DrawProgressIndicator()
        Dim g As Graphics = Panel1.CreateGraphics
        Dim portion As Double = History.Current / History.Simulations(History.Current - 1).Final_Step
        g.FillRectangle(Brushes.Blue, New Rectangle(0, 0, portion * Panel1.Width, Panel1.Height))
        g.FillRectangle(Brushes.White, New Rectangle(portion * Panel1.Width, 0, Panel1.Width - portion * Panel1.Width, Panel1.Height))
    End Sub

    Private Sub TC_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TC.MouseDown
        If Not mySimulator.Paused Then
            PauseSimulation()
        End If
    End Sub

    Private Sub TC_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TC.Scroll
        CastCurrentTrackOnDGV()
        If mySimulator.Finished Or mySimulator.Paused Then
            ReloadGraph()
        End If

    End Sub
    Sub CastCurrentTrackOnDGV()
        If History.Current > 0 Then
            If History.Current - TC.Value - 1 >= 0 Then
                dgvCurrent.Item(0, 0).Value = History.Simulations(TC.Value).Current_Step
                dgvCurrent.Item(0, 1).Value = History.Simulations(TC.Value).Time_Step
                dgvCurrent.Item(0, 2).Value = History.Simulations(TC.Value).Total_Time
                dgvCurrent.Item(0, 3).Value = History.Simulations(TC.Value).CurrentIterationCount
            Else
                TC.Value = History.Current - 1
            End If
        End If
    End Sub
#End Region


#Region "ColorBoundary"

    Private Sub ColorBoundary1_MaxClicked() Handles ColorBoundary1.MaxClicked
        Dim Max, Min As Double
        Try
            Max = Grid_Parameter(SimulationData.Model.Grids(0), ColorBoundary1.ComboBox1.Text, SimulationData.System.Units)
            Min = Max
            If History.Current = 0 Then
                GetMaxMin(ColorBoundary1.ComboBox1.Text, Max, Min, SimulationData)
            Else
                For i As Integer = 0 To History.Current - 1
                    GetMaxMin(ColorBoundary1.ComboBox1.Text, Max, Min, History.Simulations(i))
                Next
            End If
            ColorBoundary1.ValueBox2.Text = Max
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ColorBoundary1_MinClicked() Handles ColorBoundary1.MinClicked
        Dim Max, Min As Double
        Try
            Max = Grid_Parameter(SimulationData.Model.Grids(0), ColorBoundary1.ComboBox1.Text, SimulationData.System.Units)
            Min = Max
            If History.Current = 0 Then
                GetMaxMin(ColorBoundary1.ComboBox1.Text, Max, Min, SimulationData)
            Else
                For i As Integer = 0 To History.Current - 1
                    GetMaxMin(ColorBoundary1.ComboBox1.Text, Max, Min, History.Simulations(i))
                Next
            End If
            ColorBoundary1.ValueBox1.Text = Min
        Catch ex As Exception

        End Try
    End Sub


    Private Sub ColorBoundary1_Refreshed() Handles ColorBoundary1.Refreshed
        Try
            ReloadGraph()
        Catch ex As Exception

        End Try
    End Sub

#End Region


#Region "Casts"

    Sub castModel()
        CastModelGrids()
        CastModelWells()
    End Sub

    Sub castRockTypes()
        Do While TabControl4.TabPages.Count <> 1
            TabControl4.TabPages.RemoveAt(TabControl4.TabPages.Count - 1)
        Loop

        For i As Integer = 0 To SizeOf(SimulationData.Rocks) - 1
            Dim TB As New TabPage
            TB.Text = "ID " + SimulationData.Rocks(i).ID.ToString
            Dim ctlRT As New ctlRockType
            AddHandler ctlRT.Removed, AddressOf RemoveRockType
            ctlRT.ID = SimulationData.Rocks(i).ID
            ctlRT.Dock = DockStyle.Fill
            castRockType(ctlRT, SimulationData.Rocks(i))
            TB.Controls.Add(ctlRT)
            TabControl4.TabPages.Add(TB)

        Next

        Dim TB1 As New TabPage
        TB1.Text = "New"
        TabControl4.TabPages.Add(TB1)

    End Sub

    Sub castSchedules()
        Do While TabControl5.TabPages.Count <> 1
            TabControl5.TabPages.RemoveAt(TabControl5.TabPages.Count - 1)
        Loop

        For i As Integer = 0 To SizeOf(SimulationData.Schedules) - 1
            Dim TB As New TabPage
            TB.Text = "ID " + SimulationData.Schedules(i).ID.ToString
            Dim ctlSC As New ctlSchedule
            AddHandler ctlSC.Removed, AddressOf RemoveSchedule
            ctlSC.ID = SimulationData.Schedules(i).ID
            ctlSC.Dock = DockStyle.Fill
            castSchedule(ctlSC, SimulationData.Schedules(i))
            TB.Controls.Add(ctlSC)
            TabControl5.TabPages.Add(TB)

        Next

        Dim TB1 As New TabPage
        TB1.Text = "New"
        TabControl5.TabPages.Add(TB1)

    End Sub

    Sub castSchedule(ByRef ctl As ctlSchedule, ByVal SC As Simulation_Schedule)
        With ctl
            .ComboBox1.SelectedIndex = IIf(SC.Target = Simulation_Schedule_Target.Grid, 0, 1)
            .ComboBox2.SelectedIndex = SC.Type
            If .ComboBox2.Text.ToUpper.StartsWith("STATIC") Then
                .ValueBox1.Text = SC.Data(0, 0)
                .ComboBox3.SelectedIndex = SC.Position
                If SC.Position = Simulation_Schedule_Position.Surface Then
                    .ComboBox4.SelectedIndex = SC.Orientation
                Else
                    .ComboBox4.SelectedIndex = -1
                End If

            ElseIf .ComboBox2.Text.ToUpper.StartsWith("DYNAMIC") Then
                .dgvGrid.RowCount = SC.Data.GetLength(1) + 1
                For i As Integer = 0 To SC.Data.GetLength(0) - 1
                    For j As Integer = 0 To SC.Data.GetLength(1) - 1
                        .dgvGrid.Item(i, j).Value = SC.Data(i, j)
                    Next
                Next
            End If

        End With
    End Sub

    Sub castRockType(ByRef ctl As ctlRockType, ByVal Rock As Rock)
        With ctl
            .ValueBox1.Text = Rock.Permeability.Kx
            .ValueBox2.Text = Rock.Permeability.Ky
            .ValueBox3.Text = Rock.Permeability.Kz
            .ValueBox10.Text = Rock.Initial_Porosity
            .ValueBox11.Text = Rock.Cs

            If SizeOf(Rock.SCAL.Tables.Water_Oil) > 0 Then .dgvOW.RowCount = Rock.SCAL.Tables.Water_Oil.GetLength(1) + 1
            If SizeOf(Rock.SCAL.Tables.Oil_Gas) > 0 Then .dgvGO.RowCount = Rock.SCAL.Tables.Oil_Gas.GetLength(1) + 1
            If SizeOf(Rock.SCAL.Tables.Gas_Water) > 0 Then .dgvOW.RowCount = Rock.SCAL.Tables.Gas_Water.GetLength(1) + 1

            For i As Integer = 0 To .dgvOW.RowCount - 2
                .dgvOW.Item(0, i).Value = Rock.SCAL.Tables.Water_Oil(0, i)
                .dgvOW.Item(1, i).Value = Rock.SCAL.Tables.Water_Oil(1, i)
                .dgvOW.Item(2, i).Value = Rock.SCAL.Tables.Water_Oil(2, i)
                .dgvOW.Item(3, i).Value = Rock.SCAL.Tables.Water_Oil(3, i)
                .dgvOW.Item(4, i).Value = Rock.SCAL.Tables.Water_Oil(4, i)
            Next

            For i As Integer = 0 To .dgvGO.RowCount - 2
                .dgvGO.Item(0, i).Value = Rock.SCAL.Tables.Oil_Gas(0, i)
                .dgvGO.Item(1, i).Value = Rock.SCAL.Tables.Oil_Gas(1, i)
                .dgvGO.Item(2, i).Value = Rock.SCAL.Tables.Oil_Gas(2, i)
                .dgvGO.Item(3, i).Value = Rock.SCAL.Tables.Oil_Gas(3, i)
                .dgvGO.Item(4, i).Value = Rock.SCAL.Tables.Oil_Gas(4, i)
            Next

            For i As Integer = 0 To .dgvWG.RowCount - 2
                .dgvWG.Item(0, i).Value = Rock.SCAL.Tables.Gas_Water(0, i)
                .dgvWG.Item(1, i).Value = Rock.SCAL.Tables.Gas_Water(1, i)
                .dgvWG.Item(2, i).Value = Rock.SCAL.Tables.Gas_Water(2, i)
                .dgvWG.Item(3, i).Value = Rock.SCAL.Tables.Gas_Water(3, i)
                .dgvWG.Item(4, i).Value = Rock.SCAL.Tables.Gas_Water(4, i)
            Next

            .RadioButton2.Checked = (Rock.SCAL.Wettability.Water_Oil = Phase.Water)
            .RadioButton1.Checked = (Rock.SCAL.Wettability.Water_Oil = Phase.Oil)
            .RadioButton7.Checked = (Rock.SCAL.Wettability.Water_Oil = Phase.Unknown)

            .RadioButton4.Checked = (Rock.SCAL.Wettability.Oil_Gas = Phase.Gas)
            .RadioButton3.Checked = (Rock.SCAL.Wettability.Oil_Gas = Phase.Oil)
            .RadioButton8.Checked = (Rock.SCAL.Wettability.Oil_Gas = Phase.Unknown)

            .RadioButton6.Checked = (Rock.SCAL.Wettability.Water_Gas = Phase.Gas)
            .RadioButton5.Checked = (Rock.SCAL.Wettability.Water_Gas = Phase.Water)
            .RadioButton9.Checked = (Rock.SCAL.Wettability.Water_Gas = Phase.Unknown)

        End With
    End Sub


    Sub CastModelGrids()
        With dgvGrids
            .RowCount = SizeOf(SimulationData.Model.Grids) + 1
            For i As Integer = 0 To SizeOf(SimulationData.Model.Grids) - 1
                .Item(0, i).Value = SimulationData.Model.Grids(i).Location.X
                .Item(1, i).Value = SimulationData.Model.Grids(i).Location.Y
                .Item(2, i).Value = SimulationData.Model.Grids(i).Location.Z
                .Item(3, i).Value = SimulationData.Model.Grids(i).Size.X_Length
                .Item(4, i).Value = SimulationData.Model.Grids(i).Size.Y_Length
                .Item(5, i).Value = SimulationData.Model.Grids(i).Size.Z_Length
                .Item(6, i).Value = SimulationData.Model.Grids(i).Saturation.Water
                .Item(7, i).Value = SimulationData.Model.Grids(i).Saturation.Gas
                .Item(8, i).Value = SimulationData.Model.Grids(i).Pressure.Oil

                .Item(9, i).Value = SimulationData.Rocks(SimulationData.Model.Grids(i).RockType).ID

                Dim strSchedules As String = ""
                For j As Integer = 0 To SizeOf(SimulationData.Model.Grids(i).Schedules) - 1
                    strSchedules += SimulationData.Schedules(SimulationData.Model.Grids(i).Schedules(j)).ID.ToString
                    If j < SizeOf(SimulationData.Model.Grids(i).Schedules) - 1 Then strSchedules += ","
                Next
                .Item(10, i).Value = strSchedules
            Next

        End With
    End Sub
    Sub CastModelWells()
        Do While TabControl6.TabPages.Count <> 1
            TabControl6.TabPages.RemoveAt(TabControl6.TabPages.Count - 1)
        Loop

        For i As Integer = 0 To SizeOf(SimulationData.Model.Wells) - 1
            Dim TB As New TabPage
            TB.Text = SimulationData.Model.Wells(i).Name
            Dim ctlWell As New ctlWell
            AddHandler ctlWell.Removed, AddressOf RemoveWell
            AddHandler ctlWell.Update, AddressOf UpdateWell
            ctlWell.Dock = DockStyle.Fill
            castWell(ctlWell, SimulationData.Model.Wells(i))
            TB.Controls.Add(ctlWell)
            TabControl6.TabPages.Add(TB)

        Next

        Dim TB1 As New TabPage
        TB1.Text = "New"
        TabControl6.TabPages.Add(TB1)

    End Sub

    Sub castWell(ByVal ctl As ctlWell, ByVal myWell As Well)
        With ctl
            .TextBox1.Text = myWell.Name
            Dim strSchedules As String = ""
            For j As Integer = 0 To SizeOf(myWell.Schedules) - 1
                strSchedules += SimulationData.Schedules(myWell.Schedules(j)).ID.ToString
                If j < SizeOf(myWell.Schedules) - 1 Then strSchedules += ","
            Next
            .TextBox2.Text = strSchedules
            .ValueBox2.Text = myWell.DefaultDiameter
            .ValueBox1.Text = myWell.DefaultSkin

            .dgvWellPath.RowCount = SizeOf(myWell.Paths) + 1
            For i As Integer = 0 To SizeOf(myWell.Paths) - 1
                .dgvWellPath.Item(0, i).Value = myWell.Paths(i).X1.ToString + "," + myWell.Paths(i).Y1.ToString + "," + myWell.Paths(i).Z1.ToString
                .dgvWellPath.Item(1, i).Value = myWell.Paths(i).X2.ToString + "," + myWell.Paths(i).Y2.ToString + "," + myWell.Paths(i).Z2.ToString
                .dgvWellPath.Item(2, i).Value = myWell.Paths(i).Diameter
                .dgvWellPath.Item(3, i).Value = IIf(myWell.Paths(i).Flow, "True", "False")
                .dgvWellPath.Item(4, i).Value = myWell.Paths(i).Skin
            Next

        End With

    End Sub
    Sub castPVT()
        With dgvPVT
            If SizeOf(SimulationData.PVT.Table) = 0 Then Exit Sub
            .RowCount = SimulationData.PVT.Table.GetLength(1) + 1
            For i As Integer = 0 To SimulationData.PVT.Table.GetLength(0) - 1
                For j As Integer = 0 To SimulationData.PVT.Table.GetLength(1) - 1
                    .Item(i, j).Value = SimulationData.PVT.Table(i, j)
                Next
            Next
        End With
    End Sub
    Sub castSimulation()
        Try
            nudFinal.Value = SimulationData.Final_Step
        Catch ex As Exception
            nudFinal.Value = nudFinal.Minimum
        End Try

        ValueBox1.Text = SimulationData.Finish_Time
        dgvTimeSteps.RowCount = SizeOf(SimulationData.Time_Steps) + 1
        For i As Integer = 0 To SizeOf(SimulationData.Time_Steps) - 1
            dgvTimeSteps.Item(0, i).Value = SimulationData.Time_Steps(i)
        Next
        ComboBox1.SelectedIndex = IIf(SimulationData.System.Units = System_of_Units.Field, 0, 1)
        CheckBox2.Checked = SimulationData.System.Phase.Water
        CheckBox1.Checked = SimulationData.System.Phase.Oil
        CheckBox3.Checked = SimulationData.System.Phase.Gas

        Try
            NUDIteration.Value = SimulationData.Exactness.IterationMSE.Maximum
        Catch ex As Exception
            NUDIteration.Value = 10
        End Try
        Try
            orthoNUD.Value = SimulationData.Exactness.OrthominMaxIteration
        Catch ex As Exception
            orthoNUD.Value = 20
        End Try

        ValueBox2.Text = SimulationData.Exactness.IterationMSE.Sw
        ValueBox3.Text = SimulationData.Exactness.IterationMSE.Sg
        ValueBox4.Text = SimulationData.Exactness.IterationMSE.Po
        ValueBox5.Text = SimulationData.Exactness.Equation

        ComboBox2.SelectedIndex = SimulationData.Methodology.Transmissibility_Fp_Averaging
        ComboBox5.SelectedIndex = SimulationData.Methodology.Approach

        ValueBox8.Text = SimulationData.Methodology.Derivations.Water_Saturation_Epsilon
        ValueBox7.Text = SimulationData.Methodology.Derivations.Gas_Saturation_Epsilon
        ValueBox6.Text = SimulationData.Methodology.Derivations.Oil_Pressure_Epsilon

        CheckBox4.Checked = SimulationData.Options.Automatic_Grid_Adjacent_Recognition
    End Sub
#End Region


    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        SimulationData.Options.Automatic_Grid_Adjacent_Recognition = CheckBox4.Checked
    End Sub

    Private Sub dgvGrids_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvGrids.CellMouseClick
        If e.RowIndex >= SizeOf(SimulationData.Model.Grids) Then Exit Sub
    End Sub

    Private Sub dgvGrids_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvGrids.CellValidated
        If e.RowIndex >= SizeOf(SimulationData.Model.Grids) Then Exit Sub
        If isBulkGrid Then Exit Sub
        Select Case e.ColumnIndex
            Case 0
                SimulationData.Model.Grids(e.RowIndex).Location.X = Val(dgvGrids.Item(0, e.RowIndex).Value)
            Case 1
                SimulationData.Model.Grids(e.RowIndex).Location.Y = Val(dgvGrids.Item(1, e.RowIndex).Value)
            Case 2
                SimulationData.Model.Grids(e.RowIndex).Location.Z = Val(dgvGrids.Item(2, e.RowIndex).Value)
            Case 3
                SimulationData.Model.Grids(e.RowIndex).Size.X_Length = Val(dgvGrids.Item(3, e.RowIndex).Value)
            Case 4
                SimulationData.Model.Grids(e.RowIndex).Size.Y_Length = Val(dgvGrids.Item(4, e.RowIndex).Value)
            Case 5
                SimulationData.Model.Grids(e.RowIndex).Size.Z_Length = Val(dgvGrids.Item(5, e.RowIndex).Value)
            Case 6
                SimulationData.Model.Grids(e.RowIndex).Saturation.Water = Val(dgvGrids.Item(6, e.RowIndex).Value)
            Case 7
                SimulationData.Model.Grids(e.RowIndex).Saturation.Gas = Val(dgvGrids.Item(7, e.RowIndex).Value)
            Case 8
                SimulationData.Model.Grids(e.RowIndex).Pressure.Oil = Val(dgvGrids.Item(8, e.RowIndex).Value)
            Case 9
                For i As Integer = 0 To SizeOf(SimulationData.Rocks) - 1
                    If SimulationData.Rocks(i).ID = Val(dgvGrids.Item(9, e.RowIndex).Value) Then
                        SimulationData.Model.Grids(e.RowIndex).RockType = i
                    End If
                Next

            Case 10
                Try

                    If dgvGrids.Item(10, e.RowIndex).Value <> "" Then
                        Dim scd() As String = Split(dgvGrids.Item(10, e.RowIndex).Value, ",")
                        ReDim SimulationData.Model.Grids(e.RowIndex).Schedules(SizeOf(scd) - 1)
                        For i As Integer = 0 To SizeOf(scd) - 1
                            For n As Integer = 0 To SizeOf(SimulationData.Schedules) - 1
                                If SimulationData.Schedules(n).ID = Val(scd(i)) Then
                                    SimulationData.Model.Grids(e.RowIndex).Schedules(i) = n
                                End If
                            Next

                        Next
                    Else
                        SimulationData.Model.Grids(e.RowIndex).Schedules = Nothing
                    End If
                Catch ex As Exception

                End Try


        End Select

        Wpfx2.AddReplaceSingleGrid(SimulationData.Model.Grids(e.RowIndex), e.RowIndex)
    End Sub

    Private Sub dgvGrids_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvGrids.RowsAdded
        If isBulkGrid Then Exit Sub

        If dgvGrids.RowCount > SizeOf(SimulationData.Model.Grids) + 1 Then
            ReDim Preserve SimulationData.Model.Grids(SizeOf(SimulationData.Model.Grids))
            dgvGrids.Item(0, dgvGrids.RowCount - 2).Value = 0
            dgvGrids.Item(1, dgvGrids.RowCount - 2).Value = 0
            dgvGrids.Item(2, dgvGrids.RowCount - 2).Value = 0
            dgvGrids.Item(3, dgvGrids.RowCount - 2).Value = 1
            dgvGrids.Item(4, dgvGrids.RowCount - 2).Value = 1
            dgvGrids.Item(5, dgvGrids.RowCount - 2).Value = 1

            SimulationData.Model.Grids(SizeOf(SimulationData.Model.Grids) - 1).Location.X = 0
            SimulationData.Model.Grids(SizeOf(SimulationData.Model.Grids) - 1).Location.Y = 0
            SimulationData.Model.Grids(SizeOf(SimulationData.Model.Grids) - 1).Location.Z = 0
            SimulationData.Model.Grids(SizeOf(SimulationData.Model.Grids) - 1).Size.X_Length = 1
            SimulationData.Model.Grids(SizeOf(SimulationData.Model.Grids) - 1).Size.Y_Length = 1
            SimulationData.Model.Grids(SizeOf(SimulationData.Model.Grids) - 1).Size.Z_Length = 1
            SimulationData.Model.Grids(SizeOf(SimulationData.Model.Grids) - 1).Color = Windows.Media.Colors.Green

            dgvGrids.Item(10, dgvGrids.RowCount - 2).Value = 1

            Wpfx2.AddReplaceSingleGrid(SimulationData.Model.Grids(SizeOf(SimulationData.Model.Grids) - 1))
        End If
    End Sub

    Function GetNewRowID() As Integer
        Dim n As Integer
        Dim found As Boolean
        Do While Not found
            found = True
            For i As Integer = 0 To dgvGrids.RowCount - 1
                If Not String.IsNullOrEmpty(dgvGrids.Item(0, i).Value) Then
                    If dgvGrids.Item(0, i).Value = n Then
                        found = False
                    End If
                End If
            Next
            If Not found Then n += 1
        Loop
        Return n
    End Function

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim dlgOpen As New OpenFileDialog
        With dlgOpen
            .Filter = "*.xml|*.xml"
            .ShowDialog()
            If Not String.IsNullOrEmpty(.FileName) Then
                isBulkGrid = True
                OpenFile(.FileName, SimulationData)
                castSimulation()
                castModel()
                castPVT()
                castRockTypes()
                castSchedules()
                Wpfx2.LoadGridList(SimulationData.Model.Grids)
                Application.DoEvents()
                isBulkGrid = False
            End If
        End With
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click

        If frmRegular.ShowDialog = Windows.Forms.DialogResult.OK Then
            isBulkGrid = True
            CreateGrids(SimulationData.Model.Grids, frmRegular.ValueBox1.Text, frmRegular.ValueBox2.Text, frmRegular.ValueBox3.Text, frmRegular.NumericUpDown1.Value, frmRegular.NumericUpDown2.Value, frmRegular.NumericUpDown3.Value, frmRegular.ValueBox6.Text, frmRegular.ValueBox5.Text, frmRegular.ValueBox4.Text)
            CastModelGrids()
            Application.DoEvents()
            isBulkGrid = False
            Wpfx2.LoadGridList(SimulationData.Model.Grids)
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If ErrorMessage <> "" Then
            MessageBox.Show(ErrorMessage)
            ErrorMessage = ""
        End If
    End Sub

    Sub RemoveSchedule(ByVal ID As Integer)
        For i As Integer = 0 To TabControl5.TabPages.Count - 1
            If TabControl5.TabPages(i).Text = "ID " + ID.ToString Then
                TabControl5.TabPages.RemoveAt(i)
                Exit For
            End If
        Next

        Dim tempSC(SizeOf(SimulationData.Schedules) - 2) As Simulation_Schedule
        Dim n As Integer
        For i As Integer = 0 To SizeOf(SimulationData.Schedules) - 2
            If SimulationData.Schedules(i).ID <> ID Then
                tempSC(n) = SimulationData.Schedules(i)
            End If
            n += 1
        Next
        SimulationData.Schedules = tempSC
    End Sub

    Sub RemoveWell(ByVal WellName As String)
        For i As Integer = 1 To TabControl6.TabPages.Count - 2
            If CType(TabControl6.TabPages(i).Controls(0), ctlWell).TextBox1.Text = WellName Then
                TabControl6.TabPages.RemoveAt(i)
                Exit For
            End If
        Next

        Dim tempWells(SizeOf(SimulationData.Model.Wells) - 2) As Well
        Dim n As Integer
        For i As Integer = 0 To SizeOf(SimulationData.Model.Wells) - 2
            If SimulationData.Model.Wells(i).Name <> WellName Then
                tempWells(n) = SimulationData.Model.Wells(i)
            End If
            n += 1
        Next
        SimulationData.Model.Wells = tempWells
    End Sub


    Sub RemoveRockType(ByVal ID As Integer)
        If TabControl4.TabPages.Count = 3 Then
            MessageBox.Show("At least one Rock Type is required", "Black-Oil MSV", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        For i As Integer = 0 To TabControl4.TabPages.Count - 1
            If TabControl4.TabPages(i).Text = "ID " + ID.ToString Then
                TabControl4.TabPages.RemoveAt(i)
                Exit For
            End If
        Next
        Dim tempRock(SizeOf(SimulationData.Rocks) - 2) As Rock
        Dim n As Integer
        For i As Integer = 0 To SizeOf(SimulationData.Rocks) - 2
            If SimulationData.Rocks(i).ID <> ID Then
                tempRock(n) = SimulationData.Rocks(i)
            End If
            n += 1
        Next
        SimulationData.Rocks = tempRock
    End Sub

    Private Sub TabControl4_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl4.SelectedIndexChanged
        If TabControl4.SelectedTab.Text = "New" Then
            AddNewRockType()
        End If
    End Sub

    Sub AddNewRockType()
        Dim ID As Integer = GetNewRockTypeID()
        TabControl4.SelectedTab.Text = "ID " + ID.ToString

        ReDim Preserve SimulationData.Rocks(SizeOf(SimulationData.Rocks))
        With SimulationData.Rocks(SizeOf(SimulationData.Rocks) - 1)
            .ID = ID
            Dim ctlRT As New ctlRockType
            AddHandler ctlRT.Removed, AddressOf RemoveRockType
            ctlRT.ID = ID
            ctlRT.Dock = DockStyle.Fill
            TabControl4.SelectedTab.Controls.Add(ctlRT)
        End With

        Dim TBnew As New TabPage
        TBnew.Text = "New"
        TabControl4.TabPages.Add(TBnew)
    End Sub

    Function GetNewRockTypeID() As Integer
        Dim found As Boolean
        Dim n As Integer
        Do While Not found
            found = True
            n += 1
            For i As Integer = 0 To TabControl4.TabPages.Count - 1
                If TabControl4.TabPages(i).Text = "ID " + n.ToString Then
                    found = False
                End If
            Next
        Loop
        Return n.ToString
    End Function

    Function GetNewScheduleID() As Integer
        Dim found As Boolean
        Dim n As Integer
        Do While Not found
            found = True
            n += 1
            For i As Integer = 0 To TabControl5.TabPages.Count - 1
                If TabControl5.TabPages(i).Text = "ID " + n.ToString Then
                    found = False
                End If
            Next
        Loop
        Return n.ToString
    End Function

    Private Sub tmrLoadEditGraph_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrLoadEditGraph.Tick
        tmrLoadEditGraph.Enabled = False
        Application.DoEvents()
        RefreshColors(ColorBoundary1.ComboBox1.Text, ColorBoundary1.ValueBox1.Text, ColorBoundary1.ValueBox2.Text, ColorBoundary1.Button1.BackColor, ColorBoundary1.Button2.BackColor, SimulationData)
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        History.Clear()
        TC.Value = 0
        Panel1.Invalidate()

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Export.ShowDialog()
    End Sub

#Region "Updates"

    Private Sub gbxPVT_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles gbxPVT.Leave
        UpdatePVT()
        Dim nn As String
        For i As Integer = 0 To dgvPVT.RowCount - 2
            For j As Integer = 0 To dgvPVT.Columns.Count - 1
                nn += ".Table(" + j.ToString + "," + i.ToString + ")=" + dgvPVT.Item(j, i).Value.ToString + vbCrLf
            Next

        Next
        Clipboard.SetText(nn)

    End Sub
    Private Sub gbxSRockType_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles gbxSRockType.Leave
        UpdateRockBoxes()
    End Sub
    Private Sub dgvGrids_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvGrids.RowsRemoved
        If isBulkGrid Then Exit Sub
        UpdateModel()
        Wpfx2.myModel3DGroup.Children.RemoveAt(e.RowIndex + WPFX.PreRenderObjects)
    End Sub
    Private Sub gbxSchedule_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles gbxSchedule.Leave
        UpdateSchedules()
    End Sub

    Private Sub TabPage1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage1.Leave
        UpdateSimulation()
    End Sub

    Sub UpdateSimulation()
        With SimulationData
            ReDim .Time_Steps(dgvTimeSteps.RowCount - 2)
            For i As Integer = 0 To dgvTimeSteps.RowCount - 2
                .Time_Steps(i) = Val(dgvTimeSteps.Item(0, i).Value)
                .Final_Step = nudFinal.Value
                .Finish_Time = Val(ValueBox1.Text)
                .System.Phase.Water = (CheckBox1.Checked)
                .System.Phase.Oil = (CheckBox2.Checked)
                .System.Phase.Gas = (CheckBox3.Checked)
                .System.Units = ComboBox1.SelectedIndex
                .Exactness.IterationMSE.Maximum = NUDIteration.Value
                .Exactness.OrthominMaxIteration = orthoNUD.Value
                .Exactness.IterationMSE.Sw = Val(ValueBox2.Text)
                .Exactness.IterationMSE.Sg = Val(ValueBox3.Text)
                .Exactness.IterationMSE.Po = Val(ValueBox4.Text)
                .Exactness.Equation = Val(ValueBox5.Text)
                .Methodology.Derivations.Water_Saturation_Epsilon = Val(ValueBox8.Text)
                .Methodology.Derivations.Gas_Saturation_Epsilon = Val(ValueBox7.Text)
                .Methodology.Derivations.Oil_Pressure_Epsilon = Val(ValueBox6.Text)
                .Methodology.Transmissibility_Fp_Averaging = ComboBox2.SelectedIndex
                .Methodology.Approach = ComboBox5.SelectedIndex
                .Options.Automatic_Grid_Adjacent_Recognition = CheckBox4.Checked

            Next
        End With
    End Sub

    Sub UpdateSchedules()
        ReDim SimulationData.Schedules(TabControl5.TabPages.Count - 3)
        Dim n As Integer
        For i As Integer = 1 To TabControl5.TabPages.Count - 2
            Dim ctl As ctlSchedule = TabControl5.TabPages(i).Controls(0)
            With ctl
                SimulationData.Schedules(n).ID = Val(Split(TabControl5.TabPages(i).Text, " ")(1))
                SimulationData.Schedules(n).Target = .ComboBox1.SelectedIndex
                SimulationData.Schedules(n).Type = .ComboBox2.SelectedIndex
                SimulationData.Schedules(n).Position = .ComboBox3.SelectedIndex
                SimulationData.Schedules(n).Orientation = .ComboBox4.SelectedIndex

                Select Case SimulationData.Schedules(n).Target
                    Case Simulation_Schedule_Target.Grid
                        If .ComboBox2.Text.ToUpper.StartsWith("STATIC") Then
                            ReDim SimulationData.Schedules(n).Data(0, 0)
                            SimulationData.Schedules(n).Data(0, 0) = Val(.ValueBox1.Text)
                        ElseIf .ComboBox2.Text.ToUpper.StartsWith("DYNAMIC") Then
                            ReDim SimulationData.Schedules(n).Data(1, .dgvGrid.RowCount - 2)
                            For j As Integer = 0 To .dgvGrid.RowCount - 2
                                SimulationData.Schedules(n).Data(0, j) = Val(.dgvGrid(0, j).Value)
                                SimulationData.Schedules(n).Data(1, j) = Val(.dgvGrid(1, j).Value)
                            Next
                        End If

                    Case Simulation_Schedule_Target.Well

                End Select

            End With
            n += 1
        Next

    End Sub

    Sub UpdateModel()
        ReDim SimulationData.Model.Grids(dgvGrids.RowCount - 2)
        For i As Integer = 0 To dgvGrids.RowCount - 2
            With SimulationData.Model.Grids(i)
                .Location.X = Val(dgvGrids.Item(0, i).Value)
                .Location.Y = Val(dgvGrids.Item(1, i).Value)
                .Location.Z = Val(dgvGrids.Item(2, i).Value)
                .Size.X_Length = Val(dgvGrids.Item(3, i).Value)
                .Size.Y_Length = Val(dgvGrids.Item(4, i).Value)
                .Size.Z_Length = Val(dgvGrids.Item(5, i).Value)
                .Saturation.Water = Val(dgvGrids.Item(6, i).Value)
                .Saturation.Gas = Val(dgvGrids.Item(7, i).Value)
                .Pressure.Oil = Val(dgvGrids.Item(8, i).Value)
                For j As Integer = 0 To SizeOf(SimulationData.Rocks) - 1
                    If SimulationData.Rocks(j).ID = Val(dgvGrids.Item(9, i).Value) Then
                        .RockType = j
                    End If
                Next

            End With
        Next
    End Sub


    Sub UpdatePVT()
        ReDim SimulationData.PVT.Table(dgvPVT.ColumnCount - 1, dgvPVT.RowCount - 2)
        For i As Integer = 0 To dgvPVT.ColumnCount - 1
            For j As Integer = 0 To dgvPVT.RowCount - 2
                SimulationData.PVT.Table(i, j) = Val(dgvPVT.Item(i, j).Value)
            Next
        Next
    End Sub

    Sub UpdateRockBoxes()
        ReDim SimulationData.Rocks(TabControl4.TabPages.Count - 3)
        Dim n As Integer
        For i As Integer = 1 To TabControl4.TabPages.Count - 2
            Dim ctl As ctlRockType = TabControl4.TabPages(i).Controls(0)
            With ctl

                SimulationData.Rocks(n).Initial_Porosity = Val(.ValueBox10.Text)
                SimulationData.Rocks(n).Cs = Val(.ValueBox11.Text)
                SimulationData.Rocks(n).Permeability.Kx = Val(.ValueBox1.Text)
                SimulationData.Rocks(n).Permeability.Ky = Val(.ValueBox2.Text)
                SimulationData.Rocks(n).Permeability.Kz = Val(.ValueBox3.Text)

                SimulationData.Rocks(n).ID = Val(Split(TabControl4.TabPages(i).Text, " ")(1))

                If .RadioButton2.Checked Then
                    SimulationData.Rocks(n).SCAL.Wettability.Water_Oil = Phase.Water
                ElseIf .RadioButton1.Checked Then
                    SimulationData.Rocks(n).SCAL.Wettability.Water_Oil = Phase.Oil
                ElseIf .RadioButton7.Checked Then
                    SimulationData.Rocks(n).SCAL.Wettability.Water_Oil = Phase.Unknown
                End If

                If .RadioButton4.Checked Then
                    SimulationData.Rocks(n).SCAL.Wettability.Oil_Gas = Phase.Gas
                ElseIf .RadioButton3.Checked Then
                    SimulationData.Rocks(n).SCAL.Wettability.Oil_Gas = Phase.Oil
                ElseIf .RadioButton8.Checked Then
                    SimulationData.Rocks(n).SCAL.Wettability.Oil_Gas = Phase.Unknown
                End If

                If .RadioButton6.Checked Then
                    SimulationData.Rocks(n).SCAL.Wettability.Water_Gas = Phase.Water
                ElseIf .RadioButton5.Checked Then
                    SimulationData.Rocks(n).SCAL.Wettability.Water_Gas = Phase.Gas
                ElseIf .RadioButton9.Checked Then
                    SimulationData.Rocks(n).SCAL.Wettability.Water_Gas = Phase.Unknown
                End If

                ReDim Preserve SimulationData.Rocks(n).SCAL.Tables.Water_Oil(.dgvOW.ColumnCount - 1, .dgvOW.RowCount - 2)
                For i1 As Integer = 0 To .dgvOW.ColumnCount - 1
                    For i2 As Integer = 0 To .dgvOW.RowCount - 2
                        SimulationData.Rocks(n).SCAL.Tables.Water_Oil(i1, i2) = Val(.dgvOW(i1, i2).Value)
                    Next
                Next

                ReDim Preserve SimulationData.Rocks(n).SCAL.Tables.Oil_Gas(.dgvGO.ColumnCount - 1, .dgvGO.RowCount - 2)
                For i1 As Integer = 0 To .dgvGO.ColumnCount - 1
                    For i2 As Integer = 0 To .dgvGO.RowCount - 2
                        SimulationData.Rocks(n).SCAL.Tables.Oil_Gas(i1, i2) = Val(.dgvGO(i1, i2).Value)
                    Next
                Next

                ReDim Preserve SimulationData.Rocks(n).SCAL.Tables.Gas_Water(.dgvWG.ColumnCount - 1, .dgvWG.RowCount - 2)
                For i1 As Integer = 0 To .dgvWG.ColumnCount - 1
                    For i2 As Integer = 0 To .dgvWG.RowCount - 2
                        SimulationData.Rocks(n).SCAL.Tables.Gas_Water(i1, i2) = Val(.dgvWG(i1, i2).Value)
                    Next
                Next


            End With
            n += 1
        Next
    End Sub
#End Region

    Private Sub TabControl5_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl5.SelectedIndexChanged
        If TabControl5.SelectedTab.Text = "New" Then
            Dim ID As Integer = GetNewScheduleID()
            TabControl5.SelectedTab.Text = "ID " + ID.ToString
            ReDim Preserve SimulationData.Schedules(SizeOf(SimulationData.Schedules))
            With SimulationData.Schedules(SizeOf(SimulationData.Schedules) - 1)
                .ID = ID
                Dim ctlSC As New ctlSchedule
                AddHandler ctlSC.Removed, AddressOf RemoveSchedule
                ctlSC.ID = ID
                ctlSC.Dock = DockStyle.Fill
                TabControl5.SelectedTab.Controls.Add(ctlSC)

            End With

            Dim TBnew As New TabPage
            TBnew.Text = "New"
            TabControl5.TabPages.Add(TBnew)
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        About.ShowDialog()
    End Sub

    Function GetNewWellName() As String
        Dim found As Boolean
        Dim n As Integer
        Do While Not found
            found = True
            n += 1
            For i As Integer = 0 To SizeOf(SimulationData.Model.Wells) - 1
                If SimulationData.Model.Wells(i).Name.ToUpper = "WELL" + n.ToString Then
                    found = False
                End If
            Next
        Loop
        Return "Well" + n.ToString
    End Function

    Sub UpdateWell(ByRef sender As ctlWell)
        Dim nTab As Integer
        Dim tbContainer As TabPage = sender.Parent
        For i As Integer = 1 To TabControl6.TabPages.Count - 2
            If TabControl6.TabPages(i).Text = tbContainer.Text Then
                nTab = i
                Exit For
            End If
        Next

        tbContainer.Text = sender.TextBox1.Text
        SimulationData.Model.Wells(nTab - 1).Name = tbContainer.Text
        If sender.TextBox2.Text <> "" Then
            Dim scd() As String = Split(sender.TextBox2.Text, ",")
            ReDim Preserve SimulationData.Model.Wells(nTab - 1).Schedules(SizeOf(scd) - 1)
            For n As Integer = 0 To SizeOf(SimulationData.Schedules) - 1
                For m As Integer = 0 To SizeOf(scd) - 1
                    If SimulationData.Schedules(n).ID = scd(m) Then
                        SimulationData.Model.Wells(nTab - 1).Schedules(m) = n
                    End If
                Next
            Next
        Else
            SimulationData.Model.Wells(nTab - 1).Schedules = Nothing
        End If

        SimulationData.Model.Wells(nTab - 1).DefaultDiameter = Val(sender.ValueBox2.Text)
        SimulationData.Model.Wells(nTab - 1).DefaultSkin = Val(sender.ValueBox1.Text)


        Dim num As Integer
        For i As Integer = 0 To CType(TabControl6.TabPages(nTab).Controls(0), ctlWell).dgvWellPath.RowCount - 2
            If CType(TabControl6.TabPages(nTab).Controls(0), ctlWell).dgvWellPath.Item(0, i).Value <> "" And CType(TabControl6.TabPages(nTab).Controls(0), ctlWell).dgvWellPath.Item(1, i).Value <> "" Then
                num += 1
            End If
        Next

        ReDim Preserve SimulationData.Model.Wells(nTab - 1).Paths(num - 1)
        For i As Integer = 0 To num - 1
            Dim strStart() = Split(CType(TabControl6.TabPages(nTab).Controls(0), ctlWell).dgvWellPath.Item(0, i).Value, ",")
            Dim strFinish() = Split(CType(TabControl6.TabPages(nTab).Controls(0), ctlWell).dgvWellPath.Item(1, i).Value, ",")

            If SizeOf(strStart) = 3 Then
                SimulationData.Model.Wells(nTab - 1).Paths(i).X1 = strStart(0)
                SimulationData.Model.Wells(nTab - 1).Paths(i).Y1 = strStart(1)
                SimulationData.Model.Wells(nTab - 1).Paths(i).Z1 = strStart(2)
            End If

            If SizeOf(strFinish) = 3 Then
                SimulationData.Model.Wells(nTab - 1).Paths(i).X2 = strFinish(0)
                SimulationData.Model.Wells(nTab - 1).Paths(i).Y2 = strFinish(1)
                SimulationData.Model.Wells(nTab - 1).Paths(i).Z2 = strFinish(2)
            End If

            SimulationData.Model.Wells(nTab - 1).Paths(i).Diameter = Val(CType(TabControl6.TabPages(nTab).Controls(0), ctlWell).dgvWellPath.Item(2, i).Value)
            SimulationData.Model.Wells(nTab - 1).Paths(i).Flow = (CType(TabControl6.TabPages(nTab).Controls(0), ctlWell).dgvWellPath.Item(3, i).Value)
            SimulationData.Model.Wells(nTab - 1).Paths(i).Skin = Val(CType(TabControl6.TabPages(nTab).Controls(0), ctlWell).dgvWellPath.Item(4, i).Value)
        Next
    End Sub

    Private Sub TabControl6_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl6.SelectedIndexChanged
        If TabControl6.SelectedTab.Text = "New" Then
            Dim WellName As String = GetNewWellName()
            TabControl6.SelectedTab.Text = WellName.ToString

            ReDim Preserve SimulationData.Model.Wells(SizeOf(SimulationData.Model.Wells))
            With SimulationData.Model.Wells(SizeOf(SimulationData.Model.Wells) - 1)
                .Name = WellName
                Dim ctlWell As New ctlWell
                AddHandler ctlWell.Removed, AddressOf RemoveWell
                AddHandler ctlWell.Update, AddressOf UpdateWell
                ctlWell.Dock = DockStyle.Fill
                ctlWell.TextBox1.Text = .Name
                TabControl6.SelectedTab.Controls.Add(ctlWell)
            End With

            Dim TBnew As New TabPage
            TBnew.Text = "New"
            TabControl6.TabPages.Add(TBnew)
        End If

    End Sub

    Private Sub Button12_Click_1(sender As Object, e As EventArgs) Handles Button12.Click

    End Sub
End Class










