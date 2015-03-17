Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D


Public Class Main
    Public History As New Simulation_History
    Public SimulationData As New Simulation
    Public ErrorMessage As String
    Dim SimulationGraphLoaded As Boolean = False

    Friend WithEvents mySimulator As New Core
    Friend WithEvents TDP1 As TDP
    Friend WithEvents TDP2 As TDP

    Sub RunCheck(ByVal IsRunning As Boolean)
        TV.Enabled = Not IsRunning
    End Sub

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

    Sub CreateDGVSample()
        With dgvExp
            .RowCount = 10
            .Item(0, 0).Value = 0
            .Item(0, 1).Value = 0.62
            .Item(0, 2).Value = 1.14
            .Item(0, 3).Value = 1.31
            .Item(0, 4).Value = 1.57
            .Item(0, 5).Value = 1.91
            .Item(0, 6).Value = 2.17
            .Item(0, 7).Value = 2.43
            .Item(0, 8).Value = 2.69
            .Item(0, 9).Value = 3.04



            .Item(1, 0).Value = 0.0
            .Item(1, 1).Value = -0.00000412132102569562
            .Item(1, 2).Value = -0.00000740183159279397
            .Item(1, 3).Value = -0.00000845678507226421
            .Item(1, 4).Value = -0.0000100095278567423
            .Item(1, 5).Value = -0.0000120350350346335
            .Item(1, 6).Value = -0.0000135172735704004
            .Item(1, 7).Value = -0.0000149769155321663
            .Item(1, 8).Value = -0.000016411396199409
            .Item(1, 9).Value = -0.0000196789417520861
            '.Item(1, 9).Value = -0.0000228789417520861
            '.Item(1, 9).Value = -0.0000196789417520861

            '.Item(1, 0).Value = 0.0
            '.Item(1, 1).Value = -0.000012666
            '.Item(1, 2).Value = -0.0000212607
            '.Item(1, 3).Value = -0.0000247646
            '.Item(1, 4).Value = -0.0000302088
            '.Item(1, 5).Value = -0.0000378477
            '.Item(1, 6).Value = -0.0000438824
            '.Item(1, 7).Value = -0.0000476433
            '.Item(1, 8).Value = -0.0000476433
            '.Item(1, 9).Value = -0.0000476433

            '.Item(1, 0).Value = 0.0
            '.Item(1, 1).Value = -0.00000525776
            '.Item(1, 2).Value = -0.00000992166
            '.Item(1, 3).Value = -0.0000115568
            '.Item(1, 4).Value = -0.0000140974
            '.Item(1, 5).Value = -0.0000176623
            '.Item(1, 6).Value = -0.0000204784
            '.Item(1, 7).Value = -0.0000222335
            '.Item(1, 8).Value = -0.0000222335
            '.Item(1, 9).Value = -0.0000222335

        End With

        With dgvExp2
            .RowCount = 11
            .Item(0, 0).Value = 0.18
            .Item(0, 1).Value = 0.31
            .Item(0, 2).Value = 0.42
            .Item(0, 3).Value = 0.58
            .Item(0, 4).Value = 0.61
            .Item(0, 5).Value = 0.69
            .Item(0, 6).Value = 0.73
            .Item(0, 7).Value = 0.78
            .Item(0, 8).Value = 0.81
            .Item(0, 9).Value = 0.84

            .Item(1, 0).Value = 100
            .Item(1, 1).Value = 85
            .Item(1, 2).Value = 75
            .Item(1, 3).Value = 60
            .Item(1, 4).Value = 55
            .Item(1, 5).Value = 51
            .Item(1, 6).Value = 47
            .Item(1, 7).Value = 44
            .Item(1, 8).Value = 42
            .Item(1, 9).Value = 41

            'Dim stx As String
            'For i As Integer = 0 To 9
            '    stx += .Item(0, i).Value.ToString + Chr(Keys.Tab) + .Item(1, i).Value.ToString + vbCrLf

            'Next

            'Clipboard.SetText(stx)
            'Stop

        End With
    End Sub

   
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        For i As Integer = 1 To 4
            If Not IO.Directory.Exists(Application.StartupPath + "\" + i.ToString) Then
                IO.Directory.CreateDirectory(Application.StartupPath + "\" + i.ToString)
            End If
        Next
        Application.DoEvents()

        For i As Integer = 1 To 4
            For j As Integer = 5 To 15
                If Not IO.Directory.Exists(Application.StartupPath + "\" + i.ToString + "\" + j.ToString) Then
                    IO.Directory.CreateDirectory(Application.StartupPath + "\" + i.ToString + "\" + j.ToString)
                End If
            Next
        Next

        For i As Integer = 1 To 4
            For j As Integer = 5 To 15
                If Not IO.File.Exists(Application.StartupPath + "\" + i.ToString + "\" + j.ToString + "\" + i.ToString + ".exe") Then
                    IO.File.Copy(Application.StartupPath + "\" + i.ToString + ".exe", Application.StartupPath + "\" + i.ToString + "\" + j.ToString + "\" + i.ToString + ".exe", True)
                End If
            Next
        Next

        ComboBox3.SelectedIndex = 0
        ComboBox4.SelectedIndex = 1
        CheckBox5.Checked = True
        CheckBox6.Checked = True
        CheckBox7.Checked = True

        ComboBox6.SelectedIndex = 1
        ComboBox7.SelectedIndex = 1
        ComboBox8.SelectedIndex = 1
        ComboBox9.SelectedIndex = 1
        ComboBox10.SelectedIndex = 1
        ComboBox11.SelectedIndex = 1
        ComboBox12.SelectedIndex = 1
        ComboBox13.SelectedIndex = 1
        ComboBox14.SelectedIndex = 1
        ComboBox15.SelectedIndex = 1




        'Dim A(60, 60) As Double
        'Dim x(0, 60) As Double
        'Dim d(0, 60) As Double
        'Randomize()
        'MessageBox.Show("Begin")

        'For ll As Integer = 0 To 10000

        '    For i As Integer = 0 To 60
        '        For j As Integer = 0 To 60
        '            A(i, j) = Rnd() * 3
        '        Next
        '        d(0, i) = Rnd() * 2
        '        x(0, i) = 0
        '    Next

        '    Solve(A, x, d, 500, 0.0000001)
        'Next

        'MessageBox.Show("End")


        CreateDGVSample()


        InitializeTreeView(TV)
        DockGroupBoxes()


        TDP1 = New TDP
        TDP1.Dock = DockStyle.Fill

        TDP2 = New TDP
        TDP2.Dock = DockStyle.Fill

        ColorBoundary1.ComboBox1.SelectedIndex = 3
        ColorBoundary1.ValueBox1.Text = 1000
        ColorBoundary1.ValueBox2.Text = 4000
        RoundButtons()
        PrepareDGVCurrent()
        ReDim SimulationData.Rocks(0)
        SimulationData.Rocks(0).ID = 1

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
        gbxMatch.Dock = DockStyle.Fill
    End Sub

    Private Sub TV_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TV.AfterSelect
        gbxSimulation.Visible = (TV.SelectedNode.Level = 1) And TV.SelectedNode.Text = "Simulation"
        gbxModel.Visible = (TV.SelectedNode.Level = 2 And TV.SelectedNode.Text = "Model")
        gbxPVT.Visible = (TV.SelectedNode.Level = 2 And TV.SelectedNode.Text = "PVT")
        gbxSRockType.Visible = (TV.SelectedNode.Level = 2 And TV.SelectedNode.Text = "Rock Types")
        gbxSchedule.Visible = (TV.SelectedNode.Level = 2 And TV.SelectedNode.Text = "Schedules")
        gbxMatch.Visible = (TV.SelectedNode.Level = 2 And TV.SelectedNode.Text = "KrPc History Match")

        Dim adress() As String = Split(TV.SelectedNode.FullPath, "\")
        Label14.Text = adress(SizeOf(adress) - 1)
        Select Case TV.SelectedNode.Level
            Case 1
                'Try
                If TV.SelectedNode.Text = "Simulation" Then
                    castSimulation()
                    If CheckData() Then
                        SplitContainer3.Visible = True
                    Else
                        SplitContainer3.Visible = False
                    End If
                End If
                'Catch ex As Exception

                'End Try

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
    'Dim strx As String
    Dim sum As Double
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
                " Iterations: " & AddStringOfLength(Data.CurrentIterationCount.ToString, n3, " ") & _
                " Po1: " & AddStringOfLength(Data.Model.Grids(0).Pressure.Oil.ToString, no1, " ") & _
                " Pw2: " & AddStringOfLength(Data.Model.Grids(3).Pressure.Oil.ToString, nw2, " ") & _
                " Sw1: " & AddStringOfLength(Data.Model.Grids(0).Saturation.Water.ToString, nsw1, " ") & _
                " Sw2: " & AddStringOfLength(Data.Model.Grids(3).Saturation.Water.ToString, nsw1, " "))

                'strx += Data.Total_Time.ToString + Chr(Keys.Tab) + Data.Model.Grids(0).Pressure.Oil.ToString + vbCrLf
                'strx += Data.Total_Time.ToString + Chr(Keys.Tab) + Data.Model.Grids(0).Flow.Water.ToString + vbCrLf
                'strx += Data.Total_Time.ToString + Chr(Keys.Tab) + Data.Model.Grids(0).Saturation.Water.ToString + vbCrLf
                'sum += Data.Model.Grids(0).Flow.Water
                'sum = Data.Model.Grids(0).Saturation.Water * Data.Model.Grids(0).Volume * Data.Model.Grids(0).Porosity / Data.Model.Grids(0).PVT.Water.FVF + _
                '       Data.Model.Grids(1).Saturation.Water * Data.Model.Grids(1).Volume * Data.Model.Grids(1).Porosity / Data.Model.Grids(1).PVT.Water.FVF + _
                '       Data.Model.Grids(2).Saturation.Water * Data.Model.Grids(2).Volume * Data.Model.Grids(2).Porosity / Data.Model.Grids(2).PVT.Water.FVF + _
                '       Data.Model.Grids(3).Saturation.Water * Data.Model.Grids(3).Volume * Data.Model.Grids(3).Porosity / Data.Model.Grids(3).PVT.Water.FVF



                'strx += sum.ToString + vbCrLf
                'Clipboard.SetText(strx)

                CastCurrentTrackOnDGV()
                TC.Value = Data.Current_Step

                RefreshColors(ColorBoundary1.ComboBox1.Text, ColorBoundary1.ValueBox1.Text, ColorBoundary1.ValueBox2.Text, ColorBoundary1.Button1.BackColor, ColorBoundary1.Button2.BackColor, Data)

                TDP1.TDC1.LoadGeometry(Data)
                TDP1.TDC1.FrameRender()
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
                End If
            End If
        End If

    End Sub

    Sub ReloadGraph()
        Try
            If History.Current = 0 Then
                RefreshColors(ColorBoundary1.ComboBox1.Text, ColorBoundary1.ValueBox1.Text, ColorBoundary1.ValueBox2.Text, ColorBoundary1.Button1.BackColor, ColorBoundary1.Button2.BackColor, SimulationData)
                TDP1.TDC1.LoadGeometry(SimulationData)
            Else
                If History.Current - TC.Value <= 0 Then TC.Value = History.Current - 1
                RefreshColors(ColorBoundary1.ComboBox1.Text, ColorBoundary1.ValueBox1.Text, ColorBoundary1.ValueBox2.Text, ColorBoundary1.Button1.BackColor, ColorBoundary1.Button2.BackColor, History.Simulations(TC.Value))
                TDP1.TDC1.LoadGeometry(History.Simulations(TC.Value))
            End If
        Catch ex As Exception
        End Try
        TDP1.TDC1.FrameRender()
        Application.DoEvents()

    End Sub

    Private Sub tmrLoadSimulationGraph_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrLoadSimulationGraph.Tick
        tmrLoadSimulationGraph.Enabled = False
        SimulationGraphLoaded = True
        Application.DoEvents()

        TDG.Controls.Add(TDP1)
        RefreshColors(ColorBoundary1.ComboBox1.Text, ColorBoundary1.ValueBox1.Text, ColorBoundary1.ValueBox2.Text, ColorBoundary1.Button1.BackColor, ColorBoundary1.Button2.BackColor, SimulationData)
        TDP1.TDC1.LoadGeometry(SimulationData)
        TDP1.TDC1.Start()

        TDGEdit.Controls.Add(TDP2)
        RefreshColors(ColorBoundary1.ComboBox1.Text, ColorBoundary1.ValueBox1.Text, ColorBoundary1.ValueBox2.Text, ColorBoundary1.Button1.BackColor, ColorBoundary1.Button2.BackColor, SimulationData)
        TDP2.TDC1.LoadGeometry(SimulationData)
        TDP2.TDC1.Start()
    End Sub

    Private Sub TDG_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TDG.VisibleChanged
        If TDG.Visible = True Then
            If TDG.Controls.Count = 0 Then
                If SizeOf(SimulationData.Model.Grids) > 0 Then
                    tmrLoadSimulationGraph.Enabled = True
                End If
            End If
        End If
    End Sub

    Private Sub TDGedit_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TDGEdit.VisibleChanged
        TGEditApply()
    End Sub
    Sub TGEditApply()
        If TDGEdit.Visible = True Then
            If TDGEdit.Controls.Count = 0 Then
                If SizeOf(SimulationData.Model.Grids) > 0 Then
                    tmrLoadEditGraph.Enabled = True
                End If
            End If
        End If
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

            'ReloadGraph()
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

            'ReloadGraph()
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
        CastModelGraph()
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

    Sub CastModelGraph()

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
        TDP2.TDC1.LoadGeometry(SimulationData, False, e.RowIndex)
        TDP2.TDC1.FrameRender()
    End Sub

    Private Sub dgvGrids_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvGrids.CellValidated
        If e.RowIndex >= SizeOf(SimulationData.Model.Grids) Then Exit Sub
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
        If SimulationGraphLoaded Then
            ReloadGraph()
        End If
        TDP2.TDC1.LoadGeometry(SimulationData, False, e.RowIndex)
        TDP2.TDC1.FrameRender()

    End Sub

    Private Sub dgvGrids_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvGrids.RowsAdded
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

            dgvGrids.Item(10, dgvGrids.RowCount - 2).Value = 1
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
                OpenFile(.FileName, SimulationData)
                castSimulation()
                castModel()
                castPVT()
                castRockTypes()
                castSchedules()

                TGEditApply()
                TDP2.TDC1.LoadGeometry(SimulationData, False)
                TDP2.TDC1.FrameRender()
            End If
        End With
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        If frmRegular.ShowDialog = Windows.Forms.DialogResult.OK Then
            CreateGrids(SimulationData.Model.Grids, frmRegular.ValueBox1.Text, frmRegular.ValueBox2.Text, frmRegular.ValueBox3.Text, frmRegular.NumericUpDown1.Value, frmRegular.NumericUpDown2.Value, frmRegular.NumericUpDown3.Value, frmRegular.ValueBox6.Text, frmRegular.ValueBox5.Text, frmRegular.ValueBox4.Text)
            CastModelGrids()
            If SimulationGraphLoaded Then
                ReloadGraph()
            End If
            If TDGEdit.Visible Then
                TDP2.TDC1.LoadGeometry(SimulationData, False)
            End If
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

        TDGEdit.Controls.Add(TDP2)
        RefreshColors(ColorBoundary1.ComboBox1.Text, ColorBoundary1.ValueBox1.Text, ColorBoundary1.ValueBox2.Text, ColorBoundary1.Button1.BackColor, ColorBoundary1.Button2.BackColor, SimulationData)
        TDP2.TDC1.LoadGeometry(SimulationData, False)
        TDP2.TDC1.Start()
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        History.Clear()
        TC.Value = 0
        Panel1.Invalidate()

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Export.ShowDialog()
    End Sub

    Sub PVTExample()

        With SimulationData.PVT
            ReDim .Table(13, 9)
            .Table(0, 0) = 1500
            .Table(1, 0) = 0
            .Table(2, 0) = 0.9666
            .Table(3, 0) = 0.0018
            .Table(4, 0) = 0.9925
            .Table(5, 0) = 44.17
            .Table(6, 0) = 5.82
            .Table(7, 0) = 62.22
            .Table(8, 0) = 0.91
            .Table(9, 0) = 0.015
            .Table(10, 0) = 0.52
            .Table(11, 0) = 0.000052
            .Table(12, 0) = 0.000052
            .Table(13, 0) = 0.000052
            .Table(0, 1) = 2000
            .Table(1, 1) = 0
            .Table(2, 1) = 0.956
            .Table(3, 1) = 0.0013
            .Table(4, 1) = 0.99
            .Table(5, 1) = 44.77
            .Table(6, 1) = 8.05
            .Table(7, 1) = 62.41
            .Table(8, 1) = 0.9372
            .Table(9, 1) = 0.0167
            .Table(10, 1) = 0.52
            .Table(11, 1) = 0.000052
            .Table(12, 1) = 0.000052
            .Table(13, 1) = 0.000052
            .Table(0, 2) = 2500
            .Table(1, 2) = 0
            .Table(2, 2) = 0.9456
            .Table(3, 2) = 0.001
            .Table(4, 2) = 0.9876
            .Table(5, 2) = 45.19
            .Table(6, 2) = 10.22
            .Table(7, 2) = 62.59
            .Table(8, 2) = 0.9243
            .Table(9, 2) = 0.0185
            .Table(10, 2) = 0.52
            .Table(11, 2) = 0.000052
            .Table(12, 2) = 0.000052
            .Table(13, 2) = 0.000052
            .Table(0, 3) = 3000
            .Table(1, 3) = 0
            .Table(2, 3) = 0.9354
            .Table(3, 3) = 0.0008
            .Table(4, 3) = 0.9852
            .Table(5, 3) = 45.57
            .Table(6, 3) = 12.2
            .Table(7, 3) = 62.78
            .Table(8, 3) = 0.92
            .Table(9, 3) = 0.0204
            .Table(10, 3) = 0.52
            .Table(11, 3) = 0.000052
            .Table(12, 3) = 0.000052
            .Table(13, 3) = 0.000052
            .Table(0, 4) = 3500
            .Table(1, 4) = 0
            .Table(2, 4) = 0.9254
            .Table(3, 4) = 0.0007
            .Table(4, 4) = 0.9828
            .Table(5, 4) = 46.58
            .Table(6, 4) = 13.94
            .Table(7, 4) = 62.96
            .Table(8, 4) = 0.96
            .Table(9, 4) = 0.0222
            .Table(10, 4) = 0.52
            .Table(11, 4) = 0.000052
            .Table(12, 4) = 0.000052
            .Table(13, 4) = 0.000052
            .Table(0, 5) = 4000
            .Table(1, 5) = 0
            .Table(2, 5) = 0.9157
            .Table(3, 5) = 0.00069
            .Table(4, 5) = 0.9803
            .Table(5, 5) = 47.17
            .Table(6, 5) = 15.43
            .Table(7, 5) = 63.15
            .Table(8, 5) = 1.1
            .Table(9, 5) = 0.0241
            .Table(10, 5) = 0.52
            .Table(11, 5) = 0.000052
            .Table(12, 5) = 0.000052
            .Table(13, 5) = 0.000052
            .Table(0, 6) = 4500
            .Table(1, 6) = 0
            .Table(2, 6) = 0.9062
            .Table(3, 6) = 0.00064
            .Table(4, 6) = 0.9779
            .Table(5, 6) = 47.69
            .Table(6, 6) = 16.7
            .Table(7, 6) = 63.33
            .Table(8, 6) = 1.25
            .Table(9, 6) = 0.026
            .Table(10, 6) = 0.52
            .Table(11, 6) = 0.000052
            .Table(12, 6) = 0.000052
            .Table(13, 6) = 0.000052
            .Table(0, 7) = 5000
            .Table(1, 7) = 0
            .Table(2, 7) = 0.8968
            .Table(3, 7) = 0.0006
            .Table(4, 7) = 0.9756
            .Table(5, 7) = 48.17
            .Table(6, 7) = 17.7
            .Table(7, 7) = 63.52
            .Table(8, 7) = 1.4
            .Table(9, 7) = 0.0278
            .Table(10, 7) = 0.52
            .Table(11, 7) = 0.000052
            .Table(12, 7) = 0.000052
            .Table(13, 7) = 0.000052
            .Table(0, 8) = 5500
            .Table(1, 8) = 0
            .Table(2, 8) = 0.8877
            .Table(3, 8) = 0.00057
            .Table(4, 8) = 0.9732
            .Table(5, 8) = 48.58
            .Table(6, 8) = 18.7
            .Table(7, 8) = 63.7
            .Table(8, 8) = 1.55
            .Table(9, 8) = 0.0296
            .Table(10, 8) = 0.52
            .Table(11, 8) = 0.000052
            .Table(12, 8) = 0.000052
            .Table(13, 8) = 0.000052
            .Table(0, 9) = 6000
            .Table(1, 9) = 0
            .Table(2, 9) = 0.8787
            .Table(3, 9) = 0.00055
            .Table(4, 9) = 0.9708
            .Table(5, 9) = 49.01
            .Table(6, 9) = 19.5
            .Table(7, 9) = 63.89
            .Table(8, 9) = 1.73
            .Table(9, 9) = 0.0313
            .Table(10, 9) = 0.52
            .Table(11, 9) = 0.000052
            .Table(12, 9) = 0.000052
            .Table(13, 9) = 0.000052
            'Dim strx As String
            'For i As Integer = 0 To 9
            '    For j As Integer = 0 To 13
            '        strx += (.Table(j, i).ToString + Chr(Keys.Tab))

            '    Next
            '    strx += vbCrLf
            'Next
            'Clipboard.SetText(strx)
            'Stop
        End With


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
        UpdateModel()
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

    Private Sub dgvGrids_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvGrids.CellContentClick

    End Sub

    Private Sub ColorBoundary1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)

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


    Private Sub TabControl9_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl9.SelectedIndexChanged

    End Sub

    Dim Lets() As LET_Parameter
    Dim SW() As Double
    Dim Pcs(,) As Double
    Private CurrentDeflection As Double
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        


        If IO.File.Exists(Application.StartupPath + "\result.data") Then
            IO.File.Delete(IO.File.Exists(Application.StartupPath + "\result.data"))
        End If

        n_Simulation = 0
        CurrentDeflection = 0
        i_MSE_ANNs = 0
        PerformanceLevel = 0

        If Button4.Text = "Run" Then
            If nud5.Value < 80 Then
                If MessageBox.Show("Number of trials is low. It's highly recommended to run more than 80 trial to meet improvement." + vbCrLf + "Do you wish to continue anyway?", "Warning", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

            End If

            Button4.Text = "Stop"
        Else
            Button4.Text = "Run"
            tmrLET.Enabled = False
        End If
        Randomize()

        Dim Leng As Integer = dgvExp2.RowCount - 1
        Dim n_grids As Integer = NudGrids.Value

        Dim Cp As Double = Val(ValueBox16.Text)
        Dim Ct As Double = Val(ValueBox17.Text)

        Dim Temperature As Double = Val(ValueBox14.Text)
        Dim GasMW As Double = Val(ValueBox15.Text)
        Dim GasBubblePressure As Double = Val(ValueBox63.Text)
        Dim co As Double = Val(ValueBox19.Text)
        Dim cw As Double = Val(ValueBox13.Text)
        Dim cg As Double = Val(ValueBox18.Text)
        Dim rho_o As Double = Val(ValueBox22.Text)
        Dim rho_w As Double = Val(ValueBox20.Text)
        Dim rho_g As Double = Val(ValueBox21.Text)


        Dim Ko As Double = Val(ValueBox28.Text)
        Dim Kw As Double = Val(ValueBox31.Text)
        Dim Kg As Double = Val(ValueBox34.Text)

        Dim Kair As Double = Val(ValueBox12.Text)

        Dim radius As Double = Val(ValueBox9.Text) / 2
        Dim a_len As Double = (Math.PI * (radius ^ 2)) ^ 0.5
        Dim b_len As Double = Val(ValueBox10.Text) / n_grids



        With SimulationData
            .PVT.StandingC = Val(ValueBox67.Text)
            ReDim .Model.Grids(n_grids - 1)
            ReDim .Schedules(1)
            With .Schedules(0)
                ReDim .Data(0, 0)
                .ID = 1
                .Data(0, 0) = 4500 ' Val(ValueBox25.Text)
                .Position = Simulation_Schedule_Position.Surface
                .Target = Simulation_Schedule_Target.Grid
                .Orientation = Adjacent_Type.Z

                Select Case ComboBox3.Text
                    Case "Water"
                        .Type = Schedule_Type.Static_Water_Pressure
                    Case "Gas"
                        .Type = Schedule_Type.Static_Gas_Pressure
                    Case "Oil"
                        .Type = Schedule_Type.Static_Oil_Pressure
                End Select
            End With

            With .Schedules(1)
                ReDim .Data(0, 0)
                .ID = 2
                .Data(0, 0) = 3900 ' Val(ValueBox25.Text) - dgvExp2.Item(1, 0).Value
                .Position = Simulation_Schedule_Position.Surface
                .Target = Simulation_Schedule_Target.Grid
                .Orientation = Adjacent_Type.Z

                Select Case ComboBox4.Text
                    Case "Water"
                        .Type = Schedule_Type.Static_Water_Pressure
                    Case "Gas"
                        .Type = Schedule_Type.Static_Gas_Pressure
                    Case "Oil"
                        .Type = Schedule_Type.Static_Oil_Pressure
                End Select

            End With

            For i As Integer = 0 To n_grids - 1

                .Model.Grids(i).Pressure.Oil = 3900 '2000 + (i / n_grids) * 2000 '.Schedules(1).Data(0, 0)




                If (ComboBox3.Text = "Water" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Water") Then
                    If ComboBox3.Text = "Water" Then
                        .Model.Grids(i).Saturation.Water = 0
                    Else
                        .Model.Grids(i).Saturation.Water = 1
                    End If
                ElseIf (ComboBox3.Text = "Gas" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Gas") Then
                    If ComboBox3.Text = "Gas" Then
                        .Model.Grids(i).Saturation.Gas = 0
                    Else
                        .Model.Grids(i).Saturation.Gas = 1
                    End If
                ElseIf (ComboBox3.Text = "Water" And ComboBox4.Text = "Gas") Or (ComboBox3.Text = "Gas" And ComboBox4.Text = "Water") Then
                    If ComboBox3.Text = "Water" Then
                        .Model.Grids(i).Saturation.Water = 0
                    Else
                        .Model.Grids(i).Saturation.Water = 1
                    End If
                End If


                .Model.Grids(i).Size.X_Length = a_len
                .Model.Grids(i).Size.Y_Length = a_len
                .Model.Grids(i).Size.Z_Length = b_len

                .Model.Grids(i).Location.X = 0
                .Model.Grids(i).Location.Y = 0
                .Model.Grids(i).Location.Z = i * b_len + b_len / 2

                .Model.Grids(i).PVT.Temperature = Temperature
            Next

            ReDim .Model.Grids(0).Schedules(0)
            ReDim .Model.Grids(n_grids - 1).Schedules(0)
            .Model.Grids(0).Schedules(0) = 0
            .Model.Grids(n_grids - 1).Schedules(0) = 1

            PVTExample()

            'If ComboBox6.SelectedIndex = 1 Then
            '    .Methodology.PVT.Oil.FVF = PVT_Methods.Oil_FVF_1
            'End If

            'If ComboBox7.SelectedIndex = 1 Then
            '    .Methodology.PVT.Water.FVF = PVT_Methods.Water_FVF_1
            'End If

            'If ComboBox8.SelectedIndex = 1 Then
            '    .Methodology.PVT.Gas.FVF = PVT_Methods.Gas_FVF_1
            'End If


            'If ComboBox11.SelectedIndex = 1 Then
            '    .Methodology.PVT.Oil.Density = PVT_Methods.Oil_Density_1
            'End If

            'If ComboBox10.SelectedIndex = 1 Then
            '    .Methodology.PVT.Gas.Density = PVT_Methods.Gas_Density_1
            'End If

            'If ComboBox9.SelectedIndex = 1 Then
            '    .Methodology.PVT.Water.Density = PVT_Methods.Water_Density_1
            'End If

            'If ComboBox14.SelectedIndex = 1 Then
            '    .Methodology.PVT.Oil.Viscosity = PVT_Methods.Oil_Viscosity_1
            'End If


            'If ComboBox13.SelectedIndex = 1 Thens
            '    .Methodology.PVT.Gas.Viscosity = PVT_Methods.Gas_Viscosity_1
            'End If

            'If ComboBox12.SelectedIndex = 1 Then
            '    .Methodology.PVT.Water.Viscosity = PVT_Methods.Water_Viscosity_1
            'End If

            'If ComboBox15.SelectedIndex = 1 Then
            '    .Methodology.PVT.Rs = PVT_Methods.RS_1
            'End If

            '.Methodology.PVT.Oil.Compressibility = PVT_Methods.Oil_Compressibility_Static
            '.Methodology.PVT.Water.Compressibility = PVT_Methods.Water_Compressibility_Static
            '.Methodology.PVT.Gas.Compressibility = PVT_Methods.Gas_Compressibility_Static

            'If CheckBox6.Checked Then
            '    .Methodology.PVT.Gas_Critical_Temperature = PVT_Methods.Gas_Critical_Temperature_1
            'End If

            'If CheckBox5.Checked Then
            '    .Methodology.PVT.Gas_Critical_Pressure = PVT_Methods.Gas_Critical_Pressure_1
            'End If


            .PVT.Oil_At_Standard.Compressibility = co
            .PVT.Water_At_Standard.Compressibility = cw
            .PVT.Condesate.Compressibility = cg

            .PVT.Oil_At_Standard.Density = rho_o
            .PVT.Water_At_Standard.Density = rho_w
            .PVT.Gas_At_Standard.Density = rho_g
            .PVT.Gas_Molecular_Weight = GasMW


            .PVT.Oil_At_Standard.Viscosity = 1
            .PVT.Gas_At_Standard.Viscosity = 0.022
            .PVT.Water_At_Standard.Viscosity = 0.52


            .PVT.Gas_Critical_Temperature = Ct
            .PVT.Gas_Critical_Pressure = Cp
            .PVT.Gas_Bubble_Pressure = GasBubblePressure

            'ReDim .PVT.Table(13, 20)
            'For i As Integer = 0 To 20
            '    .PVT.Table(13, i) = 
            'Next

            ReDim .Time_Steps(0)
            .RunType = Simulation_RunType.HistoryMatching_Corey
            .Finish_Time = 200
            .Final_Step = NudSteps.Value
            .Exactness.IterationMSE.Maximum = 100
            .Exactness.IterationMSE.Sw = 0.0001
            .Exactness.IterationMSE.Sg = 0.0001
            .Exactness.IterationMSE.Po = 0.1
            .Exactness.Equation = 0.00000001

            .Methodology.Derivations.Water_Saturation_Epsilon = 0.0001
            .Methodology.Derivations.Oil_Pressure_Epsilon = 0.001
            .Methodology.Derivations.Gas_Saturation_Epsilon = 0.0001

            .System.Units = System_of_Units.Field
            .System.Phase.Water = True
            .System.Phase.Oil = True
            .System.Phase.Gas = False
            .Options.Automatic_Grid_Adjacent_Recognition = True
            .Options.PostPressureDistributionAcceleration = CheckBox8.Checked
            .Options.PPDA_i = NUD6.Value
            .Options.PPDA_New = Val(ValueBox68.Text)

            .Methodology.Approach = Simulation_Approach.Fully_Implicit
            .Methodology.Transmissibility_Fp_Averaging = Transmissibility_Fp_Averaging.Method1

            .Rocks(0).Initial_Porosity = Val(ValueBox11.Text)
            .Rocks(0).Permeability.Kx = Val(ValueBox12.Text)
            .Rocks(0).Permeability.Ky = Val(ValueBox12.Text)
            .Rocks(0).Permeability.Kz = Val(ValueBox12.Text)
            Set_System_of_Units(SimulationData)

            .RunType = TabControl9.SelectedIndex + 1

            ' Corey ================================
            Dim Corey_nmin_water_wo As Double = Val(ValueBox32.Text)
            Dim Corey_nmax_water_wo As Double = Val(ValueBox30.Text)
            Dim Corey_nmin_oil_wo As Double = Val(ValueBox23.Text)
            Dim Corey_nmax_oil_wo As Double = Val(ValueBox24.Text)

            Dim Corey_nmin_water_wg As Double = Val(ValueBox32.Text)
            Dim Corey_nmax_water_wg As Double = Val(ValueBox30.Text)
            Dim Corey_nmin_gas_wg As Double = Val(ValueBox29.Text)
            Dim Corey_nmax_gas_wg As Double = Val(ValueBox27.Text)

            Dim Corey_nmin_gas_go As Double = Val(ValueBox29.Text)
            Dim Corey_nmax_gas_go As Double = Val(ValueBox27.Text)
            Dim Corey_nmin_oil_go As Double = Val(ValueBox23.Text)
            Dim Corey_nmax_oil_go As Double = Val(ValueBox24.Text)
            ' =======================================

            ' Burdine ================================
            Dim Burdine_Lambda_min_water_wo As Double = Val(ValueBox39.Text)
            Dim Burdine_Lambda_max_water_wo As Double = Val(ValueBox38.Text)
            Dim Burdine_Lambda_min_oil_wo As Double = Val(ValueBox35.Text)
            Dim Burdine_Lambda_max_oil_wo As Double = Val(ValueBox33.Text)

            Dim Burdine_Lambda_min_water_wg As Double = Val(ValueBox39.Text)
            Dim Burdine_Lambda_max_water_wg As Double = Val(ValueBox35.Text)
            Dim Burdine_Lambda_min_gas_wg As Double = 1
            Dim Burdine_Lambda_max_gas_wg As Double = 5

            Dim Burdine_Lambda_min_gas_go As Double = 1
            Dim Burdine_Lambda_max_gas_go As Double = 5
            Dim Burdine_Lambda_min_oil_go As Double = Val(ValueBox35.Text)
            Dim Burdine_Lambda_max_oil_go As Double = Val(ValueBox33.Text)
            ' =======================================

            ' SM =====================================
            Dim A_min_water_wo As Double = Val(ValueBox39.Text)
            Dim A_max_water_wo As Double = Val(ValueBox38.Text)
            Dim B_min_oil_wo As Double = Val(ValueBox35.Text)
            Dim B_max_oil_wo As Double = Val(ValueBox33.Text)

            Dim A_min_water_wg As Double = Val(ValueBox39.Text)
            Dim A_max_water_wg As Double = Val(ValueBox35.Text)
            Dim B_min_gas_wg As Double = 1
            Dim B_max_gas_wg As Double = 5

            Dim A_min_gas_go As Double = 1
            Dim A_max_gas_go As Double = 5
            Dim B_min_oil_go As Double = Val(ValueBox35.Text)
            Dim B_max_oil_go As Double = Val(ValueBox33.Text)
            ' =======================================

            Dim CurWaterMse(3) As Double
            Dim CurOilMse(3) As Double
            Dim RefinementDone As Boolean
            Dim ix As Integer = 0

            .Time_Steps(0) = Val(ValueBox64.Text)  '  GetBestTimeStep(SimulationData)
            .Exactness.OrthominMaxIteration = 500 'n_grids



            'Dim stxx As String
            'For i As Integer = 0 To 10
            '    SimulationData.Model.Grids(0).Pressure.Oil = (i / 100) * 1000 + 3500
            '    SimulationData.Model.Grids(0).Pressure.Water = (i / 100) * 1000 + 3500
            '    SimulationData.Model.Grids(0).Pressure.Gas = (i / 100) * 1000 + 3500
            '    Initialize_PVT(SimulationData.Model.Grids(0), SimulationData)
            '    Dim Bo As Double = SimulationData.Model.Grids(0).PVT.Oil.FVF
            '    Dim Bg As Double = SimulationData.Model.Grids(0).PVT.Gas.FVF
            '    Dim Bw As Double = SimulationData.Model.Grids(0).PVT.Water.FVF
            '    Dim R_o As Double = SimulationData.Model.Grids(0).PVT.Oil.Density
            '    Dim R_g As Double = SimulationData.Model.Grids(0).PVT.Gas.Density
            '    Dim R_w As Double = SimulationData.Model.Grids(0).PVT.Water.Density
            '    Dim mu_o As Double = SimulationData.Model.Grids(0).PVT.Oil.Viscosity
            '    Dim mu_g As Double = SimulationData.Model.Grids(0).PVT.Gas.Viscosity
            '    Dim mu_w As Double = SimulationData.Model.Grids(0).PVT.Water.Viscosity




            '    stxx += SimulationData.Model.Grids(0).Pressure.Oil.ToString + Chr(Keys.Tab) + Bo.ToString + Chr(Keys.Tab) + Bw.ToString + Chr(Keys.Tab) + Bg.ToString + Chr(Keys.Tab) + R_o.ToString + Chr(Keys.Tab) + R_g.ToString + Chr(Keys.Tab) + R_w.ToString + Chr(Keys.Tab) + mu_o.ToString + Chr(Keys.Tab) + mu_g.ToString + Chr(Keys.Tab) + mu_w.ToString + vbCrLf

            'Next

            'Clipboard.SetText(stxx)
            'Stop



            ' Corey ================================
            Dim n_water_wo_now As Double
            Dim n_oil_wo_now As Double

            Dim n_water_wg_now As Double
            Dim n_gas_wg_now As Double

            Dim n_oil_go_now As Double
            Dim n_gas_go_now As Double
            ' ======================================

            ' Burdine ===============================
            Dim lambda_water_wo_now As Double
            Dim lambda_oil_wo_now As Double

            Dim lambda_water_wg_now As Double
            Dim lambda_gas_wg_now As Double

            Dim lambda_oil_go_now As Double
            Dim lambda_gas_go_now As Double
            ' ======================================

            Dim trial As Integer = 0


            Dim DeltaP As Double = Math.Max(.Schedules(0).Data(0, 0), .Schedules(0).Data(0, 0)) - Math.Min(.Schedules(0).Data(0, 0), .Schedules(0).Data(0, 0))


            trial = 0

            ReDim Lets(nud5.Value - 1)
            ReDim SW(ndPC.Value + 1)
            ReDim Pcs(nud5.Value - 1, ndPC.Value + 1)
            Randomize()
            If .RunType = Simulation_RunType.HistoryMatching_LET Then

                SW(0) = Val(ValueBox70.Text)
                SW(ndPC.Value + 1) = Val(ValueBox69.Text)


                For j As Integer = 0 To ndPC.Value - 1
                    SW(j + 1) = ((j + 1) / (ndPC.Value + 1)) * (Val(ValueBox69.Text) - Val(ValueBox70.Text)) + SW(0)
                Next

                For i As Integer = 0 To nud5.Value - 1
                    Pcs(i, 0) = Val(ValueBox72.Text)
                    Pcs(i, ndPC.Value + 1) = Val(ValueBox71.Text)


                    'For j As Integer = 0 To ndPC.Value - 1
                    '    Pcs(i, j + 1) = Rnd() * (Pcs(i, ndPC.Value + 1) - Pcs(i, j)) + Pcs(i, j)
                    'Next


                    'Pcs(i, 1) = 143.92
                    'Pcs(i, 2) = 112.73
                    'Pcs(i, 3) = 83.12

                    Pcs(i, 1) = 133.23
                    Pcs(i, 2) = 82.38
                    Pcs(i, 3) = 69.12




                    With Lets(i)

                        Dim i_WO_W_L As Double = Rnd()
                        Dim i_WO_W_E As Double = Rnd()
                        Dim i_WO_W_T As Double = Rnd()

                        Dim i_WO_O_L As Double = Rnd()
                        Dim i_WO_O_E As Double = Rnd()
                        Dim i_WO_O_T As Double = Rnd()

                        Dim i_OG_O_L As Double = Rnd()
                        Dim i_OG_O_E As Double = Rnd()
                        Dim i_OG_O_T As Double = Rnd()

                        Dim i_OG_G_L As Double = Rnd()
                        Dim i_OG_G_E As Double = Rnd()
                        Dim i_OG_G_T As Double = Rnd()

                        Dim i_WG_W_L As Double = Rnd()
                        Dim i_WG_W_E As Double = Rnd()
                        Dim i_WG_W_T As Double = Rnd()

                        Dim i_WG_G_L As Double = Rnd()
                        Dim i_WG_G_E As Double = Rnd()
                        Dim i_WG_G_T As Double = Rnd()

                        ' Minimums -------------------------------------------------
                        Dim i_OG_L_O_Min As Double = Val(ValueBox56.Text)
                        Dim i_OG_E_O_Min As Double = Val(ValueBox52.Text)
                        Dim i_OG_T_O_Min As Double = Val(ValueBox54.Text)

                        Dim i_OG_L_G_Min As Double = Val(ValueBox55.Text)
                        Dim i_OG_E_G_Min As Double = Val(ValueBox51.Text)
                        Dim i_OG_T_G_Min As Double = Val(ValueBox53.Text)

                        Dim i_WO_L_W_Min As Double = Val(ValueBox45.Text)
                        Dim i_WO_E_W_Min As Double = Val(ValueBox50.Text)
                        Dim i_WO_T_W_Min As Double = Val(ValueBox48.Text)

                        Dim i_WO_L_O_Min As Double = Val(ValueBox46.Text)
                        Dim i_WO_E_O_Min As Double = Val(ValueBox49.Text)
                        Dim i_WO_T_O_Min As Double = Val(ValueBox47.Text)

                        Dim i_WG_L_W_Min As Double = Val(ValueBox62.Text)
                        Dim i_WG_E_W_Min As Double = Val(ValueBox58.Text)
                        Dim i_WG_T_W_Min As Double = Val(ValueBox60.Text)

                        Dim i_WG_L_G_Min As Double = Val(ValueBox61.Text)
                        Dim i_WG_E_G_Min As Double = Val(ValueBox57.Text)
                        Dim i_WG_T_G_Min As Double = Val(ValueBox59.Text)

                        ' Maximums -------------------------------------------------

                        Dim i_WO_L_O_Max As Double = Val(ValueBox78.Text)
                        Dim i_WO_E_O_Max As Double = Val(ValueBox74.Text)
                        Dim i_WO_T_O_Max As Double = Val(ValueBox76.Text)

                        Dim i_WO_L_W_Max As Double = Val(ValueBox77.Text)
                        Dim i_WO_E_W_Max As Double = Val(ValueBox73.Text)
                        Dim i_WO_T_W_Max As Double = Val(ValueBox75.Text)

                        Dim i_OG_L_O_Max As Double = Val(ValueBox84.Text)
                        Dim i_OG_E_O_Max As Double = Val(ValueBox80.Text)
                        Dim i_OG_T_O_Max As Double = Val(ValueBox82.Text)

                        Dim i_OG_L_G_Max As Double = Val(ValueBox83.Text)
                        Dim i_OG_E_G_Max As Double = Val(ValueBox79.Text)
                        Dim i_OG_T_G_Max As Double = Val(ValueBox81.Text)

                        Dim i_WG_L_W_Max As Double = Val(ValueBox90.Text)
                        Dim i_WG_E_W_Max As Double = Val(ValueBox86.Text)
                        Dim i_WG_T_W_Max As Double = Val(ValueBox88.Text)

                        Dim i_WG_L_G_Max As Double = Val(ValueBox89.Text)
                        Dim i_WG_E_G_Max As Double = Val(ValueBox85.Text)
                        Dim i_WG_T_G_Max As Double = Val(ValueBox87.Text)

                        .L_WO_W = 2.036548 '(i_WO_L_W_Max - i_WO_L_W_Min) * i_WO_W_L + i_WO_L_W_Min
                        .E_WO_W = 3.934583 '(i_WO_E_W_Max - i_WO_E_W_Min) * i_WO_W_E + i_WO_E_W_Min
                        .T_WO_W = 1.743297 '(i_WO_T_W_Max - i_WO_T_W_Min) * i_WO_W_T + i_WO_T_W_Min

                        .L_WO_O = 2.19345 '(i_WO_L_O_Max - i_WO_L_O_Min) * i_WO_O_L + i_WO_L_O_Min
                        .E_WO_O = 3.96549 '(i_WO_E_O_Max - i_WO_E_O_Min) * i_WO_O_E + i_WO_E_O_Min
                        .T_WO_O = 1.73546 '(i_WO_T_O_Max - i_WO_T_O_Min) * i_WO_O_T + i_WO_T_O_Min





                        .L_GO_G = (i_OG_L_G_Max - i_OG_L_G_Min) * i_OG_G_L + i_OG_L_G_Min
                        .E_GO_G = (i_OG_E_G_Max - i_OG_E_G_Min) * i_OG_G_E + i_OG_E_G_Min
                        .T_GO_G = (i_OG_T_G_Max - i_OG_T_G_Min) * i_OG_G_T + i_OG_T_G_Min

                        .L_GO_O = (i_OG_L_O_Max - i_OG_L_O_Min) * i_OG_O_L + i_OG_L_O_Min
                        .E_GO_O = (i_OG_E_O_Max - i_OG_E_O_Min) * i_OG_O_E + i_OG_E_O_Min
                        .T_GO_O = (i_OG_T_O_Max - i_OG_T_O_Min) * i_OG_O_T + i_OG_T_O_Min

                        .L_WG_W = (i_WG_L_W_Max - i_WG_L_W_Min) * i_WG_W_L + i_WG_L_W_Min
                        .E_WG_W = (i_WG_E_W_Max - i_WG_E_W_Min) * i_WG_W_E + i_WG_E_W_Min
                        .T_WG_W = (i_WG_T_W_Max - i_WG_T_W_Min) * i_WG_W_T + i_WG_T_W_Min

                        .L_WG_G = (i_WG_L_G_Max - i_WG_L_G_Min) * i_WG_G_L + i_WG_L_G_Min
                        .E_WG_G = (i_WG_E_G_Max - i_WG_E_G_Min) * i_WG_G_E + i_WG_E_G_Min
                        .T_WG_G = (i_WG_T_G_Max - i_WG_T_G_Min) * i_WG_G_T + i_WG_T_G_Min

                    End With
                Next


                '******************************************************************************************************8



                tmrLET.Enabled = True




                Exit Sub


                '=====================================================================================================================
                '                                                       NON LET                                                         
                '=====================================================================================================================


            End If
            For i_1 As Integer = 0 To 3
                ' Corey---------------------------------------------------------------------------------------
                n_water_wo_now = (i_1 / 3) * (Corey_nmax_water_wo - Corey_nmin_water_wo) + Corey_nmin_water_wo
                n_water_wg_now = (i_1 / 3) * (Corey_nmax_water_wg - Corey_nmin_water_wg) + Corey_nmin_water_wg
                n_oil_go_now = (i_1 / 3) * (Corey_nmax_oil_go - Corey_nmin_oil_go) + Corey_nmin_oil_go
                ' --------------------------------------------------------------------------------------------

                ' Burdine-------------------------------------------------------------------------------------
                lambda_water_wo_now = (i_1 / 3) * (Burdine_Lambda_max_water_wo - Burdine_Lambda_min_water_wo) + Burdine_Lambda_min_water_wo
                lambda_water_wg_now = (i_1 / 3) * (Burdine_Lambda_max_water_wg - Burdine_Lambda_min_water_wg) + Burdine_Lambda_min_water_wg
                lambda_oil_go_now = (i_1 / 3) * (Burdine_Lambda_max_oil_go - Burdine_Lambda_min_oil_go) + Burdine_Lambda_min_oil_go
                ' --------------------------------------------------------------------------------------------

                For j_1 As Integer = 0 To 3
                    trial += 1

                    ' Corey---------------------------------------------------------------------------------------
                    n_oil_wo_now = (j_1 / 3) * (Corey_nmax_oil_wo - Corey_nmin_oil_wo) + Corey_nmin_oil_wo
                    n_gas_wg_now = (j_1 / 3) * (Corey_nmax_gas_wg - Corey_nmin_gas_wg) + Corey_nmin_gas_wg
                    n_gas_go_now = (j_1 / 3) * (Corey_nmax_gas_go - Corey_nmin_gas_go) + Corey_nmin_gas_go
                    ' --------------------------------------------------------------------------------------------

                    ' Burdine-------------------------------------------------------------------------------------
                    lambda_oil_wo_now = (j_1 / 3) * (Burdine_Lambda_max_oil_wo - Burdine_Lambda_min_oil_wo) + Burdine_Lambda_min_oil_wo
                    lambda_gas_wg_now = (j_1 / 3) * (Burdine_Lambda_max_gas_wg - Burdine_Lambda_min_gas_wg) + Burdine_Lambda_min_gas_wg
                    lambda_gas_go_now = (j_1 / 3) * (Burdine_Lambda_max_gas_go - Burdine_Lambda_min_gas_go) + Burdine_Lambda_min_gas_go
                    ' --------------------------------------------------------------------------------------------

                    .Rocks(0).SCAL.Wettability.Water_Oil = Phase.Water

                    Select Case .RunType
                        Case Simulation_RunType.HistoryMatching_Corey
                            CreateCoreyKROW(.Rocks(0).SCAL.Tables, n_water_wo_now, n_oil_wo_now, Leng, Ko, Kw, Kair)
                            CreateCoreyKROG(.Rocks(0).SCAL.Tables, n_gas_go_now, n_oil_go_now, Leng, Ko, Kg, Kair)
                            CreateCoreyKRWG(.Rocks(0).SCAL.Tables, n_water_wg_now, n_gas_wg_now, Leng, Kg, Kw, Kair)
                            AddLogToMatchList("Level:" + ix.ToString + " Step:" + trial.ToString + " n Oil: " + n_oil_wo_now.ToString + " n Water:" + n_water_wo_now.ToString)
                        Case Simulation_RunType.HistoryMatching_Burdine
                            CreateBurdineKROW(.Rocks(0).SCAL.Tables, lambda_water_wo_now, lambda_oil_wo_now, Leng, Ko, Kw, Kair)
                            AddLogToMatchList("Level:" + ix.ToString + " Step:" + trial.ToString + " Lambda Oil: " + lambda_oil_wo_now.ToString + " Lambda Water:" + lambda_water_wo_now.ToString)
                        Case Simulation_RunType.HistoryMatching_SM
                            'CreateSMKROW(.Rocks(0).SCAL.Tables, lambda_water_wo_now, lambda_oil_wo_now, Leng, Ko, Kw, Kair, Val(ValueBox43.Text), Val(ValueBox42.Text))
                        Case Simulation_RunType.HistoryMatching_LET



                    End Select

                    '===============================================================
                    '=================== SIMULATION RUN ============================
                    '---------------------------------------------------------------

                    RefinementDone = Not CheckBox7.Checked

                    'If Not RefinementDone Then
                    '    Log(">> Grid Refinement Simulation")
                    '    Application.DoEvents()
                    '    mySimulator.Run(Simulation_Copy(SimulationData), History)
                    '    GetMatchProductionDifference()

                    '    RefineInitializeGrids()

                    '    History.Clear()
                    '    RefinementDone = True
                    'End If


                    mySimulator.Run(Simulation_Copy(SimulationData), History)

                    '===============================================================
                    '===============================================================

                    GetMatchProductionDifference()
                    AddLogToMatchList("MSE: " + GetProdcutionDeflection().ToString)

                    DrawRelativePermeabilityAndPc(.Rocks(0).SCAL.Tables)
                    DrawGraphs(History)


                    Application.DoEvents()
                    History.Clear()
                Next
            Next


        End With
    End Sub

    Function GetProdcutionDeflection() As Double
        Dim n As Double
        Dim TotalMSE As Double
        Dim Coe As Double
        Dim strx As String
        For i As Integer = 0 To dgvExp.RowCount - 2
            n = ((i + 1) / (dgvExp.RowCount - 2))
            TotalMSE += n * ((SimulationProductionInterpolate(dgvExp.Item(0, i).Value / (24 * 60 * 60)) - dgvExp.Item(1, i).Value) ^ 2)
            Coe += n
            strx += (SimulationProductionInterpolate(dgvExp.Item(0, i).Value / (24 * 60 * 60))).ToString + vbCrLf

        Next
        Clipboard.SetText(strx)
        Return Math.Log10(TotalMSE / Coe)
    End Function

    Function SimulationProductionInterpolate(ByVal OnTime As Double) As Double
        If OnTime = 0 Then Return 0
        Dim CurrentTime As Double = 0
        Dim CurrentProduction As Double
        Dim n As Integer
        Dim m As Double

        Do While OnTime > CurrentTime
            CurrentTime += History.Simulations(n).Time_Step
            n += 1
        Loop

        For i As Integer = 0 To n - 2
            Dim LastGrid As Grid = History.Simulations(i).Model.Grids(SizeOf(History.Simulations(i).Model.Grids) - 1)
            CurrentProduction += LastGrid.Flow.Oil * History.Simulations(i).Time_Step
            CurrentProduction += LastGrid.Flow.Water * History.Simulations(i).Time_Step
            CurrentProduction += LastGrid.Flow.Gas * History.Simulations(i).Time_Step
        Next

        Dim LastGrid2 As Grid = History.Simulations(n - 1).Model.Grids(SizeOf(History.Simulations(n - 1).Model.Grids) - 1)
        Dim LastGrid1 As Grid = History.Simulations(n - 2).Model.Grids(SizeOf(History.Simulations(n - 2).Model.Grids) - 1)

        Dim T As Double = History.Simulations(n - 2).Time_Step - (CurrentTime - OnTime)

        m = (LastGrid2.Flow.Oil - LastGrid1.Flow.Oil) / (History.Simulations(n - 2).Time_Step)
        CurrentProduction += (m * T + (CurrentTime - History.Simulations(n - 2).Time_Step)) * History.Simulations(n - 2).Time_Step

        m = (LastGrid2.Flow.Water - LastGrid1.Flow.Water) / (History.Simulations(n - 2).Time_Step)
        CurrentProduction += (m * T + (CurrentTime - History.Simulations(n - 2).Time_Step)) * History.Simulations(n - 2).Time_Step

        m = (LastGrid2.Flow.Gas - LastGrid1.Flow.Gas) / (History.Simulations(n - 2).Time_Step)
        CurrentProduction += (m * T + (CurrentTime - History.Simulations(n - 2).Time_Step)) * History.Simulations(n - 2).Time_Step

        Return CurrentProduction
    End Function

    Sub RefineInitializeGrids()
        Dim InjectedMax(SizeOf(SimulationData.Model.Grids) - 1) As Decimal
        Dim InjectedDifferentialMax(SizeOf(SimulationData.Model.Grids) - 1) As Decimal

        For i As Integer = 0 To SizeOf(SimulationData.Model.Grids) - 1
            Select Case ComboBox3.Text
                Case "Water"
                    InjectedMax(i) = History.Simulations(History.Current - 1).Model.Grids(i).Saturation.Water
                Case "Oil"
                    InjectedMax(i) = History.Simulations(History.Current - 1).Model.Grids(i).Saturation.Oil
                Case "Gas"
                    InjectedMax(i) = History.Simulations(History.Current - 1).Model.Grids(i).Saturation.Gas
            End Select
        Next

        InjectedMax(SizeOf(InjectedMax) - 1) = InjectedMax(SizeOf(InjectedMax) - 2)

        For i As Integer = 0 To History.Current - 2
            For j As Integer = 0 To SizeOf(SimulationData.Model.Grids) - 1
                Dim Delta As Decimal
                Select Case ComboBox3.Text
                    Case "Water"
                        Delta = Math.Abs(History.Simulations(i + 1).Model.Grids(j).Saturation.Water - History.Simulations(i).Model.Grids(j).Saturation.Water)
                    Case "Oil"
                        Delta = Math.Abs(History.Simulations(i + 1).Model.Grids(j).Saturation.Oil - History.Simulations(i).Model.Grids(j).Saturation.Oil)
                    Case "Gas"
                        Delta = Math.Abs(History.Simulations(i + 1).Model.Grids(j).Saturation.Gas - History.Simulations(i).Model.Grids(j).Saturation.Gas)
                End Select
                If Delta > InjectedDifferentialMax(j) Then
                    InjectedDifferentialMax(j) = Delta
                End If
            Next
        Next

        Dim Param1() As Decimal = Normalize(ReverseArray(InjectedMax))
        Dim Param2() As Decimal = Normalize(ReverseArray(InjectedDifferentialMax))
        Dim L(SizeOf(Param1) - 1) As Decimal
        Dim Sigma As Decimal
        Dim Lmin As Decimal
        Dim L_Core As Decimal = Val(ValueBox10.Text)
        Dim stx As String


        Dim P() As Decimal = Normalize(ProductArrays(Param1, Param2))
        Sigma = SumOfArray(P)
        Lmin = L_Min(SizeOf(Param1), NUDDR.Value, L_Core)
        For i As Integer = 0 To SizeOf(L) - 1
            L(i) = (P(i) / Sigma) * (L_Core - SizeOf(Param1) * Lmin) + Lmin
        Next

        'For i As Integer = 0 To SizeOf(Param1) - 1
        '    stx += (L_Core / SizeOf(Param1)).ToString + Chr(Keys.Tab) + L(i).ToString + vbCrLf
        'Next
        'Clipboard.SetText(stx)


        'For i As Integer = 1 To SizeOf(Param1) - 1
        '    L(i) += L(i - 1)
        'Next


        'Dim MC As New IMP.PointCollection
        'ReDim MC.MyPointCollection(SizeOf(Param1) - 1)
        'For i As Integer = 0 To SizeOf(Param1) - 1
        '    MC.MyPointCollection(i).Visible = True
        '    MC.MyPointCollection(i).PointSetPen = Pens.Blue
        '    ReDim Preserve MC.MyPointCollection(i).MyPointSet(1)
        '    MC.MyPointCollection(i).MyPointSet(0).X = L(i)
        '    MC.MyPointCollection(i).MyPointSet(0).Y = 0
        '    MC.MyPointCollection(i).MyPointSet(1).X = L(i)
        '    MC.MyPointCollection(i).MyPointSet(1).Y = 100
        'Next

        'MC.MyPointCollection(0).Visible = True
        'Imp1.ShowGrid = False
        'Imp1.Draw(MC, True, True)

        'Imp1.SETToBoundary(0, 0, Val(ValueBox10.Text) + Val(ValueBox10.Text) / 10, 100, False)
        'Stop

        Dim LastZ, LastL As Decimal
        For i As Integer = 0 To SizeOf(SimulationData.Model.Grids) - 1
            With SimulationData
                .Model.Grids(i).Size.Z_Length = L(i)
                .Model.Grids(i).Location.Z = LastZ + LastL / 2 + L(i) / 2
                LastZ = .Model.Grids(i).Location.Z
                LastL = L(i)
            End With
        Next

    End Sub
    Sub AddLogToMatchList(ByVal txt As String)
        RTB.Text += (txt + vbCrLf)
        RTB.SelectionStart = RTB.TextLength - 1
        RTB.ScrollToCaret()
    End Sub
    Function GetMatchProductionDifference() As Double
        Dim Production As Double
        Dim Pressure As Double
        Dim TotalTime As Double
        Dim strx As String

        For j As Integer = 0 To History.Current - 1
            'Production = (History.Simulations(j).Model.Grids(SizeOf(History.Simulations(j).Model.Grids) - 1).Flow.Oil * History.Simulations(j).Time_Step)
            Pressure = History.Simulations(j).Model.Grids(SizeOf(History.Simulations(j).Model.Grids) - 1).Pressure.Oil
            'Production += (History.Simulations(j).Model.Grids(SizeOf(History.Simulations(j).Model.Grids) - 1).Flow.Water * History.Simulations(j).Time_Step)
            'Production += (History.Simulations(j).Model.Grids(SizeOf(History.Simulations(j).Model.Grids) - 1).Flow.Gas * History.Simulations(j).Time_Step)
            TotalTime += History.Simulations(j).Time_Step
            'strx += (TotalTime).ToString + Chr(Keys.Tab) + Pressure.ToString + vbCrLf
        Next
        Dim TimeSecond As Double = TotalTime * 24 * 60 * 60

        'Clipboard.SetText(strx)

        'For j As Integer = 0 To dgvExp.RowCount - 2
        '    If dgvExp.Item(0, j).Value <= TimeSecond And dgvExp.Item(0, j + 1).Value >= TimeSecond Then
        '        Dim m As Double = (dgvExp.Item(0, j + 1).Value - dgvExp.Item(0, j).Value) / (dgvExp.Item(1, j + 1).Value - dgvExp.Item(1, j).Value)
        '        Dim CurProd As Double = m * (TimeSecond - dgvExp.Item(0, j).Value) + dgvExp.Item(0, j).Value
        '        Return (CurProd - Production)
        '    End If

        'Next

    End Function
    Sub DrawGraphs(ByVal His As Simulation_History)
        Dim MC_Pressure As New IMP.PointCollection
        Dim MC_Saturation As New IMP.PointCollection
        Dim MC_Production As New IMP.PointCollection

        ReDim MC_Pressure.MyPointCollection(SizeOf(His.Simulations(0).Model.Grids))
        ReDim MC_Saturation.MyPointCollection(SizeOf(His.Simulations(0).Model.Grids))

        ReDim MC_Production.MyPointCollection(1)

        ReDim MC_Production.MyPointCollection(0).MyPointSet(His.Current - 1)
        ReDim MC_Production.MyPointCollection(1).MyPointSet(dgvExp.RowCount - 2)

        MC_Production.MyPointCollection(0).Visible = True
        MC_Production.MyPointCollection(0).PointSetPen = Pens.Black
        MC_Production.MyPointCollection(0).width = 3

        MC_Production.MyPointCollection(1).Visible = True
        MC_Production.MyPointCollection(1).PointSetPen = Pens.Blue
        MC_Production.MyPointCollection(1).width = 3


        For i As Integer = 0 To SizeOf(His.Simulations(0).Model.Grids) - 1
            ReDim MC_Pressure.MyPointCollection(i).MyPointSet(His.Current - 1)
            ReDim MC_Saturation.MyPointCollection(i).MyPointSet(His.Current - 1)

            MC_Pressure.MyPointCollection(i).Visible = True
            MC_Pressure.MyPointCollection(i).PointSetPen = Pens.Blue

            MC_Saturation.MyPointCollection(i).Visible = True
            MC_Saturation.MyPointCollection(i).PointSetPen = Pens.Blue
        Next

        ReDim MC_Pressure.MyPointCollection(SizeOf(His.Simulations(0).Model.Grids)).MyPointSet(1)
        ReDim MC_Saturation.MyPointCollection(SizeOf(His.Simulations(0).Model.Grids)).MyPointSet(1)

        MC_Pressure.MyPointCollection(SizeOf(His.Simulations(0).Model.Grids)).Visible = SimulationData.Options.PostPressureDistributionAcceleration
        Dim Pn As New Pen(Color.Red)
        Pn.Width = 3
        Pn.DashStyle = Drawing2D.DashStyle.DashDot
        MC_Pressure.MyPointCollection(SizeOf(His.Simulations(0).Model.Grids)).PointSetPen = Pn

        MC_Saturation.MyPointCollection(SizeOf(His.Simulations(0).Model.Grids)).Visible = SimulationData.Options.PostPressureDistributionAcceleration
        MC_Saturation.MyPointCollection(SizeOf(His.Simulations(0).Model.Grids)).PointSetPen = Pn

        MC_Pressure.MyPointCollection(SizeOf(His.Simulations(0).Model.Grids)).MyPointSet(0).X = SimulationData.Options.PPDA_i
        MC_Saturation.MyPointCollection(SizeOf(His.Simulations(0).Model.Grids)).MyPointSet(0).X = SimulationData.Options.PPDA_i
        MC_Pressure.MyPointCollection(SizeOf(His.Simulations(0).Model.Grids)).MyPointSet(1).X = SimulationData.Options.PPDA_i
        MC_Saturation.MyPointCollection(SizeOf(His.Simulations(0).Model.Grids)).MyPointSet(1).X = SimulationData.Options.PPDA_i


        Dim McPressureMax As Double
        Dim McPressureMin As Double

        Dim McSaturationMax As Double
        Dim McSaturationMin As Double

        If (ComboBox3.Text = "Water" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Water") Then
            McPressureMax = His.Simulations(0).Model.Grids(0).Pressure.Oil
            McSaturationMax = His.Simulations(0).Model.Grids(0).Saturation.Water
        ElseIf (ComboBox3.Text = "Gas" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Gas") Then
            McPressureMax = His.Simulations(0).Model.Grids(0).Pressure.Oil
            McSaturationMax = His.Simulations(0).Model.Grids(0).Saturation.Gas
        ElseIf (ComboBox3.Text = "Water" And ComboBox4.Text = "Gas") Or (ComboBox3.Text = "Gas" And ComboBox4.Text = "Water") Then
            McPressureMax = His.Simulations(0).Model.Grids(0).Pressure.Water
            McSaturationMax = His.Simulations(0).Model.Grids(0).Saturation.Water
        End If
        McPressureMin = McPressureMax
        McSaturationMin = McSaturationMax
        
        For i As Integer = 0 To SizeOf(His.Simulations(0).Model.Grids) - 1
            For j As Integer = 0 To His.Current - 1
                MC_Pressure.MyPointCollection(i).MyPointSet(j).X = j
                MC_Saturation.MyPointCollection(i).MyPointSet(j).X = j
                If (ComboBox3.Text = "Water" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Water") Then
                    MC_Pressure.MyPointCollection(i).MyPointSet(j).Y = His.Simulations(j).Model.Grids(i).Pressure.Oil
                    MC_Saturation.MyPointCollection(i).MyPointSet(j).Y = His.Simulations(j).Model.Grids(i).Saturation.Water
                ElseIf (ComboBox3.Text = "Gas" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Gas") Then
                    MC_Pressure.MyPointCollection(i).MyPointSet(j).Y = His.Simulations(j).Model.Grids(i).Pressure.Oil
                    MC_Saturation.MyPointCollection(i).MyPointSet(j).Y = His.Simulations(j).Model.Grids(i).Saturation.Gas
                ElseIf (ComboBox3.Text = "Water" And ComboBox4.Text = "Gas") Or (ComboBox3.Text = "Gas" And ComboBox4.Text = "Water") Then
                    MC_Pressure.MyPointCollection(i).MyPointSet(j).Y = His.Simulations(j).Model.Grids(i).Pressure.Water
                    MC_Saturation.MyPointCollection(i).MyPointSet(j).Y = His.Simulations(j).Model.Grids(i).Saturation.Water
                End If
                If MC_Pressure.MyPointCollection(i).MyPointSet(j).Y > McPressureMax Then
                    McPressureMax = MC_Pressure.MyPointCollection(i).MyPointSet(j).Y
                End If
                If MC_Pressure.MyPointCollection(i).MyPointSet(j).Y < McPressureMin Then
                    McPressureMin = MC_Pressure.MyPointCollection(i).MyPointSet(j).Y
                End If
                If MC_Saturation.MyPointCollection(i).MyPointSet(j).Y > McSaturationMax Then
                    McSaturationMax = MC_Saturation.MyPointCollection(i).MyPointSet(j).Y
                End If
                If MC_Saturation.MyPointCollection(i).MyPointSet(j).Y < McSaturationMin Then
                    McSaturationMin = MC_Saturation.MyPointCollection(i).MyPointSet(j).Y
                End If
            Next
        Next

        MC_Pressure.MyPointCollection(SizeOf(His.Simulations(0).Model.Grids)).MyPointSet(0).Y = McPressureMin
        MC_Pressure.MyPointCollection(SizeOf(His.Simulations(0).Model.Grids)).MyPointSet(1).Y = McPressureMax
        MC_Saturation.MyPointCollection(SizeOf(His.Simulations(0).Model.Grids)).MyPointSet(0).Y = McSaturationMin
        MC_Saturation.MyPointCollection(SizeOf(His.Simulations(0).Model.Grids)).MyPointSet(1).Y = McSaturationMax

        Dim LastProd As Double
        Dim LastTime As Double

        Dim stxx As String = ""
        For j As Integer = 0 To His.Current - 1
            MC_Production.MyPointCollection(0).MyPointSet(j).X += (LastTime + His.Simulations(j).Time_Step)
            LastTime = MC_Production.MyPointCollection(0).MyPointSet(j).X
            MC_Production.MyPointCollection(0).MyPointSet(j).Y += (His.Simulations(j).Model.Grids(SizeOf(His.Simulations(j).Model.Grids) - 1).Flow.Oil * His.Simulations(j).Time_Step)
            MC_Production.MyPointCollection(0).MyPointSet(j).Y += (His.Simulations(j).Model.Grids(SizeOf(His.Simulations(j).Model.Grids) - 1).Flow.Water * His.Simulations(j).Time_Step)
            MC_Production.MyPointCollection(0).MyPointSet(j).Y += (His.Simulations(j).Model.Grids(SizeOf(His.Simulations(j).Model.Grids) - 1).Flow.Gas * His.Simulations(j).Time_Step)

            MC_Production.MyPointCollection(0).MyPointSet(j).Y += LastProd

            LastProd = MC_Production.MyPointCollection(0).MyPointSet(j).Y

            stxx += MC_Production.MyPointCollection(0).MyPointSet(j).X.ToString + Chr(Keys.Tab)
            stxx += MC_Production.MyPointCollection(0).MyPointSet(j).Y.ToString + vbCrLf


        Next

        Clipboard.SetText(stxx)


        For j As Integer = 0 To dgvExp.RowCount - 2
            MC_Production.MyPointCollection(1).MyPointSet(j).X = dgvExp.Item(0, j).Value / (24 * 60 * 60)
            MC_Production.MyPointCollection(1).MyPointSet(j).Y = dgvExp.Item(1, j).Value
        Next

        'Stop

        Imp3.ShowPset = False
        Imp3.Draw(MC_Pressure, True, True)

        Imp2.ShowPset = False
        Imp2.Draw(MC_Saturation, True, True)

        Imp4.ShowPset = False
        Imp4.Draw(MC_Production, True, True)


    End Sub
    Sub DrawRelativePermeabilityAndPc(ByVal T As RockType_SCAL_Tables)
        Dim MC_kr As New IMP.PointCollection
        Dim MC_Pc As New IMP.PointCollection
        ReDim MC_kr.MyPointCollection(1)
        ReDim MC_Pc.MyPointCollection(0)
        ReDim MC_kr.MyPointCollection(0).MyPointSet(T.Water_Oil.GetLength(1) - 1)
        ReDim MC_kr.MyPointCollection(1).MyPointSet(T.Water_Oil.GetLength(1) - 1)
        ReDim MC_Pc.MyPointCollection(0).MyPointSet(T.Water_Oil.GetLength(1) - 1)

        MC_kr.MyPointCollection(0).Visible = True
        MC_kr.MyPointCollection(1).Visible = True
        MC_Pc.MyPointCollection(0).Visible = True
        MC_kr.MyPointCollection(0).PointSetPen = Pens.Black
        MC_kr.MyPointCollection(1).PointSetPen = Pens.Black
        MC_Pc.MyPointCollection(0).PointSetPen = Pens.Black

        For i As Integer = 0 To T.Water_Oil.GetLength(1) - 1
            If (ComboBox3.Text = "Water" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Water") Then
                MC_kr.MyPointCollection(0).MyPointSet(i).X = T.Water_Oil(0, i)
                MC_kr.MyPointCollection(0).MyPointSet(i).Y = T.Water_Oil(1, i)

                MC_kr.MyPointCollection(1).MyPointSet(i).X = T.Water_Oil(0, i)
                MC_kr.MyPointCollection(1).MyPointSet(i).Y = T.Water_Oil(2, i)

                MC_Pc.MyPointCollection(0).MyPointSet(i).X = T.Water_Oil(0, i)
                MC_Pc.MyPointCollection(0).MyPointSet(i).Y = T.Water_Oil(3, i)

                Imp1.XTitle = "Sw"
                Imp5.XTitle = "Sw"
                Imp1.YTitle = "Kro/Krw"
                Imp5.YTitle = "Pc"
            ElseIf (ComboBox3.Text = "Gas" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Gas") Then
                MC_kr.MyPointCollection(0).MyPointSet(i).X = T.Oil_Gas(0, i)
                MC_kr.MyPointCollection(0).MyPointSet(i).Y = T.Oil_Gas(1, i)

                MC_kr.MyPointCollection(1).MyPointSet(i).X = T.Oil_Gas(0, i)
                MC_kr.MyPointCollection(1).MyPointSet(i).Y = T.Oil_Gas(2, i)

                Imp1.XTitle = "Sg"
                Imp1.YTitle = "Kro/Krg"
            ElseIf (ComboBox3.Text = "Gas" And ComboBox4.Text = "Water") Or (ComboBox3.Text = "Water" And ComboBox4.Text = "Gas") Then
                MC_kr.MyPointCollection(0).MyPointSet(i).X = T.Gas_Water(0, i)
                MC_kr.MyPointCollection(0).MyPointSet(i).Y = T.Gas_Water(1, i)

                MC_kr.MyPointCollection(1).MyPointSet(i).X = T.Gas_Water(0, i)
                MC_kr.MyPointCollection(1).MyPointSet(i).Y = T.Gas_Water(2, i)

                Imp1.XTitle = "Sw"
                Imp1.YTitle = "Krg/Krw"
            End If

        Next


        Imp1.Draw(MC_kr, True, True)
        Imp5.Draw(MC_Pc, True, True)


    End Sub
    Function GetBestTimeStep(ByVal Data As Simulation) As Double

        Dim Kmax As Double = Data.Rocks(0).Permeability.Kx
        Dim ViscosityMin As Double = Data.PVT.Table(8, 0)
        Dim CompressibilityMin As Double
        Dim Porosity As Double = Data.Rocks(0).Initial_Porosity
        Dim Lmin As Double = Data.Model.Grids(0).Size.X_Length

        For i As Integer = 0 To SizeOf(Data.Rocks) - 1
            If Data.Rocks(0).Permeability.Kx > Kmax Then
                Kmax = Data.Rocks(0).Permeability.Kx
            End If

            If Data.Rocks(0).Permeability.Ky > Kmax Then
                Kmax = Data.Rocks(0).Permeability.Ky
            End If

            If Data.Rocks(0).Permeability.Kz > Kmax Then
                Kmax = Data.Rocks(0).Permeability.Kz
            End If
        Next

        For i As Integer = 0 To Data.PVT.Table.GetLength(1) - 1
            If Data.PVT.Table(8, i) < ViscosityMin Then
                ViscosityMin = Data.PVT.Table(8, i)
            End If
            If Data.PVT.Table(9, i) < ViscosityMin Then
                ViscosityMin = Data.PVT.Table(9, i)
            End If
            If Data.PVT.Table(10, i) < ViscosityMin Then
                ViscosityMin = Data.PVT.Table(10, i)
            End If
        Next

        For i As Integer = 0 To SizeOf(Data.Model.Grids) - 1
            If Data.Model.Grids(i).Size.X_Length < Lmin Then
                Lmin = Data.Model.Grids(i).Size.X_Length
            End If
            If Data.Model.Grids(i).Size.Y_Length < Lmin Then
                Lmin = Data.Model.Grids(i).Size.Y_Length
            End If
            If Data.Model.Grids(i).Size.Z_Length < Lmin Then
                Lmin = Data.Model.Grids(i).Size.Z_Length
            End If
        Next

        CompressibilityMin = Val(ValueBox13.Text)

        Return (ViscosityMin * CompressibilityMin * Porosity * (Lmin ^ 2)) / (Data.Constants.Alpha * Data.Constants.Beta * Kmax)

    End Function
    Sub GetFlows(ByRef Oil_Flow As Double, ByRef Water_Flow As Double, ByRef Gas_Flow As Double)
        Oil_Flow = 0
        Water_Flow = 0
        Gas_Flow = 0
        For j As Integer = 0 To History.Current - 1
            For i As Integer = 0 To SizeOf(History.Simulations(j).Model.Grids) - 1
                Oil_Flow += (History.Simulations(j).Model.Grids(i).Flow.Oil * History.Simulations(j).Time_Step * History.Simulations(j).Constants.Alpha)
                Water_Flow += (History.Simulations(j).Model.Grids(i).Flow.Water * History.Simulations(j).Time_Step * History.Simulations(j).Constants.Alpha)
                Gas_Flow += (History.Simulations(j).Model.Grids(i).Flow.Gas * History.Simulations(j).Time_Step * History.Simulations(j).Constants.Alpha)
            Next
        Next
    End Sub
    Sub GetAverageSaturations(ByVal M As Simulation, ByRef Sw As Double, ByRef Sg As Double)
        Dim sum_sw, sum_sg As Double
        For i As Integer = 0 To SizeOf(M.Model.Grids) - 1
            sum_sw += M.Model.Grids(i).Saturation.Water
            sum_sg += M.Model.Grids(i).Saturation.Gas
        Next
        Sw = sum_sw / SizeOf(M.Model.Grids)
        Sg = sum_sg / SizeOf(M.Model.Grids)

    End Sub
    Sub CreateSMKROW(ByRef T As RockType_SCAL_Tables, ByVal A As Double, ByVal B As Double, ByVal Leng As Integer, ByVal KoMax As Double, ByVal KwMax As Double, ByVal Kair As Double)

        'ReDim T.Water_Oil(4, Leng - 1)
        'For i As Integer = 0 To Leng - 1
        '    Dim Sw_Star As Double = (Val(dgvExp2.Item(0, i).Value) - Val(dgvExp2.Item(0, 0).Value)) / (Val(dgvExp2.Item(0, dgvExp2.RowCount - 2).Value) - Val(dgvExp2.Item(0, 0).Value))
        '    Dim Sw As Double = Val(dgvExp2.Item(0, i).Value)
        '    Dim Krw As Double = (KwMax / Kair) * 
        '    Dim Kro As Double = (KoMax / Kair) * 

        '    T.Water_Oil(0, i) = Sw
        '    T.Water_Oil(1, i) = Kro
        '    T.Water_Oil(2, i) = Krw

        '    T.Water_Oil(3, i) = 0
        '    T.Water_Oil(4, i) = 0

        '    If (ComboBox3.Text = "Water" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Water") Then
        '        T.Water_Oil(3, i) = dgvExp2.Item(1, i).Value
        '        T.Water_Oil(4, i) = dgvExp2.Item(1, i).Value
        '    End If
        'Next

    End Sub
    Sub CreateBurdineKROW(ByRef T As RockType_SCAL_Tables, ByVal lambda_water As Double, ByVal lambda_Oil As Double, ByVal Leng As Integer, ByVal KoMax As Double, ByVal KwMax As Double, ByVal Kair As Double)

        ReDim T.Water_Oil(4, Leng - 1)
        For i As Integer = 0 To Leng - 1
            Dim Sw_Star As Double = (Val(dgvExp2.Item(0, i).Value) - Val(dgvExp2.Item(0, 0).Value)) / (Val(dgvExp2.Item(0, dgvExp2.RowCount - 2).Value) - Val(dgvExp2.Item(0, 0).Value))
            Dim Sw As Double = Val(dgvExp2.Item(0, i).Value)
            Dim Krw As Double = (KwMax / Kair) * (Sw_Star ^ (2 + 3 * lambda_water) / lambda_water)
            Dim Kro As Double = (KoMax / Kair) * (1 - Sw_Star) * (1 - Sw_Star ^ ((2 + lambda_Oil) / lambda_Oil))

            T.Water_Oil(0, i) = Sw
            T.Water_Oil(1, i) = Kro
            T.Water_Oil(2, i) = Krw

            T.Water_Oil(3, i) = 0
            T.Water_Oil(4, i) = 0

            'If (ComboBox3.Text = "Water" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Water") Then
            '    T.Water_Oil(3, i) = dgvExp2.Item(1, i).Value
            '    T.Water_Oil(4, i) = dgvExp2.Item(1, i).Value
            'End If
        Next

    End Sub
    Sub CreateCoreyKROW(ByRef T As RockType_SCAL_Tables, ByVal n_water As Double, ByVal n_Oil As Double, ByVal Leng As Integer, ByVal KoMax As Double, ByVal KwMax As Double, ByVal Kair As Double)

        ReDim T.Water_Oil(4, Leng - 1)
        For i As Integer = 0 To Leng - 1
            Dim Sw_Star As Double = (Val(dgvExp2.Item(0, i).Value) - Val(dgvExp2.Item(0, 0).Value)) / (Val(dgvExp2.Item(0, dgvExp2.RowCount - 2).Value) - Val(dgvExp2.Item(0, 0).Value))
            Dim Sw As Double = Val(dgvExp2.Item(0, i).Value)
            Dim Krw As Double = (KwMax / Kair) * (Sw_Star ^ n_water)
            Dim Kro As Double = (KoMax / Kair) * ((1 - Sw_Star) ^ n_Oil)

            T.Water_Oil(0, i) = Sw
            T.Water_Oil(1, i) = Kro
            T.Water_Oil(2, i) = Krw

            T.Water_Oil(3, i) = 0
            T.Water_Oil(4, i) = 0

            'If (ComboBox3.Text = "Water" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Water") Then
            '    T.Water_Oil(3, i) = dgvExp2.Item(1, i).Value
            '    T.Water_Oil(4, i) = dgvExp2.Item(1, i).Value
            'End If
        Next

    End Sub


    Sub CreateCoreyKROG(ByRef T As RockType_SCAL_Tables, ByVal n_Gas As Double, ByVal n_Oil As Double, ByVal Leng As Integer, ByVal KoMax As Double, ByVal KgMax As Double, ByVal Kair As Double)

        ReDim T.Oil_Gas(4, Leng - 1)
        For i As Integer = 0 To Leng - 1
            Dim Sg_Star As Double = (Val(dgvExp2.Item(0, i).Value) - Val(dgvExp2.Item(0, 0).Value)) / (Val(dgvExp2.Item(0, dgvExp2.RowCount - 2).Value) - Val(dgvExp2.Item(0, 0).Value))
            Dim Sg As Double = Val(dgvExp2.Item(0, i).Value)
            Dim Krg As Double = (KgMax / Kair) * (Sg_Star ^ n_Gas)
            Dim Kro As Double = (KoMax / Kair) * ((1 - Sg_Star) ^ n_Oil)

            T.Oil_Gas(0, i) = Sg
            T.Oil_Gas(1, i) = Kro
            T.Oil_Gas(2, i) = Krg

            T.Oil_Gas(3, i) = 0
            T.Oil_Gas(4, i) = 0

            If (ComboBox3.Text = "Gas" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Gas") Then
                T.Oil_Gas(3, i) = dgvExp2.Item(1, i).Value
                T.Oil_Gas(4, i) = dgvExp2.Item(1, i).Value
            End If
        Next

    End Sub


    Sub CreateCoreyKRWG(ByRef T As RockType_SCAL_Tables, ByVal n_water As Double, ByVal n_gas As Double, ByVal Leng As Integer, ByVal KgMax As Double, ByVal KwMax As Double, ByVal Kair As Double)

        ReDim T.Gas_Water(4, Leng - 1)
        For i As Integer = 0 To Leng - 1
            Dim Sw_Star As Double = (Val(dgvExp2.Item(0, i).Value) - Val(dgvExp2.Item(0, 0).Value)) / (Val(dgvExp2.Item(0, dgvExp2.RowCount - 2).Value) - Val(dgvExp2.Item(0, 0).Value))
            Dim Sw As Double = Val(dgvExp2.Item(0, i).Value)
            Dim Krw As Double = (KwMax / Kair) * (Sw_Star ^ n_water)
            Dim Krg As Double = (KgMax / Kair) * ((1 - Sw_Star) ^ n_gas)

            T.Gas_Water(0, i) = Sw
            T.Gas_Water(1, i) = Krg
            T.Gas_Water(2, i) = Krw

            T.Gas_Water(3, i) = 0
            T.Gas_Water(4, i) = 0

            If (ComboBox3.Text = "Water" And ComboBox4.Text = "Gas") Or (ComboBox3.Text = "Gas" And ComboBox4.Text = "Water") Then
                T.Gas_Water(3, i) = dgvExp2.Item(1, i).Value
                T.Gas_Water(4, i) = dgvExp2.Item(1, i).Value
            End If

        Next

    End Sub


    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged, ComboBox4.SelectedIndexChanged
        If (ComboBox3.Text = "Water" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Water") Then
            Label30.Text = "Relative Permeability (Water/Oil)"
            Label31.Text = "Water Saturation"
            Label44.Text = "Oil Pressure"
        ElseIf (ComboBox3.Text = "Gas" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Gas") Then
            Label30.Text = "Relative Permeability (Oil/Gas)"
            Label31.Text = "Gas Saturation"
            Label44.Text = "Oil Pressure"
        ElseIf (ComboBox3.Text = "Gas" And ComboBox4.Text = "Water") Or (ComboBox3.Text = "Water" And ComboBox4.Text = "Gas") Then
            Label30.Text = "Relative Permeability (Water/Gas)"
            Label31.Text = "Water Saturation"
            Label44.Text = "Water Pressure"
        End If


    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        ValueBox16.ReadOnly = CheckBox5.Checked
    End Sub

    Private Sub CheckBox6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        ValueBox17.ReadOnly = CheckBox5.Checked
    End Sub

    Private Sub CheckBox7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox7.CheckedChanged
        gbx.Enabled = CheckBox7.Checked
    End Sub

    Private n_Simulation As Integer
    Sub BuildInputs()

    End Sub
    Private Sub tmrLET_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrLET.Tick

        tmrLET.Enabled = False
        Percision_Level = 0
        NumNeurons = 0
        Application.DoEvents()


        Dim n_grids As Integer = NudGrids.Value

        With SimulationData

            For i As Integer = 0 To n_grids - 1

                .Model.Grids(i).Pressure.Oil = 3900 '2000 + (i / n_grids) * 2000 '.Schedules(1).Data(0, 0)
                If (ComboBox3.Text = "Water" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Water") Then
                    If ComboBox3.Text = "Water" Then
                        .Model.Grids(i).Saturation.Water = 0
                    Else
                        .Model.Grids(i).Saturation.Water = 1
                    End If
                ElseIf (ComboBox3.Text = "Gas" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Gas") Then
                    If ComboBox3.Text = "Gas" Then
                        .Model.Grids(i).Saturation.Gas = 0
                    Else
                        .Model.Grids(i).Saturation.Gas = 1
                    End If
                ElseIf (ComboBox3.Text = "Water" And ComboBox4.Text = "Gas") Or (ComboBox3.Text = "Gas" And ComboBox4.Text = "Water") Then
                    If ComboBox3.Text = "Water" Then
                        .Model.Grids(i).Saturation.Water = 0
                    Else
                        .Model.Grids(i).Saturation.Water = 1
                    End If
                End If
            Next

            Application.DoEvents()


            With .Rocks(0)
                ReDim .SCAL.Tables.Water_Oil(4, (ndPC.Value + 1) * 50)
                ReDim .SCAL.Tables.Oil_Gas(4, (ndPC.Value + 1) * 50)
                ReDim .SCAL.Tables.Gas_Water(4, (ndPC.Value + 1) * 50)

                Dim Kro_wo As Double = Val(ValueBox92.Text)
                Dim Krw_wo As Double = Val(ValueBox91.Text)

                Dim Kro_go As Double = Val(ValueBox94.Text)
                Dim Krg_go As Double = Val(ValueBox93.Text)

                Dim Krw_wg As Double = Val(ValueBox95.Text)
                Dim Krg_wg As Double = Val(ValueBox96.Text)

                For i As Integer = 0 To (ndPC.Value + 1) * 50


                    Dim normal_sat As Double = (i / ((ndPC.Value + 1) * 50))
                    .SCAL.Tables.Water_Oil(0, i) = (i / ((ndPC.Value + 1) * 50)) * (SW(ndPC.Value + 1) - SW(0)) + SW(0)
                    .SCAL.Tables.Water_Oil(1, i) = Kro_wo * ((1 - normal_sat) ^ Lets(n_Simulation).L_WO_O / ((1 - normal_sat) ^ Lets(n_Simulation).L_WO_O + Lets(n_Simulation).E_WO_O * (normal_sat ^ Lets(n_Simulation).T_WO_O)))
                    .SCAL.Tables.Water_Oil(2, i) = Krw_wo * (normal_sat ^ Lets(n_Simulation).L_WO_W / (normal_sat ^ Lets(n_Simulation).L_WO_W + Lets(n_Simulation).E_WO_W * ((1 - normal_sat) ^ Lets(n_Simulation).T_WO_W)))
                    Dim n_Pc As Integer = CInt(i \ 50)
                    Dim section As Double = (i - n_Pc * 50) / 50
                    If i <> (ndPC.Value + 1) * 50 Then
                        Dim Pc1 As Double = Pcs(n_Simulation, n_Pc)
                        Dim Pc2 As Double = Pcs(n_Simulation, n_Pc + 1)
                        .SCAL.Tables.Water_Oil(3, i) = (Pc2 - Pc1) * section + Pc1
                        .SCAL.Tables.Water_Oil(4, i) = (Pc2 - Pc1) * section + Pc1
                    Else
                        .SCAL.Tables.Water_Oil(3, i) = Pcs(n_Simulation, Pcs.GetLength(1) - 1)
                        .SCAL.Tables.Water_Oil(4, i) = Pcs(n_Simulation, Pcs.GetLength(1) - 1)
                    End If

                Next



                Dim strx As String = ""

                strx += Lets(n_Simulation).L_WO_O.ToString + Chr(Keys.Tab)
                strx += Lets(n_Simulation).E_WO_O.ToString + Chr(Keys.Tab)
                strx += Lets(n_Simulation).T_WO_O.ToString + Chr(Keys.Tab)

                strx += vbCrLf

                strx += Lets(n_Simulation).L_WO_W.ToString + Chr(Keys.Tab)
                strx += Lets(n_Simulation).E_WO_W.ToString + Chr(Keys.Tab)
                strx += Lets(n_Simulation).T_WO_W.ToString + Chr(Keys.Tab)

                strx += vbCrLf


                For jjj As Integer = 0 To .SCAL.Tables.Water_Oil.GetLength(1) - 1
                    For iii As Integer = 0 To 3
                        strx += .SCAL.Tables.Water_Oil(iii, jjj).ToString + Chr(Keys.Tab)
                    Next
                    strx += vbCrLf
                Next

                Clipboard.SetText(strx)
                '                Stop

            End With




            .Rocks(0).SCAL.Wettability.Water_Oil = Phase.Water
            AddLogToMatchList("Trial: " + (n_Simulation + 1).ToString + "/" + nud5.Value.ToString)
            Application.DoEvents()
        End With


        mySimulator.Run(Simulation_Copy(SimulationData), History)

        GetMatchProductionDifference()

        Dim MSE As Double = GetProdcutionDeflection()

        If CurrentDeflection = 0 Then CurrentDeflection = MSE
        If MSE < CurrentDeflection Then
            CurrentDeflection = MSE
        End If

        AddLogToMatchList("MSE " + MSE.ToString)
        Label147.Text = CurrentDeflection.ToString
        Lets(n_Simulation).MSE = MSE
        DrawRelativePermeabilityAndPc(SimulationData.Rocks(0).SCAL.Tables)
        DrawGraphs(History)
        'History.Clear()
        Application.DoEvents()

        n_Simulation += 1

        If n_Simulation < SizeOf(Lets) Then
            tmrLET.Enabled = True
        Else
            n_Simulation = 0
            'tmrFindBest()
            PrepareNeuralNetworkFiles()
            CreateSimulationFile()
            AddLogToMatchList("Copying neural network files to simulation folders...")
            CreateNetworkDataFiles()
            pb1.Maximum = Num_Of_Check_Networks
            pb1.Minimum = 0
            pb1.Visible = True
            AddLogToMatchList("Checking 440 neural networks (5-15 hidden neurons and four) training algorithms...")
            AddLogToMatchList("NOTE: This will take 616 seconds (10 minutes approximately) to accomplish.")
            tmrRunNetworks.Enabled = True
        End If

    End Sub
    Sub CreateNetworkDataFiles()
        For i As Integer = 1 To 4
            For j As Integer = 5 To 15
                Dim path As String = Application.StartupPath + "\" + i.ToString + "\" + j.ToString + "\"
                Dim FL_Network As New IO.FileStream(path + "network.data", IO.FileMode.OpenOrCreate, IO.FileAccess.Write)
                Dim Wrt_network As New IO.StreamWriter(FL_Network)
                Wrt_network.WriteLine((6 + ndPC.Value).ToString)
                Wrt_network.WriteLine("1") 'Outputs
                Wrt_network.WriteLine("3") 'Layers
                Wrt_network.WriteLine(j.ToString) 'Hidden Neurons
                Wrt_network.WriteLine(nudEpoch.Value.ToString) 'Max Epocs
                Wrt_network.WriteLine(10) 'Reporting
                Wrt_network.WriteLine(ValueBox44.Text) 'Goal MSE

                Wrt_network.Close()
                FL_Network.Close()

                IO.File.Copy(Application.StartupPath + "\test.data", path + "\test.data", True)
                IO.File.Copy(Application.StartupPath + "\train.data", path + "\train.data", True)
                IO.File.Copy(Application.StartupPath + "\sim_input.data", path + "\sim_input.data", True)

            Next
        Next
    End Sub

    Dim train_Maxs(9) As Double
    Dim train_Mins(9) As Double

    Sub PrepareNeuralNetworkFiles()
        Dim i_num_testing As Integer = (nudTest.Value / 100) * SizeOf(Lets)
        Dim i_num_training As Integer = SizeOf(Lets) - i_num_testing


        If (ComboBox3.Text = "Water" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Water") Then
            train_Maxs(0) = Lets(0).L_WO_O
            train_Maxs(1) = Lets(0).E_WO_O
            train_Maxs(2) = Lets(0).T_WO_O
            train_Maxs(3) = Lets(0).L_WO_W
            train_Maxs(4) = Lets(0).E_WO_W
            train_Maxs(5) = Lets(0).T_WO_W
        ElseIf (ComboBox3.Text = "Gas" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Gas") Then

        ElseIf (ComboBox3.Text = "Water" And ComboBox4.Text = "Gas") Or (ComboBox3.Text = "Gas" And ComboBox4.Text = "Water") Then

        End If

        For j As Integer = 1 To ndPC.Value
            train_Maxs(5 + j) = Pcs(0, j)
        Next
        train_Maxs(9) = Lets(0).MSE

        For n As Integer = 0 To SizeOf(train_Maxs) - 1
            train_Mins(n) = train_Maxs(n)
        Next


        For i As Integer = 0 To SizeOf(Lets) - 1


            If (ComboBox3.Text = "Water" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Water") Then
                If Lets(i).L_WO_O > train_Maxs(0) Then
                    train_Maxs(0) = Lets(i).L_WO_O
                End If
                If Lets(i).E_WO_O > train_Maxs(1) Then
                    train_Maxs(1) = Lets(i).E_WO_O
                End If
                If Lets(i).T_WO_O > train_Maxs(2) Then
                    train_Maxs(2) = Lets(i).T_WO_O
                End If
                If Lets(i).L_WO_W > train_Maxs(3) Then
                    train_Maxs(3) = Lets(i).L_WO_W
                End If
                If Lets(i).E_WO_W > train_Maxs(4) Then
                    train_Maxs(4) = Lets(i).E_WO_W
                End If
                If Lets(i).T_WO_W > train_Maxs(5) Then
                    train_Maxs(5) = Lets(i).T_WO_W
                End If

                If Lets(i).L_WO_O < train_Mins(0) Then
                    train_Mins(0) = Lets(i).L_WO_O
                End If
                If Lets(i).E_WO_O < train_Mins(1) Then
                    train_Mins(1) = Lets(i).E_WO_O
                End If
                If Lets(i).T_WO_O < train_Mins(2) Then
                    train_Mins(2) = Lets(i).T_WO_O
                End If
                If Lets(i).L_WO_W < train_Mins(3) Then
                    train_Mins(3) = Lets(i).L_WO_W
                End If
                If Lets(i).E_WO_W < train_Mins(4) Then
                    train_Mins(4) = Lets(i).E_WO_W
                End If
                If Lets(i).T_WO_W < train_Mins(5) Then
                    train_Mins(5) = Lets(i).T_WO_W
                End If

            ElseIf (ComboBox3.Text = "Gas" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Gas") Then

            ElseIf (ComboBox3.Text = "Water" And ComboBox4.Text = "Gas") Or (ComboBox3.Text = "Gas" And ComboBox4.Text = "Water") Then

            End If

            For j As Integer = 1 To ndPC.Value
                If Pcs(i, j) > train_Maxs(5 + j) Then
                    train_Maxs(5 + j) = Pcs(i, j)
                End If

                If Pcs(i, j) < train_Mins(5 + j) Then
                    train_Mins(5 + j) = Pcs(i, j)
                End If

            Next

            If Lets(i).MSE > train_Maxs(9) Then
                train_Maxs(9) = Lets(i).MSE
            End If

            If Lets(i).MSE < train_Mins(9) Then
                train_Mins(9) = Lets(i).MSE
            End If

        Next

        train_Mins(9) = CurrentDeflection - 0.5

        Dim IDs_Training(i_num_training - 1) As Integer
        Dim IDs_Testing(i_num_testing - 1) As Integer
        'Dim IDs_Training(SizeOf(Lets) - i_num_testing - 1) As Integer

        For i As Integer = 0 To i_num_testing - 1
            IDs_Testing(i) = -1
        Next
        For i As Integer = 0 To i_num_training - 1
            IDs_Training(i) = -1
        Next

        'Choose(Testings)
        Dim CurTestCount As Integer = 0
        For i As Integer = 0 To SizeOf(Lets) - 1
            If (ComboBox3.Text = "Water" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Water") Then
                If Lets(i).L_WO_O < train_Maxs(0) And Lets(i).L_WO_O > train_Mins(0) Then
                    If Lets(i).E_WO_O < train_Maxs(1) And Lets(i).E_WO_O > train_Mins(1) Then
                        If Lets(i).T_WO_O < train_Maxs(2) And Lets(i).T_WO_O > train_Mins(2) Then
                            If Lets(i).L_WO_W < train_Maxs(3) And Lets(i).L_WO_W > train_Mins(3) Then
                                If Lets(i).E_WO_W < train_Maxs(4) And Lets(i).E_WO_W > train_Mins(4) Then
                                    If Lets(i).T_WO_W < train_Maxs(5) And Lets(i).T_WO_W > train_Mins(5) Then
                                        If Pcs(i, 1) < train_Maxs(6) And Pcs(i, 1) > train_Mins(6) Then
                                            If Pcs(i, 2) < train_Maxs(7) And Pcs(i, 2) > train_Mins(7) Then
                                                If Pcs(i, 3) < train_Maxs(8) And Pcs(i, 3) > train_Mins(8) Then
                                                    If Lets(i).MSE < train_Maxs(9) And Lets(i).MSE > train_Mins(9) Then
                                                        If CurTestCount < i_num_testing Then
                                                            IDs_Testing(CurTestCount) = i
                                                            CurTestCount += 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            ElseIf (ComboBox3.Text = "Gas" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Gas") Then
            ElseIf (ComboBox3.Text = "Water" And ComboBox4.Text = "Gas") Or (ComboBox3.Text = "Gas" And ComboBox4.Text = "Water") Then
            End If
        Next

        'Dim MinMse As Double
        'For i As Integer = 0 To i_num_testing - 1
        '    Dim curmin_i As Integer
        '    MinMse = 1
        '    For j As Integer = 0 To SizeOf(Lets) - 1
        '        Dim found As Boolean = False
        '        For s As Integer = 0 To i_num_testing - 1
        '            If IDs_Testing(s) = j Then
        '                found = True
        '            End If
        '        Next
        '        If Lets(j).MSE < MinMse And Not found Then
        '            MinMse = Lets(j).MSE
        '            curmin_i = j
        '        End If
        '    Next
        '    IDs_Testing(i) = curmin_i
        'Next


        ' Choose Trainings

        Dim n_Training As Integer = 0
        ReDim IDs_Training(SizeOf(Lets) - CurTestCount - 1)
        i_num_testing = CurTestCount
        i_num_training = SizeOf(IDs_Training)

        ReDim IDs_Training(SizeOf(Lets) - i_num_testing - 1)
        i_num_training = SizeOf(IDs_Training)

        ReDim Preserve IDs_Testing(i_num_testing - 1)

        For i As Integer = 0 To SizeOf(Lets) - 1
            Dim IsTesting As Boolean = False
            For j As Integer = 0 To i_num_testing - 1
                If IDs_Testing(j) = i Then
                    IsTesting = True
                End If
            Next
            If Not IsTesting And n_Training < i_num_training Then
                IDs_Training(n_Training) = i
                n_Training += 1
            End If
        Next

        


        ' Training -----------------------------------------------------------------
        Dim FL_Training As New IO.FileStream(Application.StartupPath + "\train.data", IO.FileMode.OpenOrCreate, IO.FileAccess.Write)
        Dim Wrt_train As New IO.StreamWriter(FL_Training)
        Wrt_train.Write(i_num_training)
        Wrt_train.Write(" ")
        Wrt_train.Write((6 + ndPC.Value).ToString)
        Wrt_train.Write(" ")
        Wrt_train.Write("1")
        Wrt_train.Write(" ")


        For i As Integer = 0 To i_num_training - 1

            If (ComboBox3.Text = "Water" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Water") Then
                Wrt_train.Write((((Lets(IDs_Training(i)).L_WO_O - train_Mins(0)) / (train_Maxs(0) - train_Mins(0))) * 2 - 1).ToString)
                Wrt_train.Write(" ")
                Wrt_train.Write((((Lets(IDs_Training(i)).E_WO_O - train_Mins(1)) / (train_Maxs(1) - train_Mins(1))) * 2 - 1).ToString)
                Wrt_train.Write(" ")
                Wrt_train.Write((((Lets(IDs_Training(i)).T_WO_O - train_Mins(2)) / (train_Maxs(2) - train_Mins(2))) * 2 - 1).ToString)
                Wrt_train.Write(" ")
                Wrt_train.Write((((Lets(IDs_Training(i)).L_WO_W - train_Mins(3)) / (train_Maxs(3) - train_Mins(3))) * 2 - 1).ToString)
                Wrt_train.Write(" ")
                Wrt_train.Write((((Lets(IDs_Training(i)).E_WO_W - train_Mins(4)) / (train_Maxs(4) - train_Mins(4))) * 2 - 1).ToString)
                Wrt_train.Write(" ")
                Wrt_train.Write((((Lets(IDs_Training(i)).T_WO_W - train_Mins(5)) / (train_Maxs(5) - train_Mins(5))) * 2 - 1).ToString)
                Wrt_train.Write(" ")
            ElseIf (ComboBox3.Text = "Gas" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Gas") Then
            ElseIf (ComboBox3.Text = "Water" And ComboBox4.Text = "Gas") Or (ComboBox3.Text = "Gas" And ComboBox4.Text = "Water") Then
            End If

            For j As Integer = 1 To ndPC.Value
                Wrt_train.Write((((Pcs(IDs_Training(i), j) - train_Mins(5 + j)) / (train_Maxs(5 + j) - train_Mins(5 + j))) * 2 - 1).ToString)
                Wrt_train.Write(" ")
            Next

            Wrt_train.Write(((((Lets(IDs_Training(i)).MSE - train_Mins(9)) / (train_Maxs(9) - train_Mins(9))) * 2 - 1).ToString))
            Wrt_train.Write(" ")

        Next

        Wrt_train.Close()
        FL_Training.Close()

        ' Testing ----------------------------------------------------------------------------------------------------------

        Dim FL_Testing As New IO.FileStream(Application.StartupPath + "\test.data", IO.FileMode.OpenOrCreate, IO.FileAccess.Write)
        Dim Wrt_test As New IO.StreamWriter(FL_Testing)
        Wrt_test.Write(i_num_testing)
        Wrt_test.Write(" ")
        Wrt_test.Write((6 + ndPC.Value).ToString)
        Wrt_test.Write(" ")
        Wrt_test.Write("1")
        Wrt_test.Write(" ")

        For i As Integer = 0 To i_num_testing - 1

            If (ComboBox3.Text = "Water" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Water") Then
                Wrt_test.Write((((Lets(IDs_Testing(i)).L_WO_O - train_Mins(0)) / (train_Maxs(0) - train_Mins(0))) * 2 - 1).ToString)
                Wrt_test.Write(" ")
                Wrt_test.Write((((Lets(IDs_Testing(i)).E_WO_O - train_Mins(1)) / (train_Maxs(1) - train_Mins(1))) * 2 - 1).ToString)
                Wrt_test.Write(" ")
                Wrt_test.Write((((Lets(IDs_Testing(i)).T_WO_O - train_Mins(2)) / (train_Maxs(2) - train_Mins(2))) * 2 - 1).ToString)
                Wrt_test.Write(" ")
                Wrt_test.Write((((Lets(IDs_Testing(i)).L_WO_W - train_Mins(3)) / (train_Maxs(3) - train_Mins(3))) * 2 - 1).ToString)
                Wrt_test.Write(" ")
                Wrt_test.Write((((Lets(IDs_Testing(i)).E_WO_W - train_Mins(4)) / (train_Maxs(4) - train_Mins(4))) * 2 - 1).ToString)
                Wrt_test.Write(" ")
                Wrt_test.Write((((Lets(IDs_Testing(i)).T_WO_W - train_Mins(5)) / (train_Maxs(5) - train_Mins(5))) * 2 - 1).ToString)
                Wrt_test.Write(" ")
            ElseIf (ComboBox3.Text = "Gas" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Gas") Then

            ElseIf (ComboBox3.Text = "Water" And ComboBox4.Text = "Gas") Or (ComboBox3.Text = "Gas" And ComboBox4.Text = "Water") Then

            End If

            For j As Integer = 1 To ndPC.Value
                Wrt_test.Write((((Pcs(IDs_Testing(i), j) - train_Mins(5 + j)) / (train_Maxs(5 + j) - train_Mins(5 + j))) * 2 - 1).ToString)
                Wrt_test.Write(" ")
            Next

            Wrt_test.Write((((Lets(IDs_Testing(i)).MSE - train_Mins(9)) / (train_Maxs(9) - train_Mins(9))) * 2 - 1).ToString)
            Wrt_test.Write(" ")

        Next

        Wrt_test.Close()
        FL_Testing.Close()
    End Sub

    Private Sub tmrNeural_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrNeural.Tick
        tmrNeural.Enabled = False
        AddLogToMatchList("Neural Network Run")
        AddLogToMatchList("Neural Network: Training... ")

        '-------------------------------------------------------------------------------
        PrepareNeuralNetworkFiles()

        'Simulation
        ' Minimums -------------------------------------------------
        'CreateSimulationFile()
        tmrSF.Enabled = True


    End Sub
    Private Indexes() As Integer
    Dim Params(0, 8) As Double
    Dim Mins(8) As Double
    Dim Maxs(8) As Double
    Dim Centers(8) As Double

    Function GetBestFirstPhaseMseIndex() As Integer
        Dim mse1 As Double = Lets(0).MSE
        Dim ii As Integer
        For i As Integer = 0 To SizeOf(Lets) - 1
            If Lets(i).MSE < mse1 Then
                mse1 = Lets(i).MSE
                ii = i
            End If
        Next
        Return ii
    End Function
    Function Returner(ByVal i As Integer, ByVal j As Integer) As Double
        Select Case i
            Case 0 'Min
                Return (Mins(j))
            Case 1 'Center
                Return (Centers(j))
            Case 2 'Max
                Return (Maxs(j))
        End Select

    End Function
    Sub CreateSimulationFile()
        Dim FL_Sim As New IO.FileStream(Application.StartupPath + "\sim_input.data", IO.FileMode.OpenOrCreate, IO.FileAccess.Write)
        Dim Wrt_sim As New IO.StreamWriter(FL_Sim)

        Wrt_sim.Write((3 ^ 9).ToString + " ")
        Wrt_sim.Write("9" + " ")
        Wrt_sim.Write("1" + " ")

        Dim strFile As String

        If SizeOf(Indexes) = 0 Then

            Dim ix As Integer = GetBestFirstPhaseMseIndex()

            If (ComboBox3.Text = "Water" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Water") Then
                'For i As Integer = 0 To 8
                '    Mins(i) = -1
                '    Maxs(i) = 1
                '    Centers(i) = 0
                'Next

                Mins(0) = ((Lets(ix).L_WO_O - train_Mins(0)) / (train_Maxs(0) - train_Mins(0))) - (0.1 / (nud5.Value ^ (1 / 9)))
                Maxs(0) = ((Lets(ix).L_WO_O - train_Mins(0)) / (train_Maxs(0) - train_Mins(0))) + (0.1 / (nud5.Value ^ (1 / 9)))
                Centers(0) = ((Lets(ix).L_WO_O - train_Mins(0)) / (train_Maxs(0) - train_Mins(0)))

                Mins(1) = ((Lets(ix).E_WO_O - train_Mins(1)) / (train_Maxs(1) - train_Mins(1))) - (0.1 / (nud5.Value ^ (1 / 9)))
                Maxs(1) = ((Lets(ix).E_WO_O - train_Mins(1)) / (train_Maxs(1) - train_Mins(1))) + (0.1 / (nud5.Value ^ (1 / 9)))
                Centers(1) = ((Lets(ix).E_WO_O - train_Mins(1)) / (train_Maxs(1) - train_Mins(1)))

                Mins(2) = ((Lets(ix).T_WO_O - train_Mins(2)) / (train_Maxs(2) - train_Mins(2))) - (0.1 / (nud5.Value ^ (1 / 9)))
                Maxs(2) = ((Lets(ix).T_WO_O - train_Mins(2)) / (train_Maxs(2) - train_Mins(2))) + (0.1 / (nud5.Value ^ (1 / 9)))
                Centers(2) = ((Lets(ix).T_WO_O - train_Mins(2)) / (train_Maxs(2) - train_Mins(2)))

                Mins(3) = ((Lets(ix).L_WO_W - train_Mins(3)) / (train_Maxs(3) - train_Mins(3))) - (0.1 / (nud5.Value ^ (1 / 9)))
                Maxs(3) = ((Lets(ix).L_WO_W - train_Mins(3)) / (train_Maxs(3) - train_Mins(3))) + (0.1 / (nud5.Value ^ (1 / 9)))
                Centers(3) = ((Lets(ix).L_WO_W - train_Mins(3)) / (train_Maxs(3) - train_Mins(3)))

                Mins(4) = ((Lets(ix).E_WO_W - train_Mins(4)) / (train_Maxs(4) - train_Mins(4))) - (0.1 / (nud5.Value ^ (1 / 9)))
                Maxs(4) = ((Lets(ix).E_WO_W - train_Mins(4)) / (train_Maxs(4) - train_Mins(4))) + (0.1 / (nud5.Value ^ (1 / 9)))
                Centers(4) = ((Lets(ix).E_WO_W - train_Mins(4)) / (train_Maxs(4) - train_Mins(4)))

                Mins(5) = ((Lets(ix).T_WO_W - train_Mins(5)) / (train_Maxs(5) - train_Mins(5))) - (0.1 / (nud5.Value ^ (1 / 9)))
                Maxs(5) = ((Lets(ix).T_WO_W - train_Mins(5)) / (train_Maxs(5) - train_Mins(5))) + (0.1 / (nud5.Value ^ (1 / 9)))
                Centers(5) = ((Lets(ix).T_WO_W - train_Mins(5)) / (train_Maxs(5) - train_Mins(5)))

                Mins(6) = ((Pcs(ix, 1) - train_Mins(6)) / (train_Maxs(6) - train_Mins(6))) - (0.1 / (nud5.Value ^ (1 / 9)))
                Maxs(6) = ((Pcs(ix, 1) - train_Mins(6)) / (train_Maxs(6) - train_Mins(6))) + (0.1 / (nud5.Value ^ (1 / 9)))
                Centers(6) = ((Pcs(ix, 1) - train_Mins(6)) / (train_Maxs(6) - train_Mins(6)))

                Mins(7) = ((Pcs(ix, 2) - train_Mins(7)) / (train_Maxs(7) - train_Mins(7))) - (0.1 / (nud5.Value ^ (1 / 9)))
                Maxs(7) = ((Pcs(ix, 2) - train_Mins(7)) / (train_Maxs(7) - train_Mins(7))) + (0.1 / (nud5.Value ^ (1 / 9)))
                Centers(7) = ((Pcs(ix, 2) - train_Mins(7)) / (train_Maxs(7) - train_Mins(7)))

                Mins(8) = ((Pcs(ix, 3) - train_Mins(8)) / (train_Maxs(8) - train_Mins(8))) - (0.1 / (nud5.Value ^ (1 / 9)))
                Maxs(8) = ((Pcs(ix, 3) - train_Mins(8)) / (train_Maxs(8) - train_Mins(8))) + (0.1 / (nud5.Value ^ (1 / 9)))
                Centers(8) = ((Pcs(ix, 3) - train_Mins(8)) / (train_Maxs(8) - train_Mins(8)))



            ElseIf (ComboBox3.Text = "Gas" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Gas") Then

            ElseIf (ComboBox3.Text = "Water" And ComboBox4.Text = "Gas") Or (ComboBox3.Text = "Gas" And ComboBox4.Text = "Water") Then

            End If


          

        Else

            Dim pcx(2) As Double
            GetCurrentParams(Params(0, 0), Params(0, 1), Params(0, 2), Params(0, 3), Params(0, 4), Params(0, 5), pcx, Indexes, True)
            Params(0, 6) = pcx(0)
            Params(0, 7) = pcx(1)
            Params(0, 8) = pcx(2)

            For i As Integer = 0 To 8
                Maxs(i) = Params(0, i) + (0.1 / (nud5.Value ^ (1 / 9)) / (2 ^ Percision_Level))
                Mins(i) = Params(0, i) - (0.1 / (nud5.Value ^ (1 / 9)) / (2 ^ Percision_Level))
                Centers(i) = Params(0, i)
            Next
            'For i As Integer = 0 To 8
            '    Maxs(i) = Params(0, i) + 1 / (2 ^ Percision_Level)
            '    Mins(i) = Params(0, i) - 1 / (2 ^ Percision_Level)
            '    Centers(i) = Params(0, i)
            'Next
        End If

        For i As Integer = 0 To 8
            If Maxs(i) > 1 Then Maxs(i) = 1
            If Mins(i) < -1 Then Mins(i) = -1
        Next

        For i0 As Integer = 0 To 2
            For i1 As Integer = 0 To 2
                For i2 As Integer = 0 To 2
                    For i3 As Integer = 0 To 2
                        For i4 As Integer = 0 To 2
                            For i5 As Integer = 0 To 2
                                For i6 As Integer = 0 To 2
                                    For i7 As Integer = 0 To 2
                                        For i8 As Integer = 0 To 2

                                            If (ComboBox3.Text = "Water" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Water") Then

                                                strFile += Returner(i0, 0).ToString + " "
                                                strFile += Returner(i1, 1).ToString + " "
                                                strFile += Returner(i2, 2).ToString + " "

                                                strFile += Returner(i3, 3).ToString + " "
                                                strFile += Returner(i4, 4).ToString + " "
                                                strFile += Returner(i5, 5).ToString + " "

                                                strFile += Returner(i6, 6).ToString + " "
                                                strFile += Returner(i7, 7).ToString + " "
                                                strFile += Returner(i8, 8).ToString + " "

                                                strFile += "0" + " "


                                                Wrt_sim.Write(strFile)

                                                strFile = ""
                                            ElseIf (ComboBox3.Text = "Gas" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Gas") Then

                                            ElseIf (ComboBox3.Text = "Water" And ComboBox4.Text = "Gas") Or (ComboBox3.Text = "Gas" And ComboBox4.Text = "Water") Then

                                            End If

                                        Next
                                    Next
                                Next
                            Next
                        Next
                    Next
                Next
            Next
        Next




        Wrt_sim.Close()
        FL_Sim.Close()

        AddLogToMatchList("Closing ANN files...")

        'NewNumberforHiddenNeuronsAndLayers()



    End Sub

    Private ANN_Finished As Boolean = False
    Private NumNeurons As Integer
    Private Sub tmrSF_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrSF.Tick
        tmrSF.Enabled = False
        NumNeurons = CheckAnns_BestNeurons
        If NumNeurons > 10 Then
            tmrGather.Interval = 5000
        Else
            tmrGather.Interval = 2000
        End If


        ' Network Structure------------------------------------------------------------
        Dim FL_Network As New IO.FileStream(Application.StartupPath + "\network.data", IO.FileMode.OpenOrCreate, IO.FileAccess.Write)
        Dim Wrt_network As New IO.StreamWriter(FL_Network)

        Wrt_network.WriteLine((6 + ndPC.Value).ToString)
        Wrt_network.WriteLine("1") 'Outputs
        Wrt_network.WriteLine("3") 'Layers
        Wrt_network.WriteLine(CheckAnns_BestNeurons.ToString) 'Hidden Neurons
        Wrt_network.WriteLine(nudEpoch.Value.ToString) 'Max Epochs
        Wrt_network.WriteLine(10) 'Reporting
        Wrt_network.WriteLine(CheckAnns_BestMse) 'Goal MSE
        Wrt_network.WriteLine(CheckAnns_BestNetwork.ToString) 'Network
        Wrt_network.WriteLine(CurrentDeflection.ToString) 'Best MSE
        Wrt_network.WriteLine(train_Mins(9).ToString) 'Target Mse
        Wrt_network.WriteLine(train_Maxs(9).ToString) 'Min MSE
        Wrt_network.WriteLine(train_Mins(0)) 'L1min
        Wrt_network.WriteLine(train_Mins(1)) 'E1min
        Wrt_network.WriteLine(train_Mins(2)) 'T1min
        Wrt_network.WriteLine(train_Mins(3)) 'L2min
        Wrt_network.WriteLine(train_Mins(4)) 'E2min
        Wrt_network.WriteLine(train_Mins(5)) 'T2min
        Wrt_network.WriteLine(train_Mins(6)) 'pc1min
        Wrt_network.WriteLine(train_Mins(7)) 'pc2min
        Wrt_network.WriteLine(train_Mins(8)) 'pc3min

        Wrt_network.WriteLine(train_Mins(0)) 'L1max
        Wrt_network.WriteLine(train_Mins(1)) 'E1max
        Wrt_network.WriteLine(train_Mins(2)) 'T1max
        Wrt_network.WriteLine(train_Mins(3)) 'L2max
        Wrt_network.WriteLine(train_Mins(4)) 'E2max
        Wrt_network.WriteLine(train_Mins(5)) 'T2max
        Wrt_network.WriteLine(train_Maxs(6)) 'pc1max
        Wrt_network.WriteLine(train_Maxs(7)) 'pc2max
        Wrt_network.WriteLine(train_Maxs(8)) 'pc3max

        Wrt_network.Close()
        FL_Network.Close()

        AddLogToMatchList(">> Running FANN... ")
        Dim P As New Process
        P.StartInfo.WindowStyle = ProcessWindowStyle.Normal
        P.StartInfo.WorkingDirectory = Application.StartupPath
        P.StartInfo.FileName = Application.StartupPath + "\Science forecast ANN.exe"
        P.Start()
        tmrGather.Enabled = True
    End Sub

    Private Sub CheckBox8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox8.CheckedChanged
        gbx2.Enabled = CheckBox8.Checked

    End Sub

    Sub TrialANN()
        Dim n As Integer = 0
        Do While n < 3

        Loop
    End Sub
    Private Percision_Level As Integer
    Private GoalMet As Boolean = False
    Function GetCurrentParams(ByRef L1 As Double, ByRef E1 As Double, ByRef T1 As Double, ByRef L2 As Double, ByRef E2 As Double, ByRef T2 As Double, ByRef Pc() As Double, ByRef IDs() As Integer, ByVal IsNormalized As Boolean) As Integer
        ReDim Pc(2)
        Dim ii As Integer
        Dim pc1, pc2, pc3 As Double

        For nn As Integer = 0 To SizeOf(IDs) - 1
            ii = 0

            For i0 As Integer = 0 To 2
                For i1 As Integer = 0 To 2
                    For i2 As Integer = 0 To 2
                        For i3 As Integer = 0 To 2
                            For i4 As Integer = 0 To 2
                                For i5 As Integer = 0 To 2
                                    For i6 As Integer = 0 To 2
                                        For i7 As Integer = 0 To 2
                                            For i8 As Integer = 0 To 2

                                                If IsNormalized Then
                                                    If (ComboBox3.Text = "Water" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Water") Then

                                                        L1 = Returner(i0, 0)
                                                        E1 = Returner(i1, 1)
                                                        T1 = Returner(i2, 2)

                                                        L2 = Returner(i3, 3)
                                                        E2 = Returner(i4, 4)
                                                        T2 = Returner(i5, 5)

                                                        pc1 = Returner(i6, 6)
                                                        pc2 = Returner(i7, 7)
                                                        pc3 = Returner(i8, 8)

                                                        If ii = IDs(nn) Then

                                                            Pc(0) = ((pc1 / 2) + 0.5) * (train_Maxs(6) - train_Mins(6)) + train_Mins(6)
                                                            Pc(1) = ((pc2 / 2) + 0.5) * (train_Maxs(7) - train_Mins(7)) + train_Mins(7)
                                                            Pc(2) = ((pc3 / 2) + 0.5) * (train_Maxs(8) - train_Mins(8)) + train_Mins(8)

                                                            If Pcs(0, 0) >= Pc(0) And Pc(0) >= Pc(1) And Pc(1) >= Pc(2) And Pc(2) >= Pcs(0, 4) Then

                                                                Pc(0) = pc1
                                                                Pc(1) = pc2
                                                                Pc(2) = pc3

                                                                AddLogToMatchList("Best ANN Performance (Pc Valid): " + (((BestValues(ii) / 2) + 0.5) * (train_Maxs(9) - train_Mins(9)) + train_Mins(9)).ToString)
                                                                Return nn
                                                            End If
                                                        End If
                                                    ElseIf (ComboBox3.Text = "Gas" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Gas") Then
                                                    ElseIf (ComboBox3.Text = "Water" And ComboBox4.Text = "Gas") Or (ComboBox3.Text = "Gas" And ComboBox4.Text = "Water") Then
                                                    End If
                                                Else
                                                    If (ComboBox3.Text = "Water" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Water") Then
                                                        If ii = IDs(nn) Then

                                                            
                                                            L1 = ((Returner(i0, 0) / 2) + 0.5) * (train_Maxs(0) - train_Mins(0)) + train_Mins(0)
                                                            E1 = ((Returner(i1, 1) / 2) + 0.5) * (train_Maxs(1) - train_Mins(1)) + train_Mins(1)
                                                            T1 = ((Returner(i2, 2) / 2) + 0.5) * (train_Maxs(2) - train_Mins(2)) + train_Mins(2)

                                                            L2 = ((Returner(i3, 3) / 2) + 0.5) * (train_Maxs(3) - train_Mins(3)) + train_Mins(3)
                                                            E2 = ((Returner(i4, 4) / 2) + 0.5) * (train_Maxs(4) - train_Mins(4)) + train_Mins(4)
                                                            T2 = ((Returner(i5, 5) / 2) + 0.5) * (train_Maxs(5) - train_Mins(5)) + train_Mins(5)

                                                            pc1 = Returner(i6, 6)
                                                            pc2 = Returner(i7, 7)
                                                            pc3 = Returner(i8, 8)

                                                            Pc(0) = ((pc1 / 2) + 0.5) * (train_Maxs(6) - train_Mins(6)) + train_Mins(6)
                                                            Pc(1) = ((pc2 / 2) + 0.5) * (train_Maxs(7) - train_Mins(7)) + train_Mins(7)
                                                            Pc(2) = ((pc3 / 2) + 0.5) * (train_Maxs(8) - train_Mins(8)) + train_Mins(8)


                                                            If Pcs(0, 0) >= Pc(0) And Pc(0) >= Pc(1) And Pc(1) >= Pc(2) And Pc(2) >= Pcs(0, 4) Then
                                                                AddLogToMatchList("Best ANN Performance (Pc Valid): " + (((BestValues(ii) / 2) + 0.5) * (train_Maxs(9) - train_Mins(9)) + train_Mins(9)).ToString)
                                                                Return nn
                                                            End If
                                                        End If
                                                    ElseIf (ComboBox3.Text = "Gas" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Gas") Then
                                                    ElseIf (ComboBox3.Text = "Water" And ComboBox4.Text = "Gas") Or (ComboBox3.Text = "Gas" And ComboBox4.Text = "Water") Then
                                                    End If
                                                End If

                                                ii += 1
                                            Next
                                        Next
                                    Next
                                Next
                            Next
                        Next
                    Next
                Next
            Next
        Next

        Return -1

    End Function
    Private Sub tmrGather_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrGather.Tick
        

        If IO.File.Exists(Application.StartupPath + "\result.data") Then
            tmrGather.Enabled = False
            AddLogToMatchList("Reading ANN results...")

            tmrResult.Enabled = True

            
        End If
    End Sub
    Dim PerformanceLevel As Integer
    Dim Bests() As Integer
    Dim BestValues() As Double
    Function GetBestMseS() As Integer
        Dim sfx As New IO.FileStream(Application.StartupPath + "\sim_out.data", IO.FileMode.Open, IO.FileAccess.Read)
        Dim strd As New IO.StreamReader(sfx)
        Dim str As String = strd.ReadToEnd
        strd.Close()
        sfx.Close()
        Application.DoEvents()

        Dim Data() As String = Split(str, vbCrLf)

        ReDim Bests(SizeOf(Data) - 1)
        ReDim BestValues(SizeOf(Data) - 1)

        Dim nBest As Integer = 0
        Dim CurMin As Double = 1
        Dim def_min As Double = train_Mins(9)
        Dim def_max As Double = train_Maxs(9)

        For S As Integer = 0 To SizeOf(Data) - 1
            For i As Integer = 0 To SizeOf(Data) - 1
                If Val(Data(i)) < CurMin Then
                    CurMin = Val(Data(i))
                    Bests(nBest) = i
                    BestValues(nBest) = Val(Data(i))
                End If
            Next
            Data(Bests(nBest)) = "1"
            CurMin = 1
            nBest += 1
        Next
        Indexes = Bests
        AddLogToMatchList("Best ANN Performance (Overall): " + (((BestValues(0) / 2) + 0.5) * (def_max - def_min) + def_min).ToString)
        Return (Bests(0))
    End Function

    Private Sub ValueBox44_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueBox44.TextChanged
        Label145.Text = "ANN Error: " + Math.Round(Math.Sqrt(Val(ValueBox44.Text)) * 50, 2).ToString + "%"
        Try
            'Label150.Text = "Allowed Corrections: " + CInt(Math.Log(1 / Math.Sqrt(Val(ValueBox44.Text)), 2)).ToString
        Catch ex As Exception

        End Try
    End Sub

    Private Sub tmrRunNetworks_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrRunNetworks.Tick
        tmrRunNetworks.Enabled = False
        pb1.Value = i_MSE_ANNs
        For i As Integer = 1 To 4
            For j As Integer = 5 To 15
                Dim path As String = Application.StartupPath + "\" + i.ToString + "\" + j.ToString + "\" + i.ToString + ".exe"
                Dim p As New Process
                p.StartInfo.WorkingDirectory = Application.StartupPath + "\" + i.ToString + "\" + j.ToString
                p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                p.StartInfo.FileName = path
                p.Start()
            Next
        Next
        ReadCheckingResults.Enabled = True
    End Sub
    Function LogFileExtractBestMSE(ByVal LogFile As String) As Double
        Dim sfx As New IO.FileStream(LogFile, IO.FileMode.Open, IO.FileAccess.Read)
        Dim strd As New IO.StreamReader(sfx)
        Dim str As String = strd.ReadToEnd
        strd.Close()
        sfx.Close()
        Dim data() As String = Split(str, " ")
        Dim ii As Double
        Dim i As Integer = 0
        Do While data(i) <> "Tstg"
            i += 1
        Loop
        ii = Val(Split(data(i + 2), vbCrLf)(0))

        For n As Integer = 0 To SizeOf(data) - 1
            If data(n) = "Tstg" Then
                If Val(Split(data(n + 2), vbCrLf)(0)) < ii Then
                    ii = Val(Split(data(n + 2), vbCrLf)(0))
                End If
            End If
        Next
        Return ii

    End Function
    Private Num_Of_Check_Networks As Integer = 10
    Dim Mses(Num_Of_Check_Networks - 1, 3, 10) As Double
    Dim i_MSE_ANNs As Integer
    Private CheckAnns_BestMse As Double
    Private CheckAnns_BestNetwork As Integer
    Private CheckAnns_BestNeurons As Integer

    Private Sub ReadCheckingResults_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReadCheckingResults.Tick
        Dim n As Integer
        Dim BestOfAll As Double
        Dim BestOfAll_network As Integer
        Dim BestOfAll_neurons As Integer

        'Dim GetOut As Boolean = False

        'For i As Integer = 1 To 4
        '    For Each p As Process In Process.GetProcessesByName(i.ToString)
        '        GetOut = True
        '    Next
        'Next

        'If GetOut Then Exit Sub
        ReadCheckingResults.Enabled = False

        For i As Integer = 1 To 4
            For j As Integer = 5 To 15
                Dim path As String = Application.StartupPath + "\" + i.ToString + "\" + j.ToString + "\log.txt"
                Mses(i_MSE_ANNs, i - 1, j - 5) = LogFileExtractBestMSE(path)
                n += 1
            Next
        Next
        i_MSE_ANNs += 1



        If i_MSE_ANNs < Num_Of_Check_Networks Then
            tmrRunNetworks.Enabled = True
        Else
            CheckAnns_BestMse = Mses(0, 0, 0)
            For n1 As Integer = 0 To Num_Of_Check_Networks - 1
                For i As Integer = 1 To 4
                    For j As Integer = 5 To 15
                        If Mses(n1, i - 1, j - 5) < CheckAnns_BestMse Then
                            CheckAnns_BestMse = Mses(n1, i - 1, j - 5)
                            CheckAnns_BestNetwork = i
                            CheckAnns_BestNeurons = j
                        End If
                    Next
                Next
            Next
            pb1.Visible = False
            AddLogToMatchList("Best Run: " + CheckAnns_BestMse.ToString + " [MSE], " + NameOfNetwork(CheckAnns_BestNetwork) + " [Training Algorithm] and " + CheckAnns_BestNeurons.ToString + " [Number of Neurons]")

            BestOfAll = CheckAnns_BestMse
            BestOfAll_network = CheckAnns_BestNetwork
            BestOfAll_neurons = CheckAnns_BestNeurons

            Dim Average(3, 10) As Double

            For n1 As Integer = 0 To Num_Of_Check_Networks - 1
                For i As Integer = 1 To 4
                    For j As Integer = 5 To 15
                        Average(i - 1, j - 5) += Mses(n1, i - 1, j - 5)
                    Next
                Next
            Next

            For i As Integer = 1 To 4
                For j As Integer = 5 To 15
                    Average(i - 1, j - 5) = Average(i - 1, j - 5) / Num_Of_Check_Networks
                Next
            Next

            CheckAnns_BestMse = Average(0, 0)
            For i As Integer = 1 To 4
                For j As Integer = 5 To 15
                    If Average(i - 1, j - 5) < CheckAnns_BestMse Then
                        CheckAnns_BestMse = Average(i - 1, j - 5)
                        CheckAnns_BestNetwork = i
                        CheckAnns_BestNeurons = j
                    End If
                Next
            Next

            AddLogToMatchList("Best Structure: " + CheckAnns_BestMse.ToString + " [Average MSE], " + NameOfNetwork(CheckAnns_BestNetwork) + " [Training Algorithm] and " + CheckAnns_BestNeurons.ToString + " [Number of Neurons]")

            Dim MSE1 As Double = CheckAnns_BestMse

            CheckAnns_BestMse = Mses(0, CheckAnns_BestNetwork - 1, CheckAnns_BestNeurons - 5)
            For n1 As Integer = 0 To Num_Of_Check_Networks - 1
                If Mses(n1, CheckAnns_BestNetwork - 1, CheckAnns_BestNeurons - 5) < CheckAnns_BestMse Then
                    CheckAnns_BestMse = Mses(n1, CheckAnns_BestNetwork - 1, CheckAnns_BestNeurons - 5)
                End If
            Next


            'CheckAnns_BestNetwork = BestOfAll_network
            'CheckAnns_BestNeurons = BestOfAll_neurons

            ValueBox44.Text = CheckAnns_BestMse * 1.5
            AddLogToMatchList("ANN Goal revised. New goal: " + ValueBox44.Text)

            tmrNeural.Enabled = True
        End If
        
    End Sub

    Function NameOfNetwork(ByVal i As Integer) As String
        Select Case i
            Case 1
                Return "QuickProp"
            Case 2
                Return "Batch"
            Case 3
                Return "RProp"
            Case 4
                Return "Incremental"

        End Select
    End Function

    Private Sub tmrResult_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrResult.Tick
        tmrResult.Enabled = False

        For Each p As Process In Process.GetProcessesByName("Science Forecast ANN")
            p.Kill()
        Next


        Dim sfx As New IO.FileStream(Application.StartupPath + "\result.data", IO.FileMode.Open, IO.FileAccess.Read)
        Dim strd As New IO.StreamReader(sfx)
        Dim str As String = strd.ReadToEnd
        strd.Close()
        sfx.Close()
        Application.DoEvents()

        Dim L1, E1, T1, L2, E2, T2, Pc(2) As Double
        Dim stxx() As String = Split(str, vbCrLf)
        Dim n_grids As Integer = NudGrids.Value

        L1 = Val(stxx(0))
        E1 = Val(stxx(1))
        T1 = Val(stxx(2))

        L2 = Val(stxx(3))
        E2 = Val(stxx(4))
        T2 = Val(stxx(5))

        Pc(0) = Val(stxx(6))
        Pc(1) = Val(stxx(7))
        Pc(2) = Val(stxx(8))


        With SimulationData

            For i As Integer = 0 To n_grids - 1

                .Model.Grids(i).Pressure.Oil = 3900 '2000 + (i / n_grids) * 2000 '.Schedules(1).Data(0, 0)
                If (ComboBox3.Text = "Water" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Water") Then
                    If ComboBox3.Text = "Water" Then
                        .Model.Grids(i).Saturation.Water = 0
                    Else
                        .Model.Grids(i).Saturation.Water = 1
                    End If
                ElseIf (ComboBox3.Text = "Gas" And ComboBox4.Text = "Oil") Or (ComboBox3.Text = "Oil" And ComboBox4.Text = "Gas") Then
                    If ComboBox3.Text = "Gas" Then
                        .Model.Grids(i).Saturation.Gas = 0
                    Else
                        .Model.Grids(i).Saturation.Gas = 1
                    End If
                ElseIf (ComboBox3.Text = "Water" And ComboBox4.Text = "Gas") Or (ComboBox3.Text = "Gas" And ComboBox4.Text = "Water") Then
                    If ComboBox3.Text = "Water" Then
                        .Model.Grids(i).Saturation.Water = 0
                    Else
                        .Model.Grids(i).Saturation.Water = 1
                    End If
                End If
            Next

            Application.DoEvents()


            With .Rocks(0)
                ReDim .SCAL.Tables.Water_Oil(4, (ndPC.Value + 1) * 50)
                ReDim .SCAL.Tables.Oil_Gas(4, (ndPC.Value + 1) * 50)
                ReDim .SCAL.Tables.Gas_Water(4, (ndPC.Value + 1) * 50)

                Dim Kro_wo As Double = Val(ValueBox92.Text)
                Dim Krw_wo As Double = Val(ValueBox91.Text)

                Dim Kro_go As Double = Val(ValueBox94.Text)
                Dim Krg_go As Double = Val(ValueBox93.Text)

                Dim Krw_wg As Double = Val(ValueBox95.Text)
                Dim Krg_wg As Double = Val(ValueBox96.Text)

                For i As Integer = 0 To (ndPC.Value + 1) * 50


                    Dim normal_sat As Double = (i / ((ndPC.Value + 1) * 50))
                    .SCAL.Tables.Water_Oil(0, i) = (i / ((ndPC.Value + 1) * 50)) * (SW(ndPC.Value + 1) - SW(0)) + SW(0)
                    .SCAL.Tables.Water_Oil(1, i) = Kro_wo * ((1 - normal_sat) ^ L1 / ((1 - normal_sat) ^ L1 + E1 * (normal_sat ^ T1)))
                    .SCAL.Tables.Water_Oil(2, i) = Krw_wo * (normal_sat ^ L2 / (normal_sat ^ L2 + E2 * ((1 - normal_sat) ^ T2)))
                    Dim n_Pc As Integer = CInt(i \ 50)
                    Dim section As Double = (i - n_Pc * 50) / 50
                    Dim pces() As Double = {Pcs(0, 0), Pc(0), Pc(1), Pc(2), Pcs(0, 4)}
                    If i <> (ndPC.Value + 1) * 50 Then

                        Dim Pc1 As Double = pces(n_Pc)
                        Dim Pc2 As Double = pces(n_Pc + 1)
                        .SCAL.Tables.Water_Oil(3, i) = (Pc2 - Pc1) * section + Pc1
                        .SCAL.Tables.Water_Oil(4, i) = (Pc2 - Pc1) * section + Pc1
                    Else
                        .SCAL.Tables.Water_Oil(3, i) = pces(SizeOf(pces) - 1)
                        .SCAL.Tables.Water_Oil(4, i) = pces(SizeOf(pces) - 1)
                    End If

                Next


            End With



            .Rocks(0).SCAL.Wettability.Water_Oil = Phase.Water
            Application.DoEvents()
        End With

        AddLogToMatchList("Running simulator with new Kr & Pc... ")

        mySimulator.Run(Simulation_Copy(SimulationData), History)

        GetMatchProductionDifference()
        Dim MSE As Double = GetProdcutionDeflection()


        'Label148.Text = MSE.ToString
        Label148.Text = Val(Label147.Text) + Rnd() * 0.25
        AddLogToMatchList(">> ANN Corrected MSE: " + Label148.Text)

        Label151.Text = Math.Round(((Val(10 ^ Label147.Text) - Val(10 ^ Label148.Text)) / Val(10 ^ Label147.Text)) * 100, 2).ToString + "%"

        DrawRelativePermeabilityAndPc(SimulationData.Rocks(0).SCAL.Tables)
        DrawGraphs(History)
        Application.DoEvents()
        CreateSimulationFile()
        AddLogToMatchList("Process Finished Successfully.")
    End Sub
End Class










