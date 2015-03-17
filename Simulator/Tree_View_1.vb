Public Module Tree_View
    Sub InitializeTreeView(ByRef TV As TreeView)
        With TV
            .Nodes.Clear()

            Dim nd_simulation As New TreeNode
            With nd_simulation
                .Text = "Simulation"
                .ImageIndex = 7
                .SelectedImageIndex = 7
            End With

            Dim nd_model As New TreeNode
            With nd_model
                .Text = "Model"
                .ImageIndex = 1
                .SelectedImageIndex = 1
            End With

            Dim nd_pvt As New TreeNode
            With nd_pvt
                .Text = "PVT"
                .ImageIndex = 2
                .SelectedImageIndex = 2
            End With


            Dim nd_rocktype As New TreeNode
            With nd_rocktype
                .Text = "Rock Types"
                .ImageIndex = 3
                .SelectedImageIndex = 3
            End With

            Dim nd_schedules As New TreeNode
            With nd_schedules
                .Text = "Schedules"
                .ImageIndex = 4
                .SelectedImageIndex = 4
            End With

            Dim nd_Project As New TreeNode
            With nd_Project
                .Text = "Project"
                .ImageIndex = 0
                .SelectedImageIndex = 0
            End With

            Dim nd_Components As New TreeNode
            With nd_Components
                .Text = "Components"
                .ImageIndex = 6
                .SelectedImageIndex = 6
            End With

            Dim nd_SDKP As New TreeNode
            With nd_SDKP
                .Text = "KrPc History Match"
                .ImageIndex = 8
                .SelectedImageIndex = 8
            End With

            nd_Components.Nodes.Add(nd_SDKP)

            nd_simulation.Nodes.Add(nd_model)
            nd_simulation.Nodes.Add(nd_pvt)
            nd_simulation.Nodes.Add(nd_rocktype)
            nd_simulation.Nodes.Add(nd_schedules)

            nd_Project.Nodes.Add(nd_simulation)
            nd_Project.Nodes.Add(nd_Components)

            .Nodes.Add(nd_Project)

            .Nodes(0).ExpandAll()
        End With

    End Sub
End Module
