Public Class TDP
    Dim TDC1 As TDC
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Tdc1.ToggleLighting()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Tdc1.ToggleTransparency()
    End Sub
    Sub Load(ByVal Data As VU)
        Tdc1.loadGeometry(Data)
    End Sub
    Sub Start(ByVal Data As VU)
        Tdc1.Start(Data)
    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        TDC1 = New TDC
        TDC1.Dock = DockStyle.Fill
        Panel1.Controls.Add(TDC1)
        ' Add any initialization after the InitializeComponent() call.

    End Sub
End Class
