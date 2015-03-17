Public Class TDP
    Public TDC1 As TDC
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        TDC1 = New TDC
        TDC1.Dock = DockStyle.Fill
        Panel1.Controls.Add(TDC1)
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    
    Private Sub TDP_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Try
            TDC1.ToggleTransparency()
        Catch ex As Exception
        End Try

    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        Try
            TDC1.ToggleLighting()
        Catch ex As Exception

        End Try

    End Sub
End Class
