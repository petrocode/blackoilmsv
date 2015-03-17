Public Class Memory_Release
    Dim n As Integer
    Private Sub Memory_Release_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Label2.Text = "5"
        tmrMR.Enabled = True
    End Sub

    Private Sub tmrMR_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrMR.Tick

    End Sub
End Class