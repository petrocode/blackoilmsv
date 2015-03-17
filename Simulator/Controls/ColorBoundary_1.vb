Public Class ColorBoundary
    Public Event Refreshed()
    Public Event MinClicked()
    Public Event MaxClicked()

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim DLG As New ColorDialog
        With DLG
            .ShowDialog()
            Try
                Button1.BackColor = .Color
                RaiseEvent Refreshed()
            Catch ex As Exception

            End Try
        End With
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim DLG As New ColorDialog
        With DLG
            .ShowDialog()
            Try
                Button2.BackColor = .Color
                RaiseEvent Refreshed()
            Catch ex As Exception

            End Try
        End With
    End Sub

    Private Sub ValueBox1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ValueBox1.Leave
        RaiseEvent Refreshed()
    End Sub

    Private Sub ValueBox1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ValueBox1.Validated

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        RaiseEvent Refreshed()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        RaiseEvent MinClicked()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        RaiseEvent MaxClicked()
    End Sub

    Sub ResetMax()
        RaiseEvent MaxClicked()
    End Sub
    Sub ResetMin()
        RaiseEvent MinClicked()
    End Sub

    Private Sub ValueBox2_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ValueBox2.Validated
        RaiseEvent Refreshed()
    End Sub

    Private Sub ValueBox2_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ValueBox2.Validating
        RaiseEvent Refreshed()
    End Sub

    Private Sub ValueBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueBox1.TextChanged

    End Sub
End Class
