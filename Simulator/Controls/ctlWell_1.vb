Public Class ctlWell
    Public Event Removed(ByVal Name As String)
    Public Event Update(ByRef sender As ctlWell)

    Private Sub dgvWellPath_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWellPath.CellValidated
        If e.RowIndex > 0 And e.ColumnIndex = 0 Then
            If dgvWellPath.Item(1, e.RowIndex - 1).Style.BackColor = Color.Red Or dgvWellPath.Item(1, e.RowIndex - 1).Value = "" Then
                dgvWellPath.Item(1, e.RowIndex - 1).Value = (dgvWellPath.Item(0, e.RowIndex).Value)
            End If
        End If

        UpdateColors()
        
        RaiseEvent Update(Me)
    End Sub
    Sub UpdateColors()
        For i As Integer = 0 To dgvWellPath.RowCount - 2
            Dim str1() As String = Split(dgvWellPath.Item(0, i).Value, ",")
            If SizeOf(str1) = 3 Then
                If dgvWellPath.Item(0, i).Style.BackColor = Color.Red Then dgvWellPath.Item(0, i).Style.BackColor = Color.White
            Else
                If dgvWellPath.Item(0, i).Style.BackColor = Color.White Then dgvWellPath.Item(0, i).Style.BackColor = Color.Red
            End If

            Dim str2() As String = Split(dgvWellPath.Item(1, i).Value, ",")
            If SizeOf(str2) = 3 Then
                If dgvWellPath.Item(1, i).Style.BackColor = Color.Red Then dgvWellPath.Item(1, i).Style.BackColor = Color.White
            Else
                If dgvWellPath.Item(1, i).Style.BackColor = Color.White Then dgvWellPath.Item(1, i).Style.BackColor = Color.Red
            End If
        Next

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If MessageBox.Show("Are you sure to remove Well " + TextBox1.Text, "Black-Oil MSV", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            RaiseEvent Removed(TextBox1.Text)
        End If
    End Sub

    Private Sub dgvWellPath_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvWellPath.RowsAdded

        dgvWellPath.Item(0, e.RowIndex).Style.BackColor = Color.White
        dgvWellPath.Item(1, e.RowIndex).Style.BackColor = Color.White

        dgvWellPath.Item(2, e.RowIndex).Value = Val(ValueBox2.Text)
        dgvWellPath.Item(3, e.RowIndex).Value = "False"
        dgvWellPath.Item(4, e.RowIndex).Value = Val(ValueBox1.Text)

    End Sub

    Private Sub TextBox1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Leave
        RaiseEvent Update(Me)
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.Leave
        RaiseEvent Update(Me)
    End Sub

    Private Sub ValueBox1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ValueBox1.Leave
        RaiseEvent Update(Me)
    End Sub

    Private Sub ValueBox2_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ValueBox2.Leave
        RaiseEvent Update(Me)
    End Sub

    Private Sub dgvWellPath_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWellPath.CellContentClick

    End Sub
End Class
