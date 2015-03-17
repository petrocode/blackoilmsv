Public Class ctlRockType
    Public ID As Integer
    Public Event Removed(ByVal ID As Integer)


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If MessageBox.Show("Are you sure to remove Rock Type ID " + ID.ToString, "Black-Oil MSV", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            RaiseEvent Removed(ID)
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Paste(dgvOW, dgvOW.Item(0, 0))
    End Sub
    Public Sub Paste(ByRef DGV As DataGridView, ByVal CurrentCell As DataGridViewCell)
        Dim i As String = Clipboard.GetText
        Dim CurRowIndex, CurColIndex As Integer
        Dim CurDel As Char
        CurDel = Chr(Keys.Tab)
        Dim MyRows() As String = Split(i, vbCrLf)
        If MyRows Is Nothing Then Exit Sub
        Dim MyCols() As String = Split(MyRows(0), CurDel)
        If CurrentCell IsNot Nothing Then
            CurRowIndex = CurrentCell.RowIndex
            CurColIndex = CurrentCell.ColumnIndex
        End If
        DGV.RowCount = CurRowIndex + MyRows.Length
        'DGV.ColumnCount = CurColIndex + MyCols.Length

        Application.DoEvents()

        For j As Integer = CurRowIndex To CurRowIndex + MyRows.Length - 1
            Dim CurCols() As String = Split(MyRows(j - CurRowIndex), CurDel)
            If CurCols(0) = vbCrLf Then Exit Sub
            For k As Integer = CurColIndex To CurColIndex + CurCols.Length - 1
                DGV.Item(k, j).Value = IIf(String.IsNullOrEmpty(CurCols(k - CurColIndex)), CurCols(k - CurColIndex), Val(CurCols(k - CurColIndex)))
            Next
        Next
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Paste(dgvGO, dgvGO.Item(0, 0))
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Paste(dgvWG, dgvWG.Item(0, 0))
    End Sub

    Private Sub dgvOW_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvOW.CellContentClick

    End Sub
End Class
