Public Class ctlSchedule
    Public ID As Integer
    Public Event Removed(ByVal ID As Integer)
    Public Event Update(ByRef sender As ctlSchedule)

    Private Sub ctlSchedule_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        
        
        GroupBox1.Dock = DockStyle.Fill
        GroupBox2.Dock = DockStyle.Fill

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        GroupBox1.Visible = (ComboBox1.SelectedIndex = 0)
        GroupBox2.Visible = (ComboBox1.SelectedIndex = 1)
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        ValueBox1.Enabled = ComboBox2.Text.ToUpper.StartsWith("STATIC")
        dgvGrid.Visible = ComboBox2.Text.ToUpper.StartsWith("DYNAMIC")
        ComboBox3.Enabled = ComboBox2.Text.ToUpper.EndsWith("PRESSURE")

        If Not ComboBox3.Enabled Then
            ComboBox3.SelectedIndex = 0
        End If

    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        If ComboBox1.SelectedIndex = -1 Then
            ComboBox1.SelectedIndex = 0
            ComboBox2.SelectedIndex = 1
        End If
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If MessageBox.Show("Are you sure to remove Schedule ID " + ID.ToString, "Black-Oil MSV", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            RaiseEvent Removed(ID)
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        
        If ComboBox3.SelectedIndex = 0 Then
            ComboBox4.Enabled = False
        Else
            ComboBox4.Enabled = True
            ComboBox4.SelectedIndex = 0
        End If

    End Sub

    Private Sub ValueBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueBox1.TextChanged

    End Sub
End Class
