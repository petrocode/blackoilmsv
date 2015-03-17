Public Class Export

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox2.Items.Clear()
        Select Case ComboBox1.SelectedIndex
            Case 0
                ComboBox2.Items.Add("Grids")
                ComboBox2.Items.Add("Wells")
                ComboBox2.Items.Add("Wells & Grids")

                ComboBox2.Enabled = True
                ComboBox2.SelectedIndex = 0

            Case 1
                ComboBox2.Items.Add("Time Steps")

                ComboBox2.SelectedIndex = 0
                ComboBox2.Enabled = False

                GroupBox2.Enabled = False
                GroupBox3.Enabled = True


            Case 2
                ComboBox2.Items.Add("Time Steps")
                ComboBox2.SelectedIndex = 0
                ComboBox2.Enabled = False

                GroupBox2.Enabled = True
                GroupBox3.Enabled = False


        End Select
    End Sub

    Private Sub Export_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If ComboBox1.SelectedIndex = -1 Then
            ComboBox1.SelectedIndex = 0
            ComboBox2.SelectedIndex = 0
        End If


    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedItem = "Grids" Then
            GroupBox3.Enabled = False
            GroupBox2.Enabled = True
        End If
        If ComboBox2.SelectedItem = "Wells" Then
            GroupBox3.Enabled = True
            GroupBox2.Enabled = False
        End If
        If ComboBox2.SelectedItem = "Wells & Grids" Then
            GroupBox3.Enabled = True
            GroupBox2.Enabled = True
        End If
        
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Hide()
    End Sub
    Sub Save()
        Select Case ComboBox1.SelectedIndex
            Case 0
                Select Case ComboBox2.SelectedIndex
                    Case 0
                        Export_TimeStep_Grids()
                    Case 1
                    Case 2
                End Select
            Case 1
            Case 2

        End Select

        
    End Sub
    Sub Export_TimeStep_Grids()
        'excelSheet = CType( _
        '       excelWorkbook.Sheets.Add(excelWorkbook.Sheets(sheetIndex), _
        '       Type.Missing, 1, XlSheetType.xlWorksheet), Worksheet)
    End Sub
End Class