<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ColorBoundary
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.ValueBox2 = New Simulator.ValueBox
        Me.ValueBox1 = New Simulator.ValueBox
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Blue
        Me.Button1.Location = New System.Drawing.Point(37, 33)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(25, 23)
        Me.Button1.TabIndex = 4
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Red
        Me.Button2.Location = New System.Drawing.Point(37, 57)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(25, 23)
        Me.Button2.TabIndex = 5
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(55, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Parameter"
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"So", "Sw", "Sg", "Po", "Pw", "Pg", "Ko - X", "Ko - Y", "Ko - Z", "Kw - X", "Kw - Y", "Kw - Z", "Kg - X", "Kg - Y", "Kg - Z", "Pcow", "Pcgo", "Bo", "Bw", "Bg", "Viscosity - Oil", "Viscosity - Water", "Viscosity - Gas", "Density - Oil", "Density - Water", "Density - Gas", "Gamma - Oil", "Gamma - Water", "Gamma - Gas", "Rs", "X", "Y", "Z", "Lx", "Ly", "Lz", "Volume"})
        Me.ComboBox1.Location = New System.Drawing.Point(64, 6)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(131, 21)
        Me.ComboBox1.TabIndex = 7
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(2, 33)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(35, 23)
        Me.Button3.TabIndex = 8
        Me.Button3.Text = "Min"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(2, 57)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(35, 23)
        Me.Button4.TabIndex = 9
        Me.Button4.Text = "Max"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'ValueBox2
        '
        Me.ValueBox2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.ValueBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
        Me.ValueBox2.IsNumeric = True
        Me.ValueBox2.Location = New System.Drawing.Point(64, 59)
        Me.ValueBox2.Max = 0
        Me.ValueBox2.MaxMinCheck = False
        Me.ValueBox2.Min = 0
        Me.ValueBox2.Name = "ValueBox2"
        Me.ValueBox2.Size = New System.Drawing.Size(131, 20)
        Me.ValueBox2.TabIndex = 3
        '
        'ValueBox1
        '
        Me.ValueBox1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.ValueBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
        Me.ValueBox1.IsNumeric = True
        Me.ValueBox1.Location = New System.Drawing.Point(64, 33)
        Me.ValueBox1.Max = 0
        Me.ValueBox1.MaxMinCheck = False
        Me.ValueBox1.Min = 0
        Me.ValueBox1.Name = "ValueBox1"
        Me.ValueBox1.Size = New System.Drawing.Size(131, 20)
        Me.ValueBox1.TabIndex = 1
        '
        'ColorBoundary
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ValueBox2)
        Me.Controls.Add(Me.ValueBox1)
        Me.Name = "ColorBoundary"
        Me.Size = New System.Drawing.Size(202, 84)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ValueBox1 As ValueBox
    Friend WithEvents ValueBox2 As ValueBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button

End Class
