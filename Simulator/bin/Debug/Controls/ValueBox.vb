''' <summary>
''' Numeric TextBox Control
''' By Amir Emamjomeh Kuhbanani
''' </summary>
''' <remarks></remarks>
Public Class ValueBox
    Inherits TextBox
    Private _Max, _Min As Double
    Private _MaxMinCheck, _IsNumeric As Boolean
    Public Property Max() As Double
        Get
            Return _Max
        End Get
        Set(ByVal value As Double)
            _Max = value
        End Set
    End Property
    Public Property Min() As Double
        Get
            Return _Min
        End Get
        Set(ByVal value As Double)
            _Min = value
        End Set
    End Property
    Public Property MaxMinCheck() As Boolean
        Get
            Return _MaxMinCheck
        End Get
        Set(ByVal value As Boolean)
            _MaxMinCheck = value
        End Set
    End Property
    Public Property IsNumeric() As Boolean
        Get
            Return _IsNumeric
        End Get
        Set(ByVal value As Boolean)
            _IsNumeric = value
        End Set
    End Property
    Private LastValidText As String
    Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.OnKeyDown(e)
    End Sub
    Protected Overrides Sub OnValidating(ByVal e As System.ComponentModel.CancelEventArgs)
        If _IsNumeric Then
            If _MaxMinCheck Then
                If Val(Text) >= _Min And Val(Text) <= _Max Then
                    e.Cancel = False
                Else
                    Dim bl As New ToolTip
                    bl.SetToolTip(Me, "Value should be between " + Min.ToString + " and " + Max.ToString)
                    bl.IsBalloon = True
                    bl.UseAnimation = True
                    bl.Show("Value should be between " + Min.ToString + " and " + Max.ToString, Me, 2000)


                    e.Cancel = True
                    Text = LastValidText
                End If
            Else
                e.Cancel = False
            End If
        End If

    End Sub
    Protected Overrides Sub OnValidated(ByVal e As System.EventArgs)
        GoValidate()
    End Sub
    Protected Overrides Sub OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If _IsNumeric Then
            SelectedText = ""
            Static k As Integer = -1
            If (Char.IsNumber(e.KeyChar) And e.KeyChar <> "0") Or e.KeyChar = Chr(Keys.Back) Or (e.KeyChar = Chr(Keys.E) Or (e.KeyChar = Chr(Asc("-"))) Or (e.KeyChar = Chr(Asc("+"))) Or (e.KeyChar = Chr(Asc("e")))) Then
                MyBase.OnKeyPress(e)
            ElseIf e.KeyChar = "." Then
                Dim j As Integer = SelectionStart
                k = Text.IndexOf(".")
                Text = Text.Replace(".", "")
                If k > j Then
                    SelectionStart = j
                ElseIf k <> -1 Then
                    SelectionStart = j - 1
                End If
                MyBase.OnKeyPress(e)
            ElseIf e.KeyChar = "0" Then
                If SelectionStart = 0 Then
                    If Len(Text) = 0 Or Text = "." Then
                        Text = "0."
                        SelectionStart = 2
                    End If
                    e.Handled = True

                End If
            Else
                e.Handled = True
            End If
        Else
            MyBase.OnKeyPress(e)
        End If
    End Sub
    Sub GoValidate()
        If Text <> "" Then
            Text = Val(Text)
        End If

        LastValidText = Text
    End Sub

    Public Sub New()
        Me.CharacterCasing = Windows.Forms.CharacterCasing.Upper
        _IsNumeric = True
        Font = New Font(FontFamily.GenericSansSerif, 8)
    End Sub


End Class

