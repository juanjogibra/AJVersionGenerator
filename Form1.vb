Imports System.ComponentModel

Public Class Form1
    Dim sDate As String
    Dim FormatValue As String
    Dim Index As Integer
    Dim Index2 As Integer
    Dim x As Integer
    Dim y As Integer


    Public Sub Main()

    End Sub



    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If (TextBox2.Text <> "") And (ComboBox1.Text <> "") And (ComboBox2.Text <> "") Then
            sDate = Format(DateTimePicker1.Value.ToOADate)
            FormatValue = ComboBox1.Text & "." & ComboBox2.Text & "." & sDate.Substring(0, sDate.IndexOf(",")) & "." & TextBox2.Text
            TextBox1.Text = FormatValue
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        FormatValue = ""
        sDate = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox1.Items.Clear()
        ComboBox2.Items.Clear()

    End Sub


    Private Sub ComboBox1_SelectedIndexFilled(sender As Object, e As EventArgs) Handles ComboBox1.DropDown
        Index = 0
        ComboBox1.Items.Clear()
        For x = 10 To 30
            ComboBox1.Items.Insert(Index, x.ToString)
            Index = Index + 1
        Next
        Index = 0
    End Sub

    Private Sub ComboBox2_SelectedIndexFilled(sender As Object, e As EventArgs) Handles ComboBox2.DropDown
        Index2 = 0
        ComboBox2.Items.Clear()
        For y = 0 To 9
            ComboBox2.Items.Insert(Index2, y.ToString)
            Index = Index + 1
        Next
        Index2 = 0
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress

        '97 - 122 = Ascii codes for simple letters
        '65 - 90  = Ascii codes for capital letters
        '48 - 57  = Ascii codes for numbers

        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If

    End Sub
    Private Sub ComboBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox1.KeyPress


        '97 - 122 = Ascii codes for simple letters
        '65 - 90  = Ascii codes for capital letters
        '48 - 57  = Ascii codes for numbers

        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If

    End Sub

    Private Sub ComboBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox2.KeyPress


        '97 - 122 = Ascii codes for simple letters
        '65 - 90  = Ascii codes for capital letters
        '48 - 57  = Ascii codes for numbers

        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        sDate = Format(DateTimePicker1.Value.ToOADate)
        TextBox3.Text = sDate.Substring(0, sDate.IndexOf(","))

        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "dd/MM/yyyy"


    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        sDate = Format(DateTimePicker1.Value.ToOADate)
        TextBox3.Text = sDate.Substring(0, sDate.IndexOf(","))
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown

        If e.KeyCode = Keys.Enter Then
            If (TextBox2.Text <> "") And (ComboBox1.Text <> "") And (ComboBox2.Text <> "") Then
                sDate = Format(DateTimePicker1.Value.ToOADate)
                FormatValue = ComboBox1.Text & "." & ComboBox2.Text & "." & sDate.Substring(0, sDate.IndexOf(",")) & "." & TextBox2.Text
                TextBox1.Text = FormatValue
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Clipboard.SetText(TextBox3.Text)
        MessageBox.Show("Copiado al portapapeles")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Clipboard.SetText(TextBox1.Text)
        MessageBox.Show("Copiado al portapapeles")
    End Sub
End Class
