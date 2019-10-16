Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'si es acces utilizamos la funcion de llenaoledb si es para sql llenasql
        DataGridView1.DataSource = GridOledb("select * from estudiantes")
    End Sub
End Class
