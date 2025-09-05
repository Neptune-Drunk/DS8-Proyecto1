Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports MySql.Data.MySqlClient

Public Class Form1
    Dim archivo As String
    Dim cm As MySqlCommand
    Dim pr As MySqlDataAdapter
    Dim ds1 As DataSet
    Dim conexion As String = "Server=localhost;database=asistencia;user=root;password='';"
    Dim miconexion As New MySqlConnection(conexion)
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            miconexion.Open()
            cm = New MySqlCommand()
            cm.CommandType = CommandType.Text
            cm.Connection = miconexion
            pr = New MySqlDataAdapter(cm)
            ds1 = New DataSet()
            pr.Fill(ds1)
        Catch ex As Exception
            MsgBox("Error de Conexión")
        End Try
    End Sub

    Private Sub Bttnarchivo_Click(sender As Object, e As EventArgs) Handles Bttnarchivo.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Title = "Seleccione el archivo"
        openFileDialog.Filter = "Archivos DAT (*.dat)|*.dat"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            archivo = openFileDialog.FileName
            MessageBox.Show("Archivo seleccionado: " & archivo)
        End If
    End Sub

    Private Sub BttnInsertar_Click(sender As Object, e As EventArgs) Handles BttnInsertar.Click
        If String.IsNullOrEmpty(archivo) Then
            MessageBox.Show("Por favor, seleccione un archivo primero.")
            Return
        End If

        Dim lines() As String = System.IO.File.ReadAllLines(archivo)
        Dim conexionbase As String = "Server=localhost;database=asistencia;user=root;password='';"

        For Each line As String In lines
            Dim parts() As String = line.Split(vbTab)
            If parts.Length >= 2 Then
                Dim id As String = parts(0).Trim()
                Dim fechaHora As String = parts(1).Trim()
                Dim fecha As String = ""
                Dim hora As String = ""

                Dim fechaHoraParts() As String = fechaHora.Split(" "c)
                If fechaHoraParts.Length = 2 Then
                    fecha = fechaHoraParts(0)
                    hora = fechaHoraParts(1)
                Else
                    MessageBox.Show("Formato de fecha y hora inválido en la línea: " & line)
                    Continue For
                End If

                Try
                    Using miconexion As New MySqlConnection(conexionbase)
                        miconexion.Open()
                        Dim query As String = "INSERT INTO asistencia (id, fecha, hora) VALUES (@id, @fecha, @hora)"
                        Using cmd As New MySqlCommand(query, miconexion)
                            cmd.Parameters.AddWithValue("@id", id)
                            cmd.Parameters.AddWithValue("@fecha", fecha)
                            cmd.Parameters.AddWithValue("@hora", hora)
                            cmd.ExecuteNonQuery()
                        End Using
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error al insertar el registro: " & ex.Message)
                End Try
            Else
                MessageBox.Show("Línea inválida en el archivo: " & line)
            End If
        Next
        MessageBox.Show("Datos insertados correctamente.")
    End Sub
End Class
