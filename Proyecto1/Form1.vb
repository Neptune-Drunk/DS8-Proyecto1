Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports MySql.Data.MySqlClient

Public Class Form1
    Dim archivo As String
    Dim conexion As String = "Server=localhost;Database=asistencia;User=root;Password=;"
    Dim miconexion As MySqlConnection

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            miconexion = New MySqlConnection(conexion)
            miconexion.Open()

            ' Test the connection with a simple query
            Using cmd As New MySqlCommand("SELECT 1", miconexion)
                cmd.ExecuteScalar()
            End Using

            MessageBox.Show("Conexión establecida exitosamente", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As MySqlException
            MessageBox.Show($"Error de Conexión: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If miconexion IsNot Nothing AndAlso miconexion.State = ConnectionState.Open Then
                miconexion.Close()
            End If
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
        Dim successCount As Integer = 0
        Dim errorCount As Integer = 0

        Try
            Using connection As New MySqlConnection(conexion)
                connection.Open()

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
                            errorCount += 1
                            Continue For
                        End If

                        Try
                            Dim query As String = "INSERT INTO asistencia (id, fecha, hora) VALUES (@id, @fecha, @hora)"
                            Using cmd As New MySqlCommand(query, connection)
                                cmd.Parameters.AddWithValue("@id", id)
                                cmd.Parameters.AddWithValue("@fecha", fecha)
                                cmd.Parameters.AddWithValue("@hora", hora)
                                cmd.ExecuteNonQuery()
                                successCount += 1
                            End Using
                        Catch ex As Exception
                            errorCount += 1
                        End Try
                    Else
                        errorCount += 1
                    End If
                Next
            End Using

            MessageBox.Show($"Proceso completado. Registros insertados: {successCount}, Errores: {errorCount}", "Resultado", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"Error al procesar el archivo: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
