Imports MySql.Data.MySqlClient
Imports System.Data
Imports System.Globalization
Imports System.Linq

Public Class Form2
	Private ReadOnly conexion As String = "Server=localhost;Database=Asistencia;User=root;Password=;"

	Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		DataGridView1.AutoGenerateColumns = True
		DataGridView1.ReadOnly = True
		DataGridView1.AllowUserToAddRows = False
		DataGridView1.AllowUserToDeleteRows = False
		DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

		DateTimePicker1.Enabled = True
		CheckBox1.Checked = False

		ActualizarContadores(0, 0, 0, 0)
	End Sub

	Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
		If CheckBox1.Checked Then
			DateTimePicker1.Visible = True
			CheckBox2.Visible = True
		Else
			DateTimePicker1.Visible = False
			CheckBox2.Visible = False
			DateTimePicker2.Visible = False
		End If

		If CheckBox2.Checked And CheckBox1.Checked = True Then
			DateTimePicker2.Visible = True
		Else
			DateTimePicker2.Visible = False
        End If

    End Sub

	Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
		Dim nombreFiltro As String = TextBox1.Text.Trim()
		If String.IsNullOrWhiteSpace(nombreFiltro) Then
			MessageBox.Show("Ingrese el nombre del empleado.", "Falta nombre", MessageBoxButtons.OK, MessageBoxIcon.Warning)
			TextBox1.Focus()
			Return
		End If

		Dim fechaFiltro As Date? = Nothing
		If CheckBox1.Checked Then
			fechaFiltro = DateTimePicker1.Value.Date
		End If

		Try
			Using cn As New MySqlConnection(conexion)
				cn.Open()

				Dim empleados As New List(Of KeyValuePair(Of Integer, String))
				Using cmdEmp As New MySqlCommand("SELECT codigo_marcacion, nombre FROM empleados WHERE nombre LIKE @n", cn)
					cmdEmp.Parameters.AddWithValue("@n", "%" & nombreFiltro & "%")
					Using rd = cmdEmp.ExecuteReader()
						While rd.Read()
							empleados.Add(New KeyValuePair(Of Integer, String)(rd.GetInt32(0), rd.GetString(1)))
						End While
					End Using
				End Using

				If empleados.Count = 0 Then
					MessageBox.Show("No se encontró ningún empleado con ese nombre.", "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information)
					DataGridView1.DataSource = Nothing
					ActualizarContadores(0, 0, 0, 0)
					Return
				End If

				Dim dt As New DataTable()
				dt.Columns.Add("Empleado", GetType(String))
				dt.Columns.Add("Código", GetType(Integer))
				dt.Columns.Add("Fecha", GetType(Date))
				dt.Columns.Add("PrimeraMarca", GetType(String))
				dt.Columns.Add("ÚltimaMarca", GetType(String))
				dt.Columns.Add("Tarde", GetType(Boolean))

				Dim totalTardanzas As Integer = 0
				Dim totalAusencias As Integer = 0

				For Each emp In empleados
					Dim horario = ObtenerHorario(emp.Key)
					If horario Is Nothing Then
						' Si no hay horario definido para el código, solo mostramos las marcas sin calcular tardanza
						Using cmd As New MySqlCommand("SELECT fecha, MIN(TIME(hora)) AS primera, MAX(TIME(hora)) AS ultima FROM marcaciones WHERE codigo_marcacion=@c" & If(fechaFiltro.HasValue, " AND fecha=@f", "") & " GROUP BY fecha ORDER BY fecha", cn)
							cmd.Parameters.AddWithValue("@c", emp.Key)
							If fechaFiltro.HasValue Then cmd.Parameters.AddWithValue("@f", fechaFiltro.Value)
							Using rd = cmd.ExecuteReader()
								While rd.Read()
									Dim fecha As Date = rd.GetDateTime(0).Date
									Dim primeraObj = rd.GetValue(1)
									Dim ultimaObj = rd.GetValue(2)
									Dim primera As TimeSpan = If(TypeOf primeraObj Is TimeSpan, DirectCast(primeraObj, TimeSpan), TimeSpan.Parse(primeraObj.ToString()))
									Dim ultima As TimeSpan = If(TypeOf ultimaObj Is TimeSpan, DirectCast(ultimaObj, TimeSpan), TimeSpan.Parse(ultimaObj.ToString()))
									dt.Rows.Add(emp.Value, emp.Key, fecha, primera.ToString("hh\:mm\:ss"), ultima.ToString("hh\:mm\:ss"), False)
								End While
							End Using
						End Using
					Else
						Using cmd As New MySqlCommand("SELECT fecha, MIN(TIME(hora)) AS primera, MAX(TIME(hora)) AS ultima FROM marcaciones WHERE codigo_marcacion=@c" & If(fechaFiltro.HasValue, " AND fecha=@f", "") & " GROUP BY fecha ORDER BY fecha", cn)
							cmd.Parameters.AddWithValue("@c", emp.Key)
							If fechaFiltro.HasValue Then cmd.Parameters.AddWithValue("@f", fechaFiltro.Value)
							Dim huboFilas As Boolean = False
							Using rd = cmd.ExecuteReader()
								While rd.Read()
									huboFilas = True
									Dim fecha As Date = rd.GetDateTime(0).Date
									Dim primeraObj = rd.GetValue(1)
									Dim ultimaObj = rd.GetValue(2)
									Dim primera As TimeSpan = If(TypeOf primeraObj Is TimeSpan, DirectCast(primeraObj, TimeSpan), TimeSpan.Parse(primeraObj.ToString()))
									Dim ultima As TimeSpan = If(TypeOf ultimaObj Is TimeSpan, DirectCast(ultimaObj, TimeSpan), TimeSpan.Parse(ultimaObj.ToString()))

									Dim esTarde As Boolean = primera > horario.Item1
									If esTarde Then totalTardanzas += 1

									dt.Rows.Add(emp.Value, emp.Key, fecha, primera.ToString("hh\:mm\:ss"), ultima.ToString("hh\:mm\:ss"), esTarde)
								End While
							End Using

							' Si se filtró por fecha y no hubo marcas ese día, contamos ausencia
							If fechaFiltro.HasValue AndAlso Not huboFilas Then
								totalAusencias += 1
							End If
						End Using
					End If
				Next

				DataGridView1.DataSource = dt

				' Ajustes visuales rápidos
				If DataGridView1.Columns.Contains("Fecha") Then
					DataGridView1.Columns("Fecha").DefaultCellStyle.Format = "dd/MM/yyyy"
				End If

				ActualizarContadores(totalAusencias, totalTardanzas, totalTardanzas, 0)
			End Using
		Catch ex As MySqlException
			MessageBox.Show($"Error de base de datos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
		Catch ex As Exception
			MessageBox.Show($"Error inesperado: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
		End Try
	End Sub

	Private Sub ActualizarContadores(ausencias As Integer, tardanzas As Integer, injustificadas As Integer, justificadas As Integer)
		Label2.Text = $"Ausencias: {ausencias}"
		Label3.Text = $"Tardanzas: {tardanzas}"
		Label4.Text = $"Injustificadas: {injustificadas}"
		Label5.Text = $"Justificadas: {justificadas}"
	End Sub

	' Devuelve inicio y fin del horario configurado para el código
	' Item1 = hora inicio, Item2 = hora fin
	Private Function ObtenerHorario(codigo As Integer) As Tuple(Of TimeSpan, TimeSpan)
		Dim manana As Integer() = {13, 2, 11, 7, 31, 3, 6, 8, 5, 30, 4, 9, 36, 12, 45}
		Dim tarde As Integer() = {41, 15, 26, 21, 22, 40, 16, 23, 18, 42, 19, 33, 44}

		If manana.Contains(codigo) Then
			Return Tuple.Create(TimeSpan.ParseExact("07:00", "hh\:mm", CultureInfo.InvariantCulture),
								TimeSpan.ParseExact("12:00", "hh\:mm", CultureInfo.InvariantCulture))
		End If
		If tarde.Contains(codigo) Then
			Return Tuple.Create(TimeSpan.ParseExact("12:20", "hh\:mm", CultureInfo.InvariantCulture),
								TimeSpan.ParseExact("17:20", "hh\:mm", CultureInfo.InvariantCulture))
		End If

		Return Nothing
	End Function

    Private Sub InsertarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InsertarToolStripMenuItem.Click
        Dim formInsertar As New Form1()
		formInsertar.Show()
        Me.Hide()
    End Sub
    Private Sub SalirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem.Click
		Application.Exit()
	End Sub

	Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked Then
            DateTimePicker2.Visible = True
        Else
            DateTimePicker2.Visible = False
        End If
	End Sub
End Class