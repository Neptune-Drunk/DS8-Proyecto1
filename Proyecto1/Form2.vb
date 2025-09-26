Imports MySql.Data.MySqlClient
Imports System.Data
Imports System.Globalization
Imports System.Linq
Imports System.Collections.Generic

Public Class Form2
	Private ReadOnly conexion As String = "Server=localhost;Database=Asistencia;User=root;Password=;"

	Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		Try
			' Configurar apariencia del formulario
			ConfigurarAparienciaFormulario()

			DataGridView1.AutoGenerateColumns = True
			DataGridView1.ReadOnly = True
			DataGridView1.AllowUserToAddRows = False
			DataGridView1.AllowUserToDeleteRows = False
			DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

			' Configurar DataGridView con estilo
			ConfigurarDataGridView()

			DateTimePicker1.Enabled = True
			CheckBox1.Checked = False

			ActualizarContadores(0, 0, 0, 0)

		Catch ex As Exception
			MessageBox.Show($"Error al cargar formulario: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
		End Try
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
		Dim useRange As Boolean = CheckBox1.Checked AndAlso CheckBox2.Checked
		Dim fechaInicio As Date = DateTimePicker1.Value.Date
		Dim fechaFin As Date = fechaInicio
		If CheckBox1.Checked AndAlso Not useRange Then
			fechaFiltro = fechaInicio
		End If
		If useRange Then
			fechaFin = DateTimePicker2.Value.Date
			If fechaFin < fechaInicio Then
				MessageBox.Show("La fecha 'Hasta' no puede ser menor que la fecha 'Desde'.", "Rango inválido", MessageBoxButtons.OK, MessageBoxIcon.Warning)
				Return
			End If
		End If

		Try
			Using cn As New MySqlConnection(conexion)
				cn.Open()

				Dim empleados As New List(Of KeyValuePair(Of Integer, String))
				Dim codigoBuscado As Integer
				Dim buscarPorCodigo As Boolean = Integer.TryParse(nombreFiltro, codigoBuscado)
				Dim sqlEmp As String = If(buscarPorCodigo,
					"SELECT codigo_marcacion, nombre FROM empleados WHERE codigo_marcacion=@v",
					"SELECT codigo_marcacion, nombre FROM empleados WHERE nombre LIKE @v")
				Using cmdEmp As New MySqlCommand(sqlEmp, cn)
					If buscarPorCodigo Then
						cmdEmp.Parameters.AddWithValue("@v", codigoBuscado)
					Else
						cmdEmp.Parameters.AddWithValue("@v", "%" & nombreFiltro & "%")
					End If
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
					Dim whereFecha As String = ""
					If useRange Then
						whereFecha = " AND fecha BETWEEN @d1 AND @d2"
					ElseIf fechaFiltro.HasValue Then
						whereFecha = " AND fecha=@f"
					End If

					Dim sql As String = "SELECT fecha, MIN(TIME(hora)) AS primera, MAX(TIME(hora)) AS ultima FROM marcaciones WHERE codigo_marcacion=@c" & whereFecha & " GROUP BY fecha ORDER BY fecha"

					If horario Is Nothing Then
						' Sin horario: solo mostramos marcas encontradas, no contamos ausencias
						Using cmd As New MySqlCommand(sql, cn)
							cmd.Parameters.AddWithValue("@c", emp.Key)
							If useRange Then
								cmd.Parameters.AddWithValue("@d1", fechaInicio)
								cmd.Parameters.AddWithValue("@d2", fechaFin)
							ElseIf fechaFiltro.HasValue Then
								cmd.Parameters.AddWithValue("@f", fechaFiltro.Value)
							End If

							Dim diasConMarca As New HashSet(Of Date)()
							Using rd = cmd.ExecuteReader()
								While rd.Read()
									Dim fecha As Date = rd.GetDateTime(0).Date
									diasConMarca.Add(fecha)
									Dim primeraObj = rd.GetValue(1)
									Dim ultimaObj = rd.GetValue(2)
									Dim primera As TimeSpan = If(TypeOf primeraObj Is TimeSpan, DirectCast(primeraObj, TimeSpan), TimeSpan.Parse(primeraObj.ToString()))
									Dim ultima As TimeSpan = If(TypeOf ultimaObj Is TimeSpan, DirectCast(ultimaObj, TimeSpan), TimeSpan.Parse(ultimaObj.ToString()))
									dt.Rows.Add(emp.Value, emp.Key, fecha, primera.ToString("hh\:mm\:ss"), ultima.ToString("hh\:mm\:ss"), False)
								End While
							End Using

							' No contamos ausencias para empleados sin horario asignado
						End Using
					Else
						Using cmd As New MySqlCommand(sql, cn)
							cmd.Parameters.AddWithValue("@c", emp.Key)
							If useRange Then
								cmd.Parameters.AddWithValue("@d1", fechaInicio)
								cmd.Parameters.AddWithValue("@d2", fechaFin)
							ElseIf fechaFiltro.HasValue Then
								cmd.Parameters.AddWithValue("@f", fechaFiltro.Value)
							End If

							Dim huboFilas As Boolean = False
							Dim diasConMarca As New HashSet(Of Date)()
							Using rd = cmd.ExecuteReader()
								While rd.Read()
									huboFilas = True
									Dim fecha As Date = rd.GetDateTime(0).Date
									diasConMarca.Add(fecha)
									Dim primeraObj = rd.GetValue(1)
									Dim ultimaObj = rd.GetValue(2)
									Dim primera As TimeSpan = If(TypeOf primeraObj Is TimeSpan, DirectCast(primeraObj, TimeSpan), TimeSpan.Parse(primeraObj.ToString()))
									Dim ultima As TimeSpan = If(TypeOf ultimaObj Is TimeSpan, DirectCast(ultimaObj, TimeSpan), TimeSpan.Parse(ultimaObj.ToString()))

									Dim esTarde As Boolean = primera > horario.Item1
									If esTarde Then totalTardanzas += 1

									dt.Rows.Add(emp.Value, emp.Key, fecha, primera.ToString("hh\:mm\:ss"), ultima.ToString("hh\:mm\:ss"), esTarde)
								End While
							End Using

							If useRange Then
								Dim totalDiasHabiles As Integer = ContarDiasHabiles(fechaInicio, fechaFin)
								Dim diasMarcadosHabiles As Integer = diasConMarca.Where(Function(d) EsDiaHabil(d)).Count()
								totalAusencias += Math.Max(0, totalDiasHabiles - diasMarcadosHabiles)
							ElseIf fechaFiltro.HasValue Then
								If EsDiaHabil(fechaInicio) AndAlso Not huboFilas Then
									totalAusencias += 1
								End If
							End If
						End Using
					End If
				Next

				DataGridView1.DataSource = dt

				' Configurar columnas específicas
				ConfigurarColumnasDataGrid()

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

	' Cuenta días hábiles (Lunes a Viernes) en el rango inclusivo
	Private Function ContarDiasHabiles(desde As Date, hasta As Date) As Integer
		Dim count As Integer = 0
		Dim d As Date = desde
		While d <= hasta
			If EsDiaHabil(d) Then count += 1
			d = d.AddDays(1)
		End While
		Return count
	End Function

	Private Function EsDiaHabil(d As Date) As Boolean
		Return d.DayOfWeek <> DayOfWeek.Saturday AndAlso d.DayOfWeek <> DayOfWeek.Sunday
	End Function

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

	' Método para configurar la apariencia general del formulario
	Private Sub ConfigurarAparienciaFormulario()
		' Configurar el formulario principal
		Me.BackColor = Color.FromArgb(236, 240, 241)
		Me.Font = New Font("Segoe UI", 9)

		' Configurar MenuStrip
		MenuStrip1.BackColor = Color.FromArgb(52, 73, 94)
		MenuStrip1.ForeColor = Color.White
		MenuToolStripMenuItem.ForeColor = Color.White

		' Configurar botones con estilo moderno
		ConfigurarBoton(Button1, Color.FromArgb(52, 152, 219), Color.White)

		' Configurar etiquetas de filtros
		Label1.Font = New Font("Segoe UI", 9, FontStyle.Bold)
		Label1.ForeColor = Color.FromArgb(52, 73, 94)

		' Configurar CheckBoxes
		CheckBox1.Font = New Font("Segoe UI", 9)
		CheckBox1.ForeColor = Color.FromArgb(52, 73, 94)
		CheckBox2.Font = New Font("Segoe UI", 9)
		CheckBox2.ForeColor = Color.FromArgb(52, 73, 94)

		' Configurar etiquetas de contadores con colores distintivos
		Label2.Font = New Font("Segoe UI", 10, FontStyle.Bold)
		Label2.ForeColor = Color.FromArgb(231, 76, 60) ' Rojo para ausencias
		Label3.Font = New Font("Segoe UI", 10, FontStyle.Bold)
		Label3.ForeColor = Color.FromArgb(230, 126, 34) ' Naranja para tardanzas
		Label4.Font = New Font("Segoe UI", 10, FontStyle.Bold)
		Label4.ForeColor = Color.FromArgb(192, 57, 43) ' Rojo oscuro para injustificadas
		Label5.Font = New Font("Segoe UI", 10, FontStyle.Bold)
		Label5.ForeColor = Color.FromArgb(39, 174, 96) ' Verde para justificadas

		' Configurar TextBox
		TextBox1.Font = New Font("Segoe UI", 9)
		TextBox1.BackColor = Color.White
		TextBox1.BorderStyle = BorderStyle.FixedSingle
	End Sub

	' Método para configurar el estilo de los botones
	Private Sub ConfigurarBoton(boton As Button, colorFondo As Color, colorTexto As Color)
		boton.BackColor = colorFondo
		boton.ForeColor = colorTexto
		boton.FlatStyle = FlatStyle.Flat
		boton.FlatAppearance.BorderSize = 0
		boton.Font = New Font("Segoe UI", 10, FontStyle.Bold)
		boton.Cursor = Cursors.Hand

		' Agregar eventos de hover
		AddHandler boton.MouseEnter, Sub() BotonMouseEnter(boton, colorFondo)
		AddHandler boton.MouseLeave, Sub() BotonMouseLeave(boton, colorFondo)
	End Sub

	' Eventos de hover para botones
	Private Sub BotonMouseEnter(boton As Button, colorOriginal As Color)
		If boton.Enabled Then
			boton.BackColor = Color.FromArgb(Math.Max(0, colorOriginal.R - 30),
										   Math.Max(0, colorOriginal.G - 30),
										   Math.Max(0, colorOriginal.B - 30))
		End If
	End Sub

	Private Sub BotonMouseLeave(boton As Button, colorOriginal As Color)
		If boton.Enabled Then
			boton.BackColor = colorOriginal
		End If
	End Sub

	' Método para configurar la apariencia del DataGridView
	Private Sub ConfigurarDataGridView()
		With DataGridView1
			' Configuraciones generales
			.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
			.RowHeadersVisible = False
			.SelectionMode = DataGridViewSelectionMode.FullRowSelect
			.MultiSelect = False
			.AllowUserToAddRows = False
			.AllowUserToDeleteRows = False
			.ReadOnly = True

			' Estilos de encabezados
			.EnableHeadersVisualStyles = False
			.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(52, 73, 94)
			.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
			.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 10, FontStyle.Bold)
			.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
			.ColumnHeadersHeight = 35

			' Estilos de celdas
			.DefaultCellStyle.BackColor = Color.White
			.DefaultCellStyle.ForeColor = Color.FromArgb(44, 62, 80)
			.DefaultCellStyle.Font = New Font("Segoe UI", 9)
			.DefaultCellStyle.SelectionBackColor = Color.FromArgb(52, 152, 219)
			.DefaultCellStyle.SelectionForeColor = Color.White

			' Filas alternadas
			.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 249, 250)

			' Bordes
			.GridColor = Color.FromArgb(189, 195, 199)
			.BorderStyle = BorderStyle.Fixed3D
		End With
	End Sub

	' Configurar columnas específicas del DataGridView
	Private Sub ConfigurarColumnasDataGrid()
		With DataGridView1
			If .Columns.Contains("Fecha") Then
				.Columns("Fecha").DefaultCellStyle.Format = "dd/MM/yyyy"
				.Columns("Fecha").FillWeight = 20
				.Columns("Fecha").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
			End If

			If .Columns.Contains("Empleado") Then
				.Columns("Empleado").FillWeight = 35
				.Columns("Empleado").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
			End If

			If .Columns.Contains("Código") Then
				.Columns("Código").FillWeight = 15
				.Columns("Código").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
			End If

			If .Columns.Contains("PrimeraMarca") Then
				.Columns("PrimeraMarca").FillWeight = 15
				.Columns("PrimeraMarca").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
			End If

			If .Columns.Contains("ÚltimaMarca") Then
				.Columns("ÚltimaMarca").FillWeight = 15
				.Columns("ÚltimaMarca").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
			End If

			If .Columns.Contains("Tarde") Then
				.Columns("Tarde").FillWeight = 10
				.Columns("Tarde").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
				
				' Colorear las filas con tardanzas
				For Each row As DataGridViewRow In .Rows
					If row.Cells("Tarde") IsNot Nothing AndAlso TypeOf row.Cells("Tarde").Value Is Boolean Then
						If CBool(row.Cells("Tarde").Value) Then
							row.DefaultCellStyle.BackColor = Color.FromArgb(255, 235, 235) ' Fondo rojizo claro
							row.DefaultCellStyle.ForeColor = Color.FromArgb(169, 68, 66) ' Texto rojizo
						End If
					End If
				Next
			End If
		End With
	End Sub

    Private Sub InsertarToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles InsertarToolStripMenuItem1.Click
        Dim formInsertar As New Form1()
		formInsertar.Show()
        Me.Hide()
    End Sub

    Private Sub DiasLibresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DiasLibresToolStripMenuItem.Click
        ' Aquí se puede implementar la funcionalidad para gestionar días libres/feriados
        MessageBox.Show("Funcionalidad de Días Libres en desarrollo", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
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