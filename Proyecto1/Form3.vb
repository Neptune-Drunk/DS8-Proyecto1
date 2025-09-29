Imports MySql.Data.MySqlClient

Public Class Form3
    ' Cadena de conexión a la base de datos xasistencia
    Private connectionString As String = "server=localhost;database=asistencia;uid=root;pwd=;"

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Configurar apariencia del formulario similar a Form1
        ConfigurarAparienciaFormulario()

        ' Configuración inicial del formulario
        Me.DateTimePicker1.Value = DateTime.Today
        Me.TextBox1.Focus()
    End Sub

    ' Método para configurar la apariencia general del formulario
    Private Sub ConfigurarAparienciaFormulario()
        ' Configurar el formulario principal con el mismo estilo que Form1
        Me.BackColor = Color.FromArgb(236, 240, 241)
        Me.Font = New Font("Segoe UI", 9)

        ' Configurar MenuStrip
        MenuStrip1.BackColor = Color.FromArgb(52, 73, 94)
        MenuStrip1.ForeColor = Color.White
        MenuToolStripMenuItem.ForeColor = Color.White
        InsertarToolStripMenuItem.ForeColor = Color.White
        ConsultaToolStripMenuItem.ForeColor = Color.White
        DiasLibresToolStripMenuItem.ForeColor = Color.White

        ' Configurar botón con estilo moderno (igual que Form1)
        ConfigurarBoton(Button1, Color.FromArgb(46, 204, 113), Color.White)
        ConfigurarBoton(Button2, Color.FromArgb(52, 152, 219), Color.White)

        ' Configurar etiquetas
        Label1.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        Label1.ForeColor = Color.FromArgb(52, 73, 94)
        Label2.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        Label2.ForeColor = Color.FromArgb(52, 73, 94)
        Label3.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        Label3.ForeColor = Color.FromArgb(52, 73, 94)

        ' Configurar TextBox
        TextBox1.Font = New Font("Segoe UI", 9)

        ' Configurar DataGridView
        ConfigurarDataGridView()
    End Sub

    ' Método para configurar el DataGridView
    Private Sub ConfigurarDataGridView()
        DataGridView1.Font = New Font("Segoe UI", 9)
        DataGridView1.BackgroundColor = Color.White
        DataGridView1.BorderStyle = BorderStyle.FixedSingle
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False
        DataGridView1.ReadOnly = True
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView1.MultiSelect = False

        ' Agregar eventos
        AddHandler DataGridView1.CellDoubleClick, AddressOf DataGridView1_CellDoubleClick
        AddHandler DataGridView1.KeyDown, AddressOf DataGridView1_KeyDown
    End Sub

    ' Evento para doble clic en una celda del DataGridView
    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 AndAlso DataGridView1.Rows.Count > 0 Then
            Dim fila As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
            Dim fecha As Date = Convert.ToDateTime(fila.Cells("Fecha").Value)
            Dim detalle As String = fila.Cells("Detalle").Value.ToString()

            ' Cargar datos en los controles para editar
            DateTimePicker1.Value = fecha
            TextBox1.Text = detalle
            TextBox1.Focus()

            MessageBox.Show($"Datos cargados para edición:{vbNewLine}{vbNewLine}" &
                          $"📅 Fecha: {fecha.ToString("dd/MM/yyyy")}{vbNewLine}" &
                          $"📝 Detalle: {detalle}{vbNewLine}{vbNewLine}" &
                          "Puede modificar el detalle y hacer clic en 'Grabar' para actualizar.",
                          "Día Libre Seleccionado", MessageBoxButtons.OK, MessageBoxIcon.Information)


        End If
    End Sub

    ' Evento para teclas en DataGridView
    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Escape Then
            ' Ocultar lista con Escape
            Button2_Click(Button2, EventArgs.Empty)
        ElseIf e.KeyCode = Keys.Enter AndAlso DataGridView1.SelectedRows.Count > 0 Then
            ' Enter para cargar datos
            Dim rowIndex As Integer = DataGridView1.SelectedRows(0).Index
            DataGridView1_CellDoubleClick(DataGridView1, New DataGridViewCellEventArgs(0, rowIndex))
        End If
    End Sub

    ' Método para configurar el estilo de los botones (igual que Form1)
    Private Sub ConfigurarBoton(boton As Button, colorFondo As Color, colorTexto As Color)
        boton.BackColor = colorFondo
        boton.ForeColor = colorTexto
        boton.FlatStyle = FlatStyle.Flat
        boton.FlatAppearance.BorderSize = 0
        boton.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        boton.Cursor = Cursors.Hand

        ' Agregar eventos de hover
        AddHandler boton.MouseEnter, Sub() BotonMouseEnter(boton, colorFondo)
        AddHandler boton.MouseLeave, Sub() BotonMouseLeave(boton, colorFondo)
    End Sub

    ' Eventos de hover para botones (igual que Form1)
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Validar que se haya ingresado un detalle
        If String.IsNullOrWhiteSpace(TextBox1.Text) Then
            MessageBox.Show("Por favor, ingrese el detalle del día libre.", "Campo requerido",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TextBox1.Focus()
            Return
        End If

        ' Obtener los valores de los controles
        Dim fechaDiaLibre As Date = DateTimePicker1.Value.Date
        Dim detalleDiaLibre As String = TextBox1.Text.Trim()

        Try
            Using connection As New MySqlConnection(connectionString)
                connection.Open()

                ' Verificar si ya existe un día libre para esa fecha
                Dim queryVerificar As String = "SELECT COUNT(*) FROM dias_libres WHERE fecha = @fecha"
                Using cmdVerificar As New MySqlCommand(queryVerificar, connection)
                    cmdVerificar.Parameters.AddWithValue("@fecha", fechaDiaLibre)
                    Dim count As Integer = Convert.ToInt32(cmdVerificar.ExecuteScalar())

                    If count > 0 Then
                        Dim resultado As DialogResult = MessageBox.Show(
                            $"Ya existe un día libre registrado para la fecha {fechaDiaLibre.ToString("dd/MM/yyyy")}.{vbNewLine}" &
                            "¿Desea actualizar el detalle?", "Fecha duplicada",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                        If resultado = DialogResult.Yes Then
                            ' Actualizar el registro existente
                            Dim queryActualizar As String = "UPDATE dias_libres SET detalle = @detalle WHERE fecha = @fecha"
                            Using cmdActualizar As New MySqlCommand(queryActualizar, connection)
                                cmdActualizar.Parameters.AddWithValue("@detalle", detalleDiaLibre)
                                cmdActualizar.Parameters.AddWithValue("@fecha", fechaDiaLibre)
                                cmdActualizar.ExecuteNonQuery()

                                MessageBox.Show($"Día libre actualizado correctamente:{vbNewLine}{vbNewLine}" &
                                              $"Fecha: {fechaDiaLibre.ToString("dd/MM/yyyy")}{vbNewLine}" &
                                              $"Detalle: {detalleDiaLibre}", "Actualización Exitosa",
                                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                            End Using
                        End If
                    Else
                        ' Insertar nuevo registro
                        Dim queryInsertar As String = "INSERT INTO dias_libres (fecha, detalle) VALUES (@fecha, @detalle)"
                        Using cmdInsertar As New MySqlCommand(queryInsertar, connection)
                            cmdInsertar.Parameters.AddWithValue("@fecha", fechaDiaLibre)
                            cmdInsertar.Parameters.AddWithValue("@detalle", detalleDiaLibre)
                            cmdInsertar.ExecuteNonQuery()

                            MessageBox.Show($"Día libre registrado exitosamente:{vbNewLine}{vbNewLine}" &
                                          $"Fecha: {fechaDiaLibre.ToString("dd/MM/yyyy")}{vbNewLine}" &
                                          $"Detalle: {detalleDiaLibre}", "Registro Exitoso",
                                          MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End Using
                    End If
                End Using
            End Using

            ' Limpiar controles después de operación exitosa
            TextBox1.Clear()
            DateTimePicker1.Value = DateTime.Today
            TextBox1.Focus()

        Catch ex As MySqlException
            MessageBox.Show($"Error de base de datos: {ex.Message}", "Error MySQL",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show($"Error inesperado: {ex.Message}", "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Alternar visibilidad del DataGridView y cargar datos
        If DataGridView1.Visible Then
            DataGridView1.Visible = False
            Button2.Text = "Ver Lista"
            Me.ClientSize = New Size(584, 200)
        Else
            CargarDiasLibres()
            DataGridView1.Visible = True
            Button2.Text = "Ocultar Lista"
            Me.ClientSize = New Size(584, 400)

            ' Mostrar estadísticas de días libres
            MostrarEstadisticasDiasLibres()
        End If
    End Sub

    Private Sub CargarDiasLibres()
        Try
            Using connection As New MySqlConnection(connectionString)
                connection.Open()

                Dim query As String = "SELECT fecha as 'Fecha', detalle as 'Detalle' FROM dias_libres ORDER BY fecha DESC"
                Using adapter As New MySqlDataAdapter(query, connection)
                    Dim dataTable As New DataTable()
                    adapter.Fill(dataTable)

                    DataGridView1.DataSource = dataTable

                    ' Configurar formato de columnas usando método común
                    ConfigurarColumnasDataGrid()
                End Using
            End Using
        Catch ex As MySqlException
            MessageBox.Show($"Error al cargar días libres: {ex.Message}", "Error MySQL",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show($"Error inesperado: {ex.Message}", "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub MostrarEstadisticasDiasLibres()
        Try
            Using connection As New MySqlConnection(connectionString)
                connection.Open()

                ' Obtener estadísticas
                Dim queryTotal As String = "SELECT COUNT(*) FROM dias_libres"
                Dim queryAnoActual As String = "SELECT COUNT(*) FROM dias_libres WHERE YEAR(fecha) = YEAR(CURDATE())"
                Dim queryProximosDias As String = "SELECT COUNT(*) FROM dias_libres WHERE fecha > CURDATE() AND YEAR(fecha) = YEAR(CURDATE())"

                Using cmdTotal As New MySqlCommand(queryTotal, connection)
                    Using cmdAnoActual As New MySqlCommand(queryAnoActual, connection)
                        Using cmdProximos As New MySqlCommand(queryProximosDias, connection)

                            Dim totalDias As Integer = Convert.ToInt32(cmdTotal.ExecuteScalar())
                            Dim diasAnoActual As Integer = Convert.ToInt32(cmdAnoActual.ExecuteScalar())
                            Dim diasProximos As Integer = Convert.ToInt32(cmdProximos.ExecuteScalar())

                            ' Estadísticas calculadas pero no mostradas en MessageBox


                        End Using
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' Error en estadísticas no es crítico, solo continúa
        End Try
    End Sub

    ' Método público para verificar si una fecha es día libre (útil para otros formularios)
    Public Function EsDiaLibre(fecha As Date) As Boolean
        Try
            Using connection As New MySqlConnection(connectionString)
                connection.Open()

                Dim query As String = "SELECT COUNT(*) FROM dias_libres WHERE fecha = @fecha"
                Using cmd As New MySqlCommand(query, connection)
                    cmd.Parameters.AddWithValue("@fecha", fecha.Date)
                    Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                    Return count > 0
                End Using
            End Using
        Catch ex As Exception
            ' En caso de error, asumir que no es día libre
            Return False
        End Try
    End Function

    ' Método público para obtener el detalle de un día libre específico
    Public Function ObtenerDetalleDiaLibre(fecha As Date) As String
        Try
            Using connection As New MySqlConnection(connectionString)
                connection.Open()

                Dim query As String = "SELECT detalle FROM dias_libres WHERE fecha = @fecha"
                Using cmd As New MySqlCommand(query, connection)
                    cmd.Parameters.AddWithValue("@fecha", fecha.Date)
                    Dim resultado = cmd.ExecuteScalar()
                    Return If(resultado IsNot Nothing, resultado.ToString(), "")
                End Using
            End Using
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Sub SalirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub InsertarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InsertarToolStripMenuItem.Click
        ' Navegar al formulario principal de inserción (Form1)
        Dim formInsertar As New Form1()
        formInsertar.Show()
        Me.Hide()
    End Sub

    Private Sub ConsultaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsultaToolStripMenuItem.Click
        ' Navegar al formulario de consulta (Form2)
        Dim consultaForm As New Form2()
        consultaForm.Show()
        Me.Hide()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        ' Permitir Enter para grabar directamente
        If e.KeyChar = Convert.ToChar(13) Then
            Button1_Click(sender, e)
        End If
    End Sub

    Private Sub DiasLibresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DiasLibresToolStripMenuItem.Click
        ' Mostrar menú de opciones avanzadas para días libres
        Dim opciones As String = "Opciones de Días Libres:" & vbNewLine & vbNewLine &
                               "• Ya está en el módulo de Días Libres" & vbNewLine &
                               "• Use 'Ver Lista' para consultar todos los días" & vbNewLine &
                               "• Doble clic en una fecha para editarla" & vbNewLine &
                               "• Presione ESC para ocultar la lista" & vbNewLine & vbNewLine &
                               "¿Desea buscar días libres por año específico?"

        Dim resultado As DialogResult = MessageBox.Show(opciones, "Días Libres - Opciones",
                                                       MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If resultado = DialogResult.Yes Then
            BuscarPorAno()
        End If
    End Sub

    Private Sub BuscarPorAno()
        Dim ano As String = InputBox("Ingrese el año para filtrar días libres:" & vbNewLine &
                                   "Ejemplo: 2024, 2025", "Buscar por Año", DateTime.Now.Year.ToString())

        If Not String.IsNullOrWhiteSpace(ano) AndAlso IsNumeric(ano) Then
            Try
                Dim anoInt As Integer = Convert.ToInt32(ano)

                Using connection As New MySqlConnection(connectionString)
                    connection.Open()

                    Dim query As String = "SELECT fecha as 'Fecha', detalle as 'Detalle' FROM dias_libres WHERE YEAR(fecha) = @ano ORDER BY fecha ASC"
                    Using adapter As New MySqlDataAdapter(query, connection)
                        adapter.SelectCommand.Parameters.AddWithValue("@ano", anoInt)
                        Dim dataTable As New DataTable()
                        adapter.Fill(dataTable)

                        If dataTable.Rows.Count > 0 Then
                            DataGridView1.DataSource = dataTable
                            DataGridView1.Visible = True
                            Button2.Text = "Ocultar Lista"
                            Me.ClientSize = New Size(584, 400)

                            ' Configurar columnas nuevamente
                            ConfigurarColumnasDataGrid()

                            MessageBox.Show($"Se encontraron {dataTable.Rows.Count} días libres para el año {ano}.",
                                          "Resultados de Búsqueda", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Else
                            MessageBox.Show($"No se encontraron días libres registrados para el año {ano}.",
                                          "Sin Resultados", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show($"Error en la búsqueda: {ex.Message}", "Error",
                              MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub ConfigurarColumnasDataGrid()
        If DataGridView1.Columns.Count > 0 Then
            DataGridView1.Columns("Fecha").Width = 120
            DataGridView1.Columns("Fecha").DefaultCellStyle.Format = "dd/MM/yyyy"
            DataGridView1.Columns("Fecha").HeaderText = "Fecha"
            DataGridView1.Columns("Detalle").Width = 320
            DataGridView1.Columns("Detalle").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            DataGridView1.Columns("Detalle").HeaderText = "Descripción del Día Libre"

            ' Estilo de encabezados
            DataGridView1.EnableHeadersVisualStyles = False
            DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(52, 73, 94)
            DataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
            DataGridView1.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)

            ' Alternar colores de filas
            DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240)
            DataGridView1.DefaultCellStyle.SelectionBackColor = Color.FromArgb(52, 152, 219)
            DataGridView1.DefaultCellStyle.SelectionForeColor = Color.White
        End If
    End Sub
End Class