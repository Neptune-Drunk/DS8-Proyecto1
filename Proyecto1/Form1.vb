Imports System.Data.SqlClient
Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions
Imports System.Globalization

Public Class Form1
    Dim archivo As String
    Dim conexion As String = "Server=localhost;Database=Asistencia;User=root;Password=;"
    Dim miconexion As MySqlConnection
    Dim archivoYaCargado As Boolean = False

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ' Configurar apariencia del formulario
            ConfigurarAparienciaFormulario()

            ' Deshabilitar el botón insertar al inicio
            BttnInsertar.Enabled = False

            miconexion = New MySqlConnection(conexion)
            miconexion.Open()

            ' Test the connection with a simple query
            Using cmd As New MySqlCommand("SELECT 1", miconexion)
                cmd.ExecuteScalar()
            End Using

            MessageBox.Show("Conexión establecida exitosamente", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
            CargarDatos()

        Catch ex As MySqlException
            MessageBox.Show($"Error de Conexión: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If miconexion IsNot Nothing AndAlso miconexion.State = ConnectionState.Open Then
                miconexion.Close()
            End If
        End Try
    End Sub

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
        ConfigurarBoton(Bttnarchivo, Color.FromArgb(52, 152, 219), Color.White)
        ConfigurarBoton(BttnInsertar, Color.FromArgb(46, 204, 113), Color.White)

        ' El botón limpiar ya tiene colores configurados en el designer
        btnLimpiar.Font = New Font("Segoe UI", 9, FontStyle.Bold)

        ' Configurar etiquetas
        lblProgreso.Font = New Font("Segoe UI", 9)
        lblProgreso.ForeColor = Color.FromArgb(52, 73, 94)

        ' Configurar barra de progreso
        progressBar.ForeColor = Color.FromArgb(52, 152, 219)
    End Sub

    ' Método para configurar el estilo de los botones
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

    Private Sub Bttnarchivo_Click(sender As Object, e As EventArgs) Handles Bttnarchivo.Click
        ' Verificar si ya hay un archivo cargado
        If archivoYaCargado Then
            MessageBox.Show("Ya tienes un archivo adjuntado. Debes limpiar la tabla antes de adjuntar un nuevo archivo.",
                           "Archivo ya cargado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Title = "Seleccione el archivo"
        openFileDialog.Filter = "Archivos DAT (*.dat)|*.dat"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            archivo = openFileDialog.FileName
            MessageBox.Show("Archivo seleccionado: " & archivo)
            ' Habilitar el botón insertar una vez que se selecciona un archivo
            BttnInsertar.Enabled = True
        End If
    End Sub

    Private Sub BttnInsertar_Click(sender As Object, e As EventArgs) Handles BttnInsertar.Click
        If String.IsNullOrEmpty(archivo) Then
            MessageBox.Show("Por favor, seleccione un archivo primero.")
            Return
        End If

        ' Validar que el archivo existe
        If Not System.IO.File.Exists(archivo) Then
            MessageBox.Show("El archivo seleccionado no existe.")
            Return
        End If

        Dim lines() As String
        Try
            lines = System.IO.File.ReadAllLines(archivo)
        Catch ex As Exception
            MessageBox.Show($"Error al leer el archivo: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try

        If lines.Length = 0 Then
            MessageBox.Show("El archivo está vacío.")
            Return
        End If

        ' Mostrar barra de progreso
        progressBar.Minimum = 0
        progressBar.Maximum = lines.Length
        progressBar.Value = 0
        progressBar.Visible = True
        lblProgreso.Visible = True
        lblProgreso.Text = "Iniciando procesamiento..."

        ' Deshabilitar botones para evitar múltiples procesos
        Bttnarchivo.Enabled = False
        BttnInsertar.Enabled = False
        btnLimpiar.Enabled = False

        Dim successCount As Integer = 0
        Dim errorCount As Integer = 0
        Dim duplicateCount As Integer = 0
        Dim skippedCount As Integer = 0
        Dim processedCount As Integer = 0

        ' NUEVAS VARIABLES DE DIAGNÓSTICO
        Dim lineCount As Integer = 0
        Dim emptyLines As Integer = 0
        Dim wrongFormat As Integer = 0
        Dim invalidCodes As Integer = 0
        Dim nonExistentEmployees As Integer = 0

        ' Para rastrear códigos no encontrados
        Dim codigosNoEncontrados As New HashSet(Of Integer)
        Dim codigosEncontrados As New HashSet(Of Integer)

        Try
            Using connection As New MySqlConnection(conexion)
                connection.Open()

                ' OPTIMIZACIÓN ULTRA-RÁPIDA: Cargar códigos de empleados válidos
                lblProgreso.Text = "Cargando empleados válidos..."
                Application.DoEvents()

                Dim validEmployeeCodes As New HashSet(Of Integer)
                Using empCmd As New MySqlCommand("SELECT codigo_marcacion FROM empleados", connection)
                    Using reader = empCmd.ExecuteReader()
                        While reader.Read()
                            validEmployeeCodes.Add(reader.GetInt32(0))
                        End While
                    End Using
                End Using

                ' Cargar duplicados existentes de una vez
                lblProgreso.Text = "Cargando registros existentes..."
                Application.DoEvents()

                Dim existingRecords As New HashSet(Of String)
                Using loadCmd As New MySqlCommand("SELECT CONCAT(codigo_marcacion, '|', DATE(fecha), '|', hora) FROM marcaciones", connection)
                    Using reader = loadCmd.ExecuteReader()
                        While reader.Read()
                            existingRecords.Add(reader.GetString(0))
                        End While
                    End Using
                End Using

                ' Procesar archivo en memoria primero (SÚPER RÁPIDO)
                lblProgreso.Text = "Procesando archivo en memoria..."
                Application.DoEvents()

                Dim validRecords As New List(Of Object())
                Dim registrosUnicos As New HashSet(Of String)

                For Each line As String In lines
                    processedCount += 1
                    lineCount += 1

                    ' Actualizar progreso cada 5000 líneas (menos frecuente = más rápido)
                    If processedCount Mod 5000 = 0 OrElse processedCount = lines.Length Then
                        progressBar.Value = processedCount
                        lblProgreso.Text = $"Procesando en memoria: {processedCount} de {lines.Length}..."
                        Application.DoEvents()
                    End If

                    ' Saltar líneas vacías
                    If String.IsNullOrWhiteSpace(line) Then
                        skippedCount += 1
                        emptyLines += 1
                        Continue For
                    End If

                    ' Limpiar espacios extra y tokenizar por cualquier espacio/tab
                    Dim cleanLine As String = line.Trim()
                    Dim tokens As List(Of String) = Regex.Matches(cleanLine, "\S+") _
                                                       .Cast(Of Match)() _
                                                       .Select(Function(m) m.Value) _
                                                       .ToList()

                    ' Diagnóstico silencioso de primera línea (sin mostrar ventana)
                    If lineCount = 1 Then
                        ' Información guardada para logs internos si es necesario
                    End If

                    ' Necesitamos al menos: código y fecha-hora (que puede venir como 1 token o 2 tokens separados)
                    If tokens.Count < 2 Then
                        skippedCount += 1
                        wrongFormat += 1
                        Continue For
                    End If

                    ' Extraer código (primer token)
                    Dim codigoStr As String = tokens(0).Trim()

                    ' Construir la cadena de fecha-hora a partir de tokens
                    Dim fechaHoraStr As String = tokens(1)
                    ' Si la fecha y hora vienen separadas (tokens[1] = fecha, tokens[2] = hora), las unimos
                    If tokens.Count >= 3 AndAlso tokens(1).Contains("-") AndAlso tokens(2).Contains(":") Then
                        fechaHoraStr = tokens(1) & " " & tokens(2)
                    End If

                    ' Validar que no estén vacías
                    If String.IsNullOrWhiteSpace(codigoStr) OrElse String.IsNullOrWhiteSpace(fechaHoraStr) Then
                        errorCount += 1
                        Continue For
                    End If

                    ' Validar que el código sea numérico
                    Dim codigo As Integer
                    If Not Integer.TryParse(codigoStr, codigo) Then
                        errorCount += 1
                        invalidCodes += 1
                        Continue For
                    End If

                    ' NUEVO: Verificar que el empleado exista en la base de datos
                    If Not validEmployeeCodes.Contains(codigo) Then
                        skippedCount += 1
                        nonExistentEmployees += 1
                        codigosNoEncontrados.Add(codigo)
                        Continue For
                    Else
                        codigosEncontrados.Add(codigo)
                    End If

                    ' Parsear fecha-hora: primero intento exacto yyyy-MM-dd HH:mm:ss; si no, intento Date.Parse
                    Dim fechaHora As DateTime
                    Dim formatos() As String = {"yyyy-MM-dd HH:mm:ss", "yyyy/M/d HH:mm:ss", "dd/MM/yyyy HH:mm:ss", "MM/dd/yyyy HH:mm:ss"}
                    If Not DateTime.TryParseExact(fechaHoraStr.Trim(), formatos, CultureInfo.InvariantCulture, DateTimeStyles.None, fechaHora) Then
                        If Not DateTime.TryParse(fechaHoraStr, fechaHora) Then
                            errorCount += 1
                            wrongFormat += 1
                            Continue For
                        End If
                    End If
                    Dim fechaDate As Date = fechaHora.Date
                    Dim horaStr As String = fechaHora.ToString("HH:mm:ss")

                    ' Crear clave única para detectar duplicados
                    Dim claveUnica As String = $"{codigo}|{fechaDate:yyyy-MM-dd}|{horaStr}"

                    ' Verificar duplicados en memoria (SÚPER RÁPIDO)
                    If registrosUnicos.Contains(claveUnica) OrElse existingRecords.Contains(claveUnica) Then
                        duplicateCount += 1
                        Continue For
                    End If

                    ' Agregar a lista de registros válidos
                    validRecords.Add({codigo, fechaDate, horaStr})
                    registrosUnicos.Add(claveUnica)

                    ' DIAGNÓSTICO: Mostrar los primeros 5 registros válidos
                    If validRecords.Count <= 5 Then
                        ' (Información guardada para el reporte final)
                    End If
                Next

                ' DIAGNÓSTICO DETALLADO ANTES DE INSERTAR
                ' (Información guardada para mostrar en el reporte final)

                ' INSERCIÓN MASIVA ULTRA-RÁPIDA
                If validRecords.Count > 0 Then
                    ' Ajustar la barra de progreso para la fase de inserción
                    Try
                        progressBar.Maximum = processedCount + validRecords.Count
                        progressBar.Value = Math.Min(progressBar.Maximum, processedCount)
                    Catch ex As Exception
                        ' Si algo falla con la barra de progreso, la ocultamos para no bloquear el proceso
                        progressBar.Visible = False
                    End Try

                    lblProgreso.Text = $"Insertando {validRecords.Count} registros en la base de datos..."
                    Application.DoEvents()

                    Using transaction = connection.BeginTransaction()
                        Try
                            Dim insertCmd As New MySqlCommand("INSERT INTO marcaciones (codigo_marcacion, fecha, hora) VALUES (@codigo, @fecha, @hora)", connection, transaction)
                            insertCmd.Parameters.Add("@codigo", MySqlDbType.Int32)
                            insertCmd.Parameters.Add("@fecha", MySqlDbType.Date)
                            insertCmd.Parameters.Add("@hora", MySqlDbType.VarChar, 20)

                            For i As Integer = 0 To validRecords.Count - 1
                                Dim record = validRecords(i)
                                insertCmd.Parameters("@codigo").Value = record(0)
                                insertCmd.Parameters("@fecha").Value = record(1)
                                insertCmd.Parameters("@hora").Value = record(2)

                                ' DIAGNÓSTICO: Mostrar primer registro que se intenta insertar
                                If i = 0 Then
                                    ' (Información guardada para el reporte final)
                                End If

                                Try
                                    insertCmd.ExecuteNonQuery()
                                    successCount += 1
                                Catch ex As MySqlException
                                    ' (Error guardado para el reporte final)
                                    errorCount += 1
                                End Try                                ' Actualizar progreso cada 1000 inserciones
                                If i Mod 1000 = 0 OrElse i = validRecords.Count - 1 Then
                                    Try
                                        Dim nuevoValor As Integer = processedCount + (i + 1)
                                        progressBar.Value = Math.Min(progressBar.Maximum, Math.Max(progressBar.Minimum, nuevoValor))
                                        lblProgreso.Text = $"Insertando: {i + 1} de {validRecords.Count}..."
                                        Application.DoEvents()
                                    Catch ex As Exception
                                        ' Si el progress bar falla, lo ocultamos para no interrumpir
                                        progressBar.Visible = False
                                    End Try
                                End If
                            Next

                            ' Confirmar transacción
                            transaction.Commit()
                            lblProgreso.Text = "Guardando cambios..."
                            Application.DoEvents()

                        Catch ex As Exception
                            transaction.Rollback()
                            Throw ex
                        End Try
                    End Using
                End If
            End Using

            ' Ocultar barra de progreso
            progressBar.Visible = False
            lblProgreso.Visible = False

            ' Mostrar resultado simplificado
            If successCount > 0 Then
                MessageBox.Show($"✅ Procesamiento completado exitosamente!{vbNewLine}{vbNewLine}" &
                               $"📊 Registros insertados: {successCount:N0}{vbNewLine}" &
                               $"❌ Duplicados omitidos: {duplicateCount:N0}{vbNewLine}" &
                               $"⚠️ Errores: {errorCount:N0}",
                               "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ' Marcar que ya se cargó un archivo exitosamente
                archivoYaCargado = True
                CargarDatos()
            Else
                MessageBox.Show("❌ No se insertaron registros. Verifique el formato del archivo.",
                               "Sin cambios", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        Catch ex As MySqlException
            progressBar.Visible = False
            lblProgreso.Visible = False
            MessageBox.Show($"Error de conexión a la base de datos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            progressBar.Visible = False
            lblProgreso.Visible = False
            MessageBox.Show($"Error inesperado: {ex.Message}{vbNewLine}Línea: {processedCount}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Rehabilitar botones
            Bttnarchivo.Enabled = True
            btnLimpiar.Enabled = True
            ' Solo habilitar Insertar si NO se cargó exitosamente el archivo
            If Not archivoYaCargado Then
                BttnInsertar.Enabled = True
            End If
        End Try
    End Sub

    ' Método para cargar los datos en el DataGridView
    Private Sub CargarDatos()
        Try
            Using connection As New MySqlConnection(conexion)
                connection.Open()
                ' Mejorar la consulta con nombres más descriptivos y orden lógico
                Dim query As String = "SELECT " &
                                     "e.nombre AS 'Empleado', " &
                                     "m.codigo_marcacion AS 'Código', " &
                                     "DATE_FORMAT(m.fecha, '%d/%m/%Y') AS 'Fecha', " &
                                     "m.hora AS 'Hora', " &
                                     "m.id AS 'ID' " &
                                     "FROM marcaciones m " &
                                     "LEFT JOIN empleados e ON m.codigo_marcacion = e.codigo_marcacion " &
                                     "ORDER BY m.fecha DESC, m.hora DESC"
                Dim adapter As New MySqlDataAdapter(query, connection)
                Dim dt As New DataTable()
                adapter.Fill(dt)
                dgvDatos.DataSource = dt

                ' Configurar apariencia del DataGridView
                ConfigurarDataGridView()
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al cargar los datos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Método para configurar la apariencia del DataGridView
    Private Sub ConfigurarDataGridView()
        With dgvDatos
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

            If .Columns.Count > 0 Then
                ' Configurar anchos y alineaciones específicas
                If .Columns.Contains("Empleado") Then
                    .Columns("Empleado").FillWeight = 35 ' Más espacio para nombres
                    .Columns("Empleado").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                End If

                If .Columns.Contains("Código") Then
                    .Columns("Código").FillWeight = 15
                    .Columns("Código").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                End If

                If .Columns.Contains("Fecha") Then
                    .Columns("Fecha").FillWeight = 20
                    .Columns("Fecha").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                End If

                If .Columns.Contains("Hora") Then
                    .Columns("Hora").FillWeight = 15
                    .Columns("Hora").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                End If

                If .Columns.Contains("ID") Then
                    .Columns("ID").FillWeight = 15
                    .Columns("ID").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                End If
            End If
        End With
    End Sub

    ' Método para limpiar la tabla de marcaciones
    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Dim result As DialogResult = MessageBox.Show("¿Está seguro de que desea eliminar TODOS los registros de marcaciones?", "Confirmar eliminación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

        If result = DialogResult.Yes Then
            Try
                Using connection As New MySqlConnection(conexion)
                    connection.Open()
                    Dim query As String = "DELETE FROM marcaciones"
                    Using cmd As New MySqlCommand(query, connection)
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                        MessageBox.Show($"Se eliminaron {rowsAffected} registros correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ' Resetear el estado para permitir nuevos archivos
                        archivoYaCargado = False
                        archivo = ""
                        BttnInsertar.Enabled = False
                        CargarDatos()
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show($"Error al limpiar la tabla: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub
    Private Sub ConsultaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsultaToolStripMenuItem.Click
        Dim consultaForm As New Form2()
        consultaForm.Show()
        Me.Hide()
    End Sub
End Class
