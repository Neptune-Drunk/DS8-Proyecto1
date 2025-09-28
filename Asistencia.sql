-- Crear base de datos
CREATE DATABASE Asistencia;
USE Asistencia;

-- Tabla de empleados (con datos precargados)
CREATE TABLE empleados (
    codigo_marcacion INT PRIMARY KEY,
    nombre VARCHAR(50) NOT NULL
);

INSERT INTO empleados (codigo_marcacion, nombre) VALUES
(13, 'Juan Pérez'),
(2, 'María González'),
(11, 'Carlos López'),
(7, 'Ana Rodríguez'),
(31, 'Pedro Sánchez'),
(3, 'Lucía Torres'),
(6, 'José Martínez'),
(8, 'Laura Herrera'),
(5, 'Miguel Díaz'),
(30, 'Sofía Jiménez'),
(4, 'Diego Castro'),
(9, 'Camila Vargas'),
(36, 'Andrés Rivas'),
(12, 'Elena Navarro'),
(45, 'Felipe Ortega'),
(41, 'Gabriela Méndez'),
(15, 'Ricardo Fuentes'),
(26, 'Isabel Romero'),
(21, 'Tomás Silva'),
(22, 'Valeria Morales'),
(40, 'Martín Cruz'),
(16, 'Claudia Vega'),
(23, 'Daniela Herrera'),
(18, 'Francisco León'),
(42, 'Paula Castillo'),
(19, 'Jorge Molina'),
(33, 'Natalia Ruiz');

-- Empleados faltantes detectados en los registros (agregados)
INSERT INTO empleados (codigo_marcacion, nombre) VALUES
(1,  'Empleado 1'),
(10, 'Empleado 10'),
(14, 'Empleado 14'),
(17, 'Empleado 17'),
(20, 'Empleado 20'),
(24, 'Empleado 24'),
(25, 'Empleado 25'),
(27, 'Empleado 27'),
(28, 'Empleado 28'),
(29, 'Empleado 29'),
(35, 'Empleado 35'),
(37, 'Empleado 37'),
(39, 'Empleado 39'),
(43, 'Empleado 43'),
(44, 'Empleado 44'),
(46, 'Empleado 46'),
(47, 'Empleado 47');

-- Tabla de marcaciones (vacía al inicio)
-- Se divide 'horario' en 'fecha' (DATE) y 'hora' (VARCHAR)
CREATE TABLE marcaciones (
    id INT AUTO_INCREMENT PRIMARY KEY,
    fecha DATE,
    hora VARCHAR(20),
    codigo_marcacion INT,
    CONSTRAINT fk_marcaciones_empleados
        FOREIGN KEY (codigo_marcacion) REFERENCES empleados(codigo_marcacion)
);

-- Tabla para registrar días libres (actualizada para compatibilidad con xasistencia)
CREATE TABLE dias_libres (
    fecha DATE NOT NULL UNIQUE,
    detalle VARCHAR(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

-- Datos reales de días libres (extraídos de xasistencia)
INSERT INTO dias_libres (fecha, detalle) VALUES
('2023-01-02', 'Año Nuevo - Día Puente'),
('2023-01-09', 'Día de los Martires'),
('2023-02-20', 'Lunes de Carnaval'),
('2023-02-21', 'Martes de Carnaval'),
('2023-02-22', 'Miércoles de Ceniza'),
('2023-03-15', 'Clases Suspendidas - MEDUCA'),
('2023-04-06', 'Jueves Santo'),
('2023-04-07', 'Viernes Santo'),
('2023-05-01', 'Día del Trabajo'),
('2023-06-12', 'Cerrado por Limpieza - Votaciones'),
('2023-12-08', 'Día de la Madre'),
('2023-12-20', 'Duelo Nacional '),
('2023-12-25', 'Navidad'),
('2024-01-01', 'Año Nuevo'),
('2024-01-09', 'Martires'),
('2024-02-12', 'Carnaval'),
('2024-02-13', 'Carnaval'),
('2024-02-14', 'Miércoles de Ceniza'),
('2024-03-01', 'Misa de Inicio de Año Escolar 2024'),
('2024-03-28', 'Jueves Santo'),
('2024-03-29', 'Viernes Santo'),
('2024-05-01', 'Día del Trabajador'),
('2024-05-06', 'Elecciones 2024 (día después)'),
('2024-07-01', 'Toma de Posesión Presidencial'),
('2024-09-12', 'Fundación - La Chorrera'),
('2024-11-04', 'Fiestas Patrias'),
('2024-11-05', 'Fiestas Patrias'),
('2024-11-06', 'Cierre de Escuelas por mal tiempo'),
('2024-11-11', 'Fiestas Patrias - Puente'),
('2024-11-28', 'Fiestas Patrias'),
('2024-11-29', 'Día del Maestro'),
('2024-12-09', 'Día Puente - Día de las Madres'),
('2024-12-20', 'Duelo Nacional'),
('2024-12-24', 'Noche Buena'),
('2024-12-25', 'Navidad'),
('2025-01-01', 'Año Nuevo');

-- Índice único para la fecha
ALTER TABLE dias_libres ADD UNIQUE KEY indice (fecha) USING BTREE;
