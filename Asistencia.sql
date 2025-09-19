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
