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
