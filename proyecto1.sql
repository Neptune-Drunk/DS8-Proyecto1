-- Crear base de datos
CREATE DATABASE proyecto1;
USE proyecto1;

-- Tabla de empleados (catálogo)
CREATE TABLE empleados (
    codigo_marcacion INT PRIMARY KEY,
    nombre VARCHAR(50) NOT NULL
);

-- Tabla de marcaciones (sin registros iniciales)
-- Se divide 'horario' en 'fecha' (DATE) y 'hora' (texto para permitir rangos como '7:00 - 12:00')
CREATE TABLE marcaciones (
    id INT AUTO_INCREMENT PRIMARY KEY,
    fecha DATE,
    hora VARCHAR(20),
    codigo_marcacion INT,
    CONSTRAINT fk_marcaciones_empleados
        FOREIGN KEY (codigo_marcacion) REFERENCES empleados(codigo_marcacion)
);
