CREATE DATABASE ProyectoFinal1_DB;
GO

USE ProyectoFinal1_DB;
GO

CREATE TABLE InformacionIA (
    Id INT IDENTITY(1,1) PRIMARY KEY,
	Prompt NVARCHAR(MAX),
    Resultado NVARCHAR(MAX) NOT NULL,
    Fecha_Investigacion DATETIME DEFAULT GETDATE()
);

SELECT * FROM InformacionIA