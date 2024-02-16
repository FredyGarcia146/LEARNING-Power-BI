CREATE DATABASE Prueba_TAWA
GO

USE Prueba_TAWA
GO

DROP TABLE [TBL_DETALLE_DATA]
GO

USE Prueba_TAWA
CREATE TABLE [dbo].[TBL_DETALLE_DATA](
	CANAL_VENTA		VARCHAR	(50)	NULL,
	CLIENTE			VARCHAR	(50)	NULL,
	NIVEL			VARCHAR	(20)	NULL,
	REGION			VARCHAR	(10)	NULL,
	DEPARTAMENTO	VARCHAR	(30)	NULL,
	CIUDAD			VARCHAR	(40)	NULL,
	LINEA			VARCHAR	(20)	NULL,
	MODELO			VARCHAR	(50)	NULL,
	STOCKActual		INT		NULL,
	SEMANA_1		INT		NULL,
	SEMANA_2		INT		NULL,
	SEMANA_3		INT		NULL,
	SEMANA_4		INT		NULL,
	SERVICIO		VARCHAR	(10)	NULL,
	SERVICIO_CALCULADO	VARCHAR	(10)	NULL
)
GO

DROP PROCEDURE SP_CARGA_DETALLE_DATOS
GO

CREATE PROCEDURE SP_CARGA_DETALLE_DATOS
@RUTA NVARCHAR(MAX) 
AS
BEGIN
	DECLARE @sentencia NVARCHAR(MAX) 
	DECLARE @campos NVARCHAR(MAX) 
	DECLARE @camcar NVARCHAR(MAX) 

	SET @sentencia = '('+''''+'Microsoft.ACE.OLEDB.12.0'++''''+','+''''+'Excel 12.0'+' Xml;HDR=YES;Database='+@RUTA+''''+','+''''+'select * from [DATA$]'+''''+');'
							
	SET @campos = 
	'(CANAL_VENTA,	
			CLIENTE,	
			NIVEL,	
			REGION,	
			DEPARTAMENTO,	
			CIUDAD,	
			LINEA,	
			MODELO,	
			STOCKActual,	
			SEMANA_1,	
			SEMANA_2,	
			SEMANA_3,	
			SEMANA_4,	
			SERVICIO,	
			SERVICIO_CALCULADO) '
	SET @camcar = 
	'SELECT ltrim(rtrim([CANAL VENTA])),
			ltrim(rtrim([CLIENTE])),
			ltrim(rtrim([NIVEL])),
			ltrim(rtrim([REGION])),
			ltrim(rtrim([DEPARTAMENTO])),
			ltrim(rtrim([CIUDAD])),
			ltrim(rtrim([LINEA])),
			ltrim(rtrim([MODELO])),
			ltrim(rtrim([STOCK_Actual])),
			ltrim(rtrim([SEMANA 1])),
			ltrim(rtrim([SEMANA 2])),
			ltrim(rtrim([SEMANA 3])),
			ltrim(rtrim([SEMANA 4])),
			ltrim(rtrim([SERVICIO])),
			ltrim(rtrim([SERVICIO1]))
			'

	SET @sentencia = 'INSERT INTO [dbo].[TBL_DETALLE_DATA] '+@campos + @camcar+' FROM OPENROWSET '+@sentencia 
	--SET @sentencia = 'SELECT * FROM OPENROWSET '+@sentencia 
	--SELECT @sentencia PRINT @sentencia
	exec (@sentencia)

	-- AJUSTE DEL CAMPO SERVICIO CON TABLA REFERIDA
	UPDATE A
	SET A.SERVICIO_CALCULADO = B.SERVICIO
	FROM [dbo].[TBL_DETALLE_DATA] A INNER JOIN TBL_MODELO_SERVICIO B 
	ON A.MODELO = B.MODELO

END

GO



--EJECUCION

TRUNCATE TABLE [dbo].[TBL_DETALLE_DATA]

EXEC SP_CARGA_DETALLE_DATOS 'C:\Users\and14\Desktop\Personal\Laboral_Trabajo\Trabajos_otros\Tawa\Prueba\Examen_Tecnico_BI.xlsx'

--CONSULTAS DE CARGADO
SELECT  * FROM [TBL_DETALLE_DATA]
SELECT * FROM TBL_MODELO_SERVICIO -- TABLA DE COMPLEMENTO PARA COMPLETAR LOS ESTADOS DE SERVICIO EN EL SP





-- Ajustes Configuracion OLEDB
EXEC master.sys.sp_MSset_oledb_prop
USE [master]  
GO  
EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'AllowInProcess', 1 
GO
EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'DynamicParameters', 1 
GO 
EXEC sp_configure 'show advanced options', 1
RECONFIGURE
GO
EXEC sp_configure 'ad hoc distributed queries', 1
RECONFIGURE
GO