/*
  Permisos de SOLO LECTURA para la cuenta de dominio CONSULTORIA\mesa-lectura-sap
  Servidor: srv-sap (10.0.0.167), SQL Server 2019
  Ticket Odoo 61139 / ERP-PARTNER SAPB122794. Autorizado por el gerente (23-nov-2026).
  Alcance autorizado: SELECT únicamente sobre OUSR (usuarios y último ingreso) y USR5 (bitácora de accesos).

  ORDEN: 1) correr en la BASE DE PRUEBAS (reemplazar <<BD_PRUEBAS>>), validar con el paso 3;
         2) repetir en producción (SBO_COMPANY_TEST).
  Este script lo ejecuta una persona en el Management Studio; el MCP no escribe nada.
*/

-- 1) Login a nivel de servidor (una sola vez por servidor)
USE [master];
IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = N'CONSULTORIA\mesa-lectura-sap')
    CREATE LOGIN [CONSULTORIA\mesa-lectura-sap] FROM WINDOWS WITH DEFAULT_DATABASE = [master];
GO

-- 2) Usuario y SELECT en la base (cambiar el nombre de la base según el ambiente)
USE [<<BD_PRUEBAS>>];            -- luego: USE [SBO_COMPANY_TEST];
GO
IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'CONSULTORIA\mesa-lectura-sap')
    CREATE USER [CONSULTORIA\mesa-lectura-sap] FOR LOGIN [CONSULTORIA\mesa-lectura-sap];
GRANT SELECT ON dbo.OUSR TO [CONSULTORIA\mesa-lectura-sap];
GRANT SELECT ON dbo.USR5 TO [CONSULTORIA\mesa-lectura-sap];
GO

-- 3) Validación (debe dar select_ousr = 1, select_usr5 = 1, insert_ousr = 0)
EXECUTE AS LOGIN = N'CONSULTORIA\mesa-lectura-sap';
SELECT SUSER_SNAME() AS login, DB_NAME() AS base,
       HAS_PERMS_BY_NAME('dbo.OUSR','OBJECT','SELECT') AS select_ousr,
       HAS_PERMS_BY_NAME('dbo.USR5','OBJECT','SELECT') AS select_usr5,
       HAS_PERMS_BY_NAME('dbo.OUSR','OBJECT','INSERT') AS insert_ousr;
REVERT;
GO

/* Para quitar el acceso:
   DROP USER [CONSULTORIA\mesa-lectura-sap];   -- en cada base
   DROP LOGIN [CONSULTORIA\mesa-lectura-sap];  -- en master
*/
