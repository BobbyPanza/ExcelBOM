-- Crea tabelle di supporto e seed mappings per ExcelBOM

-- Xt_ExcelBom_Mappings: mapping codici Excel → codici sistema
IF OBJECT_ID('dbo.Xt_ExcelBom_Mappings', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.Xt_ExcelBom_Mappings (
        IDMapping     INT IDENTITY(1,1) PRIMARY KEY,
        Ambito        NVARCHAR(50)  NOT NULL,
        CodiceExcel   NVARCHAR(100) NOT NULL,
        CodiceSistema NVARCHAR(50)  NOT NULL
    );
END

-- Xt_ExcelBom_LogRun: log delle importazioni
-- IDLog è il PK auto-generato; IDRun referenzia Xt_ExcelBom_Run
IF OBJECT_ID('dbo.Xt_ExcelBom_LogRun', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.Xt_ExcelBom_LogRun (
        IDLog      INT IDENTITY(1,1) PRIMARY KEY,
        IDRun      INT,
        Sorgente   NVARCHAR(10)  DEFAULT 'BOM',
        DataInizio DATETIME,
        DataFine   DATETIME,
        Stato      NVARCHAR(50),
        Errore     NVARCHAR(MAX)
    );
END
ELSE
BEGIN
    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Xt_ExcelBom_LogRun') AND name = 'IDRun')
        ALTER TABLE dbo.Xt_ExcelBom_LogRun ADD IDRun INT;
    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Xt_ExcelBom_LogRun') AND name = 'Sorgente')
        ALTER TABLE dbo.Xt_ExcelBom_LogRun ADD Sorgente NVARCHAR(10) DEFAULT 'BOM';
END

-- Seed mappings fasi ciclo (idempotente)
IF NOT EXISTS (SELECT 1 FROM dbo.Xt_ExcelBom_Mappings WHERE Ambito = 'FASE_CICLO' AND CodiceExcel = 'Taglio')
    INSERT INTO dbo.Xt_ExcelBom_Mappings (Ambito, CodiceExcel, CodiceSistema) VALUES ('FASE_CICLO', 'Taglio', 'TAGLIO');
IF NOT EXISTS (SELECT 1 FROM dbo.Xt_ExcelBom_Mappings WHERE Ambito = 'FASE_CICLO' AND CodiceExcel = 'Piegatura')
    INSERT INTO dbo.Xt_ExcelBom_Mappings (Ambito, CodiceExcel, CodiceSistema) VALUES ('FASE_CICLO', 'Piegatura', 'PIEGA');
IF NOT EXISTS (SELECT 1 FROM dbo.Xt_ExcelBom_Mappings WHERE Ambito = 'FASE_CICLO' AND CodiceExcel = 'Saldatura')
    INSERT INTO dbo.Xt_ExcelBom_Mappings (Ambito, CodiceExcel, CodiceSistema) VALUES ('FASE_CICLO', 'Saldatura', 'SALDA');
IF NOT EXISTS (SELECT 1 FROM dbo.Xt_ExcelBom_Mappings WHERE Ambito = 'FASE_CICLO' AND CodiceExcel = 'Verniciatura')
    INSERT INTO dbo.Xt_ExcelBom_Mappings (Ambito, CodiceExcel, CodiceSistema) VALUES ('FASE_CICLO', 'Verniciatura', 'VERNI');
