-- Crea/ricrea tabelle staging per ExcelBOM

IF OBJECT_ID('dbo.Xt_ExcelBom_Run', 'U') IS NOT NULL
    DROP TABLE dbo.Xt_ExcelBom_Run;

CREATE TABLE dbo.Xt_ExcelBom_Run (
    IDRun INT IDENTITY(1,1) PRIMARY KEY,
    ImportDate DATETIME DEFAULT GETDATE()
);

IF OBJECT_ID('dbo.Xt_ExcelBom_Parts', 'U') IS NOT NULL
    DROP TABLE dbo.Xt_ExcelBom_Parts;

CREATE TABLE dbo.Xt_ExcelBom_Parts (
    IDRun INT,
    PACOD NVARCHAR(50),
    PADSC NVARCHAR(255),
    FMCOD NVARCHAR(50),
    PAUDM NVARCHAR(10),
    PATYP NVARCHAR(10),
    PATYF NVARCHAR(10),
    PRIMARY KEY (IDRun, PACOD)
);

IF OBJECT_ID('dbo.Xt_ExcelBom_Rel', 'U') IS NOT NULL
    DROP TABLE dbo.Xt_ExcelBom_Rel;

CREATE TABLE dbo.Xt_ExcelBom_Rel (
    IDRun INT,
    DBPAR NVARCHAR(50),
    DBCHL NVARCHAR(50),
    DBSEQ NVARCHAR(10),
    DBQTA NVARCHAR(50),
    DBUDM NVARCHAR(10),
    DBNT1 NVARCHAR(255),
    PRIMARY KEY (IDRun, DBPAR, DBCHL)
);

IF OBJECT_ID('dbo.Xt_ExcelBom_Phases', 'U') IS NOT NULL
    DROP TABLE dbo.Xt_ExcelBom_Phases;

CREATE TABLE dbo.Xt_ExcelBom_Phases (
    IDRun INT,
    PACOD NVARCHAR(50),
    FASEQ NVARCHAR(10),
    FACOD NVARCHAR(50),
    FADSC NVARCHAR(255),
    TempoAtt NVARCHAR(20),
    TempoLav NVARCHAR(20),
    PRIMARY KEY (IDRun, PACOD, FASEQ)
);