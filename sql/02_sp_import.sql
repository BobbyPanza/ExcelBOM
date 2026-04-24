-- Stored Procedure per importare da staging ExcelBOM

CREATE OR ALTER PROCEDURE dbo.Xsp_ExcelBom_Import
    @IDRun INT,
    @UpdDesc BIT = 0,
    @UpdPesi BIT = 0,
    @UpdCiclo BIT = 1,
    @UpdDistinta BIT = 1,
    @UpdGrandezze BIT = 0,
    @FamigliaDefault NVARCHAR(50) = 'Generica'
AS
BEGIN
    SET NOCOUNT ON;

    INSERT INTO Xt_ExcelBom_LogRun (Sorgente, IDRun, DataInizio, Stato)
    VALUES ('EXCEL', @IDRun, GETDATE(), 'IN PROGRESS');

    BEGIN TRY

        -- Upsert Anagrafiche (UPDATE + INSERT separati: A_PAR ha trigger INSTEAD OF)
        -- FMCOD: usa il valore Excel se esiste in A_FAM, altrimenti @FamigliaDefault
        UPDATE a
        SET a.PADSC = CASE WHEN @UpdDesc = 1 THEN s.PADSC ELSE a.PADSC END,
            a.FMCOD = CASE WHEN EXISTS (SELECT 1 FROM A_FAM WHERE FMCOD = s.FMCOD) THEN s.FMCOD ELSE @FamigliaDefault END,
            a.PAUDM = s.PAUDM,
            a.PATYP = s.PATYP,
            a.PATYF = s.PATYF
        FROM A_PAR a
        INNER JOIN Xt_ExcelBom_Parts s ON a.PACOD = s.PACOD
        WHERE s.IDRun = @IDRun;

        INSERT INTO A_PAR (PACOD, PADSC, FMCOD, PAUDM, PATYP, PATYF, PAMAG, PATGM, PAQTM, PAPDR, PAFDC, PACST, PACSA, PAPNT, PACSP, TMCSP, Approved, StoreConversionFactor)
        SELECT s.PACOD, s.PADSC,
               CASE WHEN EXISTS (SELECT 1 FROM A_FAM WHERE FMCOD = s.FMCOD) THEN s.FMCOD ELSE @FamigliaDefault END,
               s.PAUDM, s.PATYP, s.PATYF, 'Y', 0, 0, 0, 1, 0, 0, 0, 0, 0, 'N', 1
        FROM Xt_ExcelBom_Parts s
        WHERE s.IDRun = @IDRun
          AND NOT EXISTS (SELECT 1 FROM A_PAR WHERE PACOD = s.PACOD);

        -- Upsert Distinta (UPDATE + INSERT separati)
        IF @UpdDistinta = 1
        BEGIN
            UPDATE a
            SET a.DBSEQ = s.DBSEQ,
                a.DBQTA = s.DBQTA,
                a.DBUDM = s.DBUDM,
                a.DBNT1 = s.DBNT1
            FROM A_DBS a
            INNER JOIN Xt_ExcelBom_Rel s ON a.DBPAR = s.DBPAR AND a.DBCHL = s.DBCHL
            WHERE s.IDRun = @IDRun;

            INSERT INTO A_DBS (DBPAR, DBCHL, DBSEQ, DBQTA, DBUDM, DBNT1, DBSFR, QTRCL)
            SELECT s.DBPAR, s.DBCHL, s.DBSEQ, s.DBQTA, s.DBUDM, s.DBNT1, 0, 0
            FROM Xt_ExcelBom_Rel s
            WHERE s.IDRun = @IDRun
              AND NOT EXISTS (SELECT 1 FROM A_DBS WHERE DBPAR = s.DBPAR AND DBCHL = s.DBCHL);
        END

        -- Upsert Ciclo
        IF @UpdCiclo = 1
        BEGIN
            -- Crea A_CCL per articoli che non ce l'hanno
            INSERT INTO A_CCL (CCCOD, CCDSC, CCABL, CCCMP)
            SELECT DISTINCT p.PACOD, p.PADSC, 'Y', 'N'
            FROM Xt_ExcelBom_Parts p
            LEFT JOIN A_CCL c ON c.CCCOD = p.PACOD
            WHERE p.IDRun = @IDRun AND c.CCCOD IS NULL;

            -- Elimina fasi obsolete: solo per cicli che hanno ALMENO UNA fase nel nuovo Excel
            -- (se un ciclo non compare in Xt_ExcelBom_Phases non viene toccato)
            -- Condizione 1: eseguito solo se @UpdCiclo = 1 (già garantito dall'IF esterno)
            -- Condizione 2: la subquery restituisce solo PACOD con COUNT >= 1
            DELETE f
            FROM A_FAC f
            WHERE f.CCCOD IN (
                      SELECT PACOD FROM Xt_ExcelBom_Phases
                      WHERE IDRun = @IDRun
                      GROUP BY PACOD HAVING COUNT(*) > 0)
              AND NOT EXISTS (
                  SELECT 1 FROM Xt_ExcelBom_Phases ph
                  WHERE ph.IDRun = @IDRun
                    AND ph.PACOD = f.CCCOD
                    AND TRY_CAST(ph.FASEQ AS INT) = f.FASEQ);

            -- Inserisci fasi nuove in A_FAC
            -- Dati risorsa/costo/utilizzo presi da A_FAS; DurataH/DurataMin dall'Excel
            -- IDFAS = progressivo per ciclo: MAX esistente + ROW_NUMBER sulle nuove righe
            INSERT INTO A_FAC (
                CCCOD, IDFAS, FASEQ, FAIDN,
                FACOD, FADSC, FAATT, FAABV,
                FATYP, FAEXE, FACST, FACOR, FAHAT, FASAT, FAHLV, FASLV,
                FAQTP, FACSO, FAIST,
                FAPAM, FAPLM, FAPAO, FAPLO, FANPA, FANPL, FANMA, FANML,
                IDPAR, FADMH, FADMS, FAMXF, IDCAT, FATRC, FARIC, FAEXM, FACSS,
                FATGR, CATGA, CAVAT, FAEXT, FAPNF, FATGT, FANPZ, FATMC, FAPRI, FAINT, FASIN, FACPG,
                FAGRA, IdJobTimeFormula, IdSetupTimeFormula, CostType, CostPerKg, FATPF,
                NumberOfControlsPercentage, EstimateManagement,
                EstimateCostType, EstimateDefaultUnitCost, EstimateCostUnitConversionFactor,
                EstimatePhaseQuantity, EstimateCostFormula,
                IdPhaseCostFormula, IdStandardCostFormula, QuickProgress,
                InternalSetupTimeHours, InternalSetupTimeSeconds, InternalJobTimeHours,
                InternalJobTimeSeconds, InternalTimeType, InternalPiecesNumber,
                ExternalJobTimeHours, ExternalSetupTimeHours, ExternalSetupTimeSeconds,
                ExternalJobTimeSeconds, ExternalTimeType, ExternalPiecesNumber,
                FAPSP
            )
            SELECT
                ph.PACOD,
                ISNULL((SELECT MAX(f2.IDFAS) FROM A_FAC f2 WHERE f2.CCCOD = ph.PACOD), 0)
                    + ROW_NUMBER() OVER (PARTITION BY ph.PACOD ORDER BY TRY_CAST(ph.FASEQ AS INT), ph.FASEQ),
                TRY_CAST(ph.FASEQ AS INT),
                0,                                          -- FAIDN
                ISNULL(m.CodiceSistema, ph.FACOD),          -- FACOD risolto
                ph.FADSC,                                   -- descrizione dall'Excel
                'S', 'Y',                                   -- FAATT, FAABV
                ISNULL(fas.FATYP, 0),
                ISNULL(fas.FAEXE, 'I'),
                ISNULL(fas.FACST, 0),
                ISNULL(fas.FACOR, 0),
                ISNULL(fas.FAHAT, 0),
                ISNULL(fas.FASAT, 0),
                TRY_CAST(ph.FAHLV AS smallint),             -- DurataH dall'Excel
                TRY_CAST(ph.FASLV AS numeric(18,6)),        -- DurataMin dall'Excel
                ISNULL(fas.FAQTP, 0),
                ISNULL(fas.FACSO, 0),
                fas.FAIST,
                fas.FAPAM, fas.FAPLM, fas.FAPAO, fas.FAPLO,
                fas.FANPA, fas.FANPL, fas.FANMA, fas.FANML,
                fas.IDPAR, fas.FADMH, fas.FADMS, fas.FAMXF, fas.IDCAT,
                fas.FATRC, fas.FARIC, fas.FAEXM, fas.FACSS,
                fas.FATGR, fas.CATGA, fas.CAVAT, fas.FAEXT, fas.FAPNF, fas.FATGT,
                fas.FANPZ, fas.FATMC, fas.FAPRI, fas.FAINT, fas.FASIN, fas.FACPG,
                fas.FAGRA, fas.IdJobTimeFormula, fas.IdSetupTimeFormula,
                fas.CostType, fas.CostPerKg, fas.FATPF,
                fas.NumberOfControlsPercentage,
                ISNULL(fas.EstimateManagement, 0),
                fas.EstimateCostType, fas.EstimateDefaultUnitCost,
                fas.EstimateCostUnitConversionFactor, fas.EstimatePhaseQuantity, fas.EstimateCostFormula,
                fas.IdPhaseCostFormula, fas.IdStandardCostFormula, fas.QuickProgress,
                fas.InternalSetupTimeHours, fas.InternalSetupTimeSeconds,
                fas.InternalJobTimeHours, fas.InternalJobTimeSeconds,
                fas.InternalTimeType, fas.InternalPiecesNumber,
                fas.ExternalJobTimeHours, fas.ExternalSetupTimeHours, fas.ExternalSetupTimeSeconds,
                fas.ExternalJobTimeSeconds, fas.ExternalTimeType, fas.ExternalPiecesNumber,
                0                                           -- FAPSP
            FROM Xt_ExcelBom_Phases ph
            LEFT JOIN Xt_ExcelBom_Mappings m  ON m.Ambito = 'FASE_CICLO' AND m.CodiceExcel = ph.FACOD
            LEFT JOIN A_FAS fas               ON fas.FACOD = ISNULL(m.CodiceSistema, ph.FACOD)
            WHERE ph.IDRun = @IDRun
              AND NOT EXISTS (
                  SELECT 1 FROM A_FAC f
                  WHERE f.CCCOD = ph.PACOD AND f.FASEQ = TRY_CAST(ph.FASEQ AS INT));

            -- Aggiorna fasi esistenti: dati risorsa/utilizzo da A_FAS, durate dall'Excel
            UPDATE f
            SET f.FACOD  = ISNULL(m.CodiceSistema, ph.FACOD),
                f.FADSC  = ph.FADSC,
                f.FAHLV  = TRY_CAST(ph.FAHLV AS smallint),
                f.FASLV  = TRY_CAST(ph.FASLV AS numeric(18,6)),
                f.FATYP  = ISNULL(fas.FATYP,  f.FATYP),
                f.FAEXE  = ISNULL(fas.FAEXE,  f.FAEXE),
                f.FACST  = ISNULL(fas.FACST,  f.FACST),
                f.FACOR  = ISNULL(fas.FACOR,  f.FACOR),
                f.FAHAT  = ISNULL(fas.FAHAT,  f.FAHAT),
                f.FASAT  = ISNULL(fas.FASAT,  f.FASAT),
                f.FAQTP  = ISNULL(fas.FAQTP,  f.FAQTP),
                f.FACSO  = ISNULL(fas.FACSO,  f.FACSO),
                f.FAIST  = fas.FAIST,
                f.FAPAM  = fas.FAPAM,  f.FAPLM = fas.FAPLM,
                f.FAPAO  = fas.FAPAO,  f.FAPLO = fas.FAPLO,
                f.FANPA  = fas.FANPA,  f.FANPL = fas.FANPL,
                f.FANMA  = fas.FANMA,  f.FANML = fas.FANML,
                f.IDPAR  = fas.IDPAR,  f.FADMH = fas.FADMH,  f.FADMS = fas.FADMS,
                f.FAMXF  = fas.FAMXF,  f.IDCAT = fas.IDCAT,
                f.FATRC  = fas.FATRC,  f.FARIC = fas.FARIC,
                f.FAEXM  = fas.FAEXM,  f.FACSS = fas.FACSS,
                f.FATGR  = fas.FATGR,  f.CATGA = fas.CATGA,  f.CAVAT = fas.CAVAT,
                f.FAEXT  = fas.FAEXT,  f.FAPNF = fas.FAPNF,  f.FATGT = fas.FATGT,
                f.FANPZ  = fas.FANPZ,  f.FATMC = fas.FATMC,  f.FAPRI = fas.FAPRI,
                f.FAINT  = fas.FAINT,  f.FASIN = fas.FASIN,  f.FACPG = fas.FACPG,
                f.FAGRA  = fas.FAGRA,
                f.IdJobTimeFormula  = fas.IdJobTimeFormula,
                f.IdSetupTimeFormula = fas.IdSetupTimeFormula,
                f.CostType  = fas.CostType,  f.CostPerKg = fas.CostPerKg,
                f.FATPF     = fas.FATPF,
                f.NumberOfControlsPercentage         = fas.NumberOfControlsPercentage,
                f.EstimateManagement                 = ISNULL(fas.EstimateManagement, f.EstimateManagement),
                f.EstimateCostType                   = fas.EstimateCostType,
                f.EstimateDefaultUnitCost            = fas.EstimateDefaultUnitCost,
                f.EstimateCostUnitConversionFactor   = fas.EstimateCostUnitConversionFactor,
                f.EstimatePhaseQuantity              = fas.EstimatePhaseQuantity,
                f.EstimateCostFormula                = fas.EstimateCostFormula,
                f.IdPhaseCostFormula                 = fas.IdPhaseCostFormula,
                f.IdStandardCostFormula              = fas.IdStandardCostFormula,
                f.QuickProgress                      = fas.QuickProgress,
                f.InternalSetupTimeHours    = fas.InternalSetupTimeHours,
                f.InternalSetupTimeSeconds  = fas.InternalSetupTimeSeconds,
                f.InternalJobTimeHours      = fas.InternalJobTimeHours,
                f.InternalJobTimeSeconds    = fas.InternalJobTimeSeconds,
                f.InternalTimeType          = fas.InternalTimeType,
                f.InternalPiecesNumber      = fas.InternalPiecesNumber,
                f.ExternalJobTimeHours      = fas.ExternalJobTimeHours,
                f.ExternalSetupTimeHours    = fas.ExternalSetupTimeHours,
                f.ExternalSetupTimeSeconds  = fas.ExternalSetupTimeSeconds,
                f.ExternalJobTimeSeconds    = fas.ExternalJobTimeSeconds,
                f.ExternalTimeType          = fas.ExternalTimeType,
                f.ExternalPiecesNumber      = fas.ExternalPiecesNumber
            FROM A_FAC f
            INNER JOIN Xt_ExcelBom_Phases ph ON ph.PACOD = f.CCCOD
                AND f.FASEQ = TRY_CAST(ph.FASEQ AS INT)
            LEFT JOIN Xt_ExcelBom_Mappings m  ON m.Ambito = 'FASE_CICLO' AND m.CodiceExcel = ph.FACOD
            LEFT JOIN A_FAS fas               ON fas.FACOD = ISNULL(m.CodiceSistema, ph.FACOD)
            WHERE ph.IDRun = @IDRun;

            -- Ricalcola FAUFC: ultima fase del ciclo = 'Y', tutte le altre = 'N'
            UPDATE A_FAC SET FAUFC = 'N'
            WHERE CCCOD IN (SELECT DISTINCT PACOD FROM Xt_ExcelBom_Phases WHERE IDRun = @IDRun);

            UPDATE f SET f.FAUFC = 'Y'
            FROM A_FAC f
            INNER JOIN (
                SELECT CCCOD, MAX(FASEQ) AS MaxSeq
                FROM A_FAC
                WHERE CCCOD IN (SELECT DISTINCT PACOD FROM Xt_ExcelBom_Phases WHERE IDRun = @IDRun)
                GROUP BY CCCOD
            ) x ON f.CCCOD = x.CCCOD AND f.FASEQ = x.MaxSeq;

            -- Associa ciclo all'articolo in L_PACC
            INSERT INTO L_PACC (PACOD, CCCOD, LPRIM)
            SELECT DISTINCT p.PACOD, p.PACOD, 'Y'
            FROM Xt_ExcelBom_Parts p
            LEFT JOIN L_PACC l ON l.PACOD = p.PACOD AND l.CCCOD = p.PACOD
            WHERE p.IDRun = @IDRun AND l.PACOD IS NULL;

            -- Aggiorna PACCP in A_PAR
            UPDATE A_PAR
            SET PACCP = p.PACOD
            FROM A_PAR a
            INNER JOIN Xt_ExcelBom_Parts p ON p.PACOD = a.PACOD
            WHERE p.IDRun = @IDRun AND (a.PACCP IS NULL OR a.PACCP != p.PACOD);
        END

        UPDATE Xt_ExcelBom_LogRun
        SET Stato = 'COMPLETED', DataFine = GETDATE()
        WHERE Sorgente = 'EXCEL' AND IDRun = @IDRun;

    END TRY
    BEGIN CATCH
        UPDATE Xt_ExcelBom_LogRun
        SET Stato = 'ERROR', Errore = ERROR_MESSAGE(), DataFine = GETDATE()
        WHERE Sorgente = 'EXCEL' AND IDRun = @IDRun;

        THROW;
    END CATCH
END
