using System.Globalization;
using System.Text;
using Dapper;
using Microsoft.Data.SqlClient;
using PortaleCEIR.Models;

namespace PortaleCEIR.Services;

/// <summary>
/// Accesso al DB gestionale ERP del cliente (Factory).
/// </summary>
public class GestionaleService(IConfiguration config)
{
    private readonly string? _connStr = config.GetConnectionString("Gestionale");

    public bool IsConfigured => !string.IsNullOrWhiteSpace(_connStr);

    private SqlConnection GetConnection() => new(_connStr);

    // ─── Auth (A_OPR) ─────────────────────────────────────────────────────────

    public async Task<(bool Ok, string? Nome)> TryLoginAsync(string opcod, string password)
    {
        using var db = GetConnection();
        var operatore = await db.QuerySingleOrDefaultAsync<(string Dsc, string Psw)>(
            "SELECT OPDSC, OPPSW FROM A_OPR WHERE OPCOD = @opcod",
            new { opcod });

        if (operatore == default) return (false, null);
        if (operatore.Psw != password) return (false, null);

        return (true, operatore.Dsc);
    }

    // ─── Coil — Anagrafiche ───────────────────────────────────────────────────

    public async Task<List<TipoMateriale>> GetTipiMaterialeAsync()
    {
        using var db = GetConnection();
        return [.. await db.QueryAsync<TipoMateriale>(
            "SELECT TMCOD, TMDSC FROM A_TMT ORDER BY TMDSC")];
    }

    public async Task<List<MisuraValore>> GetMisureValoriAsync(int grcod)
    {
        using var db = GetConnection();
        return [.. await db.QueryAsync<MisuraValore>(
            "SELECT Value, Description FROM MeasurementValues WHERE GRCOD = @grcod ORDER BY Description",
            new { grcod })];
    }

    // ─── Fasi Factory ─────────────────────────────────────────────────────────

    public async Task<List<FaseFactory>> GetFasiFactoryAsync()
    {
        using var db = GetConnection();
        return [.. await db.QueryAsync<FaseFactory>(
            "SELECT FACOD, FADSC FROM A_FAS ORDER BY FADSC")];
    }

    // ─── Coil — Lista con giacenze ────────────────────────────────────────────

    public async Task<List<CoilArticolo>> GetCoilArticoliAsync(
        string? search    = null,
        IReadOnlyCollection<string>? mtcods    = null,
        decimal? spessore = null, decimal? altezza  = null,
        IReadOnlyCollection<string>? finiture  = null,
        IReadOnlyCollection<string>? colori    = null,
        IReadOnlyCollection<string>? fornitori = null)
    {
        using var db = GetConnection();

        var where = new StringBuilder("p.FMCOD = 'Coil'");
        if (!string.IsNullOrWhiteSpace(search))
            where.Append(" AND (p.PACOD LIKE @search OR p.PADSC LIKE @search)");
        if (mtcods?.Count > 0)
            where.Append(" AND p.MTCOD IN @mtcods");
        if (spessore.HasValue)
            where.Append(" AND p.PADX1 = @spessore");
        if (altezza.HasValue)
            where.Append(" AND p.PADY1 = @altezza");
        if (finiture?.Count > 0)
            where.Append(" AND fin.GRVAA IN @finiture");
        if (colori?.Count > 0)
            where.Append(" AND col.GRVAA IN @colori");
        if (fornitori?.Count > 0)
            where.Append(" AND forni.GRVAA IN @fornitori");

        var sql = $"""
            SELECT p.PACOD,
                   p.PADSC,
                   p.MTCOD,
                   ISNULL(t.TMDSC, p.MTCOD)          AS Materiale,
                   CAST(p.PADX1 AS decimal(10,3))     AS Spessore,
                   CAST(p.PADY1 AS decimal(10,1))     AS Altezza,
                   fin.GRVAA                                    AS FinituraCod,
                   ISNULL(finMv.Description, fin.GRVAA)     AS FinituraDsc,
                   col.GRVAA                                 AS ColoreCod,
                   ISNULL(colMv.Description, col.GRVAA)     AS ColoreDsc,
                   forni.GRVAA                               AS FornitoreCod,
                   ISNULL(forniMv.Description, forni.GRVAA) AS FornitoreDsc,
                   ISNULL(ml.QTLOC, 0)                AS Giacenza,
                   ISNULL(qt.PADSP, 0)                AS Disponibile,
                   p.PAUDM                            AS UdM,
                   fam.UMALT                          AS UdM2,
                   COALESCE(NULLIF(p.StoreConversionFactor,1), fam.UMFDC) AS UdM2Fattore
            FROM A_PAR p
            LEFT JOIN A_TMT             t      ON p.MTCOD    = t.TMCOD
            LEFT JOIN L_PAGR            fin    ON p.PACOD    = fin.PACOD   AND fin.IDFGR   = (SELECT IDFGR FROM L_FMGR WHERE FMCOD='Coil' AND GRCOD=9)
            LEFT JOIN MeasurementValues finMv  ON fin.GRVAA  = finMv.Value AND finMv.GRCOD = 9
            LEFT JOIN L_PAGR            col    ON p.PACOD    = col.PACOD   AND col.IDFGR   = (SELECT IDFGR FROM L_FMGR WHERE FMCOD='Coil' AND GRCOD=11)
            LEFT JOIN MeasurementValues colMv  ON col.GRVAA  = colMv.Value AND colMv.GRCOD = 11
            LEFT JOIN L_PAGR            forni  ON p.PACOD    = forni.PACOD AND forni.IDFGR  = (SELECT IDFGR FROM L_FMGR WHERE FMCOD='Coil' AND GRCOD=13)
            LEFT JOIN MeasurementValues forniMv ON forni.GRVAA = forniMv.Value AND forniMv.GRCOD = 13
            LEFT JOIN L_MLPA            ml     ON p.PACOD    = ml.PACOD    AND ml.LCPRC    = 'Y'
            LEFT JOIN L_PAQT            qt     ON p.PACOD    = qt.PACOD
            LEFT JOIN A_FAM             fam    ON fam.FMCOD  = p.FMCOD
            WHERE {where}
            ORDER BY p.PADX1, p.PADY1, p.PADSC
            """;

        return [.. await db.QueryAsync<CoilArticolo>(sql, new {
            search    = string.IsNullOrWhiteSpace(search) ? null : $"%{search}%",
            mtcods, spessore, altezza, finiture, colori, fornitori
        })];
    }

    // ─── Coil — Creazione ─────────────────────────────────────────────────────

    public async Task<(string Codice, string Descrizione)> CalcolaCoilAsync(NuovoCoilForm f)
    {
        using var db = GetConnection();

        var tmdsc  = await db.ExecuteScalarAsync<string>("SELECT TMDSC FROM A_TMT WHERE TMCOD=@v", new { v = f.MTCOD }) ?? f.MTCOD ?? "";
        var finDsc = await db.ExecuteScalarAsync<string>("SELECT Description FROM MeasurementValues WHERE GRCOD=9  AND Value=@v", new { v = f.Finitura }) ?? f.Finitura ?? "";
        var colDsc = await db.ExecuteScalarAsync<string>("SELECT Description FROM MeasurementValues WHERE GRCOD=11 AND Value=@v", new { v = f.Colore   }) ?? f.Colore   ?? "";
        var forDsc = await db.ExecuteScalarAsync<string>("SELECT Description FROM MeasurementValues WHERE GRCOD=13 AND Value=@v", new { v = f.Fornitore}) ?? f.Fornitore?? "";

        var sp = (int)Math.Round(f.Spessore!.Value * 100);
        var h  = (int)Math.Round(f.Altezza!.Value  *  10);

        var codice = $"C{f.MTCOD}{f.Finitura}{sp:D3}H{h:D4}{f.Colore}{f.Fornitore}";
        var spStr  = f.Spessore.Value.ToString("G", CultureInfo.InvariantCulture);
        var hStr   = f.Altezza.Value .ToString("G", CultureInfo.InvariantCulture);
        var desc   = $"Coil {tmdsc} {finDsc} Sp.{spStr} H{hStr} {colDsc} {forDsc}";

        return (codice, desc);
    }

    public async Task AggiungeMisuraValoreAsync(int grcod, string value, string description)
    {
        using var db = GetConnection();
        await db.ExecuteAsync("""
            IF NOT EXISTS (SELECT 1 FROM MeasurementValues WHERE GRCOD=@grcod AND Value=@value)
            INSERT INTO MeasurementValues (GRCOD, Value, Description) VALUES (@grcod, @value, @description)
            """, new { grcod, value, description });
    }

    // Genera C001, C002… per i colori (GRCOD=11) e inserisce il nuovo valore
    public async Task<string> AggiungiColoreAsync(string descrizione)
    {
        using var db = GetConnection();
        var codici = await db.QueryAsync<string>(
            "SELECT Value FROM MeasurementValues WHERE GRCOD=11 AND Value LIKE 'C%'");
        var nums = codici
            .Where(c => c.Length > 1 && c[1..].All(char.IsDigit))
            .Select(c => int.TryParse(c[1..], out var n) ? n : 0)
            .ToList();
        var next = (nums.Count > 0 ? nums.Max() : 0) + 1;
        var code = $"C{next:D3}";
        await db.ExecuteAsync(
            "INSERT INTO MeasurementValues (GRCOD, Value, Description) VALUES (11, @code, @descrizione)",
            new { code, descrizione });
        return code;
    }

    // Genera F01, F02… per i fornitori (GRCOD=13) e inserisce il nuovo valore
    public async Task<string> AggiungiFornitoreAsync(string descrizione)
    {
        using var db = GetConnection();
        var codici = await db.QueryAsync<string>(
            "SELECT Value FROM MeasurementValues WHERE GRCOD=13 AND Value LIKE 'F%'");
        var nums = codici
            .Where(c => c.Length > 1 && c[1..].All(char.IsDigit))
            .Select(c => int.TryParse(c[1..], out var n) ? n : 0)
            .ToList();
        var next = (nums.Count > 0 ? nums.Max() : 0) + 1;
        var code = $"F{next:D2}";
        await db.ExecuteAsync(
            "INSERT INTO MeasurementValues (GRCOD, Value, Description) VALUES (13, @code, @descrizione)",
            new { code, descrizione });
        return code;
    }

    public async Task<Dictionary<string, CoilArticolo>> GetCoilByCodiciAsync(IEnumerable<string> codici)
    {
        var list = codici.Where(c => !string.IsNullOrEmpty(c)).Distinct().ToList();
        if (list.Count == 0) return [];
        using var db = GetConnection();
        var items = await db.QueryAsync<CoilArticolo>("""
            SELECT p.PACOD, p.PADSC, p.MTCOD,
                   ISNULL(t.TMDSC, p.MTCOD)          AS Materiale,
                   CAST(p.PADX1 AS decimal(10,3))     AS Spessore,
                   CAST(p.PADY1 AS decimal(10,1))     AS Altezza,
                   fin.GRVAA  AS FinituraCod, ISNULL(finMv.Description, fin.GRVAA)     AS FinituraDsc,
                   col.GRVAA  AS ColoreCod,   ISNULL(colMv.Description, col.GRVAA)     AS ColoreDsc,
                   forni.GRVAA AS FornitoreCod, ISNULL(forniMv.Description, forni.GRVAA) AS FornitoreDsc,
                   ISNULL(ml.QTLOC, 0)  AS Giacenza,
                   ISNULL(qt.PADSP, 0)  AS Disponibile,
                   p.PAUDM              AS UdM,
                   fam.UMALT            AS UdM2,
                   COALESCE(NULLIF(p.StoreConversionFactor,1), fam.UMFDC) AS UdM2Fattore
            FROM A_PAR p
            LEFT JOIN A_TMT             t       ON p.MTCOD    = t.TMCOD
            LEFT JOIN L_PAGR            fin     ON p.PACOD    = fin.PACOD   AND fin.IDFGR   = (SELECT IDFGR FROM L_FMGR WHERE FMCOD='Coil' AND GRCOD=9)
            LEFT JOIN MeasurementValues finMv   ON fin.GRVAA  = finMv.Value AND finMv.GRCOD = 9
            LEFT JOIN L_PAGR            col     ON p.PACOD    = col.PACOD   AND col.IDFGR   = (SELECT IDFGR FROM L_FMGR WHERE FMCOD='Coil' AND GRCOD=11)
            LEFT JOIN MeasurementValues colMv   ON col.GRVAA  = colMv.Value AND colMv.GRCOD = 11
            LEFT JOIN L_PAGR            forni   ON p.PACOD    = forni.PACOD AND forni.IDFGR  = (SELECT IDFGR FROM L_FMGR WHERE FMCOD='Coil' AND GRCOD=13)
            LEFT JOIN MeasurementValues forniMv ON forni.GRVAA = forniMv.Value AND forniMv.GRCOD = 13
            LEFT JOIN L_MLPA            ml      ON p.PACOD    = ml.PACOD    AND ml.LCPRC = 'Y'
            LEFT JOIN L_PAQT            qt      ON p.PACOD    = qt.PACOD
            LEFT JOIN A_FAM             fam     ON fam.FMCOD  = p.FMCOD
            WHERE p.PACOD IN @codici
            """, new { codici = list });
        return items.ToDictionary(x => x.PACOD);
    }

    // ─── Esportazione distinta in Factory (3 livelli) ────────────────────────
    //
    // Struttura BOM:
    //   Progetto  (A_PAR CodiceOrdine)
    //     └─ Posizione  (A_PAR Ordine-Pos)         qty = m²_progetto / m²_posizione
    //           └─ Lama (A_PAR Ordine-Pos-Suffisso) qty = QtaTotale per 1 pannello
    //                 └─ Materia Prima (coil)        qty = QuantitaMpMm/1000 m per 1 lama

    public async Task<List<EsportazionePosizioneResult>> EsportaOrdineAsync(
        OrdineDetail ordine,
        Dictionary<string, string?> fasMapping,
        Dictionary<string, string?> config)
    {
        var fmcod = config.GetValueOrDefault("FamigliaPannello") ?? "Generica";
        var patyp = config.GetValueOrDefault("PannelloPatyp")    ?? "F";
        var patyf = config.GetValueOrDefault("PannelloPatyf")    ?? "I";
        var udm   = config.GetValueOrDefault("PannelloUdm")      ?? "Pz";

        var risultati = new List<EsportazionePosizioneResult>();

        using var db = GetConnection();
        await db.OpenAsync();

        // ── Livello 1: articolo PROGETTO ──────────────────────────────────────
        var projCode = Trunc20(ordine.CodiceOrdine);
        var projDsc  = ordine.ClienteDescrizione is not null
            ? $"{ordine.CodiceOrdine} — {ordine.ClienteDescrizione}"
            : ordine.CodiceOrdine;
        try
        {
            await EnsureAParAsync(db, projCode, projDsc, fmcod, patyp, patyf, udm);
            risultati.Add(new("PROGETTO", projCode, "OK", null, 0, 0));
        }
        catch (Exception ex)
        {
            risultati.Add(new("PROGETTO", projCode, "ERR", ex.Message, 0, 0));
            return risultati;
        }

        // ── Livello 2: una POSIZIONE per ogni pannello ────────────────────────
        foreach (var pos in ordine.Posizioni)
        {
            try
            {
                int ok = 0, skip = 0;
                var note = new List<string>();

                var posCode = Trunc20(pos.CodicePosizione);
                var posDsc  = BuildPADSC(ordine, pos);
                await EnsureAParAsync(db, posCode, posDsc, fmcod, "S", patyf, udm);

                // Quantità posizione: manuale se impostata, altrimenti m²totali / m²pannello
                var mqPos  = pos.Larghezza * pos.Altezza / 1_000_000m;
                var qtaPos = (pos.Quantita.HasValue && pos.Quantita.Value > 0)
                             ? (decimal)pos.Quantita.Value
                             : (ordine.QuantitaMq.HasValue && mqPos > 0)
                               ? Math.Round(ordine.QuantitaMq.Value / mqPos, 0)
                               : 1m;
                await UpsertDbsAsync(db, projCode, posCode, qtaPos, udm, 0);

                // Ciclo lavorazione della POSIZIONE
                await EnsureCicloAsync(db, posCode, posDsc,
                    [.. pos.LavorazioniPos.Select(l => (l.CodiceFase, l.NomeLavorazione, l.TempoTeoricoPz))],
                    fasMapping);

                // ── Livello 3: LAMA — solo componenti master ──────────────────
                int lamaSeq = 1;
                foreach (var comp in pos.Componenti.Where(c => !c.IsRiferimento).OrderBy(c => c.SortOrder))
                {
                    var prog     = comp.Progressivo > 1 ? comp.Progressivo.ToString() : "";
                    var lamaCode = Trunc20($"{pos.CodicePosizione}-{comp.SuffissoCodice}{prog}");
                    var lamaDsc  = $"{comp.SuffissoCodice}{prog} {comp.Descrizione}";

                    await EnsureAParAsync(db, lamaCode, lamaDsc, fmcod, "S", patyf, "Pz");

                    // A_DBS Posizione → Lama: QtaTotale per 1 pannello
                    await UpsertDbsAsync(db, posCode, lamaCode, comp.QtaTotale, "Pz", lamaSeq++);

                    // Ciclo lavorazione della LAMA
                    await EnsureCicloAsync(db, lamaCode, lamaDsc,
                        [.. comp.Lavorazioni.Select(l => (l.CodiceFase, l.NomeLavorazione, l.TempoTeoricoPz))],
                        fasMapping);

                    // ── Livello 4: MATERIA PRIMA (coil) ──────────────────────
                    if (!string.IsNullOrEmpty(comp.CodiceArticoloFactory))
                    {
                        // Quantità unitaria per 1 lama: QuantitaMpMm se impostata, altrimenti LunghezzaMm
                        var qtaMp = comp.QuantitaMpMm ?? comp.LunghezzaMm ?? 0m;
                        await UpsertDbsAsync(db, lamaCode, comp.CodiceArticoloFactory, qtaMp, "mm", 1);
                        ok++;
                    }
                    else
                    {
                        skip++;
                        note.Add($"{comp.CodiceComponente}: nessun articolo Factory associato");
                    }
                }

                var msg = note.Count > 0 ? string.Join("; ", note) : null;
                risultati.Add(new(pos.CodicePosizione, posCode, "OK", msg, ok, skip));
            }
            catch (Exception ex)
            {
                risultati.Add(new(pos.CodicePosizione, null, "ERR", ex.Message, 0, 0));
            }
        }

        return risultati;
    }

    private async Task EnsureAParAsync(SqlConnection db, string pacod, string padsc,
        string fmcod, string patyp, string patyf, string udm)
    {
        var exists = await db.ExecuteScalarAsync<int>(
            "SELECT COUNT(1) FROM A_PAR WHERE PACOD=@pacod", new { pacod }) > 0;

        if (exists)
        {
            // Corregge PATYP se era stato creato con valore sbagliato
            await db.ExecuteAsync(
                "UPDATE A_PAR SET PATYP=@patyp WHERE PACOD=@pacod AND PATYP<>@patyp",
                new { pacod, patyp });
            return;
        }

        var idart = await GetProgressivoAsync(db, "dbo.A_PAR");
        await db.ExecuteAsync("""
            INSERT INTO A_PAR
                (IDART,PACOD,PADSC,FMCOD,PATYP,PATYF,PAMAG,PAUDM,
                 PATGM,PAQTM,PAPDR,PAFDC,PACST,PACSA,PAPNT,PACSP,TMCSP,
                 Approved,StoreConversionFactor)
            VALUES
                (@idart,@pacod,@padsc,@fmcod,@patyp,@patyf,'Y',@udm,
                 0,0,0,1,0,0,0,0,0,'N',1)
            """, new { idart, pacod, padsc, fmcod, patyp, patyf, udm });
    }

    private static async Task UpsertDbsAsync(SqlConnection db, string dbpar, string dbchl,
        decimal dbqta, string udm, int dbseq)
    {
        var exists = await db.ExecuteScalarAsync<int>(
            "SELECT COUNT(1) FROM A_DBS WHERE DBPAR=@p AND DBCHL=@c",
            new { p = dbpar, c = dbchl }) > 0;

        if (!exists)
        {
            var iddbs = await GetProgressivoAsync(db, "dbo.A_DBS");
            await db.ExecuteAsync("""
                INSERT INTO A_DBS (IDDBS,DBPAR,DBCHL,DBQTA,DBSFR,DBUDM,DBSEQ,QTRCL)
                VALUES (@iddbs,@dbpar,@dbchl,@dbqta,0,@udm,@dbseq,0)
                """, new { iddbs, dbpar, dbchl, dbqta, udm, dbseq });
        }
        else
        {
            await db.ExecuteAsync(
                "UPDATE A_DBS SET DBQTA=@dbqta WHERE DBPAR=@p AND DBCHL=@c",
                new { dbqta, p = dbpar, c = dbchl });
        }
    }

    private async Task EnsureCicloAsync(SqlConnection db, string pacod, string padsc,
        List<(string? CodiceFase, string NomeLavorazione, decimal TempoTeoricoPz)> fasi,
        Dictionary<string, string?> fasMapping)
    {
        // Risolve FACOD: 1) mapping esplicito  2) fallback diretto in A_FAS
        var fasiRisolte = new List<(string Facod, string Nome, decimal Tpz)>();
        foreach (var (codiceFase, nome, tpz) in fasi)
        {
            if (codiceFase is null) continue;
            string? facod = null;
            if (fasMapping.TryGetValue(codiceFase, out var mapped) && mapped != null)
                facod = mapped;
            else
            {
                var fasExists = await db.ExecuteScalarAsync<int>(
                    "SELECT COUNT(1) FROM A_FAS WHERE FACOD=@c", new { c = codiceFase }) > 0;
                if (fasExists) facod = codiceFase;
            }
            if (facod != null) fasiRisolte.Add((facod, nome, tpz));
        }
        if (fasiRisolte.Count == 0) return;

        var cclExists = await db.ExecuteScalarAsync<int>(
            "SELECT COUNT(1) FROM A_CCL WHERE CCCOD=@c", new { c = pacod }) > 0;
        if (cclExists)
        {
            // Aggiorna PACCP nel caso fosse stato resettato da un trigger
            await db.ExecuteAsync(
                "UPDATE A_PAR SET PACCP=@c WHERE PACOD=@p AND (PACCP IS NULL OR PACCP!=@c)",
                new { p = pacod, c = pacod });
            return;
        }

        var idccl = await GetProgressivoAsync(db, "dbo.A_CCL");
        await db.ExecuteAsync("""
            INSERT INTO A_CCL (CCCOD,CCDSC,CCABL,CCCMP,IDCCL)
            VALUES (@cccod,@ccdsc,'Y','N',@idccl)
            """, new { cccod = pacod, ccdsc = padsc, idccl });

        int idfas = 1;
        foreach (var (facod, nomeLavorazione, tempoPz) in fasiRisolte)
        {
            var fatyp = await db.ExecuteScalarAsync<int?>(
                "SELECT FATYP FROM A_FAS WHERE FACOD=@facod", new { facod }) ?? 1;
            var idfac = await GetProgressivoAsync(db, "dbo.A_FAC");
            await db.ExecuteAsync("""
                INSERT INTO A_FAC
                    (CCCOD,IDFAS,FACOD,FASEQ,FAIDN,FAATT,FAABV,FADSC,FATYP,FAEXE,
                     FACST,FACOR,FAHAT,FASAT,FAHLV,FASLV,FAQTP,FACSO,FAPSP,IDFAC)
                VALUES
                    (@cccod,@idfas,@facod,@faseq,0,'S','Y',@fadsc,@fatyp,'I',
                     0,0,0,0,0,0,@faqtp,0,0,@idfac)
                """, new { cccod = pacod, idfas, facod, faseq = idfas,
                           fadsc = nomeLavorazione, fatyp, faqtp = tempoPz, idfac });
            idfas++;
        }

        var laccExists = await db.ExecuteScalarAsync<int>(
            "SELECT COUNT(1) FROM L_PACC WHERE PACOD=@p AND CCCOD=@c",
            new { p = pacod, c = pacod }) > 0;
        if (!laccExists)
            await db.ExecuteAsync(
                "INSERT INTO L_PACC (PACOD,CCCOD,LPRIM) VALUES (@p,@c,'Y')",
                new { p = pacod, c = pacod });

        await db.ExecuteAsync(
            "UPDATE A_PAR SET PACCP=@c WHERE PACOD=@p",
            new { p = pacod, c = pacod });
    }

    private static async Task<int> GetProgressivoAsync(SqlConnection db, string tableName)
        => await db.ExecuteScalarAsync<int>(
            $"DECLARE @id INT; EXEC @id = dbo.GetProgressivo '{tableName}'; SELECT @id");

    private static string BuildPADSC(OrdineDetail ordine, OrdinePosizione pos)
    {
        var alias = !string.IsNullOrEmpty(pos.Alias) ? pos.Alias : pos.CodicePosizione;
        return $"{alias} {(int)pos.Larghezza}×{(int)pos.Altezza}mm";
    }

    private static string Trunc20(string s) => s.Length <= 20 ? s : s[..20];

    public async Task<bool> CoilEsisteAsync(string codice)
    {
        using var db = GetConnection();
        return await db.ExecuteScalarAsync<int>(
            "SELECT COUNT(1) FROM A_PAR WHERE PACOD=@codice", new { codice }) > 0;
    }

    public async Task CreaCoilAsync(NuovoCoilForm f, string codice, string descrizione)
    {
        using var db = GetConnection();
        await db.OpenAsync();

        // Legge MGCOD e LCCOD dalla famiglia Coil
        var fam = await db.QuerySingleOrDefaultAsync<(string MgCod, string LcCod)>(
            "SELECT MGCOD, LCCOD FROM A_FAM WHERE FMCOD='Coil'");
        var mgcod = fam.MgCod ?? "MG01";
        var lccod = fam.LcCod ?? "LC01";

        // Peso specifico dal tipo materiale
        var tmpsp = await db.ExecuteScalarAsync<decimal?>(
            "SELECT TMPSP FROM A_TMT WHERE TMCOD=@m", new { m = f.MTCOD });

        // Fattore conversione m→kg: kg/m = TMPSP × sp_mm × H_mm / 1000
        // StoreConversionFactor = m/kg = 1 / (TMPSP × sp × H / 1000)
        decimal? fattore = null;
        if (tmpsp is > 0 && f.Spessore is > 0 && f.Altezza is > 0)
            fattore = 1m / (tmpsp.Value * f.Spessore.Value * f.Altezza.Value / 1000m);

        // A_PAR — solo se non esiste già (tentativo precedente parzialmente riuscito)
        var parExists = await db.ExecuteScalarAsync<int>(
            "SELECT COUNT(1) FROM A_PAR WHERE PACOD=@codice", new { codice });
        if (parExists == 0)
        {
            await db.ExecuteAsync("""
                INSERT INTO A_PAR (PACOD,PADSC,FMCOD,MTCOD,PADX1,PADY1,PAUDM,PATYP,PATYF,PAMAG)
                VALUES (@codice,@descrizione,'Coil',@mtcod,@sp,@h,'m','S','E','Y')
                """, new { codice, descrizione, mtcod = f.MTCOD, sp = f.Spessore, h = f.Altezza });
        }

        // Assicura UdM='m', popola PAPSP e StoreConversionFactor (il trigger potrebbe aver resettato UdM)
        await db.ExecuteAsync("""
            UPDATE A_PAR
            SET PAUDM='m',
                PAPSP = CASE WHEN @psp IS NOT NULL THEN @psp ELSE PAPSP END,
                StoreConversionFactor = CASE WHEN @fatt IS NOT NULL THEN @fatt ELSE StoreConversionFactor END
            WHERE PACOD=@codice
            """, new { psp = tmpsp, fatt = fattore, codice });

        // Legge IDFGR da L_FMGR per la famiglia Coil (mappa GRCOD → IDFGR)
        var fmgr = (await db.QueryAsync<(int GrCod, int IdFgr)>(
            "SELECT GRCOD, IDFGR FROM L_FMGR WHERE FMCOD='Coil'"))
            .ToDictionary(x => x.GrCod, x => x.IdFgr);

        // L_PAGR — il trigger TRG_ON_BEFORE_INSERT_PAR le crea già vuote con A_PAR;
        //           aggiorniamo soltanto i valori
        await db.ExecuteAsync("UPDATE L_PAGR SET GRVAL=@v, UMCOD='mm' WHERE PACOD=@p AND IDFGR=@fgr",
            new { p = codice, fgr = fmgr[7],  v = f.Spessore });   // Spessore g1
        await db.ExecuteAsync("UPDATE L_PAGR SET GRVAL=@v, UMCOD='mm' WHERE PACOD=@p AND IDFGR=@fgr",
            new { p = codice, fgr = fmgr[5],  v = f.Altezza  });   // Altezza  g2
        await db.ExecuteAsync("UPDATE L_PAGR SET GRVAA=@v WHERE PACOD=@p AND IDFGR=@fgr",
            new { p = codice, fgr = fmgr[13], v = f.Fornitore });  // Fornitore a3
        await db.ExecuteAsync("UPDATE L_PAGR SET GRVAA=@v WHERE PACOD=@p AND IDFGR=@fgr",
            new { p = codice, fgr = fmgr[9],  v = f.Finitura  });  // Finitura  a4
        await db.ExecuteAsync("UPDATE L_PAGR SET GRVAA=@v WHERE PACOD=@p AND IDFGR=@fgr",
            new { p = codice, fgr = fmgr[11], v = f.Colore    });  // Colore    a5

        // L_MLPA — solo se non esiste già
        var mlpaExists = await db.ExecuteScalarAsync<int>(
            "SELECT COUNT(1) FROM L_MLPA WHERE PACOD=@p", new { p = codice });
        if (mlpaExists == 0)
        {
            await db.ExecuteAsync("""
                INSERT INTO L_MLPA(PACOD,MGCOD,LCCOD,LCPRC,LCABI,QTLOC,QTMIN,QTMAX)
                VALUES(@p,@mg,@lc,'Y','Y',0,0,0)
                """, new { p = codice, mg = mgcod, lc = lccod });
        }

        // L_PAQT — idempotente
        await db.ExecuteAsync("""
            IF NOT EXISTS (SELECT 1 FROM L_PAQT WHERE PACOD=@p)
            INSERT INTO L_PAQT(PACOD,PAIMP,PAORD,PADSP,RequiredQuantity) VALUES(@p,0,0,0,0)
            """, new { p = codice });
    }
}
