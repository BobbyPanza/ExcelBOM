# ExcelBOM — Applicazione GUI per importare distinte da Excel in Factory

## Obiettivo

Applicazione Windows Forms (.NET) portabile che importa anagrafiche articoli, distinte base, cicli di lavorazione da Excel in Factory.

L'app è GUI: sfoglia per selezionare Excel, scegli flussi da attivare, visualizza staging, importa.

Configurazione in `config.json` nella cartella dell'eseguibile.

## Struttura Excel

Excel con 3 fogli fissi:

### Foglio "Articoli"

| Codice | Descrizione | Famiglia | UdM |
|---|---|---|---|

### Foglio "Diba"

| Padre | Figlio | Seq | Quantità | UdM | Nota |
|---|---|---|---|---|---|

### Foglio "Cicli"

| Codice | SeqFase | CodFase | Descrizione | DurataH | DurataMin |
|---|---|---|---|---|---|

## Configurazione (config.json)

```json
{
  "ConnectionString": "Server=localhost;Database=Factory;User Id=sa;Password=Intsupport1;TrustServerCertificate=True;",
  "PATYP_DEFAULT": "C",
  "PATYF_DEFAULT": "I",
  "PAUDM_DEFAULT": "PZ",
  "FAMIGLIA_DEFAULT": "",
  "EnableAnagrafiche": true,
  "EnableDistinta": true,
  "EnableCiclo": true
}
```

## Architettura

GUI → Leggi Excel → Carica staging SQL → Chiama SP import.

Stesse tabelle e SP del progetto ImportBOM.

## Dipendenze NuGet

- `ClosedXML` — Lettura Excel .xlsx
- `Microsoft.Data.SqlClient` — SQL Server
- `System.Text.Json` — Config JSON

## Pubblicazione

`dotnet publish -c Release --self-contained -r win-x64` per app portabile.
