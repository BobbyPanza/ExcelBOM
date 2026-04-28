# Installazione ExcelBOM

## 1. Scarica e decomprimi

Scarica `ExcelBOM_v1.3.zip` e decomprimi in una cartella locale, ad esempio:

```
C:\Tools\ExcelBOM\
```

Troverai:
- `ExcelBOM.exe` — eseguibile principale
- `config.json` — configurazione connessione e valori di default
- `ExcelBOM.CustomForms.xml` — custom forms da importare in Factory
- `01_staging.sql`, `02_sp_import.sql`, `03_mappings_seed.sql` — script database

---

## 2. Prepara il database

Apri SQL Server Management Studio, connettiti al database **Factory** ed esegui gli script nell'ordine seguente:

| Ordine | Script | Cosa fa |
|---|---|---|
| 1 | `01_staging.sql` | Crea le tabelle di staging |
| 2 | `03_mappings_seed.sql` | Crea la tabella mappature e inserisce esempi |
| 3 | `02_sp_import.sql` | Crea la stored procedure di import |

---

## 3. Configura la connessione

Apri `config.json` con un editor di testo e modifica la `ConnectionString` con i dati del tuo server:

```json
{
  "ConnectionString": "Server=NOMESERVER;Database=Factory;User Id=sa;Password=TUA_PASSWORD;TrustServerCertificate=True;",
  "PATYP_DEFAULT": "C",
  "PATYF_DEFAULT": "I",
  "PAUDM_DEFAULT": "PZ",
  "FAMIGLIA_DEFAULT": "",
  "EnableAnagrafiche": true,
  "EnableDistinta": true,
  "EnableCiclo": true,

  "Sheet_Articoli": "Articoli",
  "Sheet_Diba": "Diba",
  "Sheet_Cicli": "Cicli",

  "Col_Articoli_Codice": "Codice",
  "Col_Articoli_Descrizione": "Descrizione",
  "Col_Articoli_Famiglia": "Famiglia",
  "Col_Articoli_UdM": "UdM",

  "Col_Diba_Padre": "Padre",
  "Col_Diba_Figlio": "Figlio",
  "Col_Diba_Seq": "Seq",
  "Col_Diba_Quantita": "Quantità",
  "Col_Diba_UdM": "UdM",
  "Col_Diba_Nota": "Nota",

  "Col_Cicli_Codice": "Codice",
  "Col_Cicli_SeqFase": "SeqFase",
  "Col_Cicli_CodFase": "CodFase",
  "Col_Cicli_Descrizione": "Descrizione",
  "Col_Cicli_TempoAtt": "TempoAtt",
  "Col_Cicli_TempoLav": "TempoLav"
}
```

`config.json` deve stare nella stessa cartella di `ExcelBOM.exe`.

### Nomi fogli e colonne

I valori `Sheet_*` e `Col_*` permettono di adattare l'applicazione a Excel con intestazioni diverse da quelle di default. Se il tuo file ha ad esempio il foglio chiamato `BOM` invece di `Diba`, basta cambiare `"Sheet_Diba": "BOM"`.

| Chiave | Default | Descrizione |
|---|---|---|
| `Sheet_Articoli` | `Articoli` | Nome del foglio anagrafiche |
| `Sheet_Diba` | `Diba` | Nome del foglio distinta base |
| `Sheet_Cicli` | `Cicli` | Nome del foglio cicli |
| `Col_Articoli_Codice` | `Codice` | Colonna codice articolo |
| `Col_Articoli_Descrizione` | `Descrizione` | Colonna descrizione |
| `Col_Articoli_Famiglia` | `Famiglia` | Colonna famiglia |
| `Col_Articoli_UdM` | `UdM` | Colonna unità di misura |
| `Col_Diba_Padre` | `Padre` | Colonna articolo padre |
| `Col_Diba_Figlio` | `Figlio` | Colonna articolo figlio |
| `Col_Diba_Seq` | `Seq` | Colonna sequenza |
| `Col_Diba_Quantita` | `Quantità` | Colonna quantità |
| `Col_Diba_UdM` | `UdM` | Colonna unità di misura distinta |
| `Col_Diba_Nota` | `Nota` | Colonna nota |
| `Col_Cicli_Codice` | `Codice` | Colonna codice articolo ciclo |
| `Col_Cicli_SeqFase` | `SeqFase` | Colonna sequenza fase |
| `Col_Cicli_CodFase` | `CodFase` | Colonna codice fase |
| `Col_Cicli_Descrizione` | `Descrizione` | Colonna descrizione fase |
| `Col_Cicli_TempoAtt` | `TempoAtt` | Colonna tempo attrezzaggio (minuti interi) |
| `Col_Cicli_TempoLav` | `TempoLav` | Colonna tempo lavorazione (minuti interi) |

---

## 4. Importa le Custom Forms in Factory (opzionale)

Per visualizzare i dati di staging e i log direttamente da Factory:

1. Apri Factory → menu **Utilità** → **Importa Custom Forms**
2. Seleziona il file `ExcelBOM.CustomForms.xml`
3. Conferma l'importazione

Le form appariranno nel gruppo ribbon **Tecnologia**:
- **ExcelBOM - Log** — log esecuzioni
- **ExcelBOM - Mappings Fasi** — gestione mappature codici fase
- **ExcelBOM - Staging Articoli / Distinta / Fasi** — anteprima dati

---

## 5. Avvia l'applicazione

Lancia `ExcelBOM.exe` direttamente dalla cartella — non richiede installazione né .NET separato.

> **Requisiti minimi:** Windows 10/11 a 64 bit, accesso di rete al SQL Server Factory.
