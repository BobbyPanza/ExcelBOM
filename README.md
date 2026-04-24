# ExcelBOM

Applicazione Windows Forms (.NET) per importare anagrafiche articoli, distinte base e cicli di lavorazione da un file Excel in **Factory** (gestionale Intesi).

---

## Requisiti

- Windows 10/11 (l'eseguibile è self-contained, non richiede .NET installato)
- SQL Server con database Factory accessibile
- Excel con i 3 fogli descritti sotto

---

## Installazione database

Eseguire gli script SQL nell'ordine indicato sul database Factory:

```
1. sql/01_staging.sql       — Crea le tabelle staging
2. sql/03_mappings_seed.sql — Crea tabelle di supporto e seed mappings fasi
3. sql/02_sp_import.sql     — Crea la stored procedure di import
```

---

## Configurazione

Copiare `config.json` nella stessa cartella di `ExcelBOM.exe` e modificare i valori:

```json
{
  "ConnectionString": "Server=localhost;Database=Factory;User Id=sa;Password=<PASSWORD>;TrustServerCertificate=True;",
  "PATYP_DEFAULT": "C",
  "PATYF_DEFAULT": "I",
  "PAUDM_DEFAULT": "PZ",
  "FAMIGLIA_DEFAULT": "Generica",
  "EnableAnagrafiche": true,
  "EnableDistinta": true,
  "EnableCiclo": true
}
```

| Chiave | Descrizione |
|---|---|
| `ConnectionString` | Stringa di connessione SQL Server |
| `PATYP_DEFAULT` | Tipo articolo di default (es. `C` = componente) |
| `PATYF_DEFAULT` | Tipo fornitura di default (es. `I` = interno) |
| `PAUDM_DEFAULT` | Unità di misura di default se non presente in Excel |
| `FAMIGLIA_DEFAULT` | Famiglia di default se non trovata in `A_FAM` |

---

## Struttura Excel

Il file `.xlsx` deve avere esattamente **3 fogli** con i nomi e le colonne seguenti:

### Foglio `Articoli`

| Codice | Descrizione | Famiglia | UdM |
|---|---|---|---|

### Foglio `Diba`

| Padre | Figlio | Seq | Quantità | UdM | Nota |
|---|---|---|---|---|---|

### Foglio `Cicli`

| Codice | SeqFase | CodFase | Descrizione | TempoAtt | TempoLav |
|---|---|---|---|---|---|

- **CodFase**: codice fase di `A_FAS` oppure nome mappato in `Xt_ExcelBom_Mappings`
- **TempoAtt**: tempo di attrezzaggio in **minuti interi** → convertito in FAHAT (ore) / FASAT (secondi residui)
- **TempoLav**: tempo di lavorazione in **minuti interi** → convertito in FAHLV (ore) / FASLV (secondi residui)

Un file di esempio multi-livello è disponibile in `samples/test_import_multi.xlsx`.

---

## Mappature fasi ciclo

Se nel foglio `Cicli` si usano nomi descrittivi invece dei codici `A_FAS` (es. `Verniciatura` invece di `FA04`), occorre popolare la tabella `Xt_ExcelBom_Mappings`:

| IDMapping | Ambito | CodiceExcel | CodiceSistema |
|---|---|---|---|
| 1 | FASE_CICLO | Verniciatura | FA04 |
| 2 | FASE_CICLO | Saldatura | FA03 |

Lo script `sql/03_mappings_seed.sql` inserisce alcuni esempi di partenza. Le mappature si possono gestire direttamente da Factory tramite la custom form **ExcelBOM - Mappings Fasi** (vedi sezione Custom Forms).

> **Nota**: se un codice fase non è in `A_FAS` e non ha una mappatura, la riga viene evidenziata in **rosso** nella griglia di staging — l'import è comunque possibile ma il ciclo non avrà i dati di risorsa da `A_FAS`.

---

## Utilizzo

1. Avviare `ExcelBOM.exe`
2. **Sfoglia** → selezionare il file Excel
3. Selezionare i flussi da attivare: **Anagrafiche**, **Distinta**, **Ciclo**
4. **Carica Staging** → i dati vengono letti dall'Excel e visualizzati in anteprima
   - Le righe vengono colorate automaticamente:
     - **Arancione** — avviso (es. famiglia non trovata in `A_FAM`, verrà usato `FAMIGLIA_DEFAULT`)
     - **Rosso** — errore (es. codice fase non in `A_FAS` e non mappato, articolo padre/figlio inesistente)
   - La status bar mostra il conteggio di errori e avvisi
5. **Importa** → esegue l'import nel database

---

## Comportamento import

### Anagrafiche (`A_PAR`)
- **Nuovo articolo**: inserito con i valori di default da config
- **Articolo esistente**: aggiornato (famiglia, UdM, tipo); descrizione aggiornata solo se `chkAnagrafiche` attivo
- Se la famiglia Excel non esiste in `A_FAM`, viene usato `FAMIGLIA_DEFAULT`

### Distinta (`A_DBS`)
- Upsert padre-figlio; le relazioni non presenti nell'Excel non vengono eliminate

### Ciclo (`A_FAC` / `A_CCL`)
- Crea il ciclo (`A_CCL`) **solo per gli articoli che hanno almeno una fase nel file Excel**
- **Le fasi vengono sincronizzate**: fasi presenti in Factory ma assenti nel nuovo Excel vengono **eliminate**
- I dati di risorsa (operatori, macchine, percentuali, costi) vengono copiati da `A_FAS`
- I tempi (`TempoAtt`, `TempoLav` in minuti) vengono presi dall'Excel e convertiti in ore/secondi
- L'ultima fase del ciclo riceve `FAUFC = 'Y'`
- La sincronizzazione avviene **solo se**: il flusso Ciclo è attivato **e** il ciclo ha almeno una fase nel file Excel

---

## Custom Forms Factory

Nella cartella `docs/` è presente il file `ExcelBOM.CustomForms.xml` da importare in Factory per visualizzare i dati direttamente dal gestionale:

| Form | Descrizione |
|---|---|
| ExcelBOM - Log | Log delle esecuzioni import (stato, errori) |
| ExcelBOM - Mappings Fasi | Gestione mappature CodiceExcel → CodiceSistema |
| ExcelBOM - Staging Articoli | Anagrafiche dell'ultimo staging |
| ExcelBOM - Staging Distinta | Distinta base dell'ultimo staging |
| ExcelBOM - Staging Fasi | Fasi ciclo dell'ultimo staging |

Tutte le form sono nel gruppo ribbon **Tecnologia**.

---

## Tabelle database

| Tabella | Descrizione |
|---|---|
| `Xt_ExcelBom_Run` | Log delle sessioni di import |
| `Xt_ExcelBom_Parts` | Staging anagrafiche |
| `Xt_ExcelBom_Rel` | Staging distinta base |
| `Xt_ExcelBom_Phases` | Staging fasi ciclo |
| `Xt_ExcelBom_Mappings` | Mapping codici Excel → codici sistema |
| `Xt_ExcelBom_LogRun` | Log esecuzione SP |

---

## Build dal sorgente

```bash
dotnet restore
dotnet build -c Release
dotnet publish -c Release --self-contained -r win-x64 -o publish/
```
