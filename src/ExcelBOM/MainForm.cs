using System.Data;
using System.Text.Json;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;

namespace ExcelBOM;

public partial class MainForm : Form
{
    private Dictionary<string, string> config = new();
    private List<Anagrafica> anagrafiche = new();
    private List<Distinta> distinte = new();
    private List<Fase> fasi = new();
    private List<Color> anagraficheColors = new();
    private List<Color> distinteColors = new();
    private List<Color> fasiColors = new();

    public MainForm()
    {
        InitializeComponent();
        LoadConfig();
    }

    private void InitializeComponent()
    {
        this.label1 = new System.Windows.Forms.Label();
        this.txtExcelPath = new System.Windows.Forms.TextBox();
        this.btnBrowse = new System.Windows.Forms.Button();
        this.chkAnagrafiche = new System.Windows.Forms.CheckBox();
        this.chkDistinta = new System.Windows.Forms.CheckBox();
        this.chkCiclo = new System.Windows.Forms.CheckBox();
        this.btnLoadStaging = new System.Windows.Forms.Button();
        this.tabControl1 = new System.Windows.Forms.TabControl();
        this.tabPage1 = new System.Windows.Forms.TabPage();
        this.dgvArticoli = new System.Windows.Forms.DataGridView();
        this.tabPage2 = new System.Windows.Forms.TabPage();
        this.dgvDiba = new System.Windows.Forms.DataGridView();
        this.tabPage3 = new System.Windows.Forms.TabPage();
        this.dgvCicli = new System.Windows.Forms.DataGridView();
        this.btnImport = new System.Windows.Forms.Button();
        this.statusStrip1 = new System.Windows.Forms.StatusStrip();
        this.statusLabel = new System.Windows.Forms.ToolStripStatusLabel();
        this.tabControl1.SuspendLayout();
        this.tabPage1.SuspendLayout();
        this.tabPage2.SuspendLayout();
        this.tabPage3.SuspendLayout();
        this.statusStrip1.SuspendLayout();
        this.SuspendLayout();
        // 
        // label1
        // 
        this.label1.AutoSize = true;
        this.label1.Location = new System.Drawing.Point(12, 15);
        this.label1.Name = "label1";
        this.label1.Size = new System.Drawing.Size(58, 15);
        this.label1.TabIndex = 0;
        this.label1.Text = "File Excel:";
        // 
        // txtExcelPath
        // 
        this.txtExcelPath.Location = new System.Drawing.Point(76, 12);
        this.txtExcelPath.Name = "txtExcelPath";
        this.txtExcelPath.Size = new System.Drawing.Size(300, 23);
        this.txtExcelPath.TabIndex = 1;
        // 
        // btnBrowse
        // 
        this.btnBrowse.Location = new System.Drawing.Point(382, 10);
        this.btnBrowse.Name = "btnBrowse";
        this.btnBrowse.Size = new System.Drawing.Size(75, 23);
        this.btnBrowse.TabIndex = 2;
        this.btnBrowse.Text = "Sfoglia";
        this.btnBrowse.UseVisualStyleBackColor = true;
        this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
        // 
        // chkAnagrafiche
        // 
        this.chkAnagrafiche.AutoSize = true;
        this.chkAnagrafiche.Checked = true;
        this.chkAnagrafiche.CheckState = System.Windows.Forms.CheckState.Checked;
        this.chkAnagrafiche.Location = new System.Drawing.Point(15, 45);
        this.chkAnagrafiche.Name = "chkAnagrafiche";
        this.chkAnagrafiche.Size = new System.Drawing.Size(90, 19);
        this.chkAnagrafiche.TabIndex = 3;
        this.chkAnagrafiche.Text = "Anagrafiche";
        this.chkAnagrafiche.UseVisualStyleBackColor = true;
        // 
        // chkDistinta
        // 
        this.chkDistinta.AutoSize = true;
        this.chkDistinta.Checked = true;
        this.chkDistinta.CheckState = System.Windows.Forms.CheckState.Checked;
        this.chkDistinta.Location = new System.Drawing.Point(111, 45);
        this.chkDistinta.Name = "chkDistinta";
        this.chkDistinta.Size = new System.Drawing.Size(65, 19);
        this.chkDistinta.TabIndex = 4;
        this.chkDistinta.Text = "Distinta";
        this.chkDistinta.UseVisualStyleBackColor = true;
        // 
        // chkCiclo
        // 
        this.chkCiclo.AutoSize = true;
        this.chkCiclo.Checked = true;
        this.chkCiclo.CheckState = System.Windows.Forms.CheckState.Checked;
        this.chkCiclo.Location = new System.Drawing.Point(182, 45);
        this.chkCiclo.Name = "chkCiclo";
        this.chkCiclo.Size = new System.Drawing.Size(50, 19);
        this.chkCiclo.TabIndex = 5;
        this.chkCiclo.Text = "Ciclo";
        this.chkCiclo.UseVisualStyleBackColor = true;
        // 
        // btnLoadStaging
        // 
        this.btnLoadStaging.Location = new System.Drawing.Point(250, 40);
        this.btnLoadStaging.Name = "btnLoadStaging";
        this.btnLoadStaging.Size = new System.Drawing.Size(100, 23);
        this.btnLoadStaging.TabIndex = 6;
        this.btnLoadStaging.Text = "Carica Staging";
        this.btnLoadStaging.UseVisualStyleBackColor = true;
        this.btnLoadStaging.Click += new System.EventHandler(this.btnLoadStaging_Click);
        // 
        // tabControl1
        // 
        this.tabControl1.Controls.Add(this.tabPage1);
        this.tabControl1.Controls.Add(this.tabPage2);
        this.tabControl1.Controls.Add(this.tabPage3);
        this.tabControl1.Location = new System.Drawing.Point(12, 75);
        this.tabControl1.Name = "tabControl1";
        this.tabControl1.SelectedIndex = 0;
        this.tabControl1.Size = new System.Drawing.Size(776, 300);
        this.tabControl1.TabIndex = 7;
        // 
        // tabPage1
        // 
        this.tabPage1.Controls.Add(this.dgvArticoli);
        this.tabPage1.Location = new System.Drawing.Point(4, 22);
        this.tabPage1.Name = "tabPage1";
        this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
        this.tabPage1.Size = new System.Drawing.Size(768, 274);
        this.tabPage1.TabIndex = 0;
        this.tabPage1.Text = "Articoli";
        this.tabPage1.UseVisualStyleBackColor = true;
        // 
        // dgvArticoli
        // 
        this.dgvArticoli.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        this.dgvArticoli.Dock = System.Windows.Forms.DockStyle.Fill;
        this.dgvArticoli.Location = new System.Drawing.Point(3, 3);
        this.dgvArticoli.Name = "dgvArticoli";
        this.dgvArticoli.ReadOnly = true;
        this.dgvArticoli.RowTemplate.Height = 25;
        this.dgvArticoli.Size = new System.Drawing.Size(762, 268);
        this.dgvArticoli.TabIndex = 0;
        // 
        // tabPage2
        // 
        this.tabPage2.Controls.Add(this.dgvDiba);
        this.tabPage2.Location = new System.Drawing.Point(4, 22);
        this.tabPage2.Name = "tabPage2";
        this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
        this.tabPage2.Size = new System.Drawing.Size(768, 274);
        this.tabPage2.TabIndex = 1;
        this.tabPage2.Text = "Diba";
        this.tabPage2.UseVisualStyleBackColor = true;
        // 
        // dgvDiba
        // 
        this.dgvDiba.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        this.dgvDiba.Dock = System.Windows.Forms.DockStyle.Fill;
        this.dgvDiba.Location = new System.Drawing.Point(3, 3);
        this.dgvDiba.Name = "dgvDiba";
        this.dgvDiba.ReadOnly = true;
        this.dgvDiba.RowTemplate.Height = 25;
        this.dgvDiba.Size = new System.Drawing.Size(762, 268);
        this.dgvDiba.TabIndex = 0;
        // 
        // tabPage3
        // 
        this.tabPage3.Controls.Add(this.dgvCicli);
        this.tabPage3.Location = new System.Drawing.Point(4, 22);
        this.tabPage3.Name = "tabPage3";
        this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
        this.tabPage3.Size = new System.Drawing.Size(768, 274);
        this.tabPage3.TabIndex = 2;
        this.tabPage3.Text = "Cicli";
        this.tabPage3.UseVisualStyleBackColor = true;
        // 
        // dgvCicli
        // 
        this.dgvCicli.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        this.dgvCicli.Dock = System.Windows.Forms.DockStyle.Fill;
        this.dgvCicli.Location = new System.Drawing.Point(3, 3);
        this.dgvCicli.Name = "dgvCicli";
        this.dgvCicli.ReadOnly = true;
        this.dgvCicli.RowTemplate.Height = 25;
        this.dgvCicli.Size = new System.Drawing.Size(762, 268);
        this.dgvCicli.TabIndex = 0;
        // 
        // btnImport
        // 
        this.btnImport.Location = new System.Drawing.Point(356, 40);
        this.btnImport.Name = "btnImport";
        this.btnImport.Size = new System.Drawing.Size(75, 23);
        this.btnImport.TabIndex = 8;
        this.btnImport.Text = "Importa";
        this.btnImport.UseVisualStyleBackColor = true;
        this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
        // 
        // statusStrip1
        // 
        this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
        this.statusLabel});
        this.statusStrip1.Location = new System.Drawing.Point(0, 381);
        this.statusStrip1.Name = "statusStrip1";
        this.statusStrip1.Size = new System.Drawing.Size(800, 22);
        this.statusStrip1.TabIndex = 9;
        this.statusStrip1.Text = "statusStrip1";
        // 
        // statusLabel
        // 
        this.statusLabel.Name = "statusLabel";
        this.statusLabel.Size = new System.Drawing.Size(118, 17);
        this.statusLabel.Text = "Pronto";
        // 
        // MainForm
        // 
        this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.ClientSize = new System.Drawing.Size(800, 403);
        this.Controls.Add(this.statusStrip1);
        this.Controls.Add(this.btnImport);
        this.Controls.Add(this.tabControl1);
        this.Controls.Add(this.btnLoadStaging);
        this.Controls.Add(this.chkCiclo);
        this.Controls.Add(this.chkDistinta);
        this.Controls.Add(this.chkAnagrafiche);
        this.Controls.Add(this.btnBrowse);
        this.Controls.Add(this.txtExcelPath);
        this.Controls.Add(this.label1);
        this.Name = "MainForm";
        this.Text = "ExcelBOM Import";
        this.tabControl1.ResumeLayout(false);
        this.tabPage1.ResumeLayout(false);
        this.tabPage2.ResumeLayout(false);
        this.tabPage3.ResumeLayout(false);
        this.statusStrip1.ResumeLayout(false);
        this.statusStrip1.PerformLayout();
        this.ResumeLayout(false);
        this.PerformLayout();

    }

    private System.Windows.Forms.Label label1;
    private System.Windows.Forms.TextBox txtExcelPath;
    private System.Windows.Forms.Button btnBrowse;
    private System.Windows.Forms.CheckBox chkAnagrafiche;
    private System.Windows.Forms.CheckBox chkDistinta;
    private System.Windows.Forms.CheckBox chkCiclo;
    private System.Windows.Forms.Button btnLoadStaging;
    private System.Windows.Forms.TabControl tabControl1;
    private System.Windows.Forms.TabPage tabPage1;
    private System.Windows.Forms.DataGridView dgvArticoli;
    private System.Windows.Forms.TabPage tabPage2;
    private System.Windows.Forms.DataGridView dgvDiba;
    private System.Windows.Forms.TabPage tabPage3;
    private System.Windows.Forms.DataGridView dgvCicli;
    private System.Windows.Forms.Button btnImport;
    private System.Windows.Forms.StatusStrip statusStrip1;
    private System.Windows.Forms.ToolStripStatusLabel statusLabel;

    private void LoadConfig()
    {
        var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
        if (File.Exists(configPath))
        {
            var json = File.ReadAllText(configPath);
            config = JsonSerializer.Deserialize<Dictionary<string, string>>(json) ?? new();
        }
    }

    private void btnBrowse_Click(object sender, EventArgs e)
    {
        using var ofd = new OpenFileDialog
        {
            Filter = "Excel files (*.xlsx)|*.xlsx",
            Title = "Seleziona file Excel"
        };
        if (ofd.ShowDialog() == DialogResult.OK)
        {
            txtExcelPath.Text = ofd.FileName;
        }
    }

    private async void btnLoadStaging_Click(object sender, EventArgs e)
    {
        if (string.IsNullOrEmpty(txtExcelPath.Text) || !File.Exists(txtExcelPath.Text))
        {
            MessageBox.Show("Seleziona un file Excel valido.");
            return;
        }

        try
        {
            statusLabel.Text = "Caricamento dati...";
            await LoadExcelDataAsync(txtExcelPath.Text);
            BindDataToGrids();
            statusLabel.Text = "Verifica dati...";
            await RunValidationAsync();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Errore nel caricamento: {ex.Message}");
            statusLabel.Text = "Errore nel caricamento.";
        }
    }

    private async Task LoadExcelDataAsync(string filePath)
    {
        anagrafiche.Clear();
        distinte.Clear();
        fasi.Clear();

        using var workbook = new XLWorkbook(filePath);

        // Load Articoli
        if (chkAnagrafiche.Checked)
        {
            var sheet = workbook.Worksheet("Articoli");
            if (sheet != null)
            {
                var headerRow = sheet.Row(1);
                var colIndexes = new Dictionary<string, int>();
                for (int col = 1; col <= headerRow.LastCellUsed().Address.ColumnNumber; col++)
                {
                    var header = headerRow.Cell(col).GetString();
                    colIndexes[header] = col;
                }

                foreach (var row in sheet.RowsUsed().Skip(1))
                {
                    var codice = GetCellValue(row, colIndexes, "Codice");
                    if (string.IsNullOrEmpty(codice)) continue;

                    var famiglia = GetCellValue(row, colIndexes, "Famiglia");
                    var udm = GetCellValue(row, colIndexes, "UdM");
                    anagrafiche.Add(new Anagrafica
                    {
                        PACOD = codice,
                        PADSC = GetCellValue(row, colIndexes, "Descrizione"),
                        FMCOD = string.IsNullOrEmpty(famiglia) ? config.GetValueOrDefault("FAMIGLIA_DEFAULT", "") : famiglia,
                        PAUDM = string.IsNullOrEmpty(udm) ? config.GetValueOrDefault("PAUDM_DEFAULT", "PZ") : udm,
                        PATYP = config.GetValueOrDefault("PATYP_DEFAULT", "C"),
                        PATYF = config.GetValueOrDefault("PATYF_DEFAULT", "I")
                    });
                }
            }
        }

        // Load Diba
        if (chkDistinta.Checked)
        {
            var sheet = workbook.Worksheet("Diba");
            if (sheet != null)
            {
                var headerRow = sheet.Row(1);
                var colIndexes = new Dictionary<string, int>();
                for (int col = 1; col <= headerRow.LastCellUsed().Address.ColumnNumber; col++)
                {
                    var header = sheet.Cell(1, col).GetString();
                    colIndexes[header] = col;
                }

                foreach (var row in sheet.RowsUsed().Skip(1))
                {
                    var padre = GetCellValue(row, colIndexes, "Padre");
                    var figlio = GetCellValue(row, colIndexes, "Figlio");
                    if (string.IsNullOrEmpty(padre) || string.IsNullOrEmpty(figlio)) continue;

                    distinte.Add(new Distinta
                    {
                        DBPAR = padre,
                        DBCHL = figlio,
                        DBSEQ = GetCellValue(row, colIndexes, "Seq"),
                        DBQTA = GetCellValue(row, colIndexes, "Quantità"),
                        DBUDM = GetCellValue(row, colIndexes, "UdM"),
                        DBNT1 = GetCellValue(row, colIndexes, "Nota")
                    });
                }
            }
        }

        // Load Cicli
        if (chkCiclo.Checked)
        {
            var sheet = workbook.Worksheet("Cicli");
            if (sheet != null)
            {
                var headerRow = sheet.Row(1);
                var colIndexes = new Dictionary<string, int>();
                for (int col = 1; col <= headerRow.LastCellUsed().Address.ColumnNumber; col++)
                {
                    var header = sheet.Cell(1, col).GetString();
                    colIndexes[header] = col;
                }

                foreach (var row in sheet.RowsUsed().Skip(1))
                {
                    var codice = GetCellValue(row, colIndexes, "Codice");
                    var codFase = GetCellValue(row, colIndexes, "CodFase");
                    if (string.IsNullOrEmpty(codice) || string.IsNullOrEmpty(codFase)) continue;

                    fasi.Add(new Fase
                    {
                        PACOD = codice,
                        FASEQ = GetCellValue(row, colIndexes, "SeqFase"),
                        FACOD = codFase,
                        FADSC = GetCellValue(row, colIndexes, "Descrizione"),
                        TempoAtt = GetCellValue(row, colIndexes, "TempoAtt"),
                        TempoLav = GetCellValue(row, colIndexes, "TempoLav")
                    });
                }
            }
        }
    }

    private void BindDataToGrids()
    {
        dgvArticoli.DataSource = anagrafiche.Select(a => new
        {
            Codice = a.PACOD,
            Descrizione = a.PADSC,
            Famiglia = a.FMCOD,
            UdM = a.PAUDM,
            Tipo = a.PATYP,
            Fornitura = a.PATYF
        }).ToList();

        dgvDiba.DataSource = distinte.Select(d => new
        {
            Padre = d.DBPAR,
            Figlio = d.DBCHL,
            Seq = d.DBSEQ,
            Quantità = d.DBQTA,
            UdM = d.DBUDM,
            Nota = d.DBNT1
        }).ToList();

        dgvCicli.DataSource = fasi.Select(f => new
        {
            Codice = f.PACOD,
            SeqFase = f.FASEQ,
            CodFase = f.FACOD,
            Descrizione = f.FADSC,
            TempoAtt = f.TempoAtt,
            TempoLav = f.TempoLav
        }).ToList();
    }

    private async void btnImport_Click(object sender, EventArgs e)
    {
        if (anagrafiche.Count == 0 && distinte.Count == 0 && fasi.Count == 0)
        {
            MessageBox.Show("Carica prima i dati di staging.");
            return;
        }

        try
        {
            statusLabel.Text = "Importazione in corso...";
            await ImportToDatabaseAsync();
            statusLabel.Text = "Importazione completata.";
            MessageBox.Show("Importazione completata con successo!");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Errore nell'importazione: {ex.Message}");
            statusLabel.Text = "Errore nell'importazione.";
        }
    }

    private async Task ImportToDatabaseAsync()
    {
        var connectionString = config.GetValueOrDefault("ConnectionString", "Server=localhost;Database=Factory;User Id=sa;Password=Intsupport1;TrustServerCertificate=True;");

        using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync();

        var runId = await InsertRunAsync(connection);

        using var transaction = connection.BeginTransaction();
        try
        {
            await BulkInsertAnagraficheAsync(connection, transaction, runId, anagrafiche);
            await BulkInsertDistinteAsync(connection, transaction, runId, distinte);
            await BulkInsertFasiAsync(connection, transaction, runId, fasi);
            var famigliaDefault = config.GetValueOrDefault("FAMIGLIA_DEFAULT", "Generica");
        await CallImportSPAsync(connection, transaction, runId, chkAnagrafiche.Checked, false, chkCiclo.Checked, chkDistinta.Checked, false, famigliaDefault);
            transaction.Commit();
        }
        catch
        {
            transaction.Rollback();
            throw;
        }
    }

    private async Task<int> InsertRunAsync(SqlConnection conn)
    {
        var cmd = new SqlCommand("INSERT INTO Xt_ExcelBom_Run (ImportDate) VALUES (GETDATE()); SELECT SCOPE_IDENTITY();", conn);
        var result = await cmd.ExecuteScalarAsync();
        return Convert.ToInt32(result);
    }

    private async Task BulkInsertAnagraficheAsync(SqlConnection conn, SqlTransaction transaction, int runId, List<Anagrafica> list)
    {
        if (list.Count == 0) return;
        var dt = new DataTable();
        dt.Columns.Add("IDRun", typeof(int));
        dt.Columns.Add("PACOD", typeof(string));
        dt.Columns.Add("PADSC", typeof(string));
        dt.Columns.Add("FMCOD", typeof(string));
        dt.Columns.Add("PAUDM", typeof(string));
        dt.Columns.Add("PATYP", typeof(string));
        dt.Columns.Add("PATYF", typeof(string));

        foreach (var a in list)
        {
            dt.Rows.Add(runId, a.PACOD, a.PADSC, a.FMCOD, a.PAUDM, a.PATYP, a.PATYF);
        }

        using var bulk = new SqlBulkCopy(conn, SqlBulkCopyOptions.Default, transaction);
        bulk.DestinationTableName = "Xt_ExcelBom_Parts";
        foreach (DataColumn col in dt.Columns) bulk.ColumnMappings.Add(col.ColumnName, col.ColumnName);
        await bulk.WriteToServerAsync(dt);
    }

    private async Task BulkInsertDistinteAsync(SqlConnection conn, SqlTransaction transaction, int runId, List<Distinta> list)
    {
        if (list.Count == 0) return;
        var dt = new DataTable();
        dt.Columns.Add("IDRun", typeof(int));
        dt.Columns.Add("DBPAR", typeof(string));
        dt.Columns.Add("DBCHL", typeof(string));
        dt.Columns.Add("DBSEQ", typeof(string));
        dt.Columns.Add("DBQTA", typeof(string));
        dt.Columns.Add("DBUDM", typeof(string));
        dt.Columns.Add("DBNT1", typeof(string));

        foreach (var d in list)
        {
            dt.Rows.Add(runId, d.DBPAR, d.DBCHL, d.DBSEQ, d.DBQTA, d.DBUDM, d.DBNT1);
        }

        using var bulk = new SqlBulkCopy(conn, SqlBulkCopyOptions.Default, transaction);
        bulk.DestinationTableName = "Xt_ExcelBom_Rel";
        foreach (DataColumn col in dt.Columns) bulk.ColumnMappings.Add(col.ColumnName, col.ColumnName);
        await bulk.WriteToServerAsync(dt);
    }

    private async Task BulkInsertFasiAsync(SqlConnection conn, SqlTransaction transaction, int runId, List<Fase> list)
    {
        if (list.Count == 0) return;
        var dt = new DataTable();
        dt.Columns.Add("IDRun", typeof(int));
        dt.Columns.Add("PACOD", typeof(string));
        dt.Columns.Add("FASEQ", typeof(string));
        dt.Columns.Add("FACOD", typeof(string));
        dt.Columns.Add("FADSC", typeof(string));
        dt.Columns.Add("TempoAtt", typeof(string));
        dt.Columns.Add("TempoLav", typeof(string));

        foreach (var f in list)
        {
            dt.Rows.Add(runId, f.PACOD, f.FASEQ, f.FACOD, f.FADSC, f.TempoAtt, f.TempoLav);
        }

        using var bulk = new SqlBulkCopy(conn, SqlBulkCopyOptions.Default, transaction);
        bulk.DestinationTableName = "Xt_ExcelBom_Phases";
        foreach (DataColumn col in dt.Columns) bulk.ColumnMappings.Add(col.ColumnName, col.ColumnName);
        await bulk.WriteToServerAsync(dt);
    }

    private async Task CallImportSPAsync(SqlConnection conn, SqlTransaction transaction, int runId, bool updDesc, bool updPesi, bool updCiclo, bool updDistinta, bool updGrandezze, string famigliaDefault)
    {
        var cmd = new SqlCommand("EXEC dbo.Xsp_ExcelBom_Import @IDRun, @UpdDesc, @UpdPesi, @UpdCiclo, @UpdDistinta, @UpdGrandezze, @FamigliaDefault", conn, transaction);
        cmd.Parameters.AddWithValue("@IDRun", runId);
        cmd.Parameters.AddWithValue("@UpdDesc", updDesc);
        cmd.Parameters.AddWithValue("@UpdPesi", updPesi);
        cmd.Parameters.AddWithValue("@UpdCiclo", updCiclo);
        cmd.Parameters.AddWithValue("@UpdDistinta", updDistinta);
        cmd.Parameters.AddWithValue("@UpdGrandezze", updGrandezze);
        cmd.Parameters.AddWithValue("@FamigliaDefault", famigliaDefault);
        await cmd.ExecuteNonQueryAsync();
    }

    private async Task RunValidationAsync()
    {
        anagraficheColors = Enumerable.Repeat(Color.White, anagrafiche.Count).ToList();
        distinteColors    = Enumerable.Repeat(Color.White, distinte.Count).ToList();
        fasiColors        = Enumerable.Repeat(Color.White, fasi.Count).ToList();

        int warnings = 0, errors = 0;
        var connectionString = config.GetValueOrDefault("ConnectionString", "Server=localhost;Database=Factory;User Id=sa;Password=Intsupport1;TrustServerCertificate=True;");

        try
        {
            using var conn = new SqlConnection(connectionString);
            await conn.OpenAsync();

            var validFamiglie = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var cmd = new SqlCommand("SELECT FMCOD FROM A_FAM", conn))
            using (var r = await cmd.ExecuteReaderAsync())
                while (await r.ReadAsync()) validFamiglie.Add(r.GetString(0));

            var validFaseCodici = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var cmd = new SqlCommand("SELECT FACOD FROM A_FAS", conn))
            using (var r = await cmd.ExecuteReaderAsync())
                while (await r.ReadAsync()) validFaseCodici.Add(r.GetString(0));

            var mappedFaseExcel = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var cmd = new SqlCommand("SELECT CodiceExcel FROM Xt_ExcelBom_Mappings WHERE Ambito = 'FASE_CICLO'", conn))
            using (var r = await cmd.ExecuteReaderAsync())
                while (await r.ReadAsync()) mappedFaseExcel.Add(r.GetString(0));

            var articoliInDb = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var cmd = new SqlCommand("SELECT PACOD FROM A_PAR", conn))
            using (var r = await cmd.ExecuteReaderAsync())
                while (await r.ReadAsync()) articoliInDb.Add(r.GetString(0));

            // Articoli: FMCOD non in A_FAM → avviso (verrà usato il default da config)
            for (int i = 0; i < anagrafiche.Count; i++)
            {
                var fmcod = anagrafiche[i].FMCOD;
                if (!string.IsNullOrEmpty(fmcod) && !validFamiglie.Contains(fmcod))
                {
                    anagraficheColors[i] = Color.LightSalmon;
                    warnings++;
                }
            }

            // Fasi: FACOD non in A_FAS e non nei mapping → errore
            for (int i = 0; i < fasi.Count; i++)
            {
                var facod = fasi[i].FACOD;
                if (!validFaseCodici.Contains(facod) && !mappedFaseExcel.Contains(facod))
                {
                    fasiColors[i] = Color.LightCoral;
                    errors++;
                }
            }

            // Distinte: padre o figlio non nell'Excel e non in DB → errore
            var articoliLocali = new HashSet<string>(anagrafiche.Select(a => a.PACOD), StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < distinte.Count; i++)
            {
                var d = distinte[i];
                bool padreOk  = articoliLocali.Contains(d.DBPAR) || articoliInDb.Contains(d.DBPAR);
                bool figlioOk = articoliLocali.Contains(d.DBCHL) || articoliInDb.Contains(d.DBCHL);
                if (!padreOk || !figlioOk)
                {
                    distinteColors[i] = Color.LightCoral;
                    errors++;
                }
            }
        }
        catch { /* DB non raggiungibile: si procede senza colori */ }

        ApplyValidationColors();

        var summary = $"{anagrafiche.Count} articoli, {distinte.Count} distinte, {fasi.Count} fasi";
        if (errors > 0)   summary += $" — {errors} errori";
        if (warnings > 0) summary += $" — {warnings} avvisi";
        if (errors == 0 && warnings == 0) summary += " — tutto OK";
        statusLabel.Text = summary;
    }

    private void ApplyValidationColors()
    {
        for (int i = 0; i < anagraficheColors.Count && i < dgvArticoli.Rows.Count; i++)
            dgvArticoli.Rows[i].DefaultCellStyle.BackColor = anagraficheColors[i];
        for (int i = 0; i < distinteColors.Count && i < dgvDiba.Rows.Count; i++)
            dgvDiba.Rows[i].DefaultCellStyle.BackColor = distinteColors[i];
        for (int i = 0; i < fasiColors.Count && i < dgvCicli.Rows.Count; i++)
            dgvCicli.Rows[i].DefaultCellStyle.BackColor = fasiColors[i];
    }

    private string GetCellValue(IXLRow row, Dictionary<string, int> colIndexes, string colName)
    {
        if (!colIndexes.TryGetValue(colName, out var colIndex)) return "";
        var cell = row.Cell(colIndex);
        if (cell.DataType == XLDataType.Number)
            return cell.GetDouble().ToString(System.Globalization.CultureInfo.InvariantCulture);
        return cell.GetString();
    }
}