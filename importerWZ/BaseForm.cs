using importerWZ.Datasources;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace importerWZ
{
    public class BaseForm : Form
    {
        private IContainer components = (IContainer)null;
        private BackgroundWorker bgWorkerImport;
        private DataTable data;
        private int selectedIndex;
        private StatusStrip statusStrip;
        private MenuStrip menuStrip;
        private ToolStripMenuItem tsmiPlik;
        private ToolStripMenuItem tsmiOdswiez;
        private ToolStripMenuItem tsmiZamknij;
        private ToolStripMenuItem tsmiPomoc;
        private ToolStripMenuItem tsmiOProgramie;
        private ToolStripStatusLabel tssLabel;
        private ToolStripStatusLabel tssStatus;
        private Button btnStop;
        private Button btnStart;
        private Button btnChangeFile;
        private TextBox txtFilePath;
        private RichTextBox rtbLog;
        private ProgressBar progressBar;
        private OpenFileDialog openFileDialog;
        private ToolStripMenuItem tsmiKodyBledow;
        private ComboBox cmbNabywca;

        public BaseForm()
        {
            this.InitializeComponent();
            this.data = new DataTable();
            this.cmbNabywca.SelectedIndex = 0;
            this.bgWorkerImport = new BackgroundWorker();
            this.bgWorkerImport.DoWork += new DoWorkEventHandler(this.BgWorkerImportStart);
            this.bgWorkerImport.ProgressChanged += new ProgressChangedEventHandler(this.BgWorkerImportProgressChanged);
            this.bgWorkerImport.RunWorkerCompleted += new RunWorkerCompletedEventHandler(this.BgWorkerImportCompleted);
            this.bgWorkerImport.WorkerReportsProgress = true;
            this.bgWorkerImport.WorkerSupportsCancellation = true;
        }

        private void BgWorkerImportCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                int num1 = (int)MessageBox.Show("ERROR " + e.Error.Message);
            }
            else
            {
                int num2 = (int)MessageBox.Show("Import zakończony.");
            }
            this.progressBar.Value = 0;
            this.btnChangeFile.Enabled = true;
            this.btnStart.Enabled = true;
            this.btnStop.Enabled = false;
            this.tssStatus.Text = "OK";
        }

        private void BgWorkerImportProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.progressBar.PerformStep();
        }

        private void BgWorkerImportStart(object sender, DoWorkEventArgs e)
        {
            UrzZewDataSource urzZewDataSource = new UrzZewDataSource("WZ", (UrzZewDataSource.Nabywca)this.selectedIndex);
            try
            {
                urzZewDataSource.Connect();
                foreach (DataRow row in (InternalDataCollectionBase)this.data.Rows)
                {
                    DataRow item = row;

                    if (this.bgWorkerImport.CancellationPending)
                    {
                        e.Cancel = true;
                        break;
                    }
                    string logMessage;
                    urzZewDataSource.AddDok(item, out string mesage);
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally
            {
                urzZewDataSource.Disconnect();
            }
        }

        private void btnChangeFile_Click(object sender, EventArgs e)
        {
            if (this.openFileDialog.ShowDialog() != DialogResult.OK)
                return;
            this.txtFilePath.Text = this.openFileDialog.FileName.ToString();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtFilePath.Text))
            {
                int num = (int)MessageBox.Show("Wybierz plik do zaimportowania.");
            }
            else
            {
                this.rtbLog.Clear();
                this.data.Clear();
                this.data = new ExcelDataSource(this.txtFilePath.Text).GetSheetAsDataTable();
                this.selectedIndex = this.cmbNabywca.SelectedIndex;
                this.progressBar.Maximum = this.data.Rows.Count;
                this.btnChangeFile.Enabled = false;
                this.btnStart.Enabled = false;
                this.btnStop.Enabled = true;
                this.tssStatus.Text = "PRZETWARZANIE...";
                this.bgWorkerImport.RunWorkerAsync();
            }
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            this.bgWorkerImport.CancelAsync();
        }

        private void tsmiZamknij_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void tsmiOProgramie_Click(object sender, EventArgs e)
        {
            int num = (int)MessageBox.Show("Importer WZ" + Environment.NewLine + "wersja 2.0 Ekspert" +
                " Marcin Strugliński");
        }

        private void tsmiKodyBledow_Click(object sender, EventArgs e)
        {
            int num = (int)MessageBox.Show("URZZEWNAGL_ADD 1 - brak w spisie urz. zewnętrznych podanego kodu urządzenia" + Environment.NewLine + "URZZEWNAGL_ADD 2 - brak w spisie podanego kontrahenta." + Environment.NewLine + "URZZEWNAGL_ADD 3 - brak zadanej definicji dokumentu" + Environment.NewLine + "URZZEWNAGL_ADD 9 - inny nieznany błąd" + Environment.NewLine + "URZZEWPOZ_ADD 1 - brak kartoteki" + Environment.NewLine + "URZZEWPOZ_ADD 9 - inny nieznany błąd" + Environment.NewLine + "UrzZewDataSource() - Zawiera opis błędu podczas przetwarzania rekordu.");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && this.components != null)
                this.components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(BaseForm));
            this.statusStrip = new StatusStrip();
            this.tssLabel = new ToolStripStatusLabel();
            this.tssStatus = new ToolStripStatusLabel();
            this.menuStrip = new MenuStrip();
            this.tsmiPlik = new ToolStripMenuItem();
            this.tsmiOdswiez = new ToolStripMenuItem();
            this.tsmiZamknij = new ToolStripMenuItem();
            this.tsmiPomoc = new ToolStripMenuItem();
            this.tsmiOProgramie = new ToolStripMenuItem();
            this.tsmiKodyBledow = new ToolStripMenuItem();
            this.btnStop = new Button();
            this.btnStart = new Button();
            this.btnChangeFile = new Button();
            this.txtFilePath = new TextBox();
            this.rtbLog = new RichTextBox();
            this.progressBar = new ProgressBar();
            this.openFileDialog = new OpenFileDialog();
            this.cmbNabywca = new ComboBox();
            this.statusStrip.SuspendLayout();
            this.menuStrip.SuspendLayout();
            this.SuspendLayout();
            this.statusStrip.Items.AddRange(new ToolStripItem[2]
            {
        (ToolStripItem) this.tssLabel,
        (ToolStripItem) this.tssStatus
            });
            this.statusStrip.Location = new Point(0, 539);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Size = new Size(784, 22);
            this.statusStrip.TabIndex = 0;
            this.statusStrip.Text = "statusStrip";
            this.tssLabel.BackColor = SystemColors.MenuBar;
            this.tssLabel.Name = "tssLabel";
            this.tssLabel.Size = new Size(42, 17);
            this.tssLabel.Text = "Status:";
            this.tssStatus.BackColor = SystemColors.MenuBar;
            this.tssStatus.Name = "tssStatus";
            this.tssStatus.Size = new Size(23, 17);
            this.tssStatus.Text = "OK";
            this.menuStrip.Items.AddRange(new ToolStripItem[2]
            {
        (ToolStripItem) this.tsmiPlik,
        (ToolStripItem) this.tsmiPomoc
            });
            this.menuStrip.Location = new Point(0, 0);
            this.menuStrip.Name = "menuStrip";
            this.menuStrip.Size = new Size(784, 24);
            this.menuStrip.TabIndex = 1;
            this.menuStrip.Text = "menuStrip";
            this.tsmiPlik.DropDownItems.AddRange(new ToolStripItem[2]
            {
        (ToolStripItem) this.tsmiOdswiez,
        (ToolStripItem) this.tsmiZamknij
            });
            this.tsmiPlik.Name = "tsmiPlik";
            this.tsmiPlik.Size = new Size(38, 20);
            this.tsmiPlik.Text = "Plik";
            this.tsmiOdswiez.Name = "tsmiOdswiez";
            this.tsmiOdswiez.Size = new Size(118, 22);
            this.tsmiOdswiez.Text = "Odśwież";
            this.tsmiZamknij.Name = "tsmiZamknij";
            this.tsmiZamknij.Size = new Size(118, 22);
            this.tsmiZamknij.Text = "Zamknij";
            this.tsmiZamknij.Click += new EventHandler(this.tsmiZamknij_Click);
            this.tsmiPomoc.DropDownItems.AddRange(new ToolStripItem[2]
            {
        (ToolStripItem) this.tsmiOProgramie,
        (ToolStripItem) this.tsmiKodyBledow
            });
            this.tsmiPomoc.Name = "tsmiPomoc";
            this.tsmiPomoc.Size = new Size(57, 20);
            this.tsmiPomoc.Text = "Pomoc";
            this.tsmiOProgramie.Name = "tsmiOProgramie";
            this.tsmiOProgramie.Size = new Size(143, 22);
            this.tsmiOProgramie.Text = "O Programie";
            this.tsmiOProgramie.Click += new EventHandler(this.tsmiOProgramie_Click);
            this.tsmiKodyBledow.Name = "tsmiKodyBledow";
            this.tsmiKodyBledow.Size = new Size(143, 22);
            this.tsmiKodyBledow.Text = "Kody Błędów";
            this.tsmiKodyBledow.Click += new EventHandler(this.tsmiKodyBledow_Click);
            this.btnStop.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.btnStop.Enabled = false;
            this.btnStop.FlatStyle = FlatStyle.Flat;
            this.btnStop.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Regular, GraphicsUnit.Point, (byte)238);
            this.btnStop.Image = (Image)componentResourceManager.GetObject("btnStop.Image");
            this.btnStop.Location = new Point(652, 56);
            this.btnStop.Name = "btnStop";
            this.btnStop.Size = new Size(120, 35);
            this.btnStop.TabIndex = 2;
            this.btnStop.Text = "Stop";
            this.btnStop.TextAlign = ContentAlignment.MiddleRight;
            this.btnStop.TextImageRelation = TextImageRelation.ImageBeforeText;
            this.btnStop.UseVisualStyleBackColor = true;
            this.btnStop.Click += new EventHandler(this.btnStop_Click);
            this.btnStart.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.btnStart.FlatStyle = FlatStyle.Flat;
            this.btnStart.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Regular, GraphicsUnit.Point, (byte)238);
            this.btnStart.Image = (Image)componentResourceManager.GetObject("btnStart.Image");
            this.btnStart.Location = new Point(526, 56);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new Size(120, 35);
            this.btnStart.TabIndex = 3;
            this.btnStart.Text = "Importuj";
            this.btnStart.TextAlign = ContentAlignment.MiddleRight;
            this.btnStart.TextImageRelation = TextImageRelation.ImageBeforeText;
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new EventHandler(this.btnStart_Click);
            this.btnChangeFile.FlatStyle = FlatStyle.Flat;
            this.btnChangeFile.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Regular, GraphicsUnit.Point, (byte)238);
            this.btnChangeFile.Location = new Point(12, 31);
            this.btnChangeFile.Name = "btnChangeFile";
            this.btnChangeFile.Size = new Size(75, 26);
            this.btnChangeFile.TabIndex = 4;
            this.btnChangeFile.Text = "Wybierz";
            this.btnChangeFile.UseVisualStyleBackColor = true;
            this.btnChangeFile.Click += new EventHandler(this.btnChangeFile_Click);
            this.txtFilePath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.txtFilePath.Font = new Font("Microsoft Sans Serif", 11.25f, FontStyle.Regular, GraphicsUnit.Point, (byte)238);
            this.txtFilePath.Location = new Point(93, 32);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.ReadOnly = true;
            this.txtFilePath.Size = new Size(427, 24);
            this.txtFilePath.TabIndex = 5;
            this.rtbLog.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.rtbLog.BackColor = SystemColors.Window;
            this.rtbLog.Location = new Point(12, 97);
            this.rtbLog.Name = "rtbLog";
            this.rtbLog.ReadOnly = true;
            this.rtbLog.ScrollBars = RichTextBoxScrollBars.ForcedBoth;
            this.rtbLog.Size = new Size(760, 439);
            this.rtbLog.TabIndex = 6;
            this.rtbLog.Text = "";
            this.progressBar.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.progressBar.Location = new Point(12, 68);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new Size(508, 23);
            this.progressBar.Step = 1;
            this.progressBar.TabIndex = 7;
            this.openFileDialog.FileName = "openFileDialog";
            this.openFileDialog.Filter = "Pliki Excel|*.xlsx";
            this.cmbNabywca.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.cmbNabywca.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbNabywca.FlatStyle = FlatStyle.Flat;
            this.cmbNabywca.FormattingEnabled = true;
            this.cmbNabywca.Items.AddRange(new object[2]
            {
        (object) "LOTOS",
        (object) "BP"
            });
            this.cmbNabywca.Location = new Point(526, 31);
            this.cmbNabywca.Name = "cmbNabywca";
            this.cmbNabywca.Size = new Size(246, 21);
            this.cmbNabywca.TabIndex = 8;
            this.AutoScaleMode = AutoScaleMode.None;
            this.BackColor = SystemColors.GradientInactiveCaption;
            this.ClientSize = new Size(784, 561);
            this.Controls.Add((Control)this.cmbNabywca);
            this.Controls.Add((Control)this.progressBar);
            this.Controls.Add((Control)this.rtbLog);
            this.Controls.Add((Control)this.txtFilePath);
            this.Controls.Add((Control)this.btnChangeFile);
            this.Controls.Add((Control)this.btnStart);
            this.Controls.Add((Control)this.btnStop);
            this.Controls.Add((Control)this.statusStrip);
            this.Controls.Add((Control)this.menuStrip);
            this.Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
            this.MainMenuStrip = this.menuStrip;
            this.Name = nameof(BaseForm);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Import WZ";
            this.statusStrip.ResumeLayout(false);
            this.statusStrip.PerformLayout();
            this.menuStrip.ResumeLayout(false);
            this.menuStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
