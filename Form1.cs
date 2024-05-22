using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data.Common;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Objets100cLib;

namespace Excel_to_Ecriture
{
    public partial class Form1 : Form
    {
        private System.Windows.Forms.Label label1, logLabel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox2, baseComptable, 
            utilisateurTextBox,txtPassword;
         private System.Windows.Forms.Button button1, btnValider;
        private OpenFileDialog openFileDialog1,openFileDialog2;
        string filePath;
        Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

        private bool isHandlingSelection = false;
        private Dictionary<ComboBox, string> previousSelections = new Dictionary<ComboBox, string>();

        private static BSCPTAApplication100c bCpta = new BSCPTAApplication100c();
        Microsoft.Office.Interop.Excel.Application excelApp;
        Workbook workbook;
        Worksheet worksheet;
        int rowCount;
        Range usedRange;
        bool isWorkbookOpen = false;
        List<string> selectedValues = new List<string>();
        Dictionary<string, string> indices = new Dictionary<string, string>();
        int erreurline;


        List<string> options = new List<string>
        {
            "Journal", "N° Pièce", "Inutilisé", "Date de pièce",
            "N° compte générale", "N° compte tiers", "Libellé écriture",
            "Référence", "Montant débit", "Montant Crédit"
        };
        List<string> columns = new List<string>
        {
                "Colonne A", "Colonne B", "Colonne C",
                "Colonne D", "Colonne E", "Colonne F",
                "Colonne G", "Colonne H", "Colonne I",
                "Colonne J", "Colonne K", "Colonne L"
        };

        List<ComboBox> comboBoxes = new List<ComboBox>();
        public Form1()
        {
            InitializeComponent(); // Initialize components first
            InitializeComponents();
        }
        private void InitializeComponents()
        {

            this.ClientSize = new System.Drawing.Size(1250, 700);
            string excelFile = ConfigurationManager.AppSettings["excelFile"];
             filePath = excelFile;

            // Label 1
            label1 = new System.Windows.Forms.Label();
            label1.Text = "Fichier d'écritures xlsx";
            label1.Location = new System.Drawing.Point(20, 20);
            label1.Size = new Size(150, 20);
            Controls.Add(label1);

            // TextBox 2
            textBox2 = new System.Windows.Forms.TextBox();
            textBox2.Location = new System.Drawing.Point(20, 50);
            textBox2.Size = new Size(250, 20);
            textBox2.Enabled = false;
            textBox2.Text = excelFile;
            Controls.Add(textBox2);

            // Button to choose file
            button1 = new System.Windows.Forms.Button();
            button1.Text = "Parcourir...";
            button1.Location = new System.Drawing.Point(280, 50);
            button1.Size = new Size(130, 30);
            button1.Click += Button1_Click;
            Controls.Add(button1);

            // Button Importer
            btnValider = new System.Windows.Forms.Button();
            btnValider.Text = "Importer";
            btnValider.Location = new System.Drawing.Point(20, 80);
            btnValider.Size = new Size(120, 30);
            btnValider.Click += Valider;
            Controls.Add(btnValider);

            // Button Quitter
            System.Windows.Forms.Button btnQuitter = new System.Windows.Forms.Button();
            btnQuitter.Text = "Fermer";
            btnQuitter.Location = new System.Drawing.Point(this.ClientSize.Width - 100, this.ClientSize.Height-50);
            btnQuitter.Size = new Size(80, 30);
            btnQuitter.Click += Quitter_Click;
            Controls.Add(btnQuitter);

            // Large Label for Log
            logLabel = new System.Windows.Forms.Label();
            logLabel.Text = "";
            logLabel.Location = new System.Drawing.Point(20, 120);
            logLabel.Size = new Size(390, 200);
            logLabel.BackColor = System.Drawing.Color.White;
            logLabel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            logLabel.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            logLabel.AutoSize = false;
            Controls.Add(logLabel);

            // Open File Dialog
            openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Files|*.xlsx;*.xls";
            openFileDialog1.Title = "Choisir un fichier Excel";
            openFileDialog1.FileOk += OpenFileDialog1_FileOk;

         
            // Right side controls
            InitializeRightSideControls();

            // TabControl for column mapping configuration
            InitializeTabControl();
        }

        private void InitializeRightSideControls()
        {
            int rightTextBoxX = this.ClientSize.Width - 600;
            int rightTextBoxY = 60;
            openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "MAE Files|*.mae";
            openFileDialog2.Title = "Choisir un fichier Comptable";
            openFileDialog2.FileOk += OpenFileDialog2_FileOk;

            // Label above the TextBox on the right side
            System.Windows.Forms.Label labelBaseComptable = new System.Windows.Forms.Label();
            labelBaseComptable.Text = "Base Comptable";
            labelBaseComptable.Location = new System.Drawing.Point(rightTextBoxX, rightTextBoxY - 20);
            labelBaseComptable.AutoSize = true;
            Controls.Add(labelBaseComptable);

            // TextBox on the right side
             baseComptable = new System.Windows.Forms.TextBox();
            baseComptable.Location = new System.Drawing.Point(rightTextBoxX, rightTextBoxY);
            baseComptable.Text = ConfigurationManager.AppSettings["bCpta"];
            baseComptable.Size = new Size(250, 20);
            baseComptable.Enabled = false;
            Controls.Add(baseComptable);

            // Button beside the TextBox
            System.Windows.Forms.Button rightButton = new System.Windows.Forms.Button();
            rightButton.Text = "Parcourir";
            rightButton.Location = new System.Drawing.Point(rightTextBoxX + baseComptable.Width + 10, rightTextBoxY);
            rightButton.Size = new Size(130, 30);
            rightButton.Click += ParcourirMae_Click;
            Controls.Add(rightButton);

            // Label and TextBox 1 below the first TextBox
            System.Windows.Forms.Label labelUtilisateur = new System.Windows.Forms.Label();
            labelUtilisateur.Text = "Utilisateur";
            labelUtilisateur.Location = new System.Drawing.Point(rightTextBoxX, rightTextBoxY + baseComptable.Height + 20);
            labelUtilisateur.AutoSize = true;
            Controls.Add(labelUtilisateur);

          utilisateurTextBox = new System.Windows.Forms.TextBox();
            utilisateurTextBox.Location = new System.Drawing.Point(rightTextBoxX, labelUtilisateur.Location.Y + labelUtilisateur.Height + 5);
            utilisateurTextBox.Text = ConfigurationManager.AppSettings["username"];
            utilisateurTextBox.Size = new Size(200, 20);
            Controls.Add(utilisateurTextBox);

            // Label and TextBox 2 below the first set
            System.Windows.Forms.Label labele = new System.Windows.Forms.Label();
            labele.Text = "Mot de passe";
            labele.Location = new System.Drawing.Point(rightTextBoxX, utilisateurTextBox.Location.Y + utilisateurTextBox.Height + 20);
            labele.AutoSize = true;
            Controls.Add(labele);

              txtPassword = new System.Windows.Forms.TextBox();
            txtPassword.Location = new System.Drawing.Point(rightTextBoxX, labele.Location.Y + labele.Height + 5);
            txtPassword.Text = ConfigurationManager.AppSettings["password"];
            txtPassword.Size = new Size(200, 20);
            txtPassword.UseSystemPasswordChar = true;
            Controls.Add(txtPassword);
        }

        // Dictionary to keep track of previous selections

        public int GetIndexByValue(Dictionary<string, string> dictionary, string value)
        {
            int index = 1;
            foreach (var kvp in dictionary)
            {
                if (kvp.Value.Equals(value))
                {
                    return index;
                }
                index++;
            }
            return -1; // Return -1 if the value is not found
        }

        private void InitializeTabControl()
        {
            System.Windows.Forms.TabControl tabControl = new System.Windows.Forms.TabControl();
            tabControl.Location = new System.Drawing.Point(20, 350);
            tabControl.Size = new System.Drawing.Size(1200, 220);

            TabPage tabPage = new TabPage();
            tabPage.Text = "Configuration de Mappage";

            // Sample column names and mapping options
    

      

            int labelY = 20;
            int labelX = 20;
            int comboBoxX = 150;
            int comboBoxY = 20;
            int columnsCount = 0;


            for (int i = 0; i < columns.Count; i++)
            {
                if(ConfigurationManager.AppSettings[columns[i]]!= "Inutilisé")
                previousSelections.Add(new ComboBox(), ConfigurationManager.AppSettings[columns[i]]);
            }

                for (int i = 0; i < columns.Count; i++)
            {
                
                // Label for column
                System.Windows.Forms.Label columnLabel = new System.Windows.Forms.Label();
                columnLabel.Text = columns[i];
                columnLabel.Location = new System.Drawing.Point(labelX, labelY);
                columnLabel.AutoSize = true;
                tabPage.Controls.Add(columnLabel);

                // ComboBox for mapping options
                ComboBox comboBox = new ComboBox();
                comboBox.Location = new System.Drawing.Point(comboBoxX, comboBoxY - 3);
                comboBox.Size = new System.Drawing.Size(200, 21);

                var filteredOptions = options.Where(option => option == "Inutilisé" || !previousSelections.Values.Contains(option)).ToArray();
                comboBox.Items.AddRange(filteredOptions);

                // Set the initial selected text, if any
                string selectedText = ConfigurationManager.AppSettings[columns[i]];
                comboBox.SelectedText = selectedText;

                // Add to previousSelections dictionary with initial value
                previousSelections[comboBox] = selectedText;

                comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;



                comboBoxes.Add(comboBox);
            
               
            
                tabPage.Controls.Add(comboBox);
                
                labelY += 30;
                comboBoxY += 30;
                columnsCount++;

                // Move to the next column after three rows
                if (columnsCount % 4 == 0)
                {
                    labelX += 400;
                    comboBoxX += 400;
                    labelY = 20;
                    comboBoxY = 20;
                }
            }
            
           


            tabControl.TabPages.Add(tabPage);

            System.Windows.Forms.Button btnSave = new System.Windows.Forms.Button();
            btnSave.Text = "Enregister";
            btnSave.Location = new System.Drawing.Point(20, tabPage.Height +50);
            btnSave.Size = new Size(100, 30);
            btnSave.Click += (sender, e) => BtnSave_Click(comboBoxes, columns);
            tabPage.Controls.Add(btnSave);

            Controls.Add(tabControl);
 
        }

        private void BtnSave_Click(List<ComboBox> comboBoxes, List<string> columns)
        {
            // Create a dictionary to hold the column mappings
            Dictionary<string, string> columnMappings = new Dictionary<string, string>();

            // Iterate through each ComboBox to gather the selected values
            for (int i = 0; i < comboBoxes.Count; i++)
            {
                ComboBox comboBox = comboBoxes[i];
                string selectedValue = comboBox.SelectedItem?.ToString();
                if (!string.IsNullOrEmpty(selectedValue))
                {
                    columnMappings[columns[i]] = selectedValue;
                }
            }

            // Save the mappings to app.config
            SaveColumnMappingsToConfig(columnMappings);

        }
        private void SaveColumnMappingsToConfig(Dictionary<string, string> columnMappings)
        {
 
            // Add or update the keys in the appSettings section
            foreach (var mapping in columnMappings)
            {
                if (config.AppSettings.Settings[mapping.Key] != null)
                {
                    config.AppSettings.Settings[mapping.Key].Value = mapping.Value;
                }
                else
                {
                    config.AppSettings.Settings.Add(mapping.Key, mapping.Value);
                }
            }

            // Save the changes to the app.config file
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
            MessageBox.Show("La configuration bien enregistré");
        }

        private void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox changedComboBox = sender as ComboBox;
            if (changedComboBox == null) return;

            // Get the new selection
            string newSelection = changedComboBox.SelectedItem?.ToString();
            if (!string.IsNullOrEmpty(newSelection))
            {
                // Remove the new selection from all other ComboBoxes
                foreach (var comboBox in comboBoxes)
                {
                    if (comboBox != changedComboBox && newSelection != "Inutilisé" && comboBox.Items.Contains(newSelection))
                    {
                        comboBox.Items.Remove(newSelection);
                    }
                }
            }

            // Get the previous selection if it exists
            if (previousSelections.TryGetValue(changedComboBox, out string previousSelection))
            {
                // Show the previous selection in a MessageBox
                //MessageBox.Show(previousSelection);

                // Add the previous selection back to all ComboBoxes
                foreach (var comboBox in comboBoxes)
                {
                    if (comboBox != changedComboBox && previousSelection != "Inutilisé" && !comboBox.Items.Contains(previousSelection))
                    {
                        comboBox.Items.Add(previousSelection);
                    }
                }
            }
            else
            {
                // Handle case where previous selection doesn't exist (optional)
               // MessageBox.Show("Previous selection not found for this ComboBox.");
            }

            // Update the previous selection for the changed ComboBox
            previousSelections[changedComboBox] = newSelection;
        }

        private static bool IsExcelFile(string filePath)
        {
            string extension = Path.GetExtension(filePath);
            return extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".xls", StringComparison.OrdinalIgnoreCase);
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            // Show the open file dialog
            openFileDialog1.ShowDialog();
        }
        private void ParcourirMae_Click(object sender, EventArgs e)
        {
            // Show the open file dialog
            openFileDialog2.ShowDialog();
        }
        private void Quitter_Click(object sender, EventArgs e)
        {
            bCpta.Close();
            // Close the form
            this.Close();
        }

        private void Valider(object sender, EventArgs e)
        {
            if (ConfigurationManager.AppSettings["username"]!= utilisateurTextBox.Text
                || ConfigurationManager.AppSettings["password"] != txtPassword.Text) { 
            config.AppSettings.Settings["username"].Value = utilisateurTextBox.Text;
            config.AppSettings.Settings["password"].Value = txtPassword.Text;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
            }
            
            for (int i = 0; i < columns.Count; i++)
            {
                indices.Remove(columns[i]);
                indices.Add(columns[i], ConfigurationManager.AppSettings[columns[i]]);
            }

            string erreur = "";

            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("Veuillez choisir un fichier excel");
                return;
            }
            if (!IsExcelFile(filePath))
            {
                MessageBox.Show("Veuillez choisir un fichier excel");
                return;
            }

            string bCptaSetting = ConfigurationManager.AppSettings["bCpta"];
            string Usernamesetting = ConfigurationManager.AppSettings["username"];
            string passwordsetting = ConfigurationManager.AppSettings["password"];

            logLabel.Text = "";
            logLabel.Text += "Ouverture de fichier Excel. \n";
             try
            {
                // Open the workbook
                OpenWorkbook(filePath);
                 // Assuming data is in the first worksheet
                worksheet = workbook.Sheets[1];
                usedRange = worksheet.UsedRange;
                rowCount = usedRange.Rows.Count;

                int success = 1;
                logLabel.Text += "Tentative d'ouverture du fichier comptable. \n";

                if (OpenBase(ref bCpta, @bCptaSetting, @Usernamesetting, @passwordsetting))
                {
                    List<DateTime> uniqueDates = new List<DateTime>();
                    HashSet<string> Journals = new HashSet<string>();

                    logLabel.Text += " Le fichier comptable est ouvert.  \n";
                    logLabel.Text += "En cours d'importation de " + rowCount + " écritures  \n";
                    int journalsIndex = GetIndexByValue(indices, "Journal");
                    int dateIndex = GetIndexByValue(indices, "Date de pièce");
                    int pieceIndex = GetIndexByValue(indices, "N° Pièce");
                    int comptegIndex = GetIndexByValue(indices, "N° compte générale");
                    int compteTiersIndex = GetIndexByValue(indices, "N° compte tiers");
                    int libelleIndex = GetIndexByValue(indices, "Libellé écriture");
                    int referenceIndex = GetIndexByValue(indices, "Référence");
                    int debitIndex = GetIndexByValue(indices, "Montant débit");
                    int creditIndex = GetIndexByValue(indices, "Montant Crédit");
                    for (int i = 1; i <= rowCount; i++)
                    {
                        string cellValue = usedRange.Cells[i, journalsIndex].Value?.ToString().Trim();
                        if (!string.IsNullOrEmpty(cellValue) && !Journals.Contains(cellValue))
                        {
                            Journals.Add(cellValue);
                        }
                    }

                    for (int i = 1; i <= rowCount; i++)
                    {
                        DateTime date = (DateTime)usedRange.Cells[i, dateIndex].Value;
                        if (!uniqueDates.Contains(date))
                        {
                            uniqueDates.Add(date);
                        }
                    }
                   
                    foreach (string journalItem in Journals)
                    {
                        foreach (DateTime date in uniqueDates)
                        {
                            IPMEncoder mProcess = bCpta.CreateProcess_Encoder();


                            for (int i = 1; i <= rowCount; i++)
                            {
                                erreurline = i;
                                float credit = 0;
                                float debit = 0;
                                IBOTiers3 tiers = null;
                                IBOCompteG3 compteg = null;

                                string piece = "";
                                string intitlule = "", reference = "";

                                DateTime rowDate = (DateTime)usedRange.Cells[i, dateIndex].Value;
                                string journ = usedRange.Cells[i, journalsIndex].Value.ToString().Trim();

                                if (rowDate.Equals(date) && journ == journalItem)
                                {
                                    piece = usedRange.Cells[i, pieceIndex].Value.ToString().Trim();

                                    compteg = bCpta.FactoryCompteG.ReadNumero(usedRange.Cells[i,comptegIndex].Value.ToString().Trim());

                                    if(usedRange.Cells[i, compteTiersIndex].Value != null) { 
                                       tiers = bCpta.FactoryTiers.ReadNumero(usedRange.Cells[i,   compteTiersIndex].Value.ToString().Trim());
                                    }

                                    intitlule = usedRange.Cells[i, libelleIndex].Value.ToString().Trim();
                                    reference = usedRange.Cells[i, referenceIndex].Value.ToString().Trim();
                                    debit = (float)usedRange.Cells[i, debitIndex].Value;
                                    credit = (float)usedRange.Cells[i, creditIndex].Value;

                                 

                                    mProcess.Journal = bCpta.FactoryJournal.ReadNumero(journ);
                                    mProcess.Date = rowDate;
                                    mProcess.EC_Piece = piece;
                                    mProcess.EC_Intitule = intitlule;
                                    mProcess.EC_Reference = reference;

                                    IBOEcriture3 ecriture = (IBOEcriture3)mProcess.FactoryEcritureIn.Create();

                                    if (compteg != null) ecriture.CompteG = compteg;
                                    if (tiers != null)
                                    {
                                        ecriture.Tiers = tiers;
                                        ecriture.EC_Echeance = DateTime.Now;
                                    }
                                    if (credit > 0)
                                    {
                                        ecriture.EC_Sens = EcritureSensType.EcritureSensTypeCredit;
                                        ecriture.EC_Montant = credit;
                                    }
                                    else if (debit > 0)
                                    {
                                        ecriture.EC_Sens = EcritureSensType.EcritureSensTypeDebit;
                                        ecriture.EC_Montant = debit;
                                    }

                                    ecriture.WriteDefault();
                                }
                            }

                            if (mProcess.CanProcess)
                            {
                                mProcess.Process();
                            }
                            else
                            {
                                success = 0;
                                for (int d = 1; d <= mProcess.Errors.Count; d++)
                                {
                                    IFailInfo iFail = mProcess.Errors[d];
                                    erreur += iFail.Text;
                                    MessageBox.Show(iFail.Text + "au journal" + journalItem + ", Date");
                                }
                            }
                        }
                    }

                    if (success == 1)
                    {
                        MessageBox.Show("La procédure est terminée");
                        logLabel.Text += "\n L'importation est terminée.\n";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur : " + ex.Message + ", Ligne : " + erreurline);
                logLabel.Text += "Erreur : " + ex.Message + ", Ligne: " + erreurline +"\n";


            }
            finally
            {
                // Ensure cleanup
                CleanupExcel();
                bCpta.Close();
            }
        }

        private void OpenWorkbook(string filePath)
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            workbook = excelApp.Workbooks.Open(filePath);
            isWorkbookOpen = true;
        }

        private void CleanupExcel()
        {
            if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            if (workbook != null)
            {
                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
            }
            if (excelApp != null)
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            worksheet = null;
            workbook = null;
            excelApp = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void OpenFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            bCpta.Close();
            logLabel.Text = "";
            workbook?.Close();
            isWorkbookOpen = false;
            // Get the selected file name and display it in TextBox
            string fileName = openFileDialog1.FileName;
            textBox2.Text = fileName;
            filePath = fileName;
            config.AppSettings.Settings["excelFile"].Value = filePath;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }

        private void OpenFileDialog2_FileOk(object sender, CancelEventArgs e)
        {
            bCpta.Close();
            logLabel.Text = "";
            workbook?.Close();
            isWorkbookOpen = false;
            // Get the selected file name and display it in TextBox
            string fileName = openFileDialog2.FileName;
            baseComptable.Text = fileName;
            filePath = fileName;
            config.AppSettings.Settings["bCpta"].Value = filePath;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }

        public static bool OpenBase(ref BSCPTAApplication100c BaseCpta, string sMae, string sUid, string sPwd)
        {
            try
            {
                BaseCpta.Name = sMae;
                BaseCpta.Loggable.UserName = sUid;
                BaseCpta.Loggable.UserPwd = sPwd;
                BaseCpta.Open();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        public static bool CloseBase(ref BSCPTAApplication100c BaseCpta)
        {
            try
            {
                if (BaseCpta.IsOpen)
                    BaseCpta.Close();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
