using System;
using System.Windows.Forms;
using System.Collections.Generic;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace WhiteWindowApp
{
    public partial class PDFsMerger : Form
    {
        Button mergeButton, addButton, upButton, downButton, deleteButton, saveLocationButton;
        ListBox fileList, historyList;
        OpenFileDialog openFileDialog;
        SaveFileDialog saveFileDialog;
        FolderBrowserDialog folderBrowserDialog;
        string csvFilePath = "mergedPDFs.csv";

        public PDFsMerger()
        {
            InitializeComponent();
            InitializeCustomComponents();
        }

        private void InitializeCustomComponents()
        {
            addButton = new Button { Text = "Ajouter PDF", Location = new Point(10, 10) };
            upButton = new Button { Text = "Monter", Location = new Point(420, 70) };
            downButton = new Button { Text = "Descendre", Location = new Point(420, 100) };
            deleteButton = new Button { Text = "Supprimer", Location = new Point(420, 130) };
            mergeButton = new Button { Text = "Fusionner PDFs", Location = new Point(10, 260) };
            saveLocationButton = new Button { Text = "Choisir le dossier de sauvegarde", Location = new Point(10, 290) };

            fileList = new ListBox { Location = new Point(10, 40), Size = new Size(400, 200) };
            historyList = new ListBox { Location = new Point(500, 40), Size = new Size(330, 200) };

            openFileDialog = new OpenFileDialog { Multiselect = true, Filter = "PDF files (*.pdf)|*.pdf" };
            saveFileDialog = new SaveFileDialog { Filter = "PDF files (*.pdf)|*.pdf", DefaultExt = "pdf" };
            folderBrowserDialog = new FolderBrowserDialog();

            addButton.Click += AddButton_Click;
            upButton.Click += UpButton_Click;
            downButton.Click += DownButton_Click;
            deleteButton.Click += DeleteButton_Click;
            mergeButton.Click += MergeButton_Click;
            saveLocationButton.Click += SaveLocationButton_Click;
            historyList.DoubleClick += HistoryList_DoubleClick;


            this.Controls.Add(addButton);
            this.Controls.Add(upButton);
            this.Controls.Add(downButton);
            this.Controls.Add(deleteButton);
            this.Controls.Add(mergeButton);
            this.Controls.Add(saveLocationButton);
            this.Controls.Add(fileList);
            this.Controls.Add(historyList);

            LoadHistoryFromCSV();
        }

        private void SaveLocationButton_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show("Dossier de sauvegarde sélectionné : " + folderBrowserDialog.SelectedPath, "Dossier Sélectionné", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void HistoryList_DoubleClick(object sender, EventArgs e)
        {       
            if (historyList.SelectedItem != null)
            {
                string filePath = historyList.SelectedItem.ToString();
                System.Diagnostics.Process.Start("explorer.exe", filePath);
            }
        }

        private void AddButton_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                fileList.Items.AddRange(openFileDialog.FileNames);
            }
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            if (fileList.SelectedIndex != -1)
            {
                fileList.Items.RemoveAt(fileList.SelectedIndex);
            }
        }

        private void UpButton_Click(object sender, EventArgs e) => MoveItem(-1);
        private void DownButton_Click(object sender, EventArgs e) => MoveItem(1);

        private void MoveItem(int direction)
        {
            int newIndex = fileList.SelectedIndex + direction;
            if (fileList.SelectedItem != null && newIndex >= 0 && newIndex < fileList.Items.Count)
            {
                object selected = fileList.SelectedItem;
                fileList.Items.Remove(selected);
                fileList.Items.Insert(newIndex, selected);
                fileList.SelectedIndex = newIndex;
            }
        }

        private void MergeButton_Click(object sender, EventArgs e)
        {
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                List<string> files = fileList.Items.Cast<string>().ToList();
                Task.Run(() => MergePDFs(files, saveFileDialog.FileName))
                    .ContinueWith(t =>
                    {
                        if (t.Exception != null)
                        {
                            MessageBox.Show("Une erreur est survenue pendant la fusion des PDFs : " + t.Exception.InnerException.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                        {
                            MessageBox.Show("Les PDF ont été fusionnés et enregistrés sous : " + saveFileDialog.FileName, "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            SaveHistory(saveFileDialog.FileName);
                        }
                    }, TaskScheduler.FromCurrentSynchronizationContext());
            }
        }

private void MergePDFs(List<string> files, string outputPath)
{
    using (PdfDocument outPdf = new PdfDocument())
    {
        foreach (string file in files)
        {
            using (PdfDocument inPdf = PdfReader.Open(file, PdfDocumentOpenMode.Import))
            {
                for (int i = 0; i < inPdf.PageCount; i++)
                {
                    outPdf.AddPage(inPdf.Pages[i]);
                }
            }
        }
        outPdf.Save(outputPath);

        // Sauvegarde dans le dossier spécifié par l'utilisateur
        if (!string.IsNullOrWhiteSpace(folderBrowserDialog.SelectedPath))
        {
            string backupPath = Path.Combine(folderBrowserDialog.SelectedPath, Path.GetFileName(outputPath));
            outPdf.Save(backupPath);
            SaveHistory(backupPath);
        }

        // Sauvegarde dans le sous-dossier du répertoire de l'exécutable
        string exeFolderPath = Path.Combine(Application.StartupPath, "Backup");
        if (!Directory.Exists(exeFolderPath))
        {
            Directory.CreateDirectory(exeFolderPath);
        }
        string exeFolderBackupPath = Path.Combine(exeFolderPath, Path.GetFileName(outputPath));
        outPdf.Save(exeFolderBackupPath);
        SaveHistory(exeFolderBackupPath);
    }
}


        private void SaveHistory(string filePath)
        {
            using (StreamWriter sw = new StreamWriter(csvFilePath, true))
            {
                sw.WriteLine(filePath);
            }
            LoadHistoryFromCSV();
        }

        private void LoadHistoryFromCSV()
        {
            if (File.Exists(csvFilePath))
            {
                historyList.Items.Clear();
                using (StreamReader sr = new StreamReader(csvFilePath))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        historyList.Items.Add(line);
                    }
                }
            }
        }
    }
}
