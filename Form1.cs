using System;
using System.Windows.Forms;
using System.Collections.Generic;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Linq;
using System.Threading.Tasks;

namespace WhiteWindowApp
{
    public partial class Form1 : Form
    {
        Button mergeButton;
        Button addButton;
        Button upButton;
        Button downButton;
        Button deleteButton;
        ListBox fileList;
        OpenFileDialog openFileDialog;
        SaveFileDialog saveFileDialog;

        public Form1()
        {
            InitializeComponent();
            this.BackColor = System.Drawing.Color.White;
            this.Text = "PDF's Merger";
            this.Size = new System.Drawing.Size(550, 350);

            addButton = new Button { Text = "Ajouter PDF", Location = new System.Drawing.Point(10, 10) };
            upButton = new Button { Text = "Monter", Location = new System.Drawing.Point(420, 70) };
            downButton = new Button { Text = "Descendre", Location = new System.Drawing.Point(420, 100) };
            deleteButton = new Button { Text = "Supprimer", Location = new System.Drawing.Point(420, 130) };
            mergeButton = new Button { Text = "Fusionner PDFs", Location = new System.Drawing.Point(10, 260) };

            fileList = new ListBox { Location = new System.Drawing.Point(10, 40), Size = new System.Drawing.Size(400, 200) };

            openFileDialog = new OpenFileDialog { Multiselect = true, Filter = "PDF files (*.pdf)|*.pdf" };
            saveFileDialog = new SaveFileDialog { Filter = "PDF files (*.pdf)|*.pdf", DefaultExt = "pdf" };

            addButton.Click += AddButton_Click;
            upButton.Click += UpButton_Click;
            downButton.Click += DownButton_Click;
            deleteButton.Click += DeleteButton_Click;
            mergeButton.Click += MergeButton_Click;

            this.Controls.Add(addButton);
            this.Controls.Add(upButton);
            this.Controls.Add(downButton);
            this.Controls.Add(deleteButton);
            this.Controls.Add(mergeButton);
            this.Controls.Add(fileList);
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
            }
        }
    }
}
