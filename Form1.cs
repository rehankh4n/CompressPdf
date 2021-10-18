using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CompressPdf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Add_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Title = "Select your PDF files";
            openFileDialog.Filter = "PDF Files (*.PDF)| *.PDF";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (var item in openFileDialog.FileNames)
                {
                    listBoxPdfFiles.Items.Add(item);
                }
            }
        }

        private bool CompressPDF(string InputFile, string OutPutFile, string CompressValue)
        {
            try
            {
                Process proc = new Process();
                ProcessStartInfo psi = new ProcessStartInfo();
                psi.CreateNoWindow = true;
                psi.ErrorDialog = false;
                psi.UseShellExecute = false;
                psi.WindowStyle = ProcessWindowStyle.Hidden;
                psi.FileName = string.Concat(Path.GetDirectoryName(Application.ExecutablePath), "\\ghsot.exe");


                string args = "-sDEVICE=pdfwrite -dCompatibilityLevel=1.4" + " -dPDFSETTINGS=/" + CompressValue + " -dNOPAUSE  -dQUIET -dBATCH" + " -sOutputFile=\"" + OutPutFile + "\" " + "\"" + InputFile + "\"";


                psi.Arguments = args;


                //start the execution
                proc.StartInfo = psi;

                proc.Start();
                proc.WaitForExit();


                return true;
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }



        string outputDir = string.Empty;

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            foreach (string item in listBoxPdfFiles.Items)
            {
                string output = $"{outputDir}\\{Path.GetFileNameWithoutExtension(item)}_Compress.pdf";

                string quality = string.Empty;
                if (comboBox1.InvokeRequired)
                {
                    comboBox1.Invoke(new Action(()=> quality = comboBox1.Text)); 
                }
                else
                {
                    quality = comboBox1.Text;
                }

                CompressPDF(item, output, quality);

                if (progressBar1.InvokeRequired)
                {
                    progressBar1.Invoke(new Action(() => progressBar1.Increment(1)));
                }
                else
                {
                    progressBar1.Increment(1);
                }
            }

            
        }

        private void Compress_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                outputDir = folderBrowserDialog.SelectedPath;
            }
            progressBar1.Maximum = listBoxPdfFiles.Items.Count;
            backgroundWorker1.RunWorkerAsync();
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Done");
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }
    }
}




