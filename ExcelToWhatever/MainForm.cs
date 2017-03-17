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
using OfficeOpenXml;

namespace ExcelToWhatever
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void Generate()
        {
            if (!string.IsNullOrWhiteSpace(tbTemplate.Text) && 
                !string.IsNullOrWhiteSpace(tbFilename.Text) &&
                File.Exists(tbFilename.Text))
            {

                var filename = tbFilename.Text;

                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {

                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();

                    var sb = new StringBuilder();

                    var values = new List<object>();

                    if (worksheet != null)
                    {
                        for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
                        {

                            for (int j = 1; j <= worksheet.Dimension.End.Column; j++)
                            {

                                values.Add(worksheet.Cells[i, j].Value);

                            }

                            sb.Append(string.Format(tbTemplate.Text, values.ToArray()) + Environment.NewLine);

                            values.Clear();

                        }

                        SaveAndOpen(sb);
                    }
                }

            }

        }

        private static void SaveAndOpen(StringBuilder sb)
        {
            var outfile = Path.Combine(Application.StartupPath,
                "whatever-" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".txt");

            File.WriteAllText(outfile, sb.ToString());

            Process.Start(outfile);
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            try
            {
                using (var fileDialog = new OpenFileDialog() { Filter = "Excel files (*.xlsx;*.xls) |*.xlsx;*.xls" })
                {
                    if (fileDialog.ShowDialog() == DialogResult.OK)
                    {
                        tbFilename.Text = fileDialog.FileName;
                    }
                }
            }
            catch
            {
                // Ignored
            }
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            
                Generate();
            
        }

    }
}
