using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using EXCEL = Microsoft.Office.Interop.Excel;

namespace workspace23_8
{
    public partial class Form1 : Form
    {
        EXCEL.Application task = new EXCEL.Application();
        EXCEL.Workbook bookafter;
        public Form1()
        {
            InitializeComponent();
            fileopener();
            FormClosed += new FormClosedEventHandler(after);
        }
        private void textBox1_Click(object sender, EventArgs e) { fileopener(); }
        private void button1_Click(object sender, EventArgs e)
        {
            processing pc = new processing("processing...", new DoWorkEventHandler(processing_doWork), null);
            DialogResult result = pc.ShowDialog(this);
            if (result == DialogResult.Cancel)
            {
                MessageBox.Show("キャンセルされました","",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
            else if (result == DialogResult.Abort)
            {
                Exception ex = pc.Error;
                MessageBox.Show("エラー:" + ex.Message,"",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            else if (result == DialogResult.OK) { MessageBox.Show("成功しました","",MessageBoxButtons.OK,MessageBoxIcon.Information); }
            pc.Dispose();
        }
        private void processing_doWork(object sender,DoWorkEventArgs e)
        {
            BackgroundWorker bw = (BackgroundWorker)sender;
            string imagename;
            imagename = textBox1.Text;
            int wid, hig;
            double prog, total, count = 0;
            string r, g, bb;
            Bitmap img = new Bitmap(imagename);
            using(var book=new XLWorkbook())
            {
                var sheet1 = book.Worksheets.Add("Sheet1");
                sheet1.Columns().Width = 0.08;
                sheet1.Rows().Height = 4.50;
                wid = img.Width;
                hig = img.Height;
                book.SaveAs(@"C:\Users\" + Environment.UserName + "\\Desktop\\testafter.xlsx");
                book.Dispose();
                total = (wid + 1) * (hig + 1);
            }
            EXCEL.Worksheet sheet;
            EXCEL.Range cell;
            bookafter = (EXCEL.Workbook)(task.Workbooks.Open(@"C:\Users\" + Environment.UserName + "\\Desktop\\testafter.xlsx"));
            task.Application.Visible = true;
            sheet = (EXCEL.Worksheet)bookafter.Sheets[1];
            sheet.Select(Type.Missing);
            for (int a = 0; a < hig; a++)
            {
                for (int b = 0; b < wid; b++)
                {
                    Color col = img.GetPixel(b, a);
                    r = Convert.ToString(col.R, 16);
                    g = Convert.ToString(col.G, 16);
                    bb = Convert.ToString(col.B, 16);
                    if (r.Length == 1) { r = "0" + r; }
                    if (g.Length == 1) { g = "0" + g; }
                    if (bb.Length == 1) { bb = "0" + bb; }
                    cell = (EXCEL.Range)sheet.Cells[a + 1, b + 1];
                    cell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#" + r + g + bb));
                    count++;
                    prog = count / total * 100;
                    prog = Math.Round(prog, 2, MidpointRounding.AwayFromZero);
                    bw.ReportProgress((int)prog, b.ToString() + "," + a.ToString() + "," + prog.ToString() + "%");
                    if (bw.CancellationPending)
                    {
                        e.Cancel = true;
                        return;
                    }
                }
            }
            bookafter.SaveAs(@"C:\Users\" + Environment.UserName + "\\Desktop\\" + textBox2.Text);
            task.Quit();
        }
        void fileopener()
        {
            button1.Enabled = false;
            textBox1.Text = "";
            dialog = new OpenFileDialog();
            dialog.FileName = "";
            dialog.InitialDirectory = "C:¥";
            dialog.Filter = "画像ファイル(*.jpg;*.jpeg;*.png)|*.jpg;*.jpeg;*.png";
            dialog.Title = "画像を選択";
            dialog.RestoreDirectory = true;
            dialog.CheckFileExists = true;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = dialog.FileName;
                button1.Enabled = true;
            }
        }
        void after(object sender,FormClosedEventArgs e)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(bookafter);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(task);
        }
    }
}
