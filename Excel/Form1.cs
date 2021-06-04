using ExcelDataReader;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private string s;
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件夹";
            dialog.Filter = "所有文件(*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                int len = dialog.FileNames.Length;
                ArrayList workload = new ArrayList();
                int[] workloads = { 0 };
                Dictionary<string, int> wl = new Dictionary<string, int>();

                //遍历每个文件
                for (int i = 0; i < len; i++)
                {
                    
                    string file = dialog.FileNames[i];

                    using (FileStream stream = File.Open(file, FileMode.Open, FileAccess.Read))
                    {
                        IExcelDataReader reader;
                        if (file[file.Length - 1] == 'x')
                        {
                            reader = ExcelReaderFactory.CreateOpenXmlReader(stream);//xlsx

                        }
                        else
                        {
                            if (file[file.Length - 1] == 's')
                            {
                                reader = ExcelReaderFactory.CreateBinaryReader(stream);//xls
                            }
                            else
                            {
                                MessageBox.Show("文件类型不支持！");
                                return;
                            }
                        }

                        var result = reader.AsDataSet();

                        if(i == 0)
                        {
                            for(int j = 0; j < reader.RowCount; j++)
                            {
                                workload.Add(0);
                            }
                            workloads = (int[])workload.ToArray(typeof(int));

                            for (int j = 0; j < reader.RowCount - 1; j++)
                            {

                                wl.Add(result.Tables[0].Rows[j + 1][2].ToString(), int.Parse(result.Tables[0].Rows[j + 1][18].ToString()));
                                

                            }
                        }
                        else
                        {
                            for(int j = 0; j < reader.RowCount - 1; j++)
                            {
                                if( wl.ContainsKey(result.Tables[0].Rows[j + 1][2].ToString()))
                                {
                                    wl[result.Tables[0].Rows[j + 1][2].ToString()] += int.Parse(result.Tables[0].Rows[j + 1][18].ToString());
                                }
                                else
                                {
                                    wl.Add(result.Tables[0].Rows[j + 1][2].ToString(), int.Parse(result.Tables[0].Rows[j + 1][18].ToString()));
                                }
                            }
                        }

                        
                        



                    }


                }


               

                //建立Excel对象 
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                //excel.Application.Workbooks.Add(true);
                Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];
                excel.Visible = true;
                //Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)excel.Worksheets[1];
                Microsoft.Office.Interop.Excel.Range range;

                int row = 1;
                worksheet.Cells[row, 1] = "name";
                worksheet.Cells[row, 2] = "workload";// use non-zero index
                row++;

                foreach (KeyValuePair<string, int> kvp in wl)
                {
                    worksheet.Cells[row, 1] = kvp.Key.ToString();
                    worksheet.Cells[row, 2] = kvp.Value.ToString();
                    row++;
                    
                }
         

                

            }



 
        }

        private void button2_Click(object sender, EventArgs e)//输出某位老师的工作量
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件夹";
            dialog.Filter = "所有文件(*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                int len = dialog.FileNames.Length;
                ArrayList workload = new ArrayList();
                int[] workloads = { 0 };
                Dictionary<string, int> wl = new Dictionary<string, int>();

                //建立Excel对象 
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                //excel.Application.Workbooks.Add(true);
                Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];
                excel.Visible = true;
                //Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)excel.Worksheets[1];
                Microsoft.Office.Interop.Excel.Range range;

                //遍历每个文件
                for (int i = 0; i < len; i++)
                {

                    string file = dialog.FileNames[i];

                    using (FileStream stream = File.Open(file, FileMode.Open, FileAccess.Read))
                    {
                        IExcelDataReader reader;
                        if (file[file.Length - 1] == 'x')
                        {
                            reader = ExcelReaderFactory.CreateOpenXmlReader(stream);//xlsx

                        }
                        else
                        {
                            if (file[file.Length - 1] == 's')
                            {
                                reader = ExcelReaderFactory.CreateBinaryReader(stream);//xls
                            }
                            else
                            {
                                MessageBox.Show("文件类型不支持！");
                                return;
                            }
                        }

                        var result = reader.AsDataSet();
                        reader.Read();
                        for(int j = 1; j <= 20; j++)//copy header
                        {
                            worksheet.Cells[1, j] = result.Tables[0].Rows[0][j-1].ToString();
                        }

                        for(int j = 1; j < reader.RowCount; j++)
                        {
                            if(result.Tables[0].Rows[j][2].ToString() == textBox1.Text)
                            {
                                for (int k = 1; k <= 20; k++)//copy header
                                {
                                    worksheet.Cells[i+2, k] = result.Tables[0].Rows[j][k-1].ToString();
                                }
                            }
                        }







                    }


                }











            }

        }



        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }
    }
}
