using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Xml;
using System.Data.OleDb;
using System.Xml.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using DevComponents.DotNetBar.Controls;
using System.Threading;
using System.Data.Common;
using System.Windows.Forms.DataVisualization.Charting;
using System.Drawing.Imaging;

namespace TableParagraph
{
    public partial class Form1 : Form
    {
        public static DataSet ds;
        int comboBoxint;
       
        public Form1()
        {
            InitializeComponent();
        }
        void openUserControl1()
        {
            UserControl1 uc1 = new UserControl1(this);
        }

        private void openFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Title = "请选择一个Excel文件";
            openDialog.Filter = "Excel文件|*.xls;*.xlsx";
            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                string excelFile = openDialog.FileName;
                if (String.IsNullOrEmpty(excelFile) || !File.Exists(excelFile))
                {
                    MessageBox.Show(this, "输入的Excel文件为空或不存在！", "错误",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                //tabControl1.Controls.Clear();
                ToDataTable(excelFile, out ds);                
                displayUserControl(ds);
               
            }
           
        }
        /// <summary>  
        /// 读取Excel文件到DataSet中  
        /// </summary>  
        /// <param name="filePath">文件路径</param>  
        /// <returns></returns>  
        public static DataSet ToDataTable(string filePath, out DataSet ds)
        {
            ds = new DataSet();
            string connStr = "";
            string fileType = Path.GetExtension(filePath);
            //if (string.IsNullOrEmpty(fileType)) return ;

            if (fileType == ".xls")
                connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + filePath + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
            else
                connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + filePath + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            string sql_F = "Select * FROM [{0}]";

            OleDbConnection conn = null;
            OleDbDataAdapter da = null;
            DataTable dtSheetName = null;
            try
            {
                // 初始化连接，并打开  
                conn = new OleDbConnection(connStr);
                conn.Open();

                // 获取数据源的表定义元数据                         
                string SheetName = "";
                dtSheetName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                // 初始化适配器  
                da = new OleDbDataAdapter();
                for (int i = 0; i < dtSheetName.Rows.Count; i++)
                {
                    SheetName = (string)dtSheetName.Rows[i]["TABLE_NAME"];

                    if (SheetName.Contains("$") && !SheetName.Replace("'", "").EndsWith("$"))
                    {
                        continue;
                    }

                    da.SelectCommand = new OleDbCommand(String.Format(sql_F, SheetName), conn);
                    DataSet dsItem = new DataSet();
                    da.Fill(dsItem, SheetName);

                    ds.Tables.Add(dsItem.Tables[0].Copy());
                }
            }
            catch (Exception ex)
            {
            }
            finally
            {
                // 关闭连接  
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    da.Dispose();
                    conn.Dispose();
                }
            }
            return ds;
        }
        private void displayUserControl(DataSet ds)
        {
            DataTable dt = new DataTable();
            String tableFile = dt.TableName;
            // handle the remained tables
             //ds.Tables.Count;
            for (int i = 0; i < ds.Tables.Count; i++)
            {                
                dt = ds.Tables[i];
                tableFile = dt.TableName;
                TabPage tp = new TabPage();
                tabControl1.Controls.Add(tp);
                tp.Name = tableFile;
                tp.Text = tableFile;
                var uc = createUC();
                uc.dataGridView1.DataSource = dt;
                createChart(dt, uc);                
                tp.Controls.Add(uc);                

            }
        }
        /// <summary>
        /// 设置DataGridView属性并实例化
        /// </summary>
        /// <param name="tableName"></param>
        /// <returns></returns>
        private DataGridView createNewDGV(String tableName)
        { 
            
            DataGridView dgv = new DataGridView();
            dgv.AutoGenerateColumns = true;
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.AllowUserToOrderColumns = false;

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            for (int i = 5; i < dgv.Columns.Count; i++)
                dgv.Columns[i].MinimumWidth = 60;

            dgv.Location = new Point(0, 0);
            dgv.Dock = DockStyle.Fill;
            return dgv;
        }
        /// <summary>
        /// 自定义组合控件设置属性
        /// </summary>
        /// <returns></returns>
        private UserControl1 createUC()
        {
            UserControl1 uc = new UserControl1(this);
            uc.Dock = System.Windows.Forms.DockStyle.Fill;
            uc.Enabled = true;
            uc.Location = new System.Drawing.Point(3, 3);            
            uc.Size = new System.Drawing.Size(734, 504);
            uc.TabIndex = 0;           
            return uc;
        }        
       private void createChart(DataTable dt,UserControl1 uc)
        {
            uc.chart1.DataSource = dt;
            uc.chart1.ChartAreas.Clear(); //图表区
            uc.chart1.Titles.Clear(); //图表标题
            uc.chart1.Series.Clear(); //图表序列
            uc.chart1.Legends.Clear(); //图表图例
            ChartArea care = new ChartArea();
            Legend leg = new Legend();
            int j = dt.Rows.Count;
            Series[] ss = new Series[j]; // 默认读入所有行的数据           
            for (int i = 0; i < j; i++)
            {
                ss[i] = new Series(dt.Rows[i][0].ToString());
                int ColCount = dt.Columns.Count;
               for(int k=0;k<ColCount;k++)
                ss[i].Points.AddY(Convert.ToDouble(dt.Rows[i][k]));

                uc.chart1.Series.Add(ss[i]);
                uc.chart1.Series[i].ChartType = SeriesChartType.Line;
            }                           
            // 显示
            uc.chart1.ChartAreas.Add(care);
            uc.chart1.Legends.Add(leg);          
        }
    }
}
