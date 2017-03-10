using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Windows.Forms.DataVisualization.Charting;

namespace TableParagraph
{
    public partial class UserControl1 : UserControl
    {
        public DataSet ds = new DataSet();
        public DataTable dt = null;
        public Form1 f1;
        int comboBOX = 3; 

        public UserControl1(Form1 f1)
        {
            this.f1 = f1;
            InitializeComponent();            
        }
        private void UserControl1_load(object sender, EventArgs e)
        {
            this.toolStripComboBox1.Items.Add("3");
            this.toolStripComboBox1.Items.Add("4");
            this.toolStripComboBox1.Items.Add("5");
            toolStripComboBox1.Text = Convert.ToString(3);
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                
                int curentIndex = e.RowIndex;
                // if (curentIndex == dataGridView1.RowCount - 2) curentIndex++;       
                dt = new DataTable();
                int count = dataGridView1.ColumnCount;//获取dataGridView1的列数量
                int rowCount = dataGridView1.RowCount;
                for (int i = 0; i < count; i++)
                {
                    dt.Columns.Add("A" + i.ToString(), typeof(int)); //数据类型
                }
                comboBOX = Convert.ToInt16(toolStripComboBox1.Text);
                DataRow[] dr = new DataRow[comboBOX];
                for (int k = 0; k < comboBOX; k++)
                    dr[k] = dt.NewRow();

                while (curentIndex + comboBOX > rowCount - 1) curentIndex--;//默认表格显示是在双击单元格后，在该单元格所在行及后面两行数据会显示为折线图。防止在双击最后两行出现越界异常
                for (int i = 0; i < comboBOX; i++)
                {

                    dataGridView1.Rows[curentIndex + i].Selected = true;
                    for (int c1 = 0; c1 < count; c1++)
                    {
                        int x = Convert.ToInt32(dataGridView1.Rows[curentIndex + i].Cells[c1].Value);
                        dr[i][c1] = x;
                    }
                    dt.Rows.Add(dr[i]);
                }
                ChartChange();
                chart1.Invalidate();
            }
        }
        private void ChartChange()
        {
            chart1.DataSource = dt;
            chart1.ChartAreas.Clear(); //图表区
            chart1.Titles.Clear(); //图表标题
            chart1.Series.Clear(); //图表序列
            chart1.Legends.Clear(); //图表图例
            ChartArea care = new ChartArea();
            Legend leg = new Legend();            
            Series[] ss = new Series[comboBOX]; // 默认读入3行数据，即生成3组折线图          
            for (int i = 0; i < comboBOX; i++)
            {
                ss[i] = new Series(dt.Rows[i][0].ToString());
                int ColCount = dt.Columns.Count;//获取列数量
                for (int j = 0; j < ColCount; j++)
                    ss[i].Points.AddY(Convert.ToDouble(dt.Rows[i][j]));

                chart1.Series.Add(ss[i]);
                chart1.Series[i].ChartType = SeriesChartType.Line;
                ss[0].BorderWidth = 5;
            }
            // 显示
            chart1.ChartAreas.Add(care);
            chart1.Legends.Add(leg);           
        }
    }
}
