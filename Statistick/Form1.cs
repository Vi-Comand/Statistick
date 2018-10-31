using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MetroFramework.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using DevExpress.XtraCharts;

namespace Statistick
{
    public partial class Form1 : MetroForm
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void but_save_db_Click(object sender, EventArgs e)
        {
            //jhk
        }

        private void but_load_excel_Click(object sender, EventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Microsoft Excel (*.xls*)|*.xls*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {

                Excel.Application excelApp = new Excel.Application();
                excelApp.Workbooks.Open(openFileDialog.FileName);
                int row = 4;
                int LishnColumn = 0;
                List<string> maping = new List<string>();

                maping.Add("faim");

                if (check_uud1.Checked == true)
                {
                    maping.Add("uud1");
                    maping.Add("uud2");
                    maping.Add("uud3");
                }
                else
                {
                    maping.Add("uud1");
                    LishnColumn += 2;
                }

                if (check_uud2.Checked == true)
                {
                    maping.Add("uud4");
                    maping.Add("uud5");
                    maping.Add("uud6");
                }
                else
                {
                    maping.Add("uud4");
                    LishnColumn += 2;
                }


                if (check_uud3.Checked == true)
                {
                    maping.Add("uud7");
                    maping.Add("uud8");
                    maping.Add("uud9");
                }
                else
                {
                    maping.Add("uud7");
                    LishnColumn += 2;
                }


                maping.Add("uud10");
                maping.Add("uud11");
                Excel.Worksheet currentSheet = (Excel.Worksheet)excelApp.Workbooks[1].Worksheets[1];
                int MyRows = 0;
                while (currentSheet.get_Range("B" + row).Value2 != null)
                {
                    Grid_Load_UUD.Rows.Add();
                    int MyCells = 0;

                    for (char column = 'B'; column < 'N' - LishnColumn; column++)
                    {


                        Excel.Range cell = currentSheet.get_Range(column.ToString() + row.ToString());

                        Grid_Load_UUD.Rows[MyRows].Cells[maping[MyCells]].Value = cell != null ? cell.Value2.ToString() : "";

                        MyCells++;


                    }
                    MyRows++;
                    row++;

                }

            }
        }

        private void check_uud1_CheckedChanged(object sender, EventArgs e)
        {
            if (check_uud1.Checked)
            {
                Three_UUD_Colums_Add(1);
            }
            else
            {
                Three_UUD_Colums_Del(1);
            }

        }

        private void check_uud2_CheckedChanged(object sender, EventArgs e)
        {
            if (check_uud2.Checked)
            {
                Three_UUD_Colums_Add(2);
            }
            else
            {
                Three_UUD_Colums_Del(2);
            }
        }

        private void check_uud3_CheckedChanged(object sender, EventArgs e)
        {
            if (check_uud3.Checked)
            {
                Three_UUD_Colums_Add(3);
            }
            else
            {
                Three_UUD_Colums_Del(3);
            }
        }

        private void Three_UUD_Colums_Add(int _nomerUUD)
        {
            switch (_nomerUUD)
            {
                case 1:
                    Grid_Load_UUD.Columns.Add(new DataGridViewTextBoxColumn() { Name = "uud2", HeaderText = "УУД1-2", Width = 100, DisplayIndex = 2 });
                    Grid_Load_UUD.Columns.Add(new DataGridViewTextBoxColumn() { Name = "uud3", HeaderText = "УУД1-3", Width = 100, DisplayIndex = 3 });
                    break;
                case 2:
                    Grid_Load_UUD.Columns.Add(new DataGridViewTextBoxColumn() { Name = "uud5", HeaderText = "УУД2-2", Width = 100, DisplayIndex = (check_uud1.Checked) ? 5 : 3 });
                    Grid_Load_UUD.Columns.Add(new DataGridViewTextBoxColumn() { Name = "uud6", HeaderText = "УУД2-3", Width = 100, DisplayIndex = (check_uud1.Checked) ? 6 : 4 });
                    break;
                case 3:
                    Grid_Load_UUD.Columns.Add(new DataGridViewTextBoxColumn() { Name = "uud8", HeaderText = "УУД3-2", Width = 100, DisplayIndex = Grid_Load_UUD.Columns["uud7"].DisplayIndex + 1 });
                    Grid_Load_UUD.Columns.Add(new DataGridViewTextBoxColumn() { Name = "uud9", HeaderText = "УУД3-3", Width = 100, DisplayIndex = Grid_Load_UUD.Columns["uud7"].DisplayIndex + 2 });
                    break;
            }
        }
        private void Three_UUD_Colums_Del(int _nomerUUD)
        {
            switch (_nomerUUD)
            {
                case 1:
                    Grid_Load_UUD.Columns.Remove(Grid_Load_UUD.Columns["uud2"]);
                    Grid_Load_UUD.Columns.Remove(Grid_Load_UUD.Columns["uud3"]);
                    break;
                case 2:
                    Grid_Load_UUD.Columns.Remove(Grid_Load_UUD.Columns["uud5"]);
                    Grid_Load_UUD.Columns.Remove(Grid_Load_UUD.Columns["uud6"]);
                    break;
                case 3:
                    Grid_Load_UUD.Columns.Remove(Grid_Load_UUD.Columns["uud8"]);
                    Grid_Load_UUD.Columns.Remove(Grid_Load_UUD.Columns["uud9"]);
                    break;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "in_statDataSet.uud". При необходимости она может быть перемещена или удалена.
            this.uudTableAdapter.Fill(this.in_statDataSet.uud);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "in_statDataSet.uud". При необходимости она может быть перемещена или удалена.
            //   this.uudTableAdapter.Fill(this.in_statDataSet.uud);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "in_statDataSet.kontrolnie". При необходимости она может быть перемещена или удалена.
            this.kontrolnieTableAdapter.Fill(this.in_statDataSet.kontrolnie);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "in_statDataSet.user". При необходимости она может быть перемещена или удалена.
            this.userTableAdapter.Fill(this.in_statDataSet.user);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "in_statDataSet.klass". При необходимости она может быть перемещена или удалена.
            this.klassTableAdapter.Fill(this.in_statDataSet.klass);

            Stat1();
        }

        private void Stat1()
        {
           // ChartControl areaChart = new ChartControl();

            // Create two area series.
            Series series1 = new Series("Series 1", ViewType.Area);
            // Series series2 = new Series("Series 2", ViewType.Area);
            int f = 1;
            int uud1 = 0;
            int uud2 = 0;
            int uud3 = 0;
            string fi = "";
            foreach (DataRow row in in_statDataSet.uud.Rows)
            {
                if (f != in_statDataSet.uud.Rows.Count)
                {
                    int uud = 0;
                    uud = Convert.ToInt32(row["uud11"].ToString()) + Convert.ToInt32(row["uud12"].ToString()) + Convert.ToInt32(row["uud13"].ToString()) + Convert.ToInt32(row["uud21"].ToString()) + Convert.ToInt32(row["uud22"].ToString()) + Convert.ToInt32(row["uud23"].ToString()) + Convert.ToInt32(row["uud31"].ToString()) + Convert.ToInt32(row["uud32"].ToString()) + Convert.ToInt32(row["uud33"].ToString()) + Convert.ToInt32(row["uud4"].ToString()) + Convert.ToInt32(row["uud5"].ToString());
                    fi = row["id_user"].ToString();
                    series1.Points.Add(new SeriesPoint(fi, uud));
                    f++;
                }
            }


            // Add points to them.
         /*   series1.Points.Add(new SeriesPoint(1, 15));
            series1.Points.Add(new SeriesPoint(2, 18));
            series1.Points.Add(new SeriesPoint(3, 25));
            series1.Points.Add(new SeriesPoint(4, 33));

            series2.Points.Add(new SeriesPoint(1, 10));
            series2.Points.Add(new SeriesPoint(2, 12));
            series2.Points.Add(new SeriesPoint(3, 14));
            series2.Points.Add(new SeriesPoint(4, 17));*/

            // Add both series to the chart.
            chartControl1.Series.AddRange(new Series[] { series1 });

            // Set the numerical argument scale types for the series,
            // as it is qualitative, by default.
            series1.ArgumentScaleType = ScaleType.Numerical;
           // series2.ArgumentScaleType = ScaleType.Numerical;

            // Access the view-type-specific options of the series.
            ((AreaSeriesView)series1.View).Transparency = 80;

            // Access the type-specific options of the diagram.
            ((XYDiagram)chartControl1.Diagram).EnableAxisXZooming = true;

            // Hide the legend (optional).
#pragma warning disable CS0618 // Тип или член устарел
            chartControl1.Legend.Visible = false;
#pragma warning restore CS0618 // Тип или член устарел

            // Add the chart to the form.
           // chartControl1.Dock = DockStyle.Fill;
           // this.Controls.Add(chartControl1);
        }

    }
}
