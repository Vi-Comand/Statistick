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
           
        }


        private void DiagPoUch_Points(int id_kontr, int id_klass, int god)
        {
            chartControl1.Series.Clear();
            chartControl1.Titles.Clear();
            int uud = 0;
            string fi = "";
            Series series1 = new Series("Ученики", ViewType.Bar);
            foreach (DataRow row in in_statDataSet.uud.Rows)
            {
                if (Convert.ToInt32(row["id_kontr"]) == id_kontr && Convert.ToInt32(row["id_klass"]) == id_klass && Convert.ToInt32(row["god"]) == god)
                {
                    uud = Convert.ToInt16(row["uud11"]) + Convert.ToInt16(row["uud12"]) + Convert.ToInt16(row["uud13"]) + Convert.ToInt16(row["uud21"]) + Convert.ToInt16(row["uud22"]) + Convert.ToInt16(row["uud23"]) + Convert.ToInt16(row["uud31"]) + Convert.ToInt16(row["uud32"]) + Convert.ToInt16(row["uud33"]) + Convert.ToInt16(row["uud4"]) + Convert.ToInt16(row["uud5"]);
                    foreach (DataRow roww in in_statDataSet.user.Rows)
                    {
                        if (Convert.ToString(row["id_user"]) == Convert.ToString(roww["id"]))
                        {
                            fi = roww["fi"].ToString();
                        }
                    }
                    series1.Points.Add(new SeriesPoint(fi, uud));
                }
            }
            // Add the series to the chart.
            chartControl1.Series.Add(series1);
            // Hide the legend (if necessary).
            chartControl1.Legend.Visibility = DevExpress.Utils.DefaultBoolean.True;
            // Rotate the diagram (if necessary).
            ((XYDiagram)chartControl1.Diagram).Rotated = false;
            // Add a title to the chart (if necessary).
            ChartTitle chartTitle1 = new ChartTitle();
            chartTitle1.Text = "Диаграмма по учащимся";
            chartControl1.Titles.Add(chartTitle1);
        }

        private void DiagPoPoziciyam_Points(int id_kontr, int id_klass, int god)
        {
            chartControl1.Series.Clear();
            chartControl1.Titles.Clear();
            int uud1 = 0 , uud2 = 0, uud3 = 0, uud4 = 0, uud5 = 0;
            
            Series series1 = new Series("УУД", ViewType.Bar);
            foreach (DataRow row in in_statDataSet.uud.Rows)
            {
                if (Convert.ToInt32(row["id_kontr"]) == id_kontr && Convert.ToInt32(row["id_klass"]) == id_klass && Convert.ToInt32(row["god"]) == god)
                {
                    uud1 += Convert.ToInt16(row["uud11"]) + Convert.ToInt16(row["uud12"]) + Convert.ToInt16(row["uud13"]);
                    uud2 += Convert.ToInt16(row["uud21"]) + Convert.ToInt16(row["uud22"]) + Convert.ToInt16(row["uud23"]);
                    uud3 += Convert.ToInt16(row["uud31"]) + Convert.ToInt16(row["uud32"]) + Convert.ToInt16(row["uud33"]);
                    uud4 += Convert.ToInt16(row["uud4"]);
                    uud5 += Convert.ToInt16(row["uud5"]);
                    
                }
            }
            for (int i=1; i<6; i++)
            {
                switch (i)
                {
                    case 1:
                        series1.Points.Add(new SeriesPoint("Может выбрать наиболее эффективные способы решения задач в зависимости от конкретных условий (УУД 1)", uud1));
                        break;
                    case 2:
                        series1.Points.Add(new SeriesPoint("Может строить логическую цепь рассуждений (выявлять причинно-следственные связи, выявлять закономерности) (УУД2)", uud2));
                        break;
                    case 3:
                        series1.Points.Add(new SeriesPoint("Может структурировать найденную информацию в нужной форме(УУД3)", uud3));
                        break;
                    case 4:
                        series1.Points.Add(new SeriesPoint("Владеет умением классификации(УУД4)", uud4));
                        break;
                    case 5:
                        series1.Points.Add(new SeriesPoint("Умеет осмысленно читать, извлекая нужную информацию(УУД5)", uud5));
                        break;
                }
            }
            // Add the series to the chart.
            chartControl1.Series.Add(series1);
            // Hide the legend (if necessary).
            chartControl1.Legend.Visibility = DevExpress.Utils.DefaultBoolean.True;
            // Rotate the diagram (if necessary).
            ((XYDiagram)chartControl1.Diagram).Rotated = false;
            // Add a title to the chart (if necessary).
            ChartTitle chartTitle1 = new ChartTitle();
            chartTitle1.Text = "Общая диограмма по позициям";
            chartControl1.Titles.Add(chartTitle1);
        }


        private void metroTile1_Click(object sender, EventArgs e)
        {
            if (ComboBox_God_Stat.SelectedIndex != -1 && ComboBox_Grafik_Stat.SelectedIndex != -1)
            {
                switch (ComboBox_Grafik_Stat.SelectedIndex)
                {
                    case 0:
                        DiagPoUch_Points(Convert.ToInt32(ComboBox_Kontrol_Stat.SelectedValue), Convert.ToInt32(ComboBox_Klass_Stat.SelectedValue), Convert.ToInt32(ComboBox_God_Stat.SelectedItem));
                        break;
                    case 1:
                        DiagPoPoziciyam_Points(Convert.ToInt32(ComboBox_Kontrol_Stat.SelectedValue), Convert.ToInt32(ComboBox_Klass_Stat.SelectedValue), Convert.ToInt32(ComboBox_God_Stat.SelectedItem));
                        break;
                    case 2:

                        break;
                    case 3:

                        break;
                    default:
                            MessageBox.Show("Неизвестный график");
                        break;

                }
            }
        }
    }
}
