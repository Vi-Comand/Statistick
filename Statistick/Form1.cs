using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.Charts.Native;
using MetroFramework.Forms;
using Excel = Microsoft.Office.Interop.Excel;

using DevExpress.XtraCharts;
using Microsoft.Office.Interop.Excel;

using System.Text.RegularExpressions;

namespace Statistick
{
    public partial class Form1 : MetroForm
    {
        public Form1()
        {
            InitializeComponent();
           
           
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
            ComboBox_God_Load.SelectedIndex = 0;
            Update_Combobox_Kontrol_Load();

        }

        private void but_save_db_Click(object sender, EventArgs e)
        {
            bool prinmatizmenenia=true;
            foreach (DataRow row1 in in_statDataSet.uud.Rows)
            {
                if ((int) row1[2] ==  Convert.ToInt32(ComboBox_Kontrol_Load.SelectedValue))
                {
                    DialogResult dialogResult = MessageBox.Show("Такая контрольная работа уже есть в системе. Обновить данные контрольной работы?", "Some Title", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        prinmatizmenenia = true;
                        break;
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        prinmatizmenenia = false;
                        break;
                    }
                }


            }


            if (prinmatizmenenia)
            {
                int kol = Est_v_BD();
                MessageBox.Show("Кол-во новых пользователей " + kol);

                for (int i = 0; i < NoviePolz.Count; i++)
                {

                    DataRow row = in_statDataSet.user.NewRow();
                    row["fi"] = Grid_Load_UUD.Rows[NoviePolz[i]].Cells[0].Value;
                    row["id_klass"] = ComboBox_Klass_Load.SelectedValue;

                    in_statDataSet.user.Rows.Add(row);
                }

                userTableAdapter.Update(in_statDataSet);



                for (int i = 0; i < Grid_Load_UUD.Rows.Count - 1; i++)
                {
                    int id = 0;
                    foreach (DataRow row1 in in_statDataSet.user.Rows)
                    {
                        if (Grid_Load_UUD.Rows[i].Cells[0].Value.ToString() == row1[1].ToString())
                        {
                            id = (int) row1[0];
                            break;
                        }
                    }







                    DataRow row = in_statDataSet.uud.NewRow();
                    for (int j = 0; j < Grid_Load_UUD.Columns.Count; j++)
                    {
                        if (Grid_Load_UUD.Columns[j].Name == "uud1")
                        {

                            row["id_user"] = id;
                            row["id_kontr"] = ComboBox_Kontrol_Load.SelectedValue;
                            row["id_klass"] = ComboBox_Klass_Load.SelectedValue;
                            row["uud11"] = Grid_Load_UUD.Rows[i].Cells[j].Value.ToString();
                        }

                        if (Grid_Load_UUD.Columns[j].Name == "uud2")
                        {
                            row["id_user"] = id;
                            row["id_kontr"] = ComboBox_Kontrol_Load.SelectedValue;
                            row["id_klass"] = ComboBox_Klass_Load.SelectedValue;
                            row["uud12"] = Grid_Load_UUD.Rows[i].Cells[j].Value.ToString();
                        }

                        if (Grid_Load_UUD.Columns[j].Name == "uud3")
                        {
                            row["id_user"] = id;
                            row["id_kontr"] = ComboBox_Kontrol_Load.SelectedValue;
                            row["id_klass"] = ComboBox_Klass_Load.SelectedValue;
                            row["uud13"] = Grid_Load_UUD.Rows[i].Cells[j].Value.ToString();
                        }

                        if (Grid_Load_UUD.Columns[j].Name == "uud4")
                        {
                            row["id_user"] = id;
                            row["id_kontr"] = ComboBox_Kontrol_Load.SelectedValue;
                            row["id_klass"] = ComboBox_Klass_Load.SelectedValue;
                            row["uud21"] = Grid_Load_UUD.Rows[i].Cells[j].Value.ToString();
                        }

                        if (Grid_Load_UUD.Columns[j].Name == "uud5")
                        {
                            row["id_user"] = id;
                            row["id_kontr"] = ComboBox_Kontrol_Load.SelectedValue;
                            row["id_klass"] = ComboBox_Klass_Load.SelectedValue;
                            row["uud22"] = Grid_Load_UUD.Rows[i].Cells[j].Value.ToString();
                        }

                        if (Grid_Load_UUD.Columns[j].Name == "uud6")
                        {
                            row["id_user"] = id;
                            row["id_kontr"] = ComboBox_Kontrol_Load.SelectedValue;
                            row["id_klass"] = ComboBox_Klass_Load.SelectedValue;
                            row["uud23"] = Grid_Load_UUD.Rows[i].Cells[j].Value.ToString();
                        }

                        if (Grid_Load_UUD.Columns[j].Name == "uud7")
                        {
                            row["id_user"] = id;
                            row["id_kontr"] = ComboBox_Kontrol_Load.SelectedValue;
                            row["id_klass"] = ComboBox_Klass_Load.SelectedValue;
                            row["uud31"] = Grid_Load_UUD.Rows[i].Cells[j].Value.ToString();
                        }

                        if (Grid_Load_UUD.Columns[j].Name == "uud8")
                        {
                            row["id_user"] = id;
                            row["id_kontr"] = ComboBox_Kontrol_Load.SelectedValue;
                            row["id_klass"] = ComboBox_Klass_Load.SelectedValue;
                            row["uud32"] = Grid_Load_UUD.Rows[i].Cells[j].Value.ToString();
                        }

                        if (Grid_Load_UUD.Columns[j].Name == "uud9")
                        {
                            row["id_user"] = id;
                            row["id_kontr"] = ComboBox_Kontrol_Load.SelectedValue;
                            row["id_klass"] = ComboBox_Klass_Load.SelectedValue;
                            row["uud33"] = Grid_Load_UUD.Rows[i].Cells[j].Value.ToString();
                        }

                        if (Grid_Load_UUD.Columns[j].Name == "uud10")
                        {
                            row["id_user"] = id;
                            row["id_kontr"] = ComboBox_Kontrol_Load.SelectedValue;
                            row["id_klass"] = ComboBox_Klass_Load.SelectedValue;
                            row["uud4"] = Grid_Load_UUD.Rows[i].Cells[j].Value.ToString();
                        }

                        if (Grid_Load_UUD.Columns[j].Name == "uud11")
                        {
                            row["id_user"] = id;
                            row["id_kontr"] = ComboBox_Kontrol_Load.SelectedValue;
                            row["id_klass"] = ComboBox_Klass_Load.SelectedValue;
                            row["uud5"] = Grid_Load_UUD.Rows[i].Cells[j].Value.ToString();
                        }
                    }

                    in_statDataSet.uud.Rows.Add(row);

                }

                uudTableAdapter.Update(in_statDataSet);
                MessageBox.Show("Измененя внесены");
            }

            //------------Фильтр пар-----------------------------
          /*  gridControl1.Visible = true;
            парыBindingSource.Filter = "id_Утп ='" + _idUtp + "'";
            tabControl1.SelectTab(0);*/
            //------------Фильтр пар-----------------------------
        }

        private void but_load_excel_Click(object sender, EventArgs e)
        {
          
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Microsoft Excel (*.xls*)|*.xls*"
            };
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
                string klass = currentSheet.get_Range("A2").Value.ToString();
                klass = Regex.Replace(klass, "[^А-Яа-я0-9]", "");
                klass=klass.ToUpper();
              DateTime data= Convert.ToDateTime(currentSheet.get_Range("B2").Value);
                string god = data.Year.ToString();
                string kontrolnie = currentSheet.get_Range("C2").Value.ToString();
                for (int i = 0; i < ComboBox_God_Load.Items.Count; i++)
                    if (ComboBox_God_Load.Items[i].ToString() == god)
                    {
                        ComboBox_God_Load.SelectedIndex = i;
                    }

                foreach (DataRow row1 in in_statDataSet.klass.Rows)
                {
                    
                    if (row1[1].ToString() == klass)
                    {
                        ComboBox_Klass_Load.SelectedValue = (int) row1[0];

                    }
                }
                foreach (DataRow row1 in in_statDataSet.kontrolnie.Rows)
                {

                    if (row1[1].ToString() == kontrolnie && Convert.ToDateTime(row1[2])==data)
                    {
                        ComboBox_Kontrol_Load.SelectedValue = row1[0].ToString();

                    }
                }

                //kontrolnieBindingSource.Filter = "data ='" + _idUtp + "'";


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
                excelApp.Quit();
            }
       
            MessageBox.Show(Est_v_BD().ToString());

        }
        List<int> NoviePolz;
        private int Est_v_BD(int kol=0)
        {
            NoviePolz = new List<int>();

            for ( int i=0;i<Grid_Load_UUD.Rows.Count;i++)
            {
                bool est = false;
                foreach (DataRow row1 in in_statDataSet.user.Rows )
                {
                    if ((Grid_Load_UUD.Rows[i].Cells[0].Value == null ? "": Grid_Load_UUD.Rows[i].Cells[0].Value.ToString()) == row1["fi"].ToString() && (int) row1["id_klass"]== Convert.ToInt32(ComboBox_Klass_Load.SelectedValue))
                    {
                        est = true;
                       break; 
                    }

                }

                if (est==false)
               {
                    kol++;
                   NoviePolz.Add(Grid_Load_UUD.Rows[i].Index);
                   Grid_Load_UUD.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;
               }
                else
                {
                    Grid_Load_UUD.Rows[i].DefaultCellStyle.BackColor = Color.White;
                }

            }

            Grid_Load_UUD.ClearSelection();
            return kol;
//<<<<<<< HEAD
        }

        private void Update_Combobox_Kontrol_Load()
        {
            var items = new List<KeyValuePair<string, string>>();
          
            DateTime nachalo=new DateTime(Convert.ToInt32(ComboBox_God_Load.Text),1,1);
            DateTime konec = new DateTime(Convert.ToInt32(ComboBox_God_Load.Text)+1, 1, 1);
            //ComboBox_Kontrol_Load.Items.Clear();
          
            foreach (DataRow row in in_statDataSet.kontrolnie.Rows)
            {
                if (nachalo < Convert.ToDateTime(row[2]) && Convert.ToDateTime(row[2]) < konec)
                {
                    var znach=new KeyValuePair<string,string>(row[0].ToString(),(Convert.ToDateTime( row[2]).ToShortDateString()).ToString()+" "+row[1].ToString());
                    items.Add(znach);
                }

            }
            ComboBox_Kontrol_Load.DataSource = items;
            ComboBox_Kontrol_Load.ValueMember = "Key";
            ComboBox_Kontrol_Load.DisplayMember = "Value";
          //  ComboBox_Kontrol_Load.SelectedIndex = 0;


        }

//=======
       

//>>>>>>> 6e168095ff9b9a19d30e617a0b07114c2a31c458

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

//<<<<<<< HEAD
      


        private void DiagPoUch_Points(int id_kontr, int id_klass, int god)
        {
            StatchartControl1.Series.Clear();
            StatchartControl1.Titles.Clear();
            int uud = 0;
            string fi = "";
            DevExpress.XtraCharts.Series series1 = new DevExpress.XtraCharts.Series("Ученики", DevExpress.XtraCharts.ViewType.Bar);
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
            StatchartControl1.Series.Add(series1);
            // Hide the legend (if necessary).
            StatchartControl1.Legend.Visibility = DevExpress.Utils.DefaultBoolean.True;
            // Rotate the diagram (if necessary).
            ((XYDiagram)StatchartControl1.Diagram).Rotated = false;
            // Add a title to the chart (if necessary).
            DevExpress.XtraCharts.ChartTitle chartTitle1 = new DevExpress.XtraCharts.ChartTitle();
            chartTitle1.Text = "Диаграмма по учащимся";
            StatchartControl1.Titles.Add(chartTitle1);
        }

        private void DiagPoPoziciyam_Points(int id_kontr, int id_klass, int god)
        {
            StatchartControl1.Series.Clear();
            StatchartControl1.Titles.Clear();
            int uud1 = 0, uud2 = 0, uud3 = 0, uud4 = 0, uud5 = 0;

            DevExpress.XtraCharts.Series series1 = new DevExpress.XtraCharts.Series("УУД", DevExpress.XtraCharts.ViewType.Bar3D);
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
            for (int i = 1; i < 6; i++)
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
            StatchartControl1.Series.Add(series1);
            //     ((BarSeriesLabel)series1.Label).Visible = true;
            ((BarSeriesLabel)series1.Label).ResolveOverlappingMode =
            ResolveOverlappingMode.Default;

            // Access the series options.
            series1.PointOptions.PointView = PointView.ArgumentAndValues;


            // Customize the view-type-specific properties of the series.
            Bar3DSeriesView myView = (Bar3DSeriesView)series1.View;
            myView.BarDepthAuto = false;
            myView.BarDepth = 1;
            myView.BarWidth = 1;
            myView.Transparency = 80;

            // Access the diagram's options.
            ((XYDiagram3D)StatchartControl1.Diagram).ZoomPercent = 110;

            // Add a title to the chart and hide the legend.
            DevExpress.XtraCharts.ChartTitle chartTitle1 = new DevExpress.XtraCharts.ChartTitle();

            chartTitle1.Text = "Общая диограмма по позициям";
            StatchartControl1.Titles.Add(chartTitle1);
            //   chartControl1.Legend.Visible = false;
        }


        private void metroTile1_Click(object sender, EventArgs e)
        {
            if (ComboBox_God_Stat.SelectedIndex != -1 && StatComboBox_Grafik_Stat.SelectedIndex != -1)
            {
                switch (StatComboBox_Grafik_Stat.SelectedIndex)
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





        //--------------------методы диограмм---------------------------------------------------------------------------------------------------------
        private Excel.Application excelapp;
        private Excel.Workbooks excelappworkbooks;
        private Excel.Workbook excelappworkbook;
        private Excel.Sheets excelsheets;
        private Excel.Worksheet excelworksheet;
        private Excel.Range excelcells;
        private Excel.Window excelWindow;
        string uud = "";
        int i_rows = 2;
        string _control = "";
        string _klass = "";
        string _god = "";
        string _control2 = "";
        string _klass2 = "";
        string _god2 = "";

        private void Excel_Diag_tab1()
        {
            excelapp = new Excel.Application();
            excelapp.Visible = true;
            excelappworkbooks = excelapp.Workbooks;
            String templatePath = System.Windows.Forms.Application.StartupPath;
            excelappworkbook = excelapp.Workbooks.Open(templatePath + @"\Шаблоны\1.xlsx", Type.Missing, Type.Missing, Type.Missing, "WWWWW", "WWWWW", Type.Missing, 
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excelsheets = excelappworkbook.Worksheets;

        
            Diad_tabl_1();
        }

        private void Excel_Diag_tab2()
        {
            excelapp = new Excel.Application();
            excelapp.Visible = true;
            excelappworkbooks = excelapp.Workbooks;
            String templatePath = System.Windows.Forms.Application.StartupPath;
            excelappworkbook = excelapp.Workbooks.Open(templatePath + @"\Шаблоны\2.xlsx", Type.Missing, Type.Missing, Type.Missing, "WWWWW", "WWWWW", Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excelsheets = excelappworkbook.Worksheets;


            Diad_tabl_2();
        }

        private void Excel_Diag_tab3()
        {
            excelapp = new Excel.Application();
            excelapp.Visible = true;
            excelappworkbooks = excelapp.Workbooks;
            String templatePath = System.Windows.Forms.Application.StartupPath;
            excelappworkbook = excelapp.Workbooks.Open(templatePath + @"\Шаблоны\3.xlsx", Type.Missing, Type.Missing, Type.Missing, "WWWWW", "WWWWW", Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excelsheets = excelappworkbook.Worksheets;


            Diad_tabl_3();
        }

        private void Excel_Diag_tab4()
        {
            excelapp = new Excel.Application();
            excelapp.Visible = true;
            excelappworkbooks = excelapp.Workbooks;
            String templatePath = System.Windows.Forms.Application.StartupPath;
            excelappworkbook = excelapp.Workbooks.Open(templatePath + @"\Шаблоны\4.xlsx", Type.Missing, Type.Missing, Type.Missing, "WWWWW", "WWWWW", Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excelsheets = excelappworkbook.Worksheets;


            Diad_tabl_4();
        }

        private void Add_Row1()
        {
            //----------------------------------------------------------заполнение строк-------------------------------------------------------------------
            i_rows = 2;
            foreach (DataRow row in in_statDataSet.uud.Rows)
            {

                if (Convert.ToInt32(row["id_kontr"]) == Convert.ToInt32(_control) && Convert.ToInt32(row["id_klass"]) == Convert.ToInt32(_klass) && Convert.ToInt32(row["god"]) == Convert.ToInt32(_god))
                {
                    Add_Rows(row);
                }

            }
        }

        private void Add_Row2()
        {
            //----------------------------------------------------------заполнение строк-------------------------------------------------------------------
            i_rows = 2;
            foreach (DataRow row in in_statDataSet.uud.Rows)
            {

                if (Convert.ToInt32(row["id_kontr"]) == Convert.ToInt32(_control2) && Convert.ToInt32(row["id_klass"]) == Convert.ToInt32(_klass2) && Convert.ToInt32(row["god"]) == Convert.ToInt32(_god2))
                {
                    Add_Rows(row);
                }

            }
        }

        private void Add_Row_3_v_1()
        {
            //----------------------------------------------------------заполнение строк-------------------------------------------------------------------
            i_rows = 2;
            foreach (DataRow row in in_statDataSet.uud.Rows)
            {

                if (Convert.ToInt32(row["id_kontr"]) == Convert.ToInt32(_control) && Convert.ToInt32(row["id_klass"]) == Convert.ToInt32(_klass) && Convert.ToInt32(row["god"]) == Convert.ToInt32(_god))
                {
                    Add_Rows_3_v_1(row);
                }

            }
        }

        private void Add_Row_3_v_1_2()
        {
            //----------------------------------------------------------заполнение строк-------------------------------------------------------------------
            i_rows = 2;
            foreach (DataRow row in in_statDataSet.uud.Rows)
            {

                if (Convert.ToInt32(row["id_kontr"]) == Convert.ToInt32(_control2) && Convert.ToInt32(row["id_klass"]) == Convert.ToInt32(_klass2) && Convert.ToInt32(row["god"]) == Convert.ToInt32(_god2))
                {
                    Add_Rows_3_v_1_2(row);
                }

            }
        }

        private void Add_Rows(DataRow row)
        {
            
            foreach (DataRow roww in in_statDataSet.user.Rows)
            {
                if (Convert.ToString(row["id_user"]) == Convert.ToString(roww["id"]))
                {
                    excelworksheet.Cells[i_rows, 2] = roww["fi"].ToString();
                }
            }

            uud = "";
            if (row["uud12"].ToString() == "" || row["uud13"].ToString() == "")
                uud += "1";
            if (row["uud22"].ToString() == "" || row["uud23"].ToString() == "")
                uud += "2";
            if (row["uud32"].ToString() == "" || row["uud33"].ToString() == "")
                uud += "3";
            switch (uud)
            {
                case "":
                    excelworksheet.Cells[i_rows, 3] = row["uud11"];
                    excelworksheet.Cells[i_rows, 4] = row["uud12"];
                    excelworksheet.Cells[i_rows, 5] = row["uud13"];
                    excelworksheet.Cells[i_rows, 6] = row["uud21"];
                    excelworksheet.Cells[i_rows, 7] = row["uud22"];
                    excelworksheet.Cells[i_rows, 8] = row["uud23"];
                    excelworksheet.Cells[i_rows, 9] = row["uud31"];
                    excelworksheet.Cells[i_rows, 10] = row["uud32"];
                    excelworksheet.Cells[i_rows, 11] = row["uud33"];
                    excelworksheet.Cells[i_rows, 12] = row["uud4"];
                    excelworksheet.Cells[i_rows, 13] = row["uud5"];
                    i_rows++;
                    break;
                case "1":
                    excelworksheet.Cells[i_rows, 3] = row["uud11"];
                    excelworksheet.Cells[i_rows, 6] = row["uud21"];
                    excelworksheet.Cells[i_rows, 7] = row["uud22"];
                    excelworksheet.Cells[i_rows, 8] = row["uud23"];
                    excelworksheet.Cells[i_rows, 9] = row["uud31"];
                    excelworksheet.Cells[i_rows, 10] = row["uud32"];
                    excelworksheet.Cells[i_rows, 11] = row["uud33"];
                    excelworksheet.Cells[i_rows, 12] = row["uud4"];
                    excelworksheet.Cells[i_rows, 13] = row["uud5"];
                    i_rows++;
                    break;
                case "2":
                    excelworksheet.Cells[i_rows, 3] = row["uud11"];
                    excelworksheet.Cells[i_rows, 4] = row["uud12"];
                    excelworksheet.Cells[i_rows, 5] = row["uud13"];
                    excelworksheet.Cells[i_rows, 6] = row["uud21"];
                    excelworksheet.Cells[i_rows, 9] = row["uud31"];
                    excelworksheet.Cells[i_rows, 10] = row["uud32"];
                    excelworksheet.Cells[i_rows, 11] = row["uud33"];
                    excelworksheet.Cells[i_rows, 12] = row["uud4"];
                    excelworksheet.Cells[i_rows, 13] = row["uud5"];
                    i_rows++;
                    break;
                case "3":
                    excelworksheet.Cells[i_rows, 3] = row["uud11"];
                    excelworksheet.Cells[i_rows, 4] = row["uud12"];
                    excelworksheet.Cells[i_rows, 5] = row["uud13"];
                    excelworksheet.Cells[i_rows, 6] = row["uud21"];
                    excelworksheet.Cells[i_rows, 7] = row["uud22"];
                    excelworksheet.Cells[i_rows, 8] = row["uud23"];
                    excelworksheet.Cells[i_rows, 9] = row["uud31"];
                    excelworksheet.Cells[i_rows, 12] = row["uud4"];
                    excelworksheet.Cells[i_rows, 13] = row["uud5"];
                    i_rows++;
                    break;
                case "12":
                    excelworksheet.Cells[i_rows, 3] = row["uud11"];
                    excelworksheet.Cells[i_rows, 6] = row["uud21"];
                    excelworksheet.Cells[i_rows, 9] = row["uud31"];
                    excelworksheet.Cells[i_rows, 10] = row["uud32"];
                    excelworksheet.Cells[i_rows, 11] = row["uud33"];
                    excelworksheet.Cells[i_rows, 12] = row["uud4"];
                    excelworksheet.Cells[i_rows, 13] = row["uud5"];
                    i_rows++;
                    break;
                case "13":
                    excelworksheet.Cells[i_rows, 3] = row["uud11"];
                    excelworksheet.Cells[i_rows, 6] = row["uud21"];
                    excelworksheet.Cells[i_rows, 7] = row["uud22"];
                    excelworksheet.Cells[i_rows, 8] = row["uud23"];
                    excelworksheet.Cells[i_rows, 9] = row["uud31"];
                    excelworksheet.Cells[i_rows, 12] = row["uud4"];
                    excelworksheet.Cells[i_rows, 13] = row["uud5"];
                    i_rows++;
                    break;
                case "123":
                    excelworksheet.Cells[i_rows, 3] = row["uud11"];
                    excelworksheet.Cells[i_rows, 6] = row["uud21"];
                    excelworksheet.Cells[i_rows, 9] = row["uud31"];
                    excelworksheet.Cells[i_rows, 12] = row["uud4"];
                    excelworksheet.Cells[i_rows, 13] = row["uud5"];
                    i_rows++;
                    break;
                case "23":
                    excelworksheet.Cells[i_rows, 3] = row["uud11"];
                    excelworksheet.Cells[i_rows, 4] = row["uud12"];
                    excelworksheet.Cells[i_rows, 5] = row["uud13"];
                    excelworksheet.Cells[i_rows, 6] = row["uud21"];
                    excelworksheet.Cells[i_rows, 9] = row["uud31"];
                    excelworksheet.Cells[i_rows, 12] = row["uud4"];
                    excelworksheet.Cells[i_rows, 13] = row["uud5"];
                    i_rows++;
                    break;
            }
        }

        private void Add_Rows_3_v_1(DataRow row)
        {
                foreach (DataRow roww in in_statDataSet.user.Rows)
                    {
                        if (Convert.ToString(row["id_user"]) == Convert.ToString(roww["id"]))
                        {
                            excelworksheet.Cells[i_rows, 2] = roww["fi"].ToString();
                        }
                    }

                    uud = "";
                    if (row["uud12"].ToString() == "" || row["uud13"].ToString() == "")
                        uud += "1";
                    if (row["uud22"].ToString() == "" || row["uud23"].ToString() == "")
                        uud += "2";
                    if (row["uud32"].ToString() == "" || row["uud33"].ToString() == "")
                        uud += "3";
                    switch (uud)
                    {
                        case "":
                            excelworksheet.Cells[i_rows, 3] = (Convert.ToInt16(row["uud11"]) + Convert.ToInt16(row["uud12"]) + Convert.ToInt16(row["uud13"]) > 1) ? 1 : 0;
                            excelworksheet.Cells[i_rows, 4] = (Convert.ToInt16(row["uud21"]) + Convert.ToInt16(row["uud22"]) + Convert.ToInt16(row["uud23"]) > 1) ? 1 : 0;
                            excelworksheet.Cells[i_rows, 5] = (Convert.ToInt16(row["uud31"]) + Convert.ToInt16(row["uud32"]) + Convert.ToInt16(row["uud33"]) > 1) ? 1 : 0;
                            excelworksheet.Cells[i_rows, 6] = row["uud4"];
                            excelworksheet.Cells[i_rows, 7] = row["uud5"];
                            i_rows++;
                            break;
                        case "1":
                            excelworksheet.Cells[i_rows, 3] = row["uud11"];
                            excelworksheet.Cells[i_rows, 4] = (Convert.ToInt16(row["uud21"]) + Convert.ToInt16(row["uud22"]) + Convert.ToInt16(row["uud23"]) > 1) ? 1 : 0;
                            excelworksheet.Cells[i_rows, 5] = (Convert.ToInt16(row["uud31"]) + Convert.ToInt16(row["uud32"]) + Convert.ToInt16(row["uud33"]) > 1) ? 1 : 0;
                            excelworksheet.Cells[i_rows, 6] = row["uud4"];
                            excelworksheet.Cells[i_rows, 7] = row["uud5"];
                            i_rows++;
                            break;
                        case "2":
                            excelworksheet.Cells[i_rows, 4] = row["uud21"];
                            excelworksheet.Cells[i_rows, 3] = (Convert.ToInt16(row["uud11"]) + Convert.ToInt16(row["uud12"]) + Convert.ToInt16(row["uud13"]) > 1) ? 1 : 0;
                            excelworksheet.Cells[i_rows, 5] = (Convert.ToInt16(row["uud31"]) + Convert.ToInt16(row["uud32"]) + Convert.ToInt16(row["uud33"]) > 1) ? 1 : 0;
                            excelworksheet.Cells[i_rows, 6] = row["uud4"];
                            excelworksheet.Cells[i_rows, 7] = row["uud5"];
                            i_rows++;
                            break;
                        case "3":
                            excelworksheet.Cells[i_rows, 5] = row["uud31"];
                            excelworksheet.Cells[i_rows, 3] = (Convert.ToInt16(row["uud11"]) + Convert.ToInt16(row["uud12"]) + Convert.ToInt16(row["uud13"]) > 1) ? 1 : 0;
                            excelworksheet.Cells[i_rows, 4] = (Convert.ToInt16(row["uud21"]) + Convert.ToInt16(row["uud22"]) + Convert.ToInt16(row["uud23"]) > 1) ? 1 : 0;
                            excelworksheet.Cells[i_rows, 6] = row["uud4"];
                            excelworksheet.Cells[i_rows, 7] = row["uud5"];
                            i_rows++;
                            break;
                        case "12":
                            excelworksheet.Cells[i_rows, 3] = row["uud11"];
                            excelworksheet.Cells[i_rows, 4] = row["uud21"];
                            excelworksheet.Cells[i_rows, 5] = (Convert.ToInt16(row["uud31"]) + Convert.ToInt16(row["uud32"]) + Convert.ToInt16(row["uud33"]) > 1) ? 1 : 0;
                            excelworksheet.Cells[i_rows, 6] = row["uud4"];
                            excelworksheet.Cells[i_rows, 7] = row["uud5"];
                            i_rows++;
                            break;
                        case "13":
                            excelworksheet.Cells[i_rows, 3] = row["uud11"];
                            excelworksheet.Cells[i_rows, 5] = row["uud31"];
                            excelworksheet.Cells[i_rows, 4] = (Convert.ToInt16(row["uud21"]) + Convert.ToInt16(row["uud22"]) + Convert.ToInt16(row["uud23"]) > 1) ? 1 : 0;
                            excelworksheet.Cells[i_rows, 6] = row["uud4"];
                            excelworksheet.Cells[i_rows, 7] = row["uud5"];
                            i_rows++;
                            break;
                        case "123":
                            excelworksheet.Cells[i_rows, 3] = row["uud11"];
                            excelworksheet.Cells[i_rows, 4] = row["uud21"];
                            excelworksheet.Cells[i_rows, 5] = row["uud31"];
                            excelworksheet.Cells[i_rows, 6] = row["uud4"];
                            excelworksheet.Cells[i_rows, 7] = row["uud5"];
                            i_rows++;
                            break;
                        case "23":
                            excelworksheet.Cells[i_rows, 3] = (Convert.ToInt16(row["uud11"]) + Convert.ToInt16(row["uud12"]) + Convert.ToInt16(row["uud13"]) > 1) ? 1 : 0;
                            excelworksheet.Cells[i_rows, 4] = row["uud21"];
                            excelworksheet.Cells[i_rows, 5] = row["uud31"];
                            excelworksheet.Cells[i_rows, 6] = row["uud4"];
                            excelworksheet.Cells[i_rows, 7] = row["uud5"];
                            i_rows++;
                            break;
                    }

                }

        private void Add_Rows_3_v_1_2(DataRow row)
        {
            foreach (DataRow roww in in_statDataSet.user.Rows)
            {
                if (Convert.ToString(row["id_user"]) == Convert.ToString(roww["id"]))
                {
                    excelworksheet.Cells[i_rows, 2] = roww["fi"].ToString();
                }
            }

            uud = "";
            if (row["uud12"].ToString() == "" || row["uud13"].ToString() == "")
                uud += "1";
            if (row["uud22"].ToString() == "" || row["uud23"].ToString() == "")
                uud += "2";
            if (row["uud32"].ToString() == "" || row["uud33"].ToString() == "")
                uud += "3";
            switch (uud)
            {
                case "":
                    excelworksheet.Cells[i_rows, 3] = (Convert.ToInt16(row["uud11"]) + Convert.ToInt16(row["uud12"]) + Convert.ToInt16(row["uud13"]) > 1) ? 1 : 0;
                    excelworksheet.Cells[i_rows, 4] = (Convert.ToInt16(row["uud21"]) + Convert.ToInt16(row["uud22"]) + Convert.ToInt16(row["uud23"]) > 1) ? 1 : 0;
                    excelworksheet.Cells[i_rows, 5] = (Convert.ToInt16(row["uud31"]) + Convert.ToInt16(row["uud32"]) + Convert.ToInt16(row["uud33"]) > 1) ? 1 : 0;
                    excelworksheet.Cells[i_rows, 6] = row["uud4"];
                    excelworksheet.Cells[i_rows, 7] = row["uud5"];
                    i_rows++;
                    break;
                case "1":
                    excelworksheet.Cells[i_rows, 3] = row["uud11"];
                    excelworksheet.Cells[i_rows, 4] = (Convert.ToInt16(row["uud21"]) + Convert.ToInt16(row["uud22"]) + Convert.ToInt16(row["uud23"]) > 1) ? 1 : 0;
                    excelworksheet.Cells[i_rows, 5] = (Convert.ToInt16(row["uud31"]) + Convert.ToInt16(row["uud32"]) + Convert.ToInt16(row["uud33"]) > 1) ? 1 : 0;
                    excelworksheet.Cells[i_rows, 6] = row["uud4"];
                    excelworksheet.Cells[i_rows, 7] = row["uud5"];
                    i_rows++;
                    break;
                case "2":
                    excelworksheet.Cells[i_rows, 4] = row["uud21"];
                    excelworksheet.Cells[i_rows, 3] = (Convert.ToInt16(row["uud11"]) + Convert.ToInt16(row["uud12"]) + Convert.ToInt16(row["uud13"]) > 1) ? 1 : 0;
                    excelworksheet.Cells[i_rows, 5] = (Convert.ToInt16(row["uud31"]) + Convert.ToInt16(row["uud32"]) + Convert.ToInt16(row["uud33"]) > 1) ? 1 : 0;
                    excelworksheet.Cells[i_rows, 6] = row["uud4"];
                    excelworksheet.Cells[i_rows, 7] = row["uud5"];
                    i_rows++;
                    break;
                case "3":
                    excelworksheet.Cells[i_rows, 5] = row["uud31"];
                    excelworksheet.Cells[i_rows, 3] = (Convert.ToInt16(row["uud11"]) + Convert.ToInt16(row["uud12"]) + Convert.ToInt16(row["uud13"]) > 1) ? 1 : 0;
                    excelworksheet.Cells[i_rows, 4] = (Convert.ToInt16(row["uud21"]) + Convert.ToInt16(row["uud22"]) + Convert.ToInt16(row["uud23"]) > 1) ? 1 : 0;
                    excelworksheet.Cells[i_rows, 6] = row["uud4"];
                    excelworksheet.Cells[i_rows, 7] = row["uud5"];
                    i_rows++;
                    break;
                case "12":
                    excelworksheet.Cells[i_rows, 3] = row["uud11"];
                    excelworksheet.Cells[i_rows, 4] = row["uud21"];
                    excelworksheet.Cells[i_rows, 5] = (Convert.ToInt16(row["uud31"]) + Convert.ToInt16(row["uud32"]) + Convert.ToInt16(row["uud33"]) > 1) ? 1 : 0;
                    excelworksheet.Cells[i_rows, 6] = row["uud4"];
                    excelworksheet.Cells[i_rows, 7] = row["uud5"];
                    i_rows++;
                    break;
                case "13":
                    excelworksheet.Cells[i_rows, 3] = row["uud11"];
                    excelworksheet.Cells[i_rows, 5] = row["uud31"];
                    excelworksheet.Cells[i_rows, 4] = (Convert.ToInt16(row["uud21"]) + Convert.ToInt16(row["uud22"]) + Convert.ToInt16(row["uud23"]) > 1) ? 1 : 0;
                    excelworksheet.Cells[i_rows, 6] = row["uud4"];
                    excelworksheet.Cells[i_rows, 7] = row["uud5"];
                    i_rows++;
                    break;
                case "123":
                    excelworksheet.Cells[i_rows, 3] = row["uud11"];
                    excelworksheet.Cells[i_rows, 4] = row["uud21"];
                    excelworksheet.Cells[i_rows, 5] = row["uud31"];
                    excelworksheet.Cells[i_rows, 6] = row["uud4"];
                    excelworksheet.Cells[i_rows, 7] = row["uud5"];
                    i_rows++;
                    break;
                case "23":
                    excelworksheet.Cells[i_rows, 3] = (Convert.ToInt16(row["uud11"]) + Convert.ToInt16(row["uud12"]) + Convert.ToInt16(row["uud13"]) > 1) ? 1 : 0;
                    excelworksheet.Cells[i_rows, 4] = row["uud21"];
                    excelworksheet.Cells[i_rows, 5] = row["uud31"];
                    excelworksheet.Cells[i_rows, 6] = row["uud4"];
                    excelworksheet.Cells[i_rows, 7] = row["uud5"];
                    i_rows++;
                    break;
            }

        }

        private void Del_Collums()
        {
            //-------------------------------------------------------удаление столбцов------------------------------------------------------------------
            switch (uud)
            {
                case "1":
                    excelworksheet.Columns[4].Delete();
                    excelworksheet.Columns[4].Delete();
                    Coloring_Diag_1(uud);
                    break;
                case "2":
                    excelworksheet.Columns[7].Delete();
                    excelworksheet.Columns[7].Delete();
                    Coloring_Diag_1(uud);
                    break;
                case "3":
                    excelworksheet.Columns[10].Delete();
                    excelworksheet.Columns[10].Delete();
                    Coloring_Diag_1(uud);
                    break;
                case "12":
                    excelworksheet.Columns[4].Delete();
                    excelworksheet.Columns[4].Delete();
                    excelworksheet.Columns[5].Delete();
                    excelworksheet.Columns[5].Delete();
                    Coloring_Diag_1(uud);
                    break;
                case "13":
                    excelworksheet.Columns[4].Delete();
                    excelworksheet.Columns[4].Delete();
                    excelworksheet.Columns[7].Delete();
                    excelworksheet.Columns[7].Delete();
                    Coloring_Diag_1(uud);
                    break;
                case "123":
                    excelworksheet.Columns[4].Delete();
                    excelworksheet.Columns[4].Delete();
                    excelworksheet.Columns[5].Delete();
                    excelworksheet.Columns[5].Delete();
                    excelworksheet.Columns[6].Delete();
                    excelworksheet.Columns[6].Delete();
                    Coloring_Diag_1(uud);
                    break;
                case "23":
                    excelworksheet.Columns[7].Delete();
                    excelworksheet.Columns[7].Delete();
                    excelworksheet.Columns[8].Delete();
                    excelworksheet.Columns[8].Delete();
                    Coloring_Diag_1(uud);
                    break;
            }
        }

        private void Del_Rows()
        {
            //----------------------------------------------удаление строк---------------------------------------------------------------------------
            for (int j = i_rows; j < 111; j++)
            {
                excelworksheet.Rows[i_rows].Delete();
            }
        }

        private void Hiden_Collums()
        {
            //-------------------------------------------------------скрытие столбцов------------------------------------------------------------------
            switch (uud)
            {
                case "1":
                    excelworksheet.Columns[4].Hidden = true;
                    excelworksheet.Columns[5].Hidden = true;
                    //   Coloring_Diag_1(uud);
                    break;
                case "2":
                    excelworksheet.Columns[7].Hidden = true;
                    excelworksheet.Columns[8].Hidden = true;
                    //  Coloring_Diag_1(uud);
                    break;
                case "3":
                    excelworksheet.Columns[10].Hidden = true;
                    excelworksheet.Columns[11].Hidden = true;
                    //   Coloring_Diag_1(uud);
                    break;
                case "12":
                    excelworksheet.Columns[4].Hidden = true;
                    excelworksheet.Columns[5].Hidden = true;
                    excelworksheet.Columns[7].Hidden = true;
                    excelworksheet.Columns[8].Hidden = true;
                    //   Coloring_Diag_1(uud);
                    break;
                case "13":
                    excelworksheet.Columns[4].Hidden = true;
                    excelworksheet.Columns[5].Hidden = true;
                    excelworksheet.Columns[7].Hidden = true;
                    excelworksheet.Columns[8].Hidden = true;
                    //  Coloring_Diag_1(uud);
                    break;
                case "123":
                    excelworksheet.Columns[4].Hidden = true;
                    excelworksheet.Columns[5].Hidden = true;
                    excelworksheet.Columns[7].Hidden = true;
                    excelworksheet.Columns[8].Hidden = true;
                    excelworksheet.Columns[10].Hidden = true;
                    excelworksheet.Columns[11].Hidden = true;
                    //  Coloring_Diag_1(uud);
                    break;
                case "23":
                    excelworksheet.Columns[7].Hidden = true;
                    excelworksheet.Columns[8].Hidden = true;
                    excelworksheet.Columns[10].Hidden = true;
                    excelworksheet.Columns[11].Hidden = true;
                    //    Coloring_Diag_1(uud);
                    break;
            }
        }

        private void Coloring_Diag_1(string uud)
        {
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(3);
            excelworksheet.Activate();
            Excel.ChartObjects chartsobjrcts = (Excel.ChartObjects)excelworksheet.ChartObjects(Type.Missing);
            Excel.Chart xlChart = excelworksheet.ChartObjects(1).Chart;
            Excel.Series ser = (Excel.Series)xlChart.SeriesCollection(1);

            switch (uud)
            {
                case "1":
                    ser.Points(1).Interior.Color = (int)XlRgbColor.rgbGreen;
                    ser.Points(2).Interior.Color = (int)XlRgbColor.rgbDarkBlue;
                    ser.Points(3).Interior.Color = (int)XlRgbColor.rgbDarkBlue;
                    ser.Points(4).Interior.Color = (int)XlRgbColor.rgbDarkBlue;
                    ser.Points(5).Interior.Color = (int)XlRgbColor.rgbOrange;
                    ser.Points(6).Interior.Color = (int)XlRgbColor.rgbOrange;
                    ser.Points(7).Interior.Color = (int)XlRgbColor.rgbOrange;
                    ser.Points(8).Interior.Color = (int)XlRgbColor.rgbBlue;
                    ser.Points(9).Interior.Color = (int)XlRgbColor.rgbPurple;
                    break;
                case "2":
                    ser.Points(1).Interior.Color = (int)XlRgbColor.rgbGreen;
                    ser.Points(2).Interior.Color = (int)XlRgbColor.rgbGreen;
                    ser.Points(3).Interior.Color = (int)XlRgbColor.rgbGreen;
                    ser.Points(4).Interior.Color = (int)XlRgbColor.rgbDarkBlue;
                    ser.Points(5).Interior.Color = (int)XlRgbColor.rgbOrange;
                    ser.Points(6).Interior.Color = (int)XlRgbColor.rgbOrange;
                    ser.Points(7).Interior.Color = (int)XlRgbColor.rgbOrange;
                    ser.Points(8).Interior.Color = (int)XlRgbColor.rgbBlue;
                    ser.Points(9).Interior.Color = (int)XlRgbColor.rgbPurple;
                    break;
                case "3":
                    ser.Points(1).Interior.Color = (int)XlRgbColor.rgbGreen;
                    ser.Points(2).Interior.Color = (int)XlRgbColor.rgbGreen;
                    ser.Points(3).Interior.Color = (int)XlRgbColor.rgbGreen;
                    ser.Points(4).Interior.Color = (int)XlRgbColor.rgbDarkBlue;
                    ser.Points(5).Interior.Color = (int)XlRgbColor.rgbDarkBlue;
                    ser.Points(6).Interior.Color = (int)XlRgbColor.rgbDarkBlue;
                    ser.Points(7).Interior.Color = (int)XlRgbColor.rgbOrange;
                    ser.Points(8).Interior.Color = (int)XlRgbColor.rgbBlue;
                    ser.Points(9).Interior.Color = (int)XlRgbColor.rgbPurple;
                    break;
                case "12":
                    ser.Points(1).Interior.Color = (int)XlRgbColor.rgbGreen;
                    ser.Points(2).Interior.Color = (int)XlRgbColor.rgbDarkBlue;
                    ser.Points(3).Interior.Color = (int)XlRgbColor.rgbOrange;
                    ser.Points(4).Interior.Color = (int)XlRgbColor.rgbOrange;
                    ser.Points(5).Interior.Color = (int)XlRgbColor.rgbOrange;
                    ser.Points(6).Interior.Color = (int)XlRgbColor.rgbBlue;
                    ser.Points(7).Interior.Color = (int)XlRgbColor.rgbPurple;
                    break;
                case "13":
                    ser.Points(1).Interior.Color = (int)XlRgbColor.rgbGreen;
                    ser.Points(2).Interior.Color = (int)XlRgbColor.rgbDarkBlue;
                    ser.Points(3).Interior.Color = (int)XlRgbColor.rgbDarkBlue;
                    ser.Points(4).Interior.Color = (int)XlRgbColor.rgbDarkBlue;
                    ser.Points(5).Interior.Color = (int)XlRgbColor.rgbOrange;
                    ser.Points(6).Interior.Color = (int)XlRgbColor.rgbBlue;
                    ser.Points(7).Interior.Color = (int)XlRgbColor.rgbPurple;
                    break;
                case "123":
                    ser.Points(1).Interior.Color = (int)XlRgbColor.rgbGreen;
                    ser.Points(2).Interior.Color = (int)XlRgbColor.rgbDarkBlue;
                    ser.Points(3).Interior.Color = (int)XlRgbColor.rgbOrange;
                    ser.Points(4).Interior.Color = (int)XlRgbColor.rgbBlue;
                    ser.Points(5).Interior.Color = (int)XlRgbColor.rgbPurple;
                    break;
                case "23":
                    ser.Points(1).Interior.Color = (int)XlRgbColor.rgbGreen;
                    ser.Points(2).Interior.Color = (int)XlRgbColor.rgbGreen;
                    ser.Points(3).Interior.Color = (int)XlRgbColor.rgbGreen;
                    ser.Points(4).Interior.Color = (int)XlRgbColor.rgbDarkBlue;
                    ser.Points(5).Interior.Color = (int)XlRgbColor.rgbOrange;
                    ser.Points(6).Interior.Color = (int)XlRgbColor.rgbBlue;
                    ser.Points(7).Interior.Color = (int)XlRgbColor.rgbPurple;
                    break;
            }
          
       
        }

        private void Diad_tabl_1()
        {
            i_rows = 2;
            try
            {
                 _control = ComboBox_Kontrol_Stat.SelectedValue.ToString();
                 _klass = ComboBox_Klass_Stat.SelectedValue.ToString();
                 _god = ComboBox_God_Stat.SelectedItem.ToString();

                
                excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

                Add_Row1();

                Del_Rows();

                Del_Collums();
            }
            catch (FormatException fEx)
            {
                MessageBox.Show(fEx.ToString());
            }

            catch (OverflowException oEx)
            {
                MessageBox.Show(oEx.ToString());
            }
            catch (NullReferenceException)
            {
                MessageBox.Show("Вы не заполнили один из комбобокс");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            finally
            {
                if(i_rows == 2)
                {
                    MessageBox.Show("Такой контрольной не найденно");
                }
            }

        }

        private void Diad_tabl_2()
        {
            i_rows = 2;
            try
            {
                 _control = ComboBox_Kontrol_Stat.SelectedValue.ToString();
                 _klass = ComboBox_Klass_Stat.SelectedValue.ToString();
                 _god = ComboBox_God_Stat.SelectedItem.ToString();


                excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

                Add_Row_3_v_1();

                Del_Rows();

            }
            catch (FormatException fEx)
            {
                MessageBox.Show(fEx.ToString());
            }

            catch (OverflowException oEx)
            {
                MessageBox.Show(oEx.ToString());
            }
            catch (NullReferenceException nEx)
            {
                MessageBox.Show("Вы не заполнили один из комбобокс");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            finally
            {
                if (i_rows == 2)
                {
                    MessageBox.Show("Такой контрольной не найденно");
                }
            }

        }

        private void Diad_tabl_3()
        {
            i_rows = 2;
            try
            {
                 _control = ComboBox_Kontrol_Stat1.SelectedValue.ToString();
                 _klass = ComboBox_Klass_Stat1.SelectedValue.ToString();
                 _god = ComboBox_God_Stat1.SelectedItem.ToString();
                 _control2 = ComboBox_Kontrol_Stat2.SelectedValue.ToString();
                 _klass2 = ComboBox_Klass_Stat2.SelectedValue.ToString();
                 _god2 = ComboBox_God_Stat2.SelectedItem.ToString();

                excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

                Add_Row1();

                Del_Rows();

                Hiden_Collums();

                excelworksheet = (Excel.Worksheet)excelsheets.get_Item(2);

                Add_Row2();

                Del_Rows();

                Hiden_Collums();
              
            }
            catch (FormatException fEx)
            {
                MessageBox.Show(fEx.ToString());
            }

            catch (OverflowException oEx)
            {
                MessageBox.Show(oEx.ToString());
            }
            catch (NullReferenceException)
            {
                MessageBox.Show("Вы не заполнили один из комбобокс");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            finally
            {
                if (i_rows == 2)
                {
                    MessageBox.Show("Такой контрольной не найденно");
                }
            }

        }

        private void Diad_tabl_4()
        {
            i_rows = 2;
            try
            {
                _control = ComboBox_Kontrol_Stat1.SelectedValue.ToString();
                _klass = ComboBox_Klass_Stat1.SelectedValue.ToString();
                _god = ComboBox_God_Stat1.SelectedItem.ToString();
                _control2 = ComboBox_Kontrol_Stat2.SelectedValue.ToString();
                _klass2 = ComboBox_Klass_Stat2.SelectedValue.ToString();
                _god2 = ComboBox_God_Stat2.SelectedItem.ToString();

                excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

                Add_Row_3_v_1();

                Del_Rows();

                excelworksheet = (Excel.Worksheet)excelsheets.get_Item(2);

                Add_Row_3_v_1_2();

                Del_Rows();

            }
            catch (FormatException fEx)
            {
                MessageBox.Show(fEx.ToString());
            }

            catch (OverflowException oEx)
            {
                MessageBox.Show(oEx.ToString());
            }
            catch (NullReferenceException)
            {
                MessageBox.Show("Вы не заполнили один из комбобокс");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            finally
            {
                if (i_rows == 2)
                {
                    MessageBox.Show("Такой контрольной не найденно");
                }
            }

        }

        /*
                private void Diag_98()
                {
                    excelworksheet = (Excel.Worksheet)excelsheets.get_Item(4);
                    excelworksheet.Activate();
                    Excel.ChartObjects chartsobjrcts = (Excel.ChartObjects)excelworksheet.ChartObjects(Type.Missing);
                    Excel.Chart xlChart = excelworksheet.ChartObjects(1).Chart;
                    xlChart.ChartWizard(excelworksheet.get_Range("h2", "h5"));

                    //  excelworksheet.ChartObjects(1).Chart.Ledend(excelworksheet.get_Range("h2"));
                    /*  Excel.ChartObjects chartsobjrcts = (Excel.ChartObjects)excelworksheet.ChartObjects(Type.Missing);
                      Excel.ChartObject chartsobjrct = chartsobjrcts.Add(10, 200, 500, 400);
                      chartsobjrct.Chart.ChartWizard(excelworksheet.get_Range("c3", "g5"),
                      Excel.XlChartType.xlColumnClustered, 2, Excel.XlRowCol.xlRows, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                      chartsobjrct.Activate();
                      Excel.SeriesCollection seriesCollection = (Excel.SeriesCollection)excelapp.ActiveChart.SeriesCollection(Type.Missing);
                      Excel.Series series = seriesCollection.Item(1);
                      series.Name = "1";
                }

                private void Diag_99()
                {
                    excelworksheet = (Excel.Worksheet)excelsheets.get_Item(4);
                    excelworksheet.Activate();
                    Excel.ChartObjects chartsobjrcts = (Excel.ChartObjects)excelworksheet.ChartObjects(Type.Missing);
                    Excel.Chart xlChart2 = excelworksheet.ChartObjects(2).Chart;
                    xlChart2.ChartWizard(excelworksheet.get_Range("c3", "g5"));
                    Excel.SeriesCollection seriesCollection = xlChart2.SeriesCollection();

                    Excel.Series series = seriesCollection.Item(1);

                    for (int i = 1; i <= seriesCollection.Count; i++)
                    {
                        series = seriesCollection.Item(i);
                        series.Name = Convert.ToString(excelworksheet.get_Range("b" + (i + 2), Type.Missing).Value2);

                    }
                 //   series.XValues = "Понедельник;Вторник;Среда;";
                    /* Aspose.Cell пока удалил
                     * 
                     * 
                     * String templatePath = System.Windows.Forms.Application.StartupPath;
                     Workbook book = new Workbook(templatePath + @"\Шаблоны\Свод 1 ш.xlsx");

                     // Access the first worksheet which contains the charts
                     Worksheet sheet = book.Worksheets[3];

                     for (int i = 0; i < sheet.Charts.Count; i++)
                     {
                         // Access the chart
                         Chart ch = sheet.Charts[i];

                         // Print chart type
                         Console.WriteLine(ch.Type);

                         // Change the title of the charts as per their types
                         ch.Title.Text = "Chart Type is " + ch.Type.ToString();

                     }
                     book.Save(templatePath + "out_excel2016Charts.xlsx");


                }

                private void Diad_3()
                {
                    int i=0;
                    switch (i)
                    {
                        case 1:
                            excelapp = new Excel.Application();
                            excelapp.Visible = true;
                            //Получаем набор объектов Workbook (массив ссылок на созданные книги)
                            excelappworkbooks = excelapp.Workbooks;
                            //Открываем книгу и получаем на нее ссылку
                            //Помним, что файл был запаралирован
                            excelappworkbook = excelapp.Workbooks.Open(@"C:\a.xls", Type.Missing,
                                                                     Type.Missing, Type.Missing,
                             "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,
                              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                              Type.Missing, Type.Missing);
                            //Если бы мы открыли несколько книг, то получили ссылку так
                            //excelappworkbook=excelappworkbooks[1];
                            //Получаем массив ссылок на листы выбранной книги
                            excelsheets = excelappworkbook.Worksheets;
                            //Получаем ссылку на лист 1
                            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
                            //Выделяем ячейки с данными  в таблице
                            excelcells = excelworksheet.get_Range("D8", "K10");
                            //И выбираем их
                            excelcells.Select();
                            //Создаем объект Excel.Chart диаграмму по умолчанию
                            Excel.Chart excelchart = (Excel.Chart)excelapp.Charts.Add(Type.Missing,
                             Type.Missing, Type.Missing, Type.Missing);
                            //Выбираем диграмму - отображаем лист с диаграммой
                            excelchart.Activate();
                            excelchart.Select(Type.Missing);
                            //Изменяем тип диаграммы
                            excelapp.ActiveChart.ChartType = Excel.XlChartType.xlConeCol;
                            //Создаем надпись - Заглавие диаграммы
                            excelapp.ActiveChart.HasTitle = true;
                            excelapp.ActiveChart.ChartTitle.Text
                               = "Продажи фирмы Рога и Копыта за неделю";
                            //Меняем шрифт, можно поменять и другие параметры шрифта
                            excelapp.ActiveChart.ChartTitle.Font.Size = 14;
                            excelapp.ActiveChart.ChartTitle.Font.Color = 255;
                            //Обрамление для надписи c тенями
                            excelapp.ActiveChart.ChartTitle.Shadow = true;
                            excelapp.ActiveChart.ChartTitle.Border.LineStyle
                                 = Excel.Constants.xlSolid;
                            //Даем названия осей
                            ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlCategory,
                                Excel.XlAxisGroup.xlPrimary)).HasTitle = true;
                            ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlCategory,
                                Excel.XlAxisGroup.xlPrimary)).AxisTitle.Text = "День недели";
                            ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlSeriesAxis,
                                Excel.XlAxisGroup.xlPrimary)).HasTitle = false;
                            ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlValue,
                                Excel.XlAxisGroup.xlPrimary)).HasTitle = true;
                            ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlValue,
                                Excel.XlAxisGroup.xlPrimary)).AxisTitle.Text = "Рогов/Копыт";
                            //Координатная сетка - оставляем только крупную сетку
                            ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlCategory,
                               Excel.XlAxisGroup.xlPrimary)).HasMajorGridlines = true;
                            ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlCategory,
                              Excel.XlAxisGroup.xlPrimary)).HasMinorGridlines = false;
                            ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlSeriesAxis,
                              Excel.XlAxisGroup.xlPrimary)).HasMajorGridlines = true;
                            ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlSeriesAxis,
                              Excel.XlAxisGroup.xlPrimary)).HasMinorGridlines = false;
                            ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlValue,
                              Excel.XlAxisGroup.xlPrimary)).HasMinorGridlines = false;
                            ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlValue,
                              Excel.XlAxisGroup.xlPrimary)).HasMajorGridlines = true;
                            //Будем отображать легенду и уберем строки, 
                            //которые отображают пустые строки таблицы
                            excelapp.ActiveChart.HasLegend = true;
                            //Расположение легенды
                            excelapp.ActiveChart.Legend.Position
                               = Excel.XlLegendPosition.xlLegendPositionLeft;
                            //Можно изменить шрифт легенды и другие параметры 
                            ((Excel.LegendEntry)excelapp.ActiveChart.Legend.LegendEntries(1)).Font.Size = 12;
                            ((Excel.LegendEntry)excelapp.ActiveChart.Legend.LegendEntries(3)).Font.Size = 12;
                            //Легенда тесно связана с подписями на осях - изменяем надписи
                            // - меняем легенду, удаляем чтото на оси - изменяется легенда
                            Excel.SeriesCollection seriesCollection =
                             (Excel.SeriesCollection)excelapp.ActiveChart.SeriesCollection(Type.Missing);
                            Excel.Series series = seriesCollection.Item(1);
                            series.Name = "Рога";
                            //Помним, что у нас объединенные ячейки, значит каждая второя строка - пустая
                            //Удаляем их из диаграммы и из легенды
                            series = seriesCollection.Item(2);
                            series.Delete();
                            //После удаления второго (пустого набора значений) третий занял его место
                            series = seriesCollection.Item(2);
                            series.Name = "Копыта";
                            series = seriesCollection.Item(3);
                            series.Delete();
                            series = seriesCollection.Item(1);
                            //Переименуем ось X
                            series.XValues = "Понедельник;Вторник;Среда;Четверг;Пятница;Суббота;Воскресенье;Итог";
                            //Если закончить код на этом месте то у нас Диаграммы на отдельном листе - Рис.9.
                            //Строку легенды можно удалить здесь, но строка на оси не изменится
                            //Поэтому мы удаляли в Excel.Series
                            //((Excel.LegendEntry)excelapp.ActiveChart.Legend.LegendEntries(2)).Delete();
                            //Перемещаем диаграмму на лист 1
                            excelapp.ActiveChart.Location(Excel.XlChartLocation.xlLocationAsObject, "Лист1");
                            //Получаем ссылку на лист 1
                            excelsheets = excelappworkbook.Worksheets;
                            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
                            //Перемещаем диаграмму в нужное место
                            excelworksheet.Shapes.Item(1).IncrementLeft(-201);
                            excelworksheet.Shapes.Item(1).IncrementTop((float)20.5);
                            //Задаем размеры диаграммы
                            excelworksheet.Shapes.Item(1).Height = 550;
                            excelworksheet.Shapes.Item(1).Width = 500;
                            //Конец кода - диаграммы на листе там где и таблица
                            break;
                        case 2:
                            excelappworkbooks = excelapp.Workbooks;
                            excelappworkbook = excelappworkbooks[1];
                            excelappworkbook.Save();
                            excelapp.Quit();
                            break;
                        default:
                            Close();
                            break;
                    }
                }

                private void Diad_4()
                {
                    excelapp = new Excel.Application();
                    excelapp.Visible = true;
                    excelappworkbooks = excelapp.Workbooks;
                    excelappworkbook = excelapp.Workbooks.Open(@"C:\a.xls", Type.Missing,
                                                               Type.Missing, Type.Missing,
                   "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                   Type.Missing, Type.Missing);
                    excelsheets = excelappworkbook.Worksheets;
                    excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
                    //Определяем диаграммы как объекты Excel.ChartObjects
                    Excel.ChartObjects chartsobjrcts =
                   (Excel.ChartObjects)excelworksheet.ChartObjects(Type.Missing);
                    //Добавляем одну диаграмму  в Excel.ChartObjects - диаграмма пока 
                    //не выбрана, но место для нее выделено в методе Add
                    Excel.ChartObject chartsobjrct = chartsobjrcts.Add(10, 200, 500, 400);
                    excelcells = excelworksheet.get_Range("D8", "K10");
                    //Получаем ссылку на созданную диаграмму
                    Excel.Chart excelchart = chartsobjrct.Chart;
                    //Устанавливаем источник данных для диаграммы
                    excelchart.SetSourceData(excelcells, Type.Missing);
                    //Далее отличия нет
                    excelchart.ChartType = Excel.XlChartType.xlConeCol;
                    excelchart.HasTitle = true;
                    excelchart.ChartTitle.Text = "Продажи фирмы Рога и Копыта за неделю";
                    excelchart.ChartTitle.Font.Size = 14;
                    excelchart.ChartTitle.Font.Color = 255;
                    excelchart.ChartTitle.Shadow = true;
                    excelchart.ChartTitle.Border.LineStyle = Excel.Constants.xlSolid;
                    ((Excel.Axis)(excelchart.Axes(Excel.XlAxisType.xlCategory,
                                  Excel.XlAxisGroup.xlPrimary)))
                                       .HasTitle = true;
                    ((Excel.Axis)excelchart.Axes(Excel.XlAxisType.xlCategory,
                      Excel.XlAxisGroup.xlPrimary)).HasTitle = true;
                    ((Excel.Axis)excelchart.Axes(Excel.XlAxisType.xlCategory,
                      Excel.XlAxisGroup.xlPrimary)).AxisTitle.Text = "День недели";
                    ((Excel.Axis)excelchart.Axes(Excel.XlAxisType.xlSeriesAxis,
                      Excel.XlAxisGroup.xlPrimary)).HasTitle = false;
                    ((Excel.Axis)excelchart.Axes(Excel.XlAxisType.xlValue,
                      Excel.XlAxisGroup.xlPrimary)).HasTitle = true;
                    ((Excel.Axis)excelchart.Axes(Excel.XlAxisType.xlValue,
                      Excel.XlAxisGroup.xlPrimary)).AxisTitle.Text = "Рогов/Копыт";
                    ((Excel.Axis)excelchart.Axes(Excel.XlAxisType.xlCategory,
                      Excel.XlAxisGroup.xlPrimary)).HasMajorGridlines = true;
                    ((Excel.Axis)excelchart.Axes(Excel.XlAxisType.xlCategory,
                      Excel.XlAxisGroup.xlPrimary)).HasMinorGridlines = false;
                    ((Excel.Axis)excelchart.Axes(Excel.XlAxisType.xlSeriesAxis,
                      Excel.XlAxisGroup.xlPrimary)).HasMajorGridlines = true;
                    ((Excel.Axis)excelchart.Axes(Excel.XlAxisType.xlSeriesAxis,
                      Excel.XlAxisGroup.xlPrimary)).HasMinorGridlines = false;
                    ((Excel.Axis)excelchart.Axes(Excel.XlAxisType.xlValue,
                      Excel.XlAxisGroup.xlPrimary)).HasMinorGridlines = false;
                    ((Excel.Axis)excelchart.Axes(Excel.XlAxisType.xlValue,
                      Excel.XlAxisGroup.xlPrimary)).HasMajorGridlines = true;
                    excelchart.HasLegend = true;
                    excelchart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionLeft;
                    ((Excel.LegendEntry)excelchart.Legend.LegendEntries(1)).Font.Size = 12;
                    ((Excel.LegendEntry)excelchart.Legend.LegendEntries(3)).Font.Size = 12;
                    Excel.SeriesCollection seriesCollection =
                     (Excel.SeriesCollection)excelchart.SeriesCollection(Type.Missing);
                    Excel.Series series = seriesCollection.Item(1);
                    series.Name = "Рога";
                    series = seriesCollection.Item(2);
                    series.Delete();
                    series = seriesCollection.Item(2);
                    series.Name = "Копыта";
                    series = seriesCollection.Item(1);
                    series.XValues = "Понедельник;Вторник;Среда;Четверг;Пятница;Суббота;Воскресенье;Итог";
                }
            */
        private void metroTile2_Click(object sender, EventArgs e)
        {
            switch (TabControl_Stat.SelectedIndex)
            {
                case 0:
                    Excel_Diag_tab1();
                    Excel_Diag_tab2();
                    break;
                case 1:
                    Excel_Diag_tab3();
                    Excel_Diag_tab4();
                    break;
            }
        }
        private void ComboBox_God_Load_SelectedIndexChanged(object sender, EventArgs e)
        { 

            Update_Combobox_Kontrol_Load();

        }

      
//=======
       

        
//>>>>>>> 6e168095ff9b9a19d30e617a0b07114c2a31c458
    }
}
