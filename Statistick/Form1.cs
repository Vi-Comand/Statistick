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
using Aspose.Cells.Charts;
using Aspose.Cells;

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
            chartControl1.Series.Add(series1);
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
            ((XYDiagram3D)chartControl1.Diagram).ZoomPercent = 110;

            // Add a title to the chart and hide the legend.
            ChartTitle chartTitle1 = new ChartTitle();

            chartTitle1.Text = "Общая диограмма по позициям";
            chartControl1.Titles.Add(chartTitle1);
            //   chartControl1.Legend.Visible = false;
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





        //--------------------методы диограмм---------------------------------------------------------------------------------------------------------
        private Excel.Application excelapp;
        private Excel.Workbooks excelappworkbooks;
        private Excel.Workbook excelappworkbook;
        private Excel.Sheets excelsheets;
        private Excel.Worksheet excelworksheet;
        private Excel.Range excelcells;
        private Excel.Window excelWindow;

        private void Excel_Diag()
        {
           /* excelapp = new Excel.Application();
            excelapp.Visible = true;
            excelappworkbooks = excelapp.Workbooks;
            String templatePath = System.Windows.Forms.Application.StartupPath;
            excelappworkbook = excelapp.Workbooks.Open(templatePath + @"\Шаблоны\Свод 1 ш.xlsx", Type.Missing, Type.Missing, Type.Missing, "WWWWW", "WWWWW", Type.Missing, 
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excelsheets = excelappworkbook.Worksheets;*/

            Diag_2();
        }

        private void Diag_1()
        {
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(4);
            excelworksheet.Activate();
            Excel.ChartObjects chartsobjrcts = (Excel.ChartObjects)excelworksheet.ChartObjects(Type.Missing);
            Excel.ChartObject chartsobjrct = chartsobjrcts.Add(10, 200, 500, 400);
            chartsobjrct.Chart.ChartWizard(excelworksheet.get_Range("c3", "g5"),
            Excel.XlChartType.xlColumnClustered, 2, Excel.XlRowCol.xlRows, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            chartsobjrct.Activate();
            Excel.SeriesCollection seriesCollection = (Excel.SeriesCollection)excelapp.ActiveChart.SeriesCollection(Type.Missing);
            Excel.Series series = seriesCollection.Item(1);
            series.Name = "1";
        }

        private void Diag_2()
        {
            String templatePath = System.Windows.Forms.Application.StartupPath;
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

            /* Excel.ChartObjects chartsobjrcts2 = (Excel.ChartObjects)excelworksheet.ChartObjects(Type.Missing);
           //  Excel.SeriesCollection seriesCollection = (Excel.SeriesCollection)excelapp.ActiveChart.SeriesCollection(Type.Missing);
             Excel.Series oSeries;
             //    Dim oSeries As Series
             oSeries = excelworksheet.ChartObjects(1).SeriesCollection.NewSeries;
             //Set oSeries = Worksheets(1).ChartObjects(1).Chart.SeriesCollection.NewSeries
             oSeries.Values = excelworksheet.Range["H3:H6"];
 //oSeries.Values = Worksheets(1).Range("B1:B10")*/

            //Excel.ChartObject chartsobjrct2 = chartsobjrcts2.Select("1");
            //chartsobjrct2.Chart.ChartWizard(excelworksheet.get_Range("h3", "h5"),
            //   Excel.XlChartType.xlColumnClustered, 2, Excel.XlRowCol.xlRows, excelworksheet.get_Range("b3", "b5"), 0, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

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

        private void metroTile2_Click(object sender, EventArgs e)
        {
            Excel_Diag();
        }


    }
}
