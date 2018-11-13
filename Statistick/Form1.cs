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
                    Grid_Load_UUD.Columns.Add(new DataGridViewTextBoxColumn() { Name = "uud5", HeaderText = "УУД2-2", Width = 100, DisplayIndex = (check_uud1.Checked) ? 5 : 3  });
                    Grid_Load_UUD.Columns.Add(new DataGridViewTextBoxColumn() { Name = "uud6", HeaderText = "УУД2-3", Width = 100, DisplayIndex = (check_uud1.Checked) ? 6 : 4  });
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

        private void ComboBox_God_Load_SelectedIndexChanged(object sender, EventArgs e)
        {
            Update_Combobox_Kontrol_Load();

        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
          
            MessageBox.Show(Est_v_BD().ToString());
        }
    }
}
