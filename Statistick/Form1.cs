using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using DevExpress.Charts.Native;
using MetroFramework.Forms;
using Excel = Microsoft.Office.Interop.Excel;

//using DevExpress.XtraCharts;
using Microsoft.Office.Interop.Excel;

using System.Text.RegularExpressions;
//using DevExpress.Utils.Extensions;
using System.IO;
using System.Net.Sockets;
using System.Net;
using System.Globalization;
using TextBox = System.Windows.Forms.TextBox;
using System.Net.Mail;

namespace Statistick
{
    public partial class Form1 : MetroForm
    {
        public Form1()
        {
            InitializeComponent();


        }
    void Update_bd()
        {
            in_statDataSet.Clear();
            this.uudTableAdapter.Fill(this.in_statDataSet.uud);
            //   this.uudTableAdapter.Fill(this.in_statDataSet.uud);
            this.kontrolnieTableAdapter.Fill(this.in_statDataSet.kontrolnie);
            this.userTableAdapter.Fill(this.in_statDataSet.user);
            this.klassTableAdapter.Fill(this.in_statDataSet.klass);
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {

                Update_bd();
            }
            catch 
            {
                DialogResult dialogResult =
                                    MessageBox.Show(
                                        "На компьюторе не устоновленна программа для работы с базами данных. Она будет установлена. Нажмите \"Да\", чтобы начать установку. Нажмите \"Нет\", если не хотите устанавливать программу.",
                                        "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialogResult == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(System.Windows.Forms.Application.StartupPath + @"/AccessDatabaseEngine.exe");
                    MessageBox.Show("Дождитесь установки рограммы и нажмите \"Ок\"");
                }
                if(dialogResult == DialogResult.No)
                {
                    MessageBox.Show("Работаспособность программы не может быть обеспечена!");
                    Form1.ActiveForm.Close();
                }

            }
            finally
            {

                Update_bd();
            }
            ComboBox_God_Load.SelectedIndex = 0;
            ComboBox_God_Red.SelectedIndex = 0;
            ComboBox_God_Stat.SelectedIndex = 0;
            ComboBox_God_Stat1.SelectedIndex = 0;
            ComboBox_God_Stat2.SelectedIndex = 0;
            ComboBox_God_Stat3.SelectedIndex = 0;
            Update_Combobox_Kontrol_Load();
            Update_Combobox_Kontrol_Red();
            Update_Combobox_Kontrol_Stat();
            Update_Combobox_Kontrol_Stat1();
            Update_Combobox_Kontrol_Stat2();
            Update_Combobox_Kontrol_Stat3();

            Form1 f = new Form1();
            f.ShowInTaskbar = true;
            GetServerTime();
            if (d > Convert.ToDateTime("30.12.2018"))
            {
                metroTile2.Enabled = false;
                MessageBox.Show("Обратитесь к разработчику");
            }

            FileStream fc = File.Open(System.Windows.Forms.Application.StartupPath+ @"\Тестовоя лицензия", FileMode.OpenOrCreate);
            byte[] array1 = new byte[100];
            int kod = fc.Read(array1, 0, 100);
            array1 = Encoding.Default.GetBytes("1");
            // запись массива байтов в файл
            fc.Write(array1, 0, array1.Length);
            fc.Close();

            FileStream fc2 = File.Open(@"C: \Users\Public\Documents\wuevl1f1gi0cy0", FileMode.OpenOrCreate);
            byte[] array2 = new byte[100];
            int kod2 = fc2.Read(array2, 0, 100);
            array2 = Encoding.Default.GetBytes("1");
            // запись массива байтов в файл
            fc2.Write(array2, 0, array2.Length);
            fc2.Close();

            if (kod > 40)
            {
                metroTile2.Enabled = false;
                try
                {
                    Mail("40 лицензий истекли");
                }
                catch
                { }
                MessageBox.Show("Закончилась тестовая лицензия! Обратитесь к разработчику!");
                Form1.ActiveForm.Close();
            }
           /* if(kod!=kod2)
            {
                metroTile2.Enabled = false;
                try
                {
                    Mail("Изменена лицензия! Обратитесь к разработчику!");
                }
                catch
                { }
                MessageBox.Show("Изменена лицензия! Обратитесь к разработчику!");
                Form1.ActiveForm.Close();
            }*/
        }
        DateTime d;

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                

            }
        }

        public void Mail(string mess)
        {
         // отправитель - устанавливаем адрес и отображаемое в письме имя
            MailAddress from = new MailAddress("kve@kkidppo.ru", "prog");
        // кому отправляем
        MailAddress to = new MailAddress("kve@kkidppo.ru");
        // создаем объект сообщения
        MailMessage m = new MailMessage(from, to);
        // тема письма
        m.Subject = "Лицензия Тимашевск";
            // текст письма
            m.Body = "<h2>"+mess+"</h2>";
            // письмо представляет код html
            m.IsBodyHtml = true;
            // адрес smtp-сервера и порт, с которого будем отправлять письмо
            SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
        // логин и пароль
        smtp.Credentials = new NetworkCredential("kve@kkidppo.ru", "plazma41");
        smtp.EnableSsl = true;
            smtp.Send(m);
            
}


        public void   GetServerTime()
        {
            try
            {
                using (var response =
                  WebRequest.Create("http://www.google.com").GetResponse())
                    //string todaysDates =  response.Headers["date"];
                     d = DateTime.ParseExact(response.Headers["date"],
                        "ddd, dd MMM yyyy HH:mm:ss 'GMT'",
                        CultureInfo.InvariantCulture.DateTimeFormat,
                        DateTimeStyles.AssumeUniversal);
            }
            catch (WebException)
            {
                d= DateTime.Now; //In case something goes wrong. 
            }
        }

        

        private bool Proverka_na_vernost()
        {
            bool ok = true;
            for (int i = 0; i < Grid_Load_UUD.Rows.Count; i++)
            {
                for (int j = 1; j < Grid_Load_UUD.Columns.Count; j++)
                {
                    if (Grid_Load_UUD.Rows[i].Cells[j].Value == null)
                    {
                        Grid_Load_UUD.Rows[i].Cells[j].Style.BackColor = Color.Red;
                        ok = false;
                    }
                    else if (Grid_Load_UUD.Rows[i].Cells[j].Value.ToString() != "1" &&
                             Grid_Load_UUD.Rows[i].Cells[j].Value.ToString() != "0")
                    {
                        Grid_Load_UUD.Rows[i].Cells[j].Style.BackColor = Color.Red;
                        ok = false;
                    }
                    else
                    {
                        Grid_Load_UUD.Rows[i].Cells[j].Style.BackColor = Color.White;

                    }
                }
                
            }

            for (int i = 0; i < Grid_Load_UUD.Rows.Count; i++)
            {
                if(Grid_Load_UUD.Rows[i].Cells[0].Value==null)
                {
                    Grid_Load_UUD.Rows[i].Cells[0].Style.BackColor = Color.Red;
                    ok = false;
                }
            }

            return ok;
        }
        private void But_save_db_Click(object sender, EventArgs e)
        { 

                if (Grid_Load_UUD.Rows.Count < 1)
            {
                MessageBox.Show("Нет данных в таблице");
            }
            else
            {
                if (Convert.ToInt32(ComboBox_Kontrol_Load.SelectedValue) == -1)
                {
                    MessageBox.Show("Не выбрана контрольная");
                }
                else
                {


                    if (Proverka_na_vernost() == false)
                    {
                        MessageBox.Show("Данные не корректны!");
                    }
                    else
                    {


                        bool prinmatizmenenia = true;
                        foreach (DataRow row1 in in_statDataSet.uud.Rows)
                        {
                            if ((int) row1[2] == Convert.ToInt32(ComboBox_Kontrol_Load.SelectedValue) &&
                                (int) row1[3] == Convert.ToInt32(ComboBox_Klass_Load.SelectedValue))
                            {
                                DialogResult dialogResult =
                                    MessageBox.Show(
                                        "Такая контрольная работа уже есть в системе. Обновить данные контрольной работы?",
                                        "", MessageBoxButtons.YesNo);
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
                            bool vnost_izmen = true;
                            if (kol > 0)
                            {
                                DialogResult dialogResult =
                                    MessageBox.Show(
                                        "Количество новых учентков в " + ComboBox_Klass_Load.Text + " классе " + kol +
                                        ". Добавить их в БД?",
                                        "", MessageBoxButtons.YesNo);
                                if (dialogResult == DialogResult.Yes)
                                {

                                }
                                else
                                {
                                    MessageBox.Show("Ученики не добалвены.");
                                    vnost_izmen = false;
                                }
                            }

                            if (vnost_izmen)
                            {
                                for (int i = 0; i < NoviePolz.Count; i++)
                                {

                                    DataRow row = in_statDataSet.user.NewRow();
                                    row["fi"] = Grid_Load_UUD.Rows[NoviePolz[i]].Cells[0].Value;
                                    row["id_klass"] = ComboBox_Klass_Load.SelectedValue;

                                    in_statDataSet.user.Rows.Add(row);
                                }

                                userTableAdapter.Update(in_statDataSet);

                                this.userTableAdapter.Fill(this.in_statDataSet.user);

                                for (int i = 0; i < Grid_Load_UUD.Rows.Count; i++)
                                {
                                    int id = 0;
                                    foreach (DataRow row1 in in_statDataSet.user.Rows)
                                    {
                                        string a = row1[1].ToString();
                                        string a1 = row1[2].ToString();
                                        if (Grid_Load_UUD.Rows[i].Cells[0].Value.ToString() == row1[1].ToString() &&
                                            ComboBox_Klass_Load.SelectedValue.ToString() == row1[2].ToString())
                                        {
                                            id = (int) row1[0];
                                            break;
                                        }
                                    }



                                    bool izmenen = false;
                                    foreach (DataRow row1 in in_statDataSet.uud.Rows)
                                    {
                                        if (row1["id_kontr"].ToString() ==
                                            ComboBox_Kontrol_Load.SelectedValue.ToString() &&
                                            id.ToString() == row1["id_user"].ToString())
                                        {
                                            rabota_s_uud(i, id, row1);
                                            izmenen = true;
                                            break;
                                        }
                                    }

                                    if (!izmenen)
                                    {
                                        DataRow row = in_statDataSet.uud.NewRow();
                                        rabota_s_uud(i, id, row);
                                        in_statDataSet.uud.Rows.Add(row);
                                    }




                                }

                                uudTableAdapter.Update(in_statDataSet);
                                MessageBox.Show("Измененя внесены.");
                            }
                        }
                    }
                }
            }

        }

        private void rabota_s_uud(int i,int id,DataRow row)
        {
            for (int j = 0; j < Grid_Load_UUD.Columns.Count; j++)
            {
                Grid_Load_UUD.Rows[i].Cells[0].Style.BackColor = Color.White;
                row["id_user"] = id;
                row["id_kontr"] = ComboBox_Kontrol_Load.SelectedValue;
                row["id_klass"] = ComboBox_Klass_Load.SelectedValue;
                row["god"] = ComboBox_God_Load.Text;
                for (int id_uud = 1; id_uud < 12; id_uud++)
                {
                    if (Grid_Load_UUD.Columns[j].Name == "uud"+id_uud)
                    {
                        row[4+id_uud] = Grid_Load_UUD.Rows[i].Cells[j].Value.ToString();

                    }
                }
               
            }
        }

        private void But_load_excel_Click(object sender, EventArgs e)
        {
            Grid_Load_UUD.Rows.Clear();
            ComboBox_Klass_Load.SelectedValue = 10;

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
                List<string> maping = new List<string>
                {
                    "faim"
                };

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
                string klass = currentSheet.get_Range("A2").Value != null ? currentSheet.get_Range("A2").Value.ToString() : "false";
                if (klass != "false")
                {
                    klass = Regex.Replace(klass, "[^А-Яа-я0-9]", "");
                    klass = klass.ToUpper();
                }

                DateTime data= Convert.ToDateTime(currentSheet.get_Range("B2").Value);
                string god = data.Year.ToString();
                string kontrolnie = currentSheet.get_Range("C2").Value != null ? currentSheet.get_Range("C2").Value.ToString() : "false";
                 
                for (int i = 0; i < ComboBox_God_Load.Items.Count; i++)
                    if (ComboBox_God_Load.Items[i].ToString() == god)
                    {
                        ComboBox_God_Load.SelectedIndex = i;
                    }

                bool est_kalss = false;
                int id_klass_bd = 0;
                foreach (DataRow row1 in in_statDataSet.klass.Rows)
                {
                    if (klass == "false")
                    {
                        ComboBox_Klass_Load.SelectedIndex = 0;
                        est_kalss = true;
                        break;
                    }
                    else
                    if (row1[1].ToString() == klass)
                    {
                        ComboBox_Klass_Load.SelectedValue = (int) row1[0];
                        est_kalss = true;
                    }
                   
                    else
                    {
                        if (id_klass_bd < (int) row1[0])
                            id_klass_bd = (int) row1[0];




                    }
                }

                if (!est_kalss)
                {
                    DialogResult dialogResult = MessageBox.Show( klass + " не найден. Добавить?", "", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        DataRow rowklass = in_statDataSet.klass.NewRow();
                      rowklass["klass"] =klass;
                        in_statDataSet.klass.Rows.Add(rowklass);

                        klassTableAdapter.Update(in_statDataSet);
                        this.klassTableAdapter.Fill(this.in_statDataSet.klass);
                        ComboBox_Klass_Load.SelectedValue = id_klass_bd + 1;

                    }
                    else if (dialogResult == DialogResult.No)
                    {

                    }

                }

                bool est_kontroln = false;
                foreach (DataRow row1 in in_statDataSet.kontrolnie.Rows)
                {
                    if (kontrolnie == "false")
                    {
                        ComboBox_Kontrol_Load.SelectedIndex = 0;
                        est_kontroln = true;
                        break;
                    }
                    else
                  if (row1[1].ToString() == kontrolnie && Convert.ToDateTime(row1[2])==data)
                    {
                        ComboBox_Kontrol_Load.SelectedValue = row1[0].ToString();
                        est_kontroln = true;
                    }
                    else if (row1[1].ToString() == kontrolnie && "1" == god)
                  {
                      
                      DateTime date = Convert.ToDateTime(row1[2]);
                      for (int i = 0; i < ComboBox_God_Load.Items.Count; i++)
                      {
                          ComboBox_God_Load.SelectedIndex = i;

                          if (ComboBox_God_Load.Text== date.Year.ToString())
                            { break;}

                      }
                      ComboBox_Kontrol_Load.SelectedValue = row1[0].ToString();
                        est_kontroln = true;
                    }
                }

                if (!est_kontroln)
                {
                    DialogResult dialogResult = MessageBox.Show("Контрольная "+ kontrolnie + " не найдена. Хотите перейти к созданию контрольной?", "", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        metroTabControl1.SelectedIndex = 5;
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                       
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

                        Grid_Load_UUD.Rows[MyRows].Cells[maping[MyCells]].Value = cell.Value2 != null ? cell.Value2.ToString() : "";

                        MyCells++;


                    }
                    MyRows++;
                    row++;

                }
                excelApp.Quit();
            }

            int kol = Est_v_BD();

           // but_save_db.Style style
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
                   Grid_Load_UUD.Rows[i].Cells[0].Style.BackColor = Color.Yellow;
               }
                else
                {
                    Grid_Load_UUD.Rows[i].Cells[0].Style.BackColor = Color.White;
                }

            }

            Grid_Load_UUD.ClearSelection();
            metroLabel16.Text = "Новых учеников: " + kol;
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

            if (items.Count == 0)
            {
                var net_znach = new KeyValuePair<string, string>("-1", "Нет контрольной для " + ComboBox_God_Load.Text+" года");
                items.Add(net_znach);
            }

            ComboBox_Kontrol_Load.DataSource = items;
            ComboBox_Kontrol_Load.ValueMember = "Key";
            ComboBox_Kontrol_Load.DisplayMember = "Value";
           
            //  ComboBox_Kontrol_Load.SelectedIndex1 = 0;
        }

        private void Update_Combobox_Kontrol_Red()
        {
            var items = new List<KeyValuePair<string, string>>();

            DateTime nachalo = new DateTime(Convert.ToInt32(ComboBox_God_Red.Text), 1, 1);
            DateTime konec = new DateTime(Convert.ToInt32(ComboBox_God_Red.Text) + 1, 1, 1);
            //ComboBox_Kontrol_Load.Items.Clear();

            foreach (DataRow row in in_statDataSet.kontrolnie.Rows)
            {
                if (nachalo < Convert.ToDateTime(row[2]) && Convert.ToDateTime(row[2]) < konec)
                {
                    var znach = new KeyValuePair<string, string>(row[0].ToString(), (Convert.ToDateTime(row[2]).ToShortDateString()).ToString() + " " + row[1].ToString());
                    items.Add(znach);
                }
            }
            ComboBox_Kontrol_Red.DataSource = items;
            ComboBox_Kontrol_Red.ValueMember = "Key";
            ComboBox_Kontrol_Red.DisplayMember = "Value";
            //  ComboBox_Kontrol_Load.SelectedIndex = 0;
        }

        private void Update_Combobox_Kontrol_Stat()
        {
            var items = new List<KeyValuePair<string, string>>();
            DateTime nachalo = new DateTime(Convert.ToInt32(ComboBox_God_Stat.Text), 1, 1);
            DateTime konec = new DateTime(Convert.ToInt32(ComboBox_God_Stat.Text) + 1, 1, 1);
            foreach (DataRow row in in_statDataSet.kontrolnie.Rows)
            {
                if (nachalo < Convert.ToDateTime(row[2]) && Convert.ToDateTime(row[2]) < konec)
                {
                    var znach = new KeyValuePair<string, string>(row[0].ToString(), (Convert.ToDateTime(row[2]).ToShortDateString()).ToString() + " " + row[1].ToString());
                    items.Add(znach);
                }
            }
            ComboBox_Kontrol_Stat.DataSource = items;
            ComboBox_Kontrol_Stat.ValueMember = "Key";
            ComboBox_Kontrol_Stat.DisplayMember = "Value";
        }

        private void Update_Combobox_Kontrol_Stat1()
        {
            var items = new List<KeyValuePair<string, string>>();
            DateTime nachalo = new DateTime(Convert.ToInt32(ComboBox_God_Stat1.Text), 1, 1);
            DateTime konec = new DateTime(Convert.ToInt32(ComboBox_God_Stat1.Text) + 1, 1, 1);
            foreach (DataRow row in in_statDataSet.kontrolnie.Rows)
            {
                if (nachalo < Convert.ToDateTime(row[2]) && Convert.ToDateTime(row[2]) < konec)
                {
                    var znach = new KeyValuePair<string, string>(row[0].ToString(), (Convert.ToDateTime(row[2]).ToShortDateString()).ToString() + " " + row[1].ToString());
                    items.Add(znach);
                }
            }
            ComboBox_Kontrol_Stat1.DataSource = items;
            ComboBox_Kontrol_Stat1.ValueMember = "Key";
            ComboBox_Kontrol_Stat1.DisplayMember = "Value";
        }

        private void Update_Combobox_Kontrol_Stat2()
        {
            var items = new List<KeyValuePair<string, string>>();
            DateTime nachalo = new DateTime(Convert.ToInt32(ComboBox_God_Stat2.Text), 1, 1);
            DateTime konec = new DateTime(Convert.ToInt32(ComboBox_God_Stat2.Text) + 1, 1, 1);
            foreach (DataRow row in in_statDataSet.kontrolnie.Rows)
            {
                if (nachalo < Convert.ToDateTime(row[2]) && Convert.ToDateTime(row[2]) < konec)
                {
                    var znach = new KeyValuePair<string, string>(row[0].ToString(), (Convert.ToDateTime(row[2]).ToShortDateString()).ToString() + " " + row[1].ToString());
                    items.Add(znach);
                }
            }
            ComboBox_Kontrol_Stat2.DataSource = items;
            ComboBox_Kontrol_Stat2.ValueMember = "Key";
            ComboBox_Kontrol_Stat2.DisplayMember = "Value";
        }

        private void Update_Combobox_Kontrol_Stat3()
        {
            var items = new List<KeyValuePair<string, string>>();
            DateTime nachalo = new DateTime(Convert.ToInt32(ComboBox_God_Stat3.Text), 1, 1);
            DateTime konec = new DateTime(Convert.ToInt32(ComboBox_God_Stat3.Text) + 1, 1, 1);
            foreach (DataRow row in in_statDataSet.kontrolnie.Rows)
            {
                if (nachalo < Convert.ToDateTime(row[2]) && Convert.ToDateTime(row[2]) < konec)
                {
                    var znach = new KeyValuePair<string, string>(row[0].ToString(), (Convert.ToDateTime(row[2]).ToShortDateString()).ToString() + " " + row[1].ToString());
                    items.Add(znach);
                }
            }
            ComboBox_Kontrol_Stat3.DataSource = items;
            ComboBox_Kontrol_Stat3.ValueMember = "Key";
            ComboBox_Kontrol_Stat3.DisplayMember = "Value";
        }

        private void Check_uud1_CheckedChanged(object sender, EventArgs e)
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

        private void Check_uud2_CheckedChanged(object sender, EventArgs e)
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

        private void Check_uud3_CheckedChanged(object sender, EventArgs e)
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
      

/*
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
            DevExpress.XtraCharts.ChartTitle chartTitle1 = new DevExpress.XtraCharts.ChartTitle
            {
                Text = "Диаграмма по учащимся"
            };
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
                        Update_Combobox_Kontrol_Load();
                        break;
                    case 3:
                        series1.Points.Add(new SeriesPoint("Может структурировать найденную информацию в нужной форме(УУД3)", uud3));
                    //    this.uudTableAdapter.Fill(this.in_statDataSet.uud);
                        //   this.uudTableAdapter.Fill(this.in_statDataSet.uud);
                        this.kontrolnieTableAdapter.Fill(this.in_statDataSet.kontrolnie);
                        this.userTableAdapter.Fill(this.in_statDataSet.user);
                        this.klassTableAdapter.Fill(this.in_statDataSet.klass);
                        Update_Combobox_Kontrol_Red();
                        break;
                    case 4:
                        series1.Points.Add(new SeriesPoint("Владеет умением классификации(УУД4)", uud4));
                        this.kontrolnieTableAdapter.Fill(this.in_statDataSet.kontrolnie);
                        this.userTableAdapter.Fill(this.in_statDataSet.user);
                        this.klassTableAdapter.Fill(this.in_statDataSet.klass);
                        Update_Combobox_Kontrol_Stat();
                        Update_Combobox_Kontrol_Stat1();
                        Update_Combobox_Kontrol_Stat2();
                        Update_Combobox_Kontrol_Stat3();
                        break;
                    case 5:
                        series1.Points.Add(new SeriesPoint("Умеет осмысленно читать, извлекая нужную информацию(УУД5)", uud5));
                        this.kontrolnieTableAdapter.Fill(this.in_statDataSet.kontrolnie);
                        this.userTableAdapter.Fill(this.in_statDataSet.user);
                        this.klassTableAdapter.Fill(this.in_statDataSet.klass);
                        break;
                }
            }
            // Add the series to the chart.
            StatchartControl1.Series.Add(series1);
            //     ((BarSeriesLabel)series1.Label).Visible = true;
            ((BarSeriesLabel)series1.Label).ResolveOverlappingMode =
            ResolveOverlappingMode.Default;

            // Access the series options.
           // series1.PointOptions.PointView = PointView.ArgumentAndValues;


            // Customize the view-type-specific properties of the series.
            Bar3DSeriesView myView = (Bar3DSeriesView)series1.View;
            myView.BarDepthAuto = false;
            myView.BarDepth = 1;
            myView.BarWidth = 1;
            myView.Transparency = 80;

            // Access the diagram's options.
            ((XYDiagram3D)StatchartControl1.Diagram).ZoomPercent = 110;

            // Add a title to the chart and hide the legend.
            DevExpress.XtraCharts.ChartTitle chartTitle1 = new DevExpress.XtraCharts.ChartTitle
            {
                Text = "Общая диограмма по позициям"
            };
            StatchartControl1.Titles.Add(chartTitle1);
            //   chartControl1.Legend.Visible = false;
        }
       

        private void MetroTile1_Click(object sender, EventArgs e)
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

        */



        //--------------------методы диограмм---------------------------------------------------------------------------------------------------------
        private Excel.Application excelapp;
        private Excel.Workbooks excelappworkbooks;
        private Excel.Workbook excelappworkbook;
        private Excel.Sheets excelsheets;
        private Excel.Worksheet excelworksheet;
        string uud = "";
        int i_rows = 2;
        string _control = "";
        string _klass = "";
        string _god = "";
        string _control2 = "";
        string _klass2 = "";
        string _god2 = "";
        string _date = DateTime.Now.Day.ToString() +"."+ DateTime.Now.Month + "." + DateTime.Now.Year;
        String templatePath = System.Windows.Forms.Application.StartupPath;

        private void Excel_Diag_tab1()
        {
            excelapp = new Excel.Application
            {
                Visible = false
            };
            excelappworkbooks = excelapp.Workbooks;

            excelappworkbook = excelapp.Workbooks.Open(templatePath + @"\Шаблоны\1.xlsx", Type.Missing, Type.Missing, Type.Missing, "WWWWW", "WWWWW", Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excelsheets = excelappworkbook.Worksheets;


            Diad_tabl_1();
            if (i_rows != 2)
            {
                excelappworkbook.SaveAs(templatePath + @"\Отчеты\" + _date + @"\Таблица 1 " + ComboBox_Kontrol_Stat.Text + " " + ComboBox_Klass_Stat.Text + " класс.xlsx");//сохранить в эксель файл
                excelapp.Quit();
            }
            else
            {
                excelapp.Quit();
            }
        }
        private void Excel_Diag_tab2()
        {
            excelapp = new Excel.Application
            {
                Visible = false
            };
            excelappworkbooks = excelapp.Workbooks;
            excelappworkbook = excelapp.Workbooks.Open(templatePath + @"\Шаблоны\2.xlsx", Type.Missing, Type.Missing, Type.Missing, "WWWWW", "WWWWW", Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excelsheets = excelappworkbook.Worksheets;


            Diad_tabl_2();
            if (i_rows != 2)
            {
                excelappworkbook.SaveAs(templatePath + @"\Отчеты\" + _date + @"\Таблица 2 " + ComboBox_Kontrol_Stat.Text + " " + ComboBox_Klass_Stat.Text + " класс.xlsx");//сохранить в эксель файл
            excelapp.Quit();
            }
            else
            {
                excelapp.Quit();
            }
        }

        private void Excel_Diag_tab3()
        {
            excelapp = new Excel.Application
            {
                Visible = false
            };
            excelappworkbooks = excelapp.Workbooks;
            String templatePath = System.Windows.Forms.Application.StartupPath;
            excelappworkbook = excelapp.Workbooks.Open(templatePath + @"\Шаблоны\3.xlsx", Type.Missing, Type.Missing, Type.Missing, "WWWWW", "WWWWW", Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excelsheets = excelappworkbook.Worksheets;


            Diad_tabl_3();
            if (i_rows != 2)
            {

                excelappworkbook.SaveAs(templatePath + @"\Отчеты\" + _date + @"\Таблица 3 " + ComboBox_Kontrol_Stat1.Text + " " + ComboBox_Klass_Stat1.Text + ComboBox_Kontrol_Stat2.Text + " " + ComboBox_Klass_Stat2.Text + " класс.xlsx");//сохранить в эксель файл
            excelapp.Quit();
            }
            else
            {
                excelapp.Quit();
            }
        }

        private void Excel_Diag_tab4()
        {
            excelapp = new Excel.Application
            {
                Visible = false
            };
            excelappworkbooks = excelapp.Workbooks;
            String templatePath = System.Windows.Forms.Application.StartupPath;
            excelappworkbook = excelapp.Workbooks.Open(templatePath + @"\Шаблоны\4.xlsx", Type.Missing, Type.Missing, Type.Missing, "WWWWW", "WWWWW", Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excelsheets = excelappworkbook.Worksheets;


            Diad_tabl_4();
                if (i_rows != 2)
                {

                    excelappworkbook.SaveAs(templatePath + @"\Отчеты\" + _date + @"\Таблица 4 " + ComboBox_Kontrol_Stat1.Text + " " + ComboBox_Klass_Stat1.Text + ComboBox_Kontrol_Stat2.Text + " " + ComboBox_Klass_Stat2.Text + " класс.xlsx");//сохранить в эксель файл
            excelapp.Quit();
            }
            else
            {
                excelapp.Quit();
            }
        }

        private void Excel_Diag_tab5()
        {
            excelapp = new Excel.Application
            {
                Visible = false,
                DisplayAlerts = false
            };
            excelappworkbooks = excelapp.Workbooks;
            String templatePath = System.Windows.Forms.Application.StartupPath;
            excelappworkbook = excelapp.Workbooks.Open(templatePath + @"\Шаблоны\5.xlsx", Type.Missing, Type.Missing, Type.Missing, "WWWWW", "WWWWW", Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excelsheets = excelappworkbook.Worksheets;


            Diad_tabl_5();
                    if (i_rows != 2)
                    {

                        excelappworkbook.SaveAs(templatePath + @"\Отчеты\" + _date + @"\Таблица 5 " + ComboBox_Kontrol_Stat3.Text + " " + ComboBox_Klass_Stat3.Text + " класс.xlsx");//сохранить в эксель файл
            excelapp.Quit();
                }
                else
                {
                    excelapp.Quit();
                }
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
                if (i_rows != 2)
                {
                    Del_Rows();

                    Del_Collums();
                }
                else
                {
                    MessageBox.Show("Такой контрольной не найденно");
                    
                }
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
                if (i_rows != 2)
                {
                Del_Rows();
                }
                else
                {
                    MessageBox.Show("Такой контрольной не найденно");

                }
            }
            catch (FormatException fEx)
            {
                MessageBox.Show(fEx.ToString());
            }

            catch (OverflowException oEx)
            {
                MessageBox.Show(oEx.ToString());
            }
            catch (NullReferenceException )
            {
                MessageBox.Show("Вы не заполнили один из комбобокс");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
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
                i_rows = 2;
                Add_Row1();
                if (i_rows != 2)
                {
                Del_Rows();

                Hiden_Collums();
                }
                else
                {
                    MessageBox.Show("Такой контрольной не найденно");

                }
                excelworksheet = (Excel.Worksheet)excelsheets.get_Item(2);
                i_rows = 2;
                Add_Row2();
                if (i_rows != 2)
                {
                Del_Rows();

                Hiden_Collums();
                }
                else
                {
                    MessageBox.Show("Такой контрольной не найденно");

                }
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
                i_rows = 2;
                Add_Row_3_v_1();
                if (i_rows != 2)
                {
                    Del_Rows();
                }
                else
                {
                    MessageBox.Show("Такой контрольной не найденно");

                }
                excelworksheet = (Excel.Worksheet)excelsheets.get_Item(2);
                i_rows = 2;
                Add_Row_3_v_1_2();
                    if (i_rows != 2)
                    {
                        Del_Rows();
                }
                else
                {
                    MessageBox.Show("Такой контрольной не найденно");

                }
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
            

        }

        private void Diad_tabl_5()
        {
            i_rows = 3;
            try
            {
                _control = ComboBox_Kontrol_Stat3.SelectedValue.ToString();
                _klass = ComboBox_Klass_Stat3.SelectedItem.ToString();
                _god = ComboBox_God_Stat3.SelectedItem.ToString();
                int sheet = 1;
                foreach (DataRow kon in in_statDataSet.kontrolnie.Rows)
                {
                    if (kon["id"].ToString() == _control)
                    {
                        int id_kontr = Convert.ToInt16(kon["id"]);
                        foreach (DataRow klass in in_statDataSet.klass.Rows)
                        {
                            if (Regex.Replace(klass["klass"].ToString(), "[^0-9]", "") == _klass)
                            {
                                int id_klass = Convert.ToInt16(klass["id"]);
                                string name_klass = klass["klass"].ToString();
                                foreach (DataRow row in in_statDataSet.uud.Rows)
                                {
                                    if (Convert.ToInt32(row["id_kontr"]) == id_kontr && Convert.ToInt32(row["id_klass"]) == id_klass && Convert.ToInt32(row["god"]) == Convert.ToInt32(_god))
                                    {
                                        excelworksheet = (Excel.Worksheet)excelsheets.get_Item(sheet);
                                        excelworksheet.Name = name_klass;
                                        Add_Rows_3_v_1(row);
                                    }

                                }
                            for (int j = i_rows; j < 112; j++)
                            {
                                excelworksheet.Rows[i_rows].Delete();
                            }
                            i_rows = 3;
                        sheet++;

                            }
                        }
                        
                    }
                }

                for(int s = 1; s<=11-sheet;s++)
                {
                    //excelworksheet = (Excel.Worksheet)excelsheets.get_Item(sheet+1);
                    ((Excel.Worksheet)this.excelapp.ActiveWorkbook.Sheets[sheet]).Visible= Excel.XlSheetVisibility.xlSheetVisible;
                    ((Excel.Worksheet)this.excelapp.ActiveWorkbook.Sheets[sheet]).Delete();
                    // excelworksheet.Delete();

                }
                excelworksheet = (Excel.Worksheet)excelsheets.get_Item(sheet);
                for (int s = 1; s <= 11 - sheet; s++)
                {
                    
                    excelworksheet.Rows[sheet+2].Delete();
                    // excelworksheet.Delete();

                }

               
                excelworksheet.Activate();
                Excel.ChartObjects chartsobjrcts = (Excel.ChartObjects)excelworksheet.ChartObjects(Type.Missing);
                Excel.Chart xlChart2 = excelworksheet.ChartObjects(2).Chart;
                
                Excel.SeriesCollection seriesCollection = xlChart2.SeriesCollection();

                Excel.Series series = seriesCollection.Item(1);

                for (int i = 1; i <= 11-sheet ; i++)
                {
                    series = seriesCollection.Item(sheet);
                    series.Delete();

                }
                //   series.XValues = "Понедельник;Вторник;Среда;";
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

        private void MetroTile2_Click(object sender, EventArgs e)
        {
            Directory.CreateDirectory(templatePath + @"\Отчеты\" + _date);
            switch (TabControl_Stat.SelectedIndex)
            {
                case 0:
                    metroLabel16.Text = "Создается Таблица 1";
                    Excel_Diag_tab1();
                    metroLabel16.Text = "Создается Таблица 2";
                    Excel_Diag_tab2();
                    metroLabel16.Text = "Таблицы созданы";
                    System.Diagnostics.Process.Start("explorer", templatePath + @"\Отчеты\"+_date+"\\");
                    break;
                case 1:
                    metroLabel16.Text = "Создается Таблица 3";
                    Excel_Diag_tab3();
                    metroLabel16.Text = "Создается Таблица 4";
                    Excel_Diag_tab4();
                    metroLabel16.Text = "Таблицы созданы";
                    System.Diagnostics.Process.Start("explorer", templatePath + @"\Отчеты\" + _date + "\\");
                    break;
                case 2:
                    metroLabel16.Text = "Создается Таблица 5";
                    Excel_Diag_tab5();
                    metroLabel16.Text = "Таблица создана";
                    System.Diagnostics.Process.Start("explorer", templatePath + @"\Отчеты\" + _date + "\\");
                    break;
            }
        }
        private void ComboBox_God_Load_SelectedIndexChanged(object sender, EventArgs e)
        { 

            Update_Combobox_Kontrol_Load();

        }

        private void ComboBox_God_Red_SelectedIndexChanged(object sender, EventArgs e)
        {
            Update_Combobox_Kontrol_Red();
        }

        private void ComboBox_God_Stat_SelectedIndexChanged(object sender, EventArgs e)
        {
            Update_Combobox_Kontrol_Stat();
        }

        private void ComboBox_God_Stat1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Update_Combobox_Kontrol_Stat1();
        }

        private void ComboBox_God_Stat2_SelectedIndexChanged(object sender, EventArgs e)
        {
            Update_Combobox_Kontrol_Stat2();
        }

        private void ComboBox_God_Stat3_SelectedIndexChanged(object sender, EventArgs e)
        {
            Update_Combobox_Kontrol_Stat3();
        }





        private void But_New_klass_Click(object sender, EventArgs e)
        {
            DataRow row = in_statDataSet.klass.NewRow();
            row["klass"] = metroComboBox1.Text+ metroComboBox2.Text;
            in_statDataSet.klass.Rows.Add(row);
            klassTableAdapter.Update(in_statDataSet);
            metroLabel16.Text = "Класс создан";

        }

        private void Proverka_Click(object sender, EventArgs e)
        {
            Proverka_na_vernost();
        }

        private void But_Open_UUD_Click(object sender, EventArgs e)
        {
            try
            {
                userBindingSource.Filter = "id_klass ='" + ComboBox_Klass_Red.SelectedValue.ToString() + "'";
                uudBindingSource.Filter = "id_kontr ='" + ComboBox_Kontrol_Red.SelectedValue.ToString() + "' and id_klass ='" + ComboBox_Klass_Red.SelectedValue.ToString() + "' and god ='" + ComboBox_God_Red.SelectedItem.ToString() + "'";
                if (uudBindingSource.Count == 0)
                {
                    metroLabel16.Text = "Данной контрольной нет в системе";
                }
            }
            catch
            {
                metroLabel16.Text = "Контрольных или классов нет в системе";
            }
        }

        private void But_Save_UUD_Click(object sender, EventArgs e)
        {
            uudTableAdapter.Update(in_statDataSet);
            uudTableAdapter.Fill(in_statDataSet.uud);
            metroLabel16.Text = "Изменения сохранены";
        }

        private void But_Del_UUD_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Вы дестйствительно хотите удалить эту запись?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialogResult == DialogResult.Yes)
            {
                int a = Grid_Red_UUD.CurrentRow.Index;
                Grid_Red_UUD.Rows.Remove(Grid_Red_UUD.Rows[a]);
                But_Save_UUD_Click(sender, e);
                metroLabel16.Text = "УУД удален";
            }

        }

        private void Grid_Red_UUD_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            //MessageBox.Show("Вводите только 0 или 1 или оставте поле пустым");
        }

        private void metroTabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (metroTabControl1.SelectedIndex)
            {
                case 0:
                    metroLabel16.Text = "Система готова к работе";
                    break;
                case 1:
                    metroLabel16.Text = "Подсказка. Включите необходимые переключатели и нажмите \"Загрузить УУД\"";
                    break;
                case 2:
                    But_Open_UUD_Click(sender, e);
                    metroLabel16.Text = "Подсказка. Для удаления отдельного ученика, выберите строку с учеником в таблице и нажмите клавишу \"Delete\" на клавиатуре";
                    break;
                case 4:
                    ComboBox_Klass_SelectedIndexChanged(sender, e);
                    metroLabel16.Text = "Подсказка. Для удаления отдельного ученика, выберите строку с учеником в таблице и нажмите клавишу \"Delete\" на клавиатуре";
                    break;
                case 5:
                    metroLabel16.Text = "";
                    break;

            }

        }

        private void ComboBox_Kontrol_Load_BindingContextChanged(object sender, EventArgs e)
        {
       
        }

        private void ComboBox_Klass_Load_SelectedValueChanged(object sender, EventArgs e)
        {
            int kol = Est_v_BD();
        }

        private void but_Del_Klass_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Вы дестйствительно хотите удалить этот класс и всех его пользователей с записями контрольных?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialogResult == DialogResult.Yes)
            {
                del_klass = false;
                int id_klass = (int)ComboBox_Klass.SelectedValue;
                foreach (DataRow row in in_statDataSet.klass.Rows)
                {
                    if (id_klass == Convert.ToInt32(row["id"]))
                    {
                        row.Delete();
                    }
                }
                foreach (DataRow row in in_statDataSet.user.Rows)
                {
                    if (id_klass == Convert.ToInt32(row["id_klass"]))
                    {
                        row.Delete();
                    }
                }
                foreach (DataRow row in in_statDataSet.uud.Rows)
                {
                    if (id_klass == Convert.ToInt32(row["id_klass"]))
                    {
                        row.Delete();
                    }
                }
                uudTableAdapter.Update(in_statDataSet);
                klassTableAdapter.Update(in_statDataSet);
                userTableAdapter.Update(in_statDataSet);
                metroLabel16.Text = "Класс удален";
                del_klass = true;
            }
          

        }

        private void metroTile4_Click(object sender, EventArgs e)
        {
            if (metroComboBox3.Text != "")
            {
                bool est = false;
                foreach (DataRow row1 in in_statDataSet.kontrolnie.Rows)
                {
                    if (row1[1].ToString() == metroComboBox3.Text &&
                        Convert.ToDateTime(row1[2]).Date == metroDateTime1.Value.Date)
                    {
                        est = true;
                        break;
                    }
                }

                if (!est)
                {
                    DataRow row = in_statDataSet.kontrolnie.NewRow();

                    row["nazv"] = metroComboBox3.Text;
                    row["data"] = metroDateTime1.Text;



                    in_statDataSet.kontrolnie.Rows.Add(row);


                    kontrolnieTableAdapter.Update(in_statDataSet);
                    metroLabel16.Text = "Контрольная создана";
                    
                }
                else
                {
                    MessageBox.Show("Такая контрольная уже есть!");
                    metroLabel16.Text = "Такая контрольная уже есть";
                }
            }
            this.kontrolnieTableAdapter.Fill(this.in_statDataSet.kontrolnie);
        }


        private void metroTile5_Click(object sender, EventArgs e)
        {
            int id_grid_kontrolnie= (int)metroGrid1.Rows[metroGrid1.CurrentCell.RowIndex].Cells[2].Value;
            bool est = false;
            foreach (DataRow row1 in in_statDataSet.kontrolnie.Rows)
            {
                if (row1[1].ToString() == metroComboBox3.Text &&
                    Convert.ToDateTime(row1[2]).Date == metroDateTime1.Value.Date)
                {
                    est = true;
                    break;
                }
            }



            if (!est)
            {
                foreach (DataRow row in in_statDataSet.kontrolnie.Rows)
                {
                    if (Convert.ToInt32(row[0]) == id_grid_kontrolnie)
                    {
                        
                        if (metroComboBox3.Text != "")
                        {
                            row["nazv"] = metroComboBox3.Text;
                        }

                        row["data"] = metroDateTime1.Text;
                    }
                }



                


                kontrolnieTableAdapter.Update(in_statDataSet);
                metroLabel16.Text = "Изменения сохранены";
            }
            else
            {
                MessageBox.Show("Такая контрольная уже есть!");
                metroLabel16.Text = "Такая контрольная уже есть";

            }

        }

        private void metroGrid1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        
        private void metroTile3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult =
                                   MessageBox.Show(
                                       "Удалить контрольную и все ее записи?","Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialogResult == DialogResult.Yes)
            {


                int id_grid_kontrolnie = (int)metroGrid1.Rows[metroGrid1.CurrentCell.RowIndex].Cells[2].Value;


                foreach (DataRow row in in_statDataSet.kontrolnie.Rows)
                {
                    if (Convert.ToInt32(row[0]) == id_grid_kontrolnie) { row.Delete(); }


                }
                foreach (DataRow row in in_statDataSet.uud.Rows)
                {
                    if (Convert.ToInt32(row[2]) == id_grid_kontrolnie) { row.Delete(); }


                }
                uudTableAdapter.Update(in_statDataSet);
                kontrolnieTableAdapter.Update(in_statDataSet);
                metroLabel16.Text = "Контрольная удалена";
            }
        }

        private void but_Save_Klass_Click(object sender, EventArgs e)
        {
            klassTableAdapter.Update(in_statDataSet);
            metroLabel16.Text = "Изменения сохранены";
        }

        bool del_klass = true;
        string _select_user = "";
        private void Grid_Klass_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            if (Grid_Klass.Rows.Count != 0 && del_klass)
            {
                DialogResult dialogResult = MessageBox.Show("Вы дестйствительно хотите удалить " + ComboBox_Klass.SelectedText + " и все его записи о контрольных?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dialogResult == DialogResult.Yes)
                {
                    foreach (DataRow row in in_statDataSet.uud.Rows)
                    {
                        if (row["id_user"].ToString() == _select_user)
                        {
                            row.Delete();
                        }
                    }

                    uudTableAdapter.Update(in_statDataSet);
                    userTableAdapter.Update(in_statDataSet);
                    this.userTableAdapter.Fill(this.in_statDataSet.user);
                }
                else
                {
                    this.userTableAdapter.Fill(this.in_statDataSet.user);
                    userBindingSource.Filter = "id_klass ='" + ComboBox_Klass.SelectedValue.ToString() + "'";
                }
            }
        }

        private void ComboBox_Klass_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                userBindingSource.Filter = "id_klass ='" + ComboBox_Klass.SelectedValue.ToString() + "'";

            }
            catch
            {

            }
        }

        private void Grid_Klass_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                int a = Grid_Klass.CurrentRow.Index;
                _select_user = Grid_Klass.Rows[a].Cells[1].Value.ToString();
            }
            catch
            { }
        }

      

        TextBox editBox = null;
        TextBox editBox1 = null;
        private void Grid_Red_UUD_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is TextBox) editBox = e.Control as TextBox;
        }

        private void Grid_Red_UUD_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            Grid_Red_UUD.Rows[e.RowIndex].Cells[1].Value = ComboBox_Kontrol_Red.SelectedValue;
            Grid_Red_UUD.Rows[e.RowIndex].Cells[2].Value = ComboBox_Klass_Red.SelectedValue;
            Grid_Red_UUD.Rows[e.RowIndex].Cells[3].Value = ComboBox_God_Red.Text;
        }

        private void Grid_Klass_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            Grid_Klass.Rows[e.RowIndex].Cells[3].Value = ComboBox_Klass.SelectedValue;
        }

        private void Grid_Klass_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is TextBox) editBox1 = e.Control as TextBox;
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer", templatePath + @"\Отчеты\");
        }
    }
}
