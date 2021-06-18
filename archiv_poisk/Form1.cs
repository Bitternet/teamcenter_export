using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.Collections;

namespace archiv_poisk
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            /*
            dataGridView1.ColumnCount = 7;
            dataGridView1.ColumnHeadersVisible = true;


           DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();

            columnHeaderStyle.BackColor = Color.Beige;
            columnHeaderStyle.Font = new Font("Verdana", 10, FontStyle.Bold);
            dataGridView1.ColumnHeadersDefaultCellStyle = columnHeaderStyle;

            dataGridView1.Columns[0].Name = "Дец. номер";
            dataGridView1.Columns[0].Width = 150;
            dataGridView1.Columns[1].Name = "Изменение";
            dataGridView1.Columns[1].Width = 110;
            dataGridView1.Columns[2].Name = "Лист";
            dataGridView1.Columns[2].Width = 70;
            dataGridView1.Columns[3].Name = "Имя файла";
            dataGridView1.Columns[3].Width = 120;
            dataGridView1.Columns[4].Name = "Путь";
            dataGridView1.Columns[4].Width = 300;
            dataGridView1.Columns[5].Name = "Применяемость";
            dataGridView1.Columns[5].Width = 150;
            dataGridView1.Columns[6].Name = "Изделие";
            dataGridView1.Columns[6].Width = 150;*/

            addItem();

          /*  string connection = @"Data Source = 192.168.12.190; Initial Catalog = archiv; User ID = sa; Password = Sql12345678;";
            SqlConnection conn = new SqlConnection(connection);

            try
            {
                conn.Open();
                label2.ForeColor = System.Drawing.Color.Green;
                label2.Text = "Активно";
            }
            catch (SqlException ex)
            {
                label2.ForeColor = System.Drawing.Color.Red;
                label2.Text = "Ошибка";
            }
            finally
            {
                conn.Close();
            }*/
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dse_archive_000();
            dse_000();
            reviziya_backup();

            /*var fb = new FolderBrowserDialog();
            DialogResult result = fb.ShowDialog();

            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fb.SelectedPath))
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    string path = dataGridView1.Rows[i].Cells[4].Value.ToString();
                    string name = dataGridView1.Rows[i].Cells[0].Value.ToString();

                    if (File.Exists(fb.SelectedPath + "\\" + name) == false)
                    {
                        File.Copy(path, fb.SelectedPath +"\\"+ name, true);
                    }
                }
            }*/
        }

        private List<string> GetFolderList(string path)
        {
            List<string> dirList = new List<string>();
            string[] dirs = Directory.GetDirectories(path);
            foreach (string subdirectory in dirs)
            {
                dirList.Add(subdirectory);
                try
                {
                    dirList.AddRange(GetFolderList(subdirectory));
                }
                catch { }
            }
            return dirList;
        }

        public void insertrow(string name, string izm, string list, string data, string path)
        {
            string[] row = new string[] { name, izm, list, data, path + "\\" + name };
            dataGridView1.Rows.Add(row);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string connection = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source = localDB.mdb";//Provider=Microsoft.Jet.OLEDB.4.0=Microsoft.ACE.OLEDB.12.0

            string isdel = "";

            if (comboBox1.Text != "Выберите изделие:")
            {
                isdel = comboBox1.Text;


                if (radioButton1.Checked == true)
                {
                    string sql = "SELECT archiv.cod, archiv.izm, archiv.list, archiv.filename, archiv.path, archiv.prim, archiv.isdel  FROM archiv WHERE isdel ='" + isdel + "'";

                    using (OleDbConnection con = new OleDbConnection(connection))
                    {
                        con.Open();
                        // Создаем объект DataAdapter
                        OleDbDataAdapter adapter = new OleDbDataAdapter(sql, connection);
                        // Создаем объект Dataset
                        DataSet ds = new DataSet();
                        // Заполняем Dataset
                        adapter.Fill(ds);
                        // Отображаем данные
                        dataGridView1.DataSource = ds.Tables[0];

                        dataGridView1.Columns[0].HeaderText = "Дец. номер";
                        dataGridView1.Columns[0].Width = 150;
                        dataGridView1.Columns[1].HeaderText = "Изм.";
                        dataGridView1.Columns[1].Width = 80;
                        dataGridView1.Columns[2].HeaderText = "Лист";
                        dataGridView1.Columns[2].Width = 80;
                        dataGridView1.Columns[3].HeaderText = "Имя файла";
                        dataGridView1.Columns[3].Width = 150;
                        dataGridView1.Columns[4].HeaderText = "Путь";
                        dataGridView1.Columns[4].Width = 300;
                        dataGridView1.Columns[5].HeaderText = "Применяемость";
                        dataGridView1.Columns[5].Width = 150;
                        dataGridView1.Columns[6].HeaderText = "Изделие";
                        dataGridView1.Columns[6].Width = 100;


                        DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();

                        columnHeaderStyle.BackColor = Color.Beige;
                        columnHeaderStyle.Font = new Font("Verdana", 10, FontStyle.Bold);
                        dataGridView1.ColumnHeadersDefaultCellStyle = columnHeaderStyle;

                        label1.Text = dataGridView1.RowCount.ToString();
                    }
                }
                else if (radioButton2.Checked == true)
                {
                    string sql1 = "SELECT * FROM archiv";
                    MessageBox.Show(sql1);

                    using (OleDbConnection con = new OleDbConnection(connection))
                    {

                        con.Open();
                        // Создаем объект DataAdapter
                        OleDbDataAdapter adapter = new OleDbDataAdapter(sql1, connection);
                        // Создаем объект Dataset
                        DataSet ds = new DataSet();
                        // Заполняем Dataset
                        adapter.Fill(ds);
                        // Отображаем данные
                        dataGridView2.DataSource = ds.Tables[0];
                    }

                    string sql2 = "SELECT ismr.cod, ismr.ism, filestifi.list, filestifi.filename, filestifi.isv, ismr.isdel, filestifi.path0, isvr.prim, filestifi.dat FROM ismr INNER JOIN filestifi ON filestifi.isv = ismr.isv INNER JOIN isvr ON filestifi.isv = isvr.isv WHERE ismr.isdel = '" + textBox3.Text + "'";
                    MessageBox.Show(sql2);

                    using (OleDbConnection con = new OleDbConnection(connection))
                    {

                        con.Open();
                        // Создаем объект DataAdapter
                        OleDbDataAdapter adapter = new OleDbDataAdapter(sql2, connection);
                        // Создаем объект Dataset
                        DataSet ds = new DataSet();
                        // Заполняем Dataset
                        adapter.Fill(ds);
                        // Отображаем данные
                        dataGridView3.DataSource = ds.Tables[0];
                    }
                }
                else
                {
                    textBox3.Text = "Укажите тип поиска";
                }
            }
            else MessageBox.Show("Выберите изделие из списка!");
        }

    


        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "CSV (*.csv)|*.csv";
                sfd.FileName = "Отчет.csv";
                bool fileError = false;
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    if (File.Exists(sfd.FileName))
                    {
                        try
                        {
                            File.Delete(sfd.FileName);
                        }
                        catch (IOException ex)
                        {
                            fileError = true;
                            MessageBox.Show("Нет доступа к диску" + ex.Message);
                        }
                    }
                    if (!fileError)
                    {
                        try
                        {
                            int columnCount = dataGridView1.Columns.Count;
                            string columnNames = "";
                            string[] outputCsv = new string[dataGridView1.Rows.Count + 1];
                            for (int i = 0; i < columnCount; i++)
                            {
                                columnNames += dataGridView1.Columns[i].HeaderText.ToString() + ";";
                            }
                            outputCsv[0] += columnNames;

                            for (int i = 1; (i - 1) < dataGridView1.Rows.Count; i++)
                            {
                                for (int j = 0; j < columnCount; j++)
                                {
                                    outputCsv[i] += dataGridView1.Rows[i - 1].Cells[j].Value.ToString() + ";";
                                }
                            }

                            File.WriteAllLines(sfd.FileName, outputCsv, Encoding.UTF8);
                            MessageBox.Show("Сохранение выполнено успешно", "Инфо");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Ошибка :" + ex.Message);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Нет записей на экспорт", "Info");
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
             //.Filter = "Имя файла LIKE'" + textBox2.Text + "%'";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        void addItem()
        {
            comboBox1.Items.Add("104Ш6");
            comboBox1.Items.Add("16ОКС1-8");
            comboBox1.Items.Add("17В311");
            comboBox1.Items.Add("17В311 М");
            comboBox1.Items.Add("17В311М");
            comboBox1.Items.Add("17В319");
            comboBox1.Items.Add("17В320");
            comboBox1.Items.Add("17В37");
            comboBox1.Items.Add("17В39");
            comboBox1.Items.Add("17В39 1ТР");
            comboBox1.Items.Add("1А35");
            comboBox1.Items.Add("1А35М");
            comboBox1.Items.Add("1А35М-1");
            comboBox1.Items.Add("1А38");
            comboBox1.Items.Add("1А40");
            comboBox1.Items.Add("1А40-1");
            comboBox1.Items.Add("1А40-1М");
            comboBox1.Items.Add("1А40М");
            comboBox1.Items.Add("1А45");
            comboBox1.Items.Add("1А45-1");
            comboBox1.Items.Add("1А45М");
            comboBox1.Items.Add("1Г43");
            comboBox1.Items.Add("1Д22");
        }

        void dse_archive_000()
        {
            ArrayList list = new ArrayList();

            label8.Text = "Создание файла archive_doc.csv";
            progressBar1.Value = 0;
            progressBar1.Maximum = dataGridView1.RowCount;

            dataGridView2.ColumnCount = 7;
            dataGridView2.ColumnHeadersVisible = true;
            dataGridView2.Columns[0].HeaderText = "Дец. номер";
            dataGridView2.Columns[0].Width = 150;
            dataGridView2.Columns[1].HeaderText = "Изм.";
            dataGridView2.Columns[1].Width = 80;
            dataGridView2.Columns[2].HeaderText = "Лист";
            dataGridView2.Columns[2].Width = 80;
            dataGridView2.Columns[3].HeaderText = "Имя файла";
            dataGridView2.Columns[3].Width = 150;
            dataGridView2.Columns[4].HeaderText = "Путь";
            dataGridView2.Columns[4].Width = 300;
            dataGridView2.Columns[5].HeaderText = "Применяемость";
            dataGridView2.Columns[5].Width = 150;
            dataGridView2.Columns[6].HeaderText = "Изделие";
            dataGridView2.Columns[6].Width = 150;

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                progressBar1.Value++;

                if (dataGridView1.Rows[i].Cells[1].Value.ToString() == "000")
                {
                        string[] row = new string[] { dataGridView1.Rows[i].Cells[0].Value.ToString(),
                        dataGridView1.Rows[i].Cells[1].Value.ToString(), dataGridView1.Rows[i].Cells[2].Value.ToString(),
                        dataGridView1.Rows[i].Cells[3].Value.ToString(), dataGridView1.Rows[i].Cells[4].Value.ToString(),
                        dataGridView1.Rows[i].Cells[5].Value.ToString(), dataGridView1.Rows[i].Cells[6].Value.ToString()};
                        dataGridView2.Rows.Add(row);
                }
            }
            label3.Text = dataGridView2.RowCount.ToString();
        }



        void dse_000()
        {
            ArrayList list = new ArrayList();

            label8.Text = "Создание файла archive_doc_000.csv";
            progressBar1.Value = 0;
            progressBar1.Maximum = dataGridView1.RowCount;

            dataGridView3.ColumnCount = 7;
            dataGridView3.ColumnHeadersVisible = true;
            dataGridView3.Columns[0].HeaderText = "Дец. номер";
            dataGridView3.Columns[0].Width = 150;
            dataGridView3.Columns[1].HeaderText = "Изм.";
            dataGridView3.Columns[1].Width = 80;
            dataGridView3.Columns[2].HeaderText = "Лист";
            dataGridView3.Columns[2].Width = 80;
            dataGridView3.Columns[3].HeaderText = "Имя файла";
            dataGridView3.Columns[3].Width = 150;
            dataGridView3.Columns[4].HeaderText = "Путь";
            dataGridView3.Columns[4].Width = 300;
            dataGridView3.Columns[5].HeaderText = "Применяемость";
            dataGridView3.Columns[5].Width = 150;
            dataGridView3.Columns[6].HeaderText = "Изделие";
            dataGridView3.Columns[6].Width = 150;

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                progressBar1.Value++;

                    if (dataGridView1.Rows[i].Cells[1].Value.ToString() != "000")
                    {
                        string[] row = new string[] { dataGridView1.Rows[i].Cells[0].Value.ToString(), dataGridView1.Rows[i].Cells[1].Value.ToString(), dataGridView1.Rows[i].Cells[2].Value.ToString(),
                        dataGridView1.Rows[i].Cells[3].Value.ToString(), dataGridView1.Rows[i].Cells[4].Value.ToString(),
                        dataGridView1.Rows[i].Cells[5].Value.ToString(), dataGridView1.Rows[i].Cells[6].Value.ToString()};
                        dataGridView3.Rows.Add(row);
                    }
            }
            label7.Text = dataGridView3.RowCount.ToString();
        }

        void reviziya_backup()
        {
            Dictionary<string, string[]> map = new Dictionary<string, string[]>();

            ArrayList list = new ArrayList();
            List<int> ch = new List<int>();

            label8.Text = "Создание файла archive_doc_000.csv";
            progressBar1.Value = 0;
            progressBar1.Maximum = dataGridView1.RowCount;

            dataGridView4.ColumnCount = 7;
            dataGridView4.ColumnHeadersVisible = true;
            dataGridView4.Columns[0].HeaderText = "Дец. номер";
            dataGridView4.Columns[0].Width = 150;
            dataGridView4.Columns[1].HeaderText = "Изм.";
            dataGridView4.Columns[1].Width = 80;
            dataGridView4.Columns[2].HeaderText = "Лист";
            dataGridView4.Columns[2].Width = 80;
            dataGridView4.Columns[3].HeaderText = "Имя файла";
            dataGridView4.Columns[3].Width = 150;
            dataGridView4.Columns[4].HeaderText = "Путь";
            dataGridView4.Columns[4].Width = 300;
            dataGridView4.Columns[5].HeaderText = "Применяемость";
            dataGridView4.Columns[5].Width = 150;
            dataGridView4.Columns[6].HeaderText = "Изделие";
            dataGridView4.Columns[6].Width = 150;

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                progressBar1.Value++;

                string[] tmp = { dataGridView1.Rows[i].Cells[0].Value.ToString(), dataGridView1.Rows[i].Cells[1].Value.ToString(), dataGridView1.Rows[i].Cells[2].Value.ToString(),
                        dataGridView1.Rows[i].Cells[3].Value.ToString(), dataGridView1.Rows[i].Cells[4].Value.ToString(),
                        dataGridView1.Rows[i].Cells[5].Value.ToString(), dataGridView1.Rows[i].Cells[6].Value.ToString() };

                for (int j = 0; j < dataGridView1.RowCount; j++)
                {
                    if ((dataGridView1.Rows[i].Cells[0].Value.ToString() == dataGridView1.Rows[j].Cells[0].Value.ToString() && dataGridView1.Rows[i].Cells[2].Value.ToString() == dataGridView1.Rows[j].Cells[2].Value.ToString()) && dataGridView1.Rows[i].Cells[1].Value.ToString() != dataGridView1.Rows[j].Cells[1].Value.ToString())
                    {
                        string[] tmp2 = { dataGridView1.Rows[j].Cells[0].Value.ToString(), dataGridView1.Rows[j].Cells[1].Value.ToString(), dataGridView1.Rows[j].Cells[2].Value.ToString(),
                        dataGridView1.Rows[j].Cells[3].Value.ToString(), dataGridView1.Rows[j].Cells[4].Value.ToString(),
                        dataGridView1.Rows[j].Cells[5].Value.ToString(), dataGridView1.Rows[j].Cells[6].Value.ToString() };

                        if (!map.ContainsKey(dataGridView1.Rows[j].Cells[0].Value.ToString() + dataGridView1.Rows[j].Cells[1].Value.ToString() + dataGridView1.Rows[j].Cells[2].Value.ToString()))
                        {
                            map.Add((dataGridView1.Rows[j].Cells[0].Value.ToString() + dataGridView1.Rows[j].Cells[1].Value.ToString() + dataGridView1.Rows[j].Cells[2].Value.ToString()), tmp2);
                            dataGridView4.Rows.Add(tmp2);
                            ch.Add(Convert.ToInt32(dataGridView1.Rows[j].Cells[1].Value.ToString()));
                            dataGridView1.Rows.RemoveAt(j);
                        }
                    }
                }

                ch.Sort();
              /*  
                foreach (KeyValuePair<string, string[]> keyValue in map)
                {
                    MessageBox.Show(keyValue.Key + " - " + keyValue.Value);
                }
          */

                 for (int g = 0; g < ch.Count(); g++)
                 {
                    // if (arr[g] == 0)
                    //    MessageBox.Show("Изменение: " + arr[g]);

                    if (ch[g] != g)
                        for (int n = 0; n < ch[g]; n++)
                        {
                            if (ch.Contains(n) == false)
                                ch.Add(n);
                            for (int c = 0; c < dataGridView4.RowCount; c++)
                            {
                                string s = ch[g].ToString();
                                string st = "";
                                if (s.Length == 1)
                                    st = "00" + s;
                                else if (s.Length == 2)
                                    st = "0" + s;
                                else
                                    st = s;

                             /*   foreach (KeyValuePair<string, string[]> keyValue in map)
                                {
                                    string t = dataGridView1.Rows[i].Cells[0].Value.ToString() + st + dataGridView1.Rows[i].Cells[2].Value.ToString();
                                    if (keyValue.Key == t)
                                    dataGridView4.Rows.Add(keyValue.Value);
                                }
                                // MessageBox.Show("Добавлено измение: " + n);*/
                            }
                        }
                    else
                    {
                        // MessageBox.Show("Изменение: " + arr[g]);
                    }
                 }
                 ch.Sort();

                 string str = "";
                 foreach (int a in ch)
                 {
                     str += " " + a.ToString();
                 }
                 //MessageBox.Show(str);

                 ch.Clear();
            }
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}

