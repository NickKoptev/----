namespace Курсовая_Коптев
{
    public partial class UserControl4 : UserControl
    {
        Excel.Application xlexcel;
        Excel.Workbook xlworkbook;
        Excel.Worksheet xlworksheet;

        private readonly string dorogakfaily = @"C:\Users\Nic\Desktop\Курсовая Коптев шаблон накладной.docx";

        public UserControl4()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (numericUpDown1.Value != 0)
            {
                dataGridView1.Rows.Add(dateTimePicker1.Value.Date.ToShortDateString(), textBox1.Text, textBox2.Text, maskedTextBox1.Text, comboBox1.Text, "Ёлка Снежная королева", numericUpDown1.Value, numericUpDown1.Value * 41900);
            }

            if (numericUpDown2.Value != 0)
            {
                dataGridView1.Rows.Add(dateTimePicker1.Value.Date.ToShortDateString(), textBox1.Text, textBox2.Text, maskedTextBox1.Text, comboBox1.Text, "Ёлка Скандинавская", numericUpDown2.Value, numericUpDown2.Value * 24900);
            }

            if (numericUpDown3.Value != 0)
            {
                dataGridView1.Rows.Add(dateTimePicker1.Value.Date.ToShortDateString(), textBox1.Text, textBox2.Text, maskedTextBox1.Text, comboBox1.Text, "Ёлка Снежная LED", numericUpDown3.Value, numericUpDown3.Value * 29600);
            }

            if (numericUpDown4.Value != 0)
            {
                dataGridView1.Rows.Add(dateTimePicker1.Value.Date.ToShortDateString(), textBox1.Text, textBox2.Text, maskedTextBox1.Text, comboBox1.Text, "Ёлка Версальская", numericUpDown4.Value, numericUpDown4.Value * 51500);
            }

            if (numericUpDown5.Value != 0)
            {
                dataGridView1.Rows.Add(dateTimePicker1.Value.Date.ToShortDateString(), textBox1.Text, textBox2.Text, maskedTextBox1.Text, comboBox1.Text, "Ёлка Скандинавская белая", numericUpDown5.Value, numericUpDown5.Value * 12900);
            }

            if (numericUpDown6.Value != 0)
            {
                dataGridView1.Rows.Add(dateTimePicker1.Value.Date.ToShortDateString(), textBox1.Text, textBox2.Text, maskedTextBox1.Text, comboBox1.Text, "Шарики Брызги шампанского", numericUpDown6.Value, numericUpDown6.Value * 1590);
            }

            if (numericUpDown7.Value != 0)
            {
                dataGridView1.Rows.Add(dateTimePicker1.Value.Date.ToShortDateString(), textBox1.Text, textBox2.Text, maskedTextBox1.Text, comboBox1.Text, "Шарики Ягодный смузи", numericUpDown7.Value, numericUpDown7.Value * 1990);
            }

            if (numericUpDown8.Value != 0)
            {
                dataGridView1.Rows.Add(dateTimePicker1.Value.Date.ToShortDateString(), textBox1.Text, textBox2.Text, maskedTextBox1.Text, comboBox1.Text, "Шарики Мятная свежесть", numericUpDown8.Value, numericUpDown8.Value * 1990);
            }

            if (numericUpDown9.Value != 0)
            {
                dataGridView1.Rows.Add(dateTimePicker1.Value.Date.ToShortDateString(), textBox1.Text, textBox2.Text, maskedTextBox1.Text, comboBox1.Text, "Шарики Шоколадное суфле", numericUpDown9.Value, numericUpDown9.Value * 2490);
            }

            if (numericUpDown10.Value != 0)
            {
                dataGridView1.Rows.Add(dateTimePicker1.Value.Date.ToShortDateString(), textBox1.Text, textBox2.Text, maskedTextBox1.Text, comboBox1.Text, "Шарики Розовое шампанское", numericUpDown10.Value, numericUpDown10.Value * 2490);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox2.Clear();
            maskedTextBox1.Clear();
            comboBox1.SelectedIndex = -1;
            dateTimePicker1.Value = DateTime.Now;
            numericUpDown1.Value = 0;
            numericUpDown2.Value = 0;
            numericUpDown3.Value = 0;
            numericUpDown4.Value = 0;
            numericUpDown5.Value = 0;
            numericUpDown6.Value = 0;
            numericUpDown7.Value = 0;
            numericUpDown8.Value = 0;
            numericUpDown9.Value = 0;
            numericUpDown10.Value = 0;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow item in dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.RemoveAt(item.Index);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult iExit;
            iExit = MessageBox.Show("Сохранить заявку?", "Сохранить", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (iExit == DialogResult.Yes)
            {
                xlexcel = new Excel.Application();
                xlworkbook = xlexcel.Workbooks.Open("C:\\Users\\Nic\\Desktop\\Курсовая Коптев.xlsx", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlworksheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item(1);
                xlexcel.Visible = false;

                xlworksheet = xlworkbook.Sheets["Лист1"];
                xlworksheet = xlworkbook.ActiveSheet;
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    int _lasrRow = xlworksheet.Range["A" + xlworksheet.Rows.Count].End[Excel.XlDirection.xlUp].Row + 1;
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        xlworksheet.Cells[_lasrRow, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }
                xlworkbook.Close(true);
                xlexcel.Quit();
             
                var a = dateTimePicker1.Text;
                var b = textBox1.Text;
                var c = textBox2.Text;
                var d = maskedTextBox1.Text.ToString();
                var f = comboBox1.Text;
                var g = "Ёлка Снежная королева " + numericUpDown1.Value + " шт.";
                var n = "Ёлка Скандинавская " + numericUpDown2.Value + " шт.";
                var q = "Ёлка Снежная LED " + numericUpDown3.Value + " шт.";
                var w = "Ёлка Версальская " + numericUpDown4.Value + " шт.";
                var r = "Ёлка Скандинавская белая " + numericUpDown5.Value + " шт.";
                var y = "Шарики Брызги шампанского " + numericUpDown6.Value + " шт.";
                var h = "Шарики Ягодный смузи " + numericUpDown7.Value + " шт.";
                var x = "Шарики Мятная свежесть " + numericUpDown8.Value + " шт.";
                var p = "Шарики Шоколадное суфле " + numericUpDown9.Value + " шт.";
                var s = "Шарики Розовое шампанское " + numericUpDown10.Value + " шт.";
                var k = numericUpDown1.Value * 41900 + numericUpDown2.Value * 24900 + numericUpDown3.Value * 29600 + numericUpDown4.Value * 51500 + numericUpDown5.Value * 12900 + numericUpDown6.Value * 1590 + numericUpDown7.Value * 11990
                    + numericUpDown8.Value * 1990 + numericUpDown9.Value * 2490 + numericUpDown10.Value * 2490 + " руб";

                var wordApp = new Word.Application();
                wordApp.Visible = false;

                var wordDocument = wordApp.Documents.Open(dorogakfaily);
                zamena("{a}", a, wordDocument);
                zamena("{b}", b, wordDocument);
                zamena("{c}", c, wordDocument);
                zamena("{d}", d, wordDocument);
                zamena("{f}", f, wordDocument);
                zamena("{g}", g, wordDocument);
                zamena("{n}", n, wordDocument);
                zamena("{q}", q, wordDocument);
                zamena("{w}", w, wordDocument);
                zamena("{r}", r, wordDocument);
                zamena("{y}", y, wordDocument);
                zamena("{h}", h, wordDocument);
                zamena("{x}", x, wordDocument);
                zamena("{p}", p, wordDocument);
                zamena("{s}", s, wordDocument);
                zamena("{k}", k, wordDocument);
              

                wordDocument.SaveAs2(@"C:\Users\Nic\Desktop\Курсовая Коптев шаблон накладной" + a + g + ".docx");
                wordApp.Quit();

                textBox1.Clear();
                textBox2.Clear();
                maskedTextBox1.Clear();
                comboBox1.SelectedIndex = -1;
                dateTimePicker1.Value = DateTime.Now;
                numericUpDown1.Value = 0;
                numericUpDown2.Value = 0;
                numericUpDown3.Value = 0;
                numericUpDown4.Value = 0;
                numericUpDown5.Value = 0;
                numericUpDown6.Value = 0;
                numericUpDown7.Value = 0;
                numericUpDown8.Value = 0;
                numericUpDown9.Value = 0;
                numericUpDown10.Value = 0;
                dataGridView1.Rows.Clear();

            }

        }
        private void zamena(string chtoIshem, string naChtoMenyaem, Word.Document wordDocument)
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: chtoIshem, ReplaceWith: naChtoMenyaem);
        }

    }
}


