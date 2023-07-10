namespace Курсовая_Коптев
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            userControl11.BringToFront();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            userControl11.BringToFront();
        }

        private void button2_Click(object sender, EventArgs e)
        {
           userControl21.BringToFront();
        }

        private void button3_Click(object sender, EventArgs e)
        {
           userControl31.BringToFront();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            userControl41.BringToFront();
        }
    }
}

