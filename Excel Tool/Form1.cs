using System;
using System.Windows.Forms;
namespace Excel_Tool
{
    public partial class Form1 : Form
    {
        private Form2 form;
        public Form1()
        {
            InitializeComponent();
        }

        public void SelectExcelFile(int fileseq)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Filter= "Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx";

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string file = dialog.FileName;
                if (fileseq == 1)
                    textBox1.Text = file;
                else textBox2.Text = file;

            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
           
            string file1 = textBox1.Text;
            string file2 = textBox2.Text;
            if (file1.Equals("") || file2.Equals(""))
                MessageBox.Show("请选择Excel文件！", "Excel文件路径不完整", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
            {
                
                if(form==null)
                    form= new Form2(file1, file2);
                form.Show();
            }
            
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SelectExcelFile(1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SelectExcelFile(2);
        }

        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            string path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            textBox1.Text = path;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
          
        }

        private void textBox1_DragEnter(object sender, DragEventArgs e)    //实现文件拖拽获得路径功能
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Link;

            else e.Effect = DragDropEffects.None;
          
           
        }

        private void textBox2_DragDrop(object sender, DragEventArgs e)
        {
            string path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            textBox2.Text = path;
        }

        private void textBox2_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Link;

            else e.Effect = DragDropEffects.None;
        }
    }
}
