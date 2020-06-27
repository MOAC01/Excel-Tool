using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace Excel_Tool
{
    public partial class Form2 : Form
    {
        private string source, destination;
        private IWorkbook Workbook1,Workbook2;

        public Form2(string src, string des)
        {
            this.source = src;
            this.destination = des;
            InitializeComponent();
        }
        public Form2()
        {
            InitializeComponent();
        }

        public int CheckFileType(string filePath)
        {
            if (!filePath.EndsWith(".xls") || !filePath.EndsWith(".xlsx"))
                return 0;
            return 1;
        }

        public void ReadExcel(string filepath,int select)
        {

            try
            {
                FileStream fileStream = new FileStream(filepath, FileMode.Open, FileAccess.Read);

                if (select == 1)
                {
                    if (filepath.IndexOf(".xlsx") > 0)
                    {
                        Workbook1 = new XSSFWorkbook(fileStream);
                    }

                    else if (filepath.IndexOf(".xls") > 0)
                    {
                        Workbook1 = new HSSFWorkbook(fileStream);  //xls数据读入workbook
                    }
                }

                else if (select == 2)
                {
                    if (filepath.IndexOf(".xlsx") > 0)
                    {
                        Workbook2 = new XSSFWorkbook(fileStream);
                    }

                    else if (filepath.IndexOf(".xls") > 0)
                    {
                        Workbook2 = new HSSFWorkbook(fileStream);  //xls数据读入workbook
                    }

                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("在打开文件时遇到错误，请检查路径是否正确、文件是否可用或文件是否可用", "打开文件时错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);

                MessageBox.Show(Convert.ToString(ex), "打开文件时错误",MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        /*
         *   根据sheet设置不同sheet表的列名到下拉框中
         * 
         */
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
         
            
        }

        /*
         *   下拉框选择改变事件
         */
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = comboBox3.SelectedIndex;
            ISheet sheet = Workbook2.GetSheetAt(index);
            IRow row = sheet.GetRow(0);
            int ColunmCount = row.LastCellNum;
            ICell cell;
            for (int i = 0; i < ColunmCount; i++)
            {
                cell = row.GetCell(i);
                comboBox2.Items.Add(cell.StringCellValue);
                comboBox4.Items.Add(cell.StringCellValue);
            }
        }

        /*
         * 窗体加载事件
         */
        private void Form2_Load(object sender, EventArgs e)
        {
            string filepath1 = source;
            string filepath2 = destination;

            ReadExcel(source, 1);           //读取匹配源
            ReadExcel(destination, 2);     //读取需要操作的文件

            int Sheeetcount1 = Workbook1.NumberOfSheets;
            int Sheeetcount2 = Workbook2.NumberOfSheets;
            string FileName1= System.IO.Path.GetFileName(filepath1);
            string FileName2= System.IO.Path.GetFileName(filepath2);

            label2.Text = "文件名:" + FileName1 + "；表(sheet)个数："+Convert.ToString(Sheeetcount1);
            label4.Text = "文件名:" + FileName2 + "；表(sheet)个数：" + Convert.ToString(Sheeetcount2);

            for(int i = 0; i < Sheeetcount1; i++)
            {
                comboBox1.Items.Add(Workbook1.GetSheetName(i));
            }

            for(int j=0;j<Sheeetcount2;j++)
            {
                comboBox3.Items.Add(Workbook2.GetSheetName(j));
            }

        }

        /*
         *  点击事件
         */ 
        private void button1_Click(object sender, EventArgs e)
        {
            
            int conditon1 = comboBox2.SelectedIndex;
            int condition2 = comboBox4.SelectedIndex;
            int select = comboBox1.SelectedIndex;
            string p1 = comboBox2.Text;
            string p2 = comboBox4.Text;
            if(conditon1 < 0 || condition2 < 0 || select<0)
            {
                MessageBox.Show("请选择匹配条件或匹配源！", "未选择条件列或匹配源", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (p1.Equals(p2))
            {
                MessageBox.Show("根据条件列与需要填充的列不能是同一个，请重新选择", "列选择错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //if(!CheckFileType())

            int c1 = GetPachIndex(p1);      
            int c2 = GetPachIndex(p2);

            if (c1 < 0 || c2 < 0)
            {
                MessageBox.Show("错误,在匹配源中找不到选择的列,请检查匹配源的sheet或选择匹配的条件列", "操作失败", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            List<CellObject> CellList = GetPachExcelObject(c1,c2);          //获取符合条件的Cell对象
            ISheet sheet = Workbook2.GetSheetAt(comboBox3.SelectedIndex);  //操作此Sheet
            IRow row;
            ICell cell1,cell2;
            int OpColumn = comboBox4.SelectedIndex;             //填充此列
            int finished=0;
            progressBar1.Maximum = sheet.LastRowNum;
            progressBar1.Value = 0;
            progressBar1.Step = 1;
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                row = sheet.GetRow(i);                   //行移动
                cell1 = row.GetCell(conditon1);          //操作的文件的匹配条件
                cell2 = row.GetCell(OpColumn);          //本行需要填充的列
                cell1.SetCellType(CellType.String);
                if (cell2 == null)                      //如果此单元格为空则创建
                    cell2 = row.CreateCell(OpColumn);
                cell2.SetCellType(CellType.String);     //强制设置为字符串类型
                foreach(CellObject co in CellList)     //遍历list
                {
                    if(cell1.StringCellValue.Equals(co.Param1))  //如果根据条件的单元格和list中的字段1相同
                    {
                        cell2.SetCellValue(co.Param2);     //把list中的字段2写入这个单元格
                        CellList.Remove(co);             //删除已经匹配中的对象
                        break;                          //提前结束循环，提高运行效率
                    }
                }
                finished++;
                if (progressBar1.Value > progressBar1.Maximum)   //防止进度条越界
                    progressBar1.Value -= 1;
                progressBar1.Value += progressBar1.Step;
                label7.Text = "已完成:" + Convert.ToString(finished) + "/" + Convert.ToString(progressBar1.Maximum).ToString();
                label7.Refresh();
                System.Threading.Thread.Sleep(10);        //线程休眠
            }

            FileStream fs = File.Create(destination);
            Workbook2.Write(fs);        //保存修改
            Workbook1.Close();
            Workbook2.Close();
            fs.Close();
            MessageBox.Show("执行完成，已成功匹配所有记录", "操作成功", MessageBoxButtons.OK,MessageBoxIcon.Information);

        }

        /*获取有效填充行数*/
        public int GetValidRows()
        {
            ISheet sheet = Workbook2.GetSheetAt(comboBox3.SelectedIndex);
            IRow row = sheet.GetRow(1);
            int CelIndex = comboBox2.SelectedIndex;
            ICell cell;
            int sum = 0;
            for(int i = 1; i <= sheet.LastRowNum;i++)
            {
                cell = row.GetCell(CelIndex);
                if (cell != null && !cell.StringCellValue.Equals(""))
                    sum++;
            }
            return sum;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ISheet sheet = Workbook2.GetSheetAt(1);
            int count = GetValidRows();

            MessageBox.Show(Convert.ToString(count), "测试", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        public int GetPachIndex(string _Cellvalue)   /*获取匹配条件列在匹配源中的位置*/
        {
            ISheet sheet = Workbook1.GetSheetAt(comboBox1.SelectedIndex);
            IRow row = sheet.GetRow(0);
            ICell cell;
            for(int i=0;i<row.LastCellNum;i++)
            {
                cell = row.GetCell(i);
                if (cell.StringCellValue.Equals(_Cellvalue))
                    return i;    //返回参数列在匹配源中的列数
            }

            return -1;        //找不到，返回空
        }

        /*
         * 从匹配源中找出所有符合条件的行列并组成对象，返回对象列表
         */
        private List<CellObject> GetPachExcelObject(int CellIndex1,int CellIndex2)
        {
            ISheet sheet = Workbook1.GetSheetAt(comboBox1.SelectedIndex);
            IRow row;
            ICell cell1,cell2;
            List<CellObject> cellObjects = new List<CellObject>();

            for(int i=1;i<=sheet.LastRowNum;i++)
            {
                row = sheet.GetRow(i);
                cell1 = row.GetCell(CellIndex1);
                cell1.SetCellType(CellType.String);
                cell2 = row.GetCell(CellIndex2);
                CellType type = cell2.CellType;
                cell2.SetCellType(CellType.String);
                CellObject cellObject = new CellObject(cell1.StringCellValue,cell2.StringCellValue);
                cellObjects.Add(cellObject);
            }
            return cellObjects;

        }
    }
}
