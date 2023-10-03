using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using 高中職學校班級及學生概況.DAO;

namespace 高中職學校班級及學生概況
{
    public partial class Form1 : BaseForm
    {
        // 部別與科別
        Dictionary<string, List<string>> DeptGroupNameDict;

        public Form1()
        {
            InitializeComponent();
            DeptGroupNameDict = new Dictionary<string, List<string>>();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(comboDeptGroup.Text))
            {
                MsgBox.Show("請選擇部別");
                return;
            }       
            
            btnRun.Enabled = false;

            List<string> deptNameList = new List<string>();
            if (DeptGroupNameDict.ContainsKey(comboDeptGroup.Text))
            {
                deptNameList = DeptGroupNameDict[comboDeptGroup.Text];
            }

            //運用Aspose元件來新增活頁簿
            Aspose.Cells.Workbook ScoreWorkBook = new Aspose.Cells.Workbook();

            System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();

            saveFileDialog.FileName = "高中職學校班級及學生概況.xls";
            saveFileDialog.AddExtension = true;
            saveFileDialog.DefaultExt = "xls";

            string filename = (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) ? saveFileDialog.FileName : "";

            // ScoreWorkBook.Open(System.Windows.Forms.Application.StartupPath + "\\Customize\\高中職學校班級及學生概況.xlt");
            ScoreWorkBook.Open(new System.IO.MemoryStream(Properties.Resources.高中職學校班級及學生概況));

            Program.ErrorList.Clear();

            #region 普通科
            ClassStatistics ClassSat = new ClassStatistics();

            ClassSat.StartStatistics(0, deptNameList);
            ScoreWorkBook.Worksheets[0].Name = "高中職學校班級及學生概況-普通科";
            ScoreWorkBook.Worksheets[0].Cells[4, 0].PutValue(SmartSchool.Customization.Data.SystemInformation.SchoolCode);

            //一年級班級數總計
            ScoreWorkBook.Worksheets[0].Cells[7, 5].PutValue(ClassSat.Level1Count24);
            ScoreWorkBook.Worksheets[0].Cells[7, 7].PutValue(ClassSat.Level1Count34);
            ScoreWorkBook.Worksheets[0].Cells[7, 9].PutValue(ClassSat.Level1Count44);
            ScoreWorkBook.Worksheets[0].Cells[7, 11].PutValue(ClassSat.Level1Count54);
            ScoreWorkBook.Worksheets[0].Cells[7, 13].PutValue(ClassSat.Level1Count55);

            //二年級班級數總計
            ScoreWorkBook.Worksheets[0].Cells[8, 5].PutValue(ClassSat.Level2Count24);
            ScoreWorkBook.Worksheets[0].Cells[8, 7].PutValue(ClassSat.Level2Count34);
            ScoreWorkBook.Worksheets[0].Cells[8, 9].PutValue(ClassSat.Level2Count44);
            ScoreWorkBook.Worksheets[0].Cells[8, 11].PutValue(ClassSat.Level2Count54);
            ScoreWorkBook.Worksheets[0].Cells[8, 13].PutValue(ClassSat.Level2Count55);

            //三年級班級數總計
            ScoreWorkBook.Worksheets[0].Cells[9, 5].PutValue(ClassSat.Level3Count24);
            ScoreWorkBook.Worksheets[0].Cells[9, 7].PutValue(ClassSat.Level3Count34);
            ScoreWorkBook.Worksheets[0].Cells[9, 9].PutValue(ClassSat.Level3Count44);
            ScoreWorkBook.Worksheets[0].Cells[9, 11].PutValue(ClassSat.Level3Count54);
            ScoreWorkBook.Worksheets[0].Cells[9, 13].PutValue(ClassSat.Level3Count55);

            //四年級班級數總計
            ScoreWorkBook.Worksheets[0].Cells[10, 5].PutValue(ClassSat.Level4Count24);
            ScoreWorkBook.Worksheets[0].Cells[10, 7].PutValue(ClassSat.Level4Count34);
            ScoreWorkBook.Worksheets[0].Cells[10, 9].PutValue(ClassSat.Level4Count44);
            ScoreWorkBook.Worksheets[0].Cells[10, 11].PutValue(ClassSat.Level4Count54);
            ScoreWorkBook.Worksheets[0].Cells[10, 13].PutValue(ClassSat.Level4Count55);

            //一年級學生數總計
            ScoreWorkBook.Worksheets[0].Cells[15, 5].PutValue(ClassSat.SR15C5);
            ScoreWorkBook.Worksheets[0].Cells[15, 6].PutValue(ClassSat.SR15C6);
            ScoreWorkBook.Worksheets[0].Cells[15, 7].PutValue(ClassSat.SR15C7);
            ScoreWorkBook.Worksheets[0].Cells[15, 8].PutValue(ClassSat.SR15C8);
            ScoreWorkBook.Worksheets[0].Cells[15, 9].PutValue(ClassSat.SR15C9);
            ScoreWorkBook.Worksheets[0].Cells[15, 10].PutValue(ClassSat.SR15C10);
            ScoreWorkBook.Worksheets[0].Cells[15, 11].PutValue(ClassSat.SR15C11);
            ScoreWorkBook.Worksheets[0].Cells[15, 12].PutValue(ClassSat.SR15C12);
            ScoreWorkBook.Worksheets[0].Cells[15, 13].PutValue(ClassSat.SR15C13);
            ScoreWorkBook.Worksheets[0].Cells[15, 14].PutValue(ClassSat.SR15C14);

            ScoreWorkBook.Worksheets[0].Cells[16, 5].PutValue(ClassSat.SR16C5);
            ScoreWorkBook.Worksheets[0].Cells[16, 6].PutValue(ClassSat.SR16C6);
            ScoreWorkBook.Worksheets[0].Cells[16, 7].PutValue(ClassSat.SR16C7);
            ScoreWorkBook.Worksheets[0].Cells[16, 8].PutValue(ClassSat.SR16C8);
            ScoreWorkBook.Worksheets[0].Cells[16, 9].PutValue(ClassSat.SR16C9);
            ScoreWorkBook.Worksheets[0].Cells[16, 10].PutValue(ClassSat.SR16C10);
            ScoreWorkBook.Worksheets[0].Cells[16, 11].PutValue(ClassSat.SR16C11);
            ScoreWorkBook.Worksheets[0].Cells[16, 12].PutValue(ClassSat.SR16C12);
            ScoreWorkBook.Worksheets[0].Cells[16, 13].PutValue(ClassSat.SR16C13);
            ScoreWorkBook.Worksheets[0].Cells[16, 14].PutValue(ClassSat.SR16C14);

            //二年級學生數總計
            ScoreWorkBook.Worksheets[0].Cells[17, 5].PutValue(ClassSat.SR17C5);
            ScoreWorkBook.Worksheets[0].Cells[17, 6].PutValue(ClassSat.SR17C6);
            ScoreWorkBook.Worksheets[0].Cells[17, 7].PutValue(ClassSat.SR17C7);
            ScoreWorkBook.Worksheets[0].Cells[17, 8].PutValue(ClassSat.SR17C8);
            ScoreWorkBook.Worksheets[0].Cells[17, 9].PutValue(ClassSat.SR17C9);
            ScoreWorkBook.Worksheets[0].Cells[17, 10].PutValue(ClassSat.SR17C10);
            ScoreWorkBook.Worksheets[0].Cells[17, 11].PutValue(ClassSat.SR17C11);
            ScoreWorkBook.Worksheets[0].Cells[17, 12].PutValue(ClassSat.SR17C12);
            ScoreWorkBook.Worksheets[0].Cells[17, 13].PutValue(ClassSat.SR17C13);
            ScoreWorkBook.Worksheets[0].Cells[17, 14].PutValue(ClassSat.SR17C14);


            ScoreWorkBook.Worksheets[0].Cells[18, 5].PutValue(ClassSat.SR18C5);
            ScoreWorkBook.Worksheets[0].Cells[18, 6].PutValue(ClassSat.SR18C6);
            ScoreWorkBook.Worksheets[0].Cells[18, 7].PutValue(ClassSat.SR18C7);
            ScoreWorkBook.Worksheets[0].Cells[18, 8].PutValue(ClassSat.SR18C8);
            ScoreWorkBook.Worksheets[0].Cells[18, 9].PutValue(ClassSat.SR18C9);
            ScoreWorkBook.Worksheets[0].Cells[18, 10].PutValue(ClassSat.SR18C10);
            ScoreWorkBook.Worksheets[0].Cells[18, 11].PutValue(ClassSat.SR18C11);
            ScoreWorkBook.Worksheets[0].Cells[18, 12].PutValue(ClassSat.SR18C12);
            ScoreWorkBook.Worksheets[0].Cells[18, 13].PutValue(ClassSat.SR18C13);
            ScoreWorkBook.Worksheets[0].Cells[18, 14].PutValue(ClassSat.SR18C14);

            //三年級學生數總計
            ScoreWorkBook.Worksheets[0].Cells[19, 5].PutValue(ClassSat.SR19C5);
            ScoreWorkBook.Worksheets[0].Cells[19, 6].PutValue(ClassSat.SR19C6);
            ScoreWorkBook.Worksheets[0].Cells[19, 7].PutValue(ClassSat.SR19C7);
            ScoreWorkBook.Worksheets[0].Cells[19, 8].PutValue(ClassSat.SR19C8);
            ScoreWorkBook.Worksheets[0].Cells[19, 9].PutValue(ClassSat.SR19C9);
            ScoreWorkBook.Worksheets[0].Cells[19, 10].PutValue(ClassSat.SR19C10);
            ScoreWorkBook.Worksheets[0].Cells[19, 11].PutValue(ClassSat.SR19C11);
            ScoreWorkBook.Worksheets[0].Cells[19, 12].PutValue(ClassSat.SR19C12);
            ScoreWorkBook.Worksheets[0].Cells[19, 13].PutValue(ClassSat.SR19C13);
            ScoreWorkBook.Worksheets[0].Cells[19, 14].PutValue(ClassSat.SR19C14);

            ScoreWorkBook.Worksheets[0].Cells[20, 5].PutValue(ClassSat.SR20C5);
            ScoreWorkBook.Worksheets[0].Cells[20, 6].PutValue(ClassSat.SR20C6);
            ScoreWorkBook.Worksheets[0].Cells[20, 7].PutValue(ClassSat.SR20C7);
            ScoreWorkBook.Worksheets[0].Cells[20, 8].PutValue(ClassSat.SR20C8);
            ScoreWorkBook.Worksheets[0].Cells[20, 9].PutValue(ClassSat.SR20C9);
            ScoreWorkBook.Worksheets[0].Cells[20, 10].PutValue(ClassSat.SR20C10);
            ScoreWorkBook.Worksheets[0].Cells[20, 11].PutValue(ClassSat.SR20C11);
            ScoreWorkBook.Worksheets[0].Cells[20, 12].PutValue(ClassSat.SR20C12);
            ScoreWorkBook.Worksheets[0].Cells[20, 13].PutValue(ClassSat.SR20C13);
            ScoreWorkBook.Worksheets[0].Cells[20, 14].PutValue(ClassSat.SR20C14);

            //四年級學生數總計
            ScoreWorkBook.Worksheets[0].Cells[21, 5].PutValue(ClassSat.SR21C5);
            ScoreWorkBook.Worksheets[0].Cells[21, 6].PutValue(ClassSat.SR21C6);
            ScoreWorkBook.Worksheets[0].Cells[21, 7].PutValue(ClassSat.SR21C7);
            ScoreWorkBook.Worksheets[0].Cells[21, 8].PutValue(ClassSat.SR21C8);
            ScoreWorkBook.Worksheets[0].Cells[21, 9].PutValue(ClassSat.SR21C9);
            ScoreWorkBook.Worksheets[0].Cells[21, 10].PutValue(ClassSat.SR21C10);
            ScoreWorkBook.Worksheets[0].Cells[21, 11].PutValue(ClassSat.SR21C11);
            ScoreWorkBook.Worksheets[0].Cells[21, 12].PutValue(ClassSat.SR21C12);
            ScoreWorkBook.Worksheets[0].Cells[21, 13].PutValue(ClassSat.SR21C13);
            ScoreWorkBook.Worksheets[0].Cells[21, 14].PutValue(ClassSat.SR21C14);

            ScoreWorkBook.Worksheets[0].Cells[22, 5].PutValue(ClassSat.SR22C5);
            ScoreWorkBook.Worksheets[0].Cells[22, 6].PutValue(ClassSat.SR22C6);
            ScoreWorkBook.Worksheets[0].Cells[22, 7].PutValue(ClassSat.SR22C7);
            ScoreWorkBook.Worksheets[0].Cells[22, 8].PutValue(ClassSat.SR22C8);
            ScoreWorkBook.Worksheets[0].Cells[22, 9].PutValue(ClassSat.SR22C9);
            ScoreWorkBook.Worksheets[0].Cells[22, 10].PutValue(ClassSat.SR22C10);
            ScoreWorkBook.Worksheets[0].Cells[22, 11].PutValue(ClassSat.SR22C11);
            ScoreWorkBook.Worksheets[0].Cells[22, 12].PutValue(ClassSat.SR22C12);
            ScoreWorkBook.Worksheets[0].Cells[22, 13].PutValue(ClassSat.SR22C13);
            ScoreWorkBook.Worksheets[0].Cells[22, 14].PutValue(ClassSat.SR22C14);
            //延修生學生數總計
            ScoreWorkBook.Worksheets[0].Cells[23, 5].PutValue(ClassSat.SR23C5);
            ScoreWorkBook.Worksheets[0].Cells[23, 6].PutValue(ClassSat.SR23C6);
            ScoreWorkBook.Worksheets[0].Cells[23, 7].PutValue(ClassSat.SR23C7);
            ScoreWorkBook.Worksheets[0].Cells[23, 8].PutValue(ClassSat.SR23C8);
            ScoreWorkBook.Worksheets[0].Cells[23, 9].PutValue(ClassSat.SR23C9);
            ScoreWorkBook.Worksheets[0].Cells[23, 10].PutValue(ClassSat.SR23C10);
            ScoreWorkBook.Worksheets[0].Cells[23, 11].PutValue(ClassSat.SR23C11);
            ScoreWorkBook.Worksheets[0].Cells[23, 12].PutValue(ClassSat.SR23C12);
            ScoreWorkBook.Worksheets[0].Cells[23, 13].PutValue(ClassSat.SR23C13);
            ScoreWorkBook.Worksheets[0].Cells[23, 14].PutValue(ClassSat.SR23C14);

            ScoreWorkBook.Worksheets[0].Cells[24, 5].PutValue(ClassSat.SR24C5);
            ScoreWorkBook.Worksheets[0].Cells[24, 6].PutValue(ClassSat.SR24C6);
            ScoreWorkBook.Worksheets[0].Cells[24, 7].PutValue(ClassSat.SR24C7);
            ScoreWorkBook.Worksheets[0].Cells[24, 8].PutValue(ClassSat.SR24C8);
            ScoreWorkBook.Worksheets[0].Cells[24, 9].PutValue(ClassSat.SR24C9);
            ScoreWorkBook.Worksheets[0].Cells[24, 10].PutValue(ClassSat.SR24C10);
            ScoreWorkBook.Worksheets[0].Cells[24, 11].PutValue(ClassSat.SR24C11);
            ScoreWorkBook.Worksheets[0].Cells[24, 12].PutValue(ClassSat.SR24C12);
            ScoreWorkBook.Worksheets[0].Cells[24, 13].PutValue(ClassSat.SR24C13);
            ScoreWorkBook.Worksheets[0].Cells[24, 14].PutValue(ClassSat.SR24C14);

            #endregion


            #region 職業科
            ClassStatistics ClassSat1 = new ClassStatistics();

            ClassSat1.StartStatistics(1, deptNameList);
            ScoreWorkBook.Worksheets[1].Name = "高中職學校班級及學生概況-職業科";
            ScoreWorkBook.Worksheets[1].Cells[4, 0].PutValue(SmartSchool.Customization.Data.SystemInformation.SchoolCode);

            //一年級班級數總計
            ScoreWorkBook.Worksheets[1].Cells[7, 5].PutValue(ClassSat1.Level1Count24);
            ScoreWorkBook.Worksheets[1].Cells[7, 7].PutValue(ClassSat1.Level1Count34);
            ScoreWorkBook.Worksheets[1].Cells[7, 9].PutValue(ClassSat1.Level1Count44);
            ScoreWorkBook.Worksheets[1].Cells[7, 11].PutValue(ClassSat1.Level1Count54);
            ScoreWorkBook.Worksheets[1].Cells[7, 13].PutValue(ClassSat1.Level1Count55);

            //二年級班級數總計
            ScoreWorkBook.Worksheets[1].Cells[8, 5].PutValue(ClassSat1.Level2Count24);
            ScoreWorkBook.Worksheets[1].Cells[8, 7].PutValue(ClassSat1.Level2Count34);
            ScoreWorkBook.Worksheets[1].Cells[8, 9].PutValue(ClassSat1.Level2Count44);
            ScoreWorkBook.Worksheets[1].Cells[8, 11].PutValue(ClassSat1.Level2Count54);
            ScoreWorkBook.Worksheets[1].Cells[8, 13].PutValue(ClassSat1.Level2Count55);

            //三年級班級數總計
            ScoreWorkBook.Worksheets[1].Cells[9, 5].PutValue(ClassSat1.Level3Count24);
            ScoreWorkBook.Worksheets[1].Cells[9, 7].PutValue(ClassSat1.Level3Count34);
            ScoreWorkBook.Worksheets[1].Cells[9, 9].PutValue(ClassSat1.Level3Count44);
            ScoreWorkBook.Worksheets[1].Cells[9, 11].PutValue(ClassSat1.Level3Count54);
            ScoreWorkBook.Worksheets[1].Cells[9, 13].PutValue(ClassSat1.Level3Count55);

            //四年級班級數總計
            ScoreWorkBook.Worksheets[1].Cells[10, 5].PutValue(ClassSat1.Level4Count24);
            ScoreWorkBook.Worksheets[1].Cells[10, 7].PutValue(ClassSat1.Level4Count34);
            ScoreWorkBook.Worksheets[1].Cells[10, 9].PutValue(ClassSat1.Level4Count44);
            ScoreWorkBook.Worksheets[1].Cells[10, 11].PutValue(ClassSat1.Level4Count54);
            ScoreWorkBook.Worksheets[1].Cells[10, 13].PutValue(ClassSat1.Level4Count55);

            //一年級學生數總計
            ScoreWorkBook.Worksheets[1].Cells[15, 5].PutValue(ClassSat1.SR15C5);
            ScoreWorkBook.Worksheets[1].Cells[15, 6].PutValue(ClassSat1.SR15C6);
            ScoreWorkBook.Worksheets[1].Cells[15, 7].PutValue(ClassSat1.SR15C7);
            ScoreWorkBook.Worksheets[1].Cells[15, 8].PutValue(ClassSat1.SR15C8);
            ScoreWorkBook.Worksheets[1].Cells[15, 9].PutValue(ClassSat1.SR15C9);
            ScoreWorkBook.Worksheets[1].Cells[15, 10].PutValue(ClassSat1.SR15C10);
            ScoreWorkBook.Worksheets[1].Cells[15, 11].PutValue(ClassSat1.SR15C11);
            ScoreWorkBook.Worksheets[1].Cells[15, 12].PutValue(ClassSat1.SR15C12);
            ScoreWorkBook.Worksheets[1].Cells[15, 13].PutValue(ClassSat1.SR15C13);
            ScoreWorkBook.Worksheets[1].Cells[15, 14].PutValue(ClassSat1.SR15C14);

            ScoreWorkBook.Worksheets[1].Cells[16, 5].PutValue(ClassSat1.SR16C5);
            ScoreWorkBook.Worksheets[1].Cells[16, 6].PutValue(ClassSat1.SR16C6);
            ScoreWorkBook.Worksheets[1].Cells[16, 7].PutValue(ClassSat1.SR16C7);
            ScoreWorkBook.Worksheets[1].Cells[16, 8].PutValue(ClassSat1.SR16C8);
            ScoreWorkBook.Worksheets[1].Cells[16, 9].PutValue(ClassSat1.SR16C9);
            ScoreWorkBook.Worksheets[1].Cells[16, 10].PutValue(ClassSat1.SR16C10);
            ScoreWorkBook.Worksheets[1].Cells[16, 11].PutValue(ClassSat1.SR16C11);
            ScoreWorkBook.Worksheets[1].Cells[16, 12].PutValue(ClassSat1.SR16C12);
            ScoreWorkBook.Worksheets[1].Cells[16, 13].PutValue(ClassSat1.SR16C13);
            ScoreWorkBook.Worksheets[1].Cells[16, 14].PutValue(ClassSat1.SR16C14);

            //二年級學生數總計
            ScoreWorkBook.Worksheets[1].Cells[17, 5].PutValue(ClassSat1.SR17C5);
            ScoreWorkBook.Worksheets[1].Cells[17, 6].PutValue(ClassSat1.SR17C6);
            ScoreWorkBook.Worksheets[1].Cells[17, 7].PutValue(ClassSat1.SR17C7);
            ScoreWorkBook.Worksheets[1].Cells[17, 8].PutValue(ClassSat1.SR17C8);
            ScoreWorkBook.Worksheets[1].Cells[17, 9].PutValue(ClassSat1.SR17C9);
            ScoreWorkBook.Worksheets[1].Cells[17, 10].PutValue(ClassSat1.SR17C10);
            ScoreWorkBook.Worksheets[1].Cells[17, 11].PutValue(ClassSat1.SR17C11);
            ScoreWorkBook.Worksheets[1].Cells[17, 12].PutValue(ClassSat1.SR17C12);
            ScoreWorkBook.Worksheets[1].Cells[17, 13].PutValue(ClassSat1.SR17C13);
            ScoreWorkBook.Worksheets[1].Cells[17, 14].PutValue(ClassSat1.SR17C14);


            ScoreWorkBook.Worksheets[1].Cells[18, 5].PutValue(ClassSat1.SR18C5);
            ScoreWorkBook.Worksheets[1].Cells[18, 6].PutValue(ClassSat1.SR18C6);
            ScoreWorkBook.Worksheets[1].Cells[18, 7].PutValue(ClassSat1.SR18C7);
            ScoreWorkBook.Worksheets[1].Cells[18, 8].PutValue(ClassSat1.SR18C8);
            ScoreWorkBook.Worksheets[1].Cells[18, 9].PutValue(ClassSat1.SR18C9);
            ScoreWorkBook.Worksheets[1].Cells[18, 10].PutValue(ClassSat1.SR18C10);
            ScoreWorkBook.Worksheets[1].Cells[18, 11].PutValue(ClassSat1.SR18C11);
            ScoreWorkBook.Worksheets[1].Cells[18, 12].PutValue(ClassSat1.SR18C12);
            ScoreWorkBook.Worksheets[1].Cells[18, 13].PutValue(ClassSat1.SR18C13);
            ScoreWorkBook.Worksheets[1].Cells[18, 14].PutValue(ClassSat1.SR18C14);

            //三年級學生數總計
            ScoreWorkBook.Worksheets[1].Cells[19, 5].PutValue(ClassSat1.SR19C5);
            ScoreWorkBook.Worksheets[1].Cells[19, 6].PutValue(ClassSat1.SR19C6);
            ScoreWorkBook.Worksheets[1].Cells[19, 7].PutValue(ClassSat1.SR19C7);
            ScoreWorkBook.Worksheets[1].Cells[19, 8].PutValue(ClassSat1.SR19C8);
            ScoreWorkBook.Worksheets[1].Cells[19, 9].PutValue(ClassSat1.SR19C9);
            ScoreWorkBook.Worksheets[1].Cells[19, 10].PutValue(ClassSat1.SR19C10);
            ScoreWorkBook.Worksheets[1].Cells[19, 11].PutValue(ClassSat1.SR19C11);
            ScoreWorkBook.Worksheets[1].Cells[19, 12].PutValue(ClassSat1.SR19C12);
            ScoreWorkBook.Worksheets[1].Cells[19, 13].PutValue(ClassSat1.SR19C13);
            ScoreWorkBook.Worksheets[1].Cells[19, 14].PutValue(ClassSat1.SR19C14);

            ScoreWorkBook.Worksheets[1].Cells[20, 5].PutValue(ClassSat1.SR20C5);
            ScoreWorkBook.Worksheets[1].Cells[20, 6].PutValue(ClassSat1.SR20C6);
            ScoreWorkBook.Worksheets[1].Cells[20, 7].PutValue(ClassSat1.SR20C7);
            ScoreWorkBook.Worksheets[1].Cells[20, 8].PutValue(ClassSat1.SR20C8);
            ScoreWorkBook.Worksheets[1].Cells[20, 9].PutValue(ClassSat1.SR20C9);
            ScoreWorkBook.Worksheets[1].Cells[20, 10].PutValue(ClassSat1.SR20C10);
            ScoreWorkBook.Worksheets[1].Cells[20, 11].PutValue(ClassSat1.SR20C11);
            ScoreWorkBook.Worksheets[1].Cells[20, 12].PutValue(ClassSat1.SR20C12);
            ScoreWorkBook.Worksheets[1].Cells[20, 13].PutValue(ClassSat1.SR20C13);
            ScoreWorkBook.Worksheets[1].Cells[20, 14].PutValue(ClassSat1.SR20C14);

            //四年級學生數總計
            ScoreWorkBook.Worksheets[1].Cells[21, 5].PutValue(ClassSat1.SR21C5);
            ScoreWorkBook.Worksheets[1].Cells[21, 6].PutValue(ClassSat1.SR21C6);
            ScoreWorkBook.Worksheets[1].Cells[21, 7].PutValue(ClassSat1.SR21C7);
            ScoreWorkBook.Worksheets[1].Cells[21, 8].PutValue(ClassSat1.SR21C8);
            ScoreWorkBook.Worksheets[1].Cells[21, 9].PutValue(ClassSat1.SR21C9);
            ScoreWorkBook.Worksheets[1].Cells[21, 10].PutValue(ClassSat1.SR21C10);
            ScoreWorkBook.Worksheets[1].Cells[21, 11].PutValue(ClassSat1.SR21C11);
            ScoreWorkBook.Worksheets[1].Cells[21, 12].PutValue(ClassSat1.SR21C12);
            ScoreWorkBook.Worksheets[1].Cells[21, 13].PutValue(ClassSat1.SR21C13);
            ScoreWorkBook.Worksheets[1].Cells[21, 14].PutValue(ClassSat1.SR21C14);

            ScoreWorkBook.Worksheets[1].Cells[22, 5].PutValue(ClassSat1.SR22C5);
            ScoreWorkBook.Worksheets[1].Cells[22, 6].PutValue(ClassSat1.SR22C6);
            ScoreWorkBook.Worksheets[1].Cells[22, 7].PutValue(ClassSat1.SR22C7);
            ScoreWorkBook.Worksheets[1].Cells[22, 8].PutValue(ClassSat1.SR22C8);
            ScoreWorkBook.Worksheets[1].Cells[22, 9].PutValue(ClassSat1.SR22C9);
            ScoreWorkBook.Worksheets[1].Cells[22, 10].PutValue(ClassSat1.SR22C10);
            ScoreWorkBook.Worksheets[1].Cells[22, 11].PutValue(ClassSat1.SR22C11);
            ScoreWorkBook.Worksheets[1].Cells[22, 12].PutValue(ClassSat1.SR22C12);
            ScoreWorkBook.Worksheets[1].Cells[22, 13].PutValue(ClassSat1.SR22C13);
            ScoreWorkBook.Worksheets[1].Cells[22, 14].PutValue(ClassSat1.SR22C14);

            //延修生學生數總計
            ScoreWorkBook.Worksheets[1].Cells[23, 5].PutValue(ClassSat1.SR23C5);
            ScoreWorkBook.Worksheets[1].Cells[23, 6].PutValue(ClassSat1.SR23C6);
            ScoreWorkBook.Worksheets[1].Cells[23, 7].PutValue(ClassSat1.SR23C7);
            ScoreWorkBook.Worksheets[1].Cells[23, 8].PutValue(ClassSat1.SR23C8);
            ScoreWorkBook.Worksheets[1].Cells[23, 9].PutValue(ClassSat1.SR23C9);
            ScoreWorkBook.Worksheets[1].Cells[23, 10].PutValue(ClassSat1.SR23C10);
            ScoreWorkBook.Worksheets[1].Cells[23, 11].PutValue(ClassSat1.SR23C11);
            ScoreWorkBook.Worksheets[1].Cells[23, 12].PutValue(ClassSat1.SR23C12);
            ScoreWorkBook.Worksheets[1].Cells[23, 13].PutValue(ClassSat1.SR23C13);
            ScoreWorkBook.Worksheets[1].Cells[23, 14].PutValue(ClassSat1.SR23C14);

            ScoreWorkBook.Worksheets[1].Cells[24, 5].PutValue(ClassSat1.SR24C5);
            ScoreWorkBook.Worksheets[1].Cells[24, 6].PutValue(ClassSat1.SR24C6);
            ScoreWorkBook.Worksheets[1].Cells[24, 7].PutValue(ClassSat1.SR24C7);
            ScoreWorkBook.Worksheets[1].Cells[24, 8].PutValue(ClassSat1.SR24C8);
            ScoreWorkBook.Worksheets[1].Cells[24, 9].PutValue(ClassSat1.SR24C9);
            ScoreWorkBook.Worksheets[1].Cells[24, 10].PutValue(ClassSat1.SR24C10);
            ScoreWorkBook.Worksheets[1].Cells[24, 11].PutValue(ClassSat1.SR24C11);
            ScoreWorkBook.Worksheets[1].Cells[24, 12].PutValue(ClassSat1.SR24C12);
            ScoreWorkBook.Worksheets[1].Cells[24, 13].PutValue(ClassSat1.SR24C13);
            ScoreWorkBook.Worksheets[1].Cells[24, 14].PutValue(ClassSat1.SR24C14);
            #endregion

            #region 綜合高中
            ClassStatistics ClassSat2 = new ClassStatistics();

            ClassSat2.StartStatistics(2, deptNameList);
            ScoreWorkBook.Worksheets[2].Name = "高中職學校班級及學生概況-綜合高中";
            ScoreWorkBook.Worksheets[2].Cells[4, 0].PutValue(SmartSchool.Customization.Data.SystemInformation.SchoolCode);

            //一年級班級數總計
            ScoreWorkBook.Worksheets[2].Cells[7, 5].PutValue(ClassSat2.Level1Count24);
            ScoreWorkBook.Worksheets[2].Cells[7, 7].PutValue(ClassSat2.Level1Count34);
            ScoreWorkBook.Worksheets[2].Cells[7, 9].PutValue(ClassSat2.Level1Count44);
            ScoreWorkBook.Worksheets[2].Cells[7, 11].PutValue(ClassSat2.Level1Count54);
            ScoreWorkBook.Worksheets[2].Cells[7, 13].PutValue(ClassSat2.Level1Count55);

            //二年級班級數總計
            ScoreWorkBook.Worksheets[2].Cells[8, 5].PutValue(ClassSat2.Level2Count24);
            ScoreWorkBook.Worksheets[2].Cells[8, 7].PutValue(ClassSat2.Level2Count34);
            ScoreWorkBook.Worksheets[2].Cells[8, 9].PutValue(ClassSat2.Level2Count44);
            ScoreWorkBook.Worksheets[2].Cells[8, 11].PutValue(ClassSat2.Level2Count54);
            ScoreWorkBook.Worksheets[2].Cells[8, 13].PutValue(ClassSat2.Level2Count55);

            //三年級班級數總計
            ScoreWorkBook.Worksheets[2].Cells[9, 5].PutValue(ClassSat2.Level3Count24);
            ScoreWorkBook.Worksheets[2].Cells[9, 7].PutValue(ClassSat2.Level3Count34);
            ScoreWorkBook.Worksheets[2].Cells[9, 9].PutValue(ClassSat2.Level3Count44);
            ScoreWorkBook.Worksheets[2].Cells[9, 11].PutValue(ClassSat2.Level3Count54);
            ScoreWorkBook.Worksheets[2].Cells[9, 13].PutValue(ClassSat2.Level3Count55);

            //四年級班級數總計
            ScoreWorkBook.Worksheets[2].Cells[10, 5].PutValue(ClassSat2.Level4Count24);
            ScoreWorkBook.Worksheets[2].Cells[10, 7].PutValue(ClassSat2.Level4Count34);
            ScoreWorkBook.Worksheets[2].Cells[10, 9].PutValue(ClassSat2.Level4Count44);
            ScoreWorkBook.Worksheets[2].Cells[10, 11].PutValue(ClassSat2.Level4Count54);
            ScoreWorkBook.Worksheets[2].Cells[10, 13].PutValue(ClassSat2.Level4Count55);

            //一年級學生數總計
            ScoreWorkBook.Worksheets[2].Cells[15, 5].PutValue(ClassSat2.SR15C5);
            ScoreWorkBook.Worksheets[2].Cells[15, 6].PutValue(ClassSat2.SR15C6);
            ScoreWorkBook.Worksheets[2].Cells[15, 7].PutValue(ClassSat2.SR15C7);
            ScoreWorkBook.Worksheets[2].Cells[15, 8].PutValue(ClassSat2.SR15C8);
            ScoreWorkBook.Worksheets[2].Cells[15, 9].PutValue(ClassSat2.SR15C9);
            ScoreWorkBook.Worksheets[2].Cells[15, 10].PutValue(ClassSat2.SR15C10);
            ScoreWorkBook.Worksheets[2].Cells[15, 11].PutValue(ClassSat2.SR15C11);
            ScoreWorkBook.Worksheets[2].Cells[15, 12].PutValue(ClassSat2.SR15C12);
            ScoreWorkBook.Worksheets[2].Cells[15, 13].PutValue(ClassSat2.SR15C13);
            ScoreWorkBook.Worksheets[2].Cells[15, 14].PutValue(ClassSat2.SR15C14);

            ScoreWorkBook.Worksheets[2].Cells[16, 5].PutValue(ClassSat2.SR16C5);
            ScoreWorkBook.Worksheets[2].Cells[16, 6].PutValue(ClassSat2.SR16C6);
            ScoreWorkBook.Worksheets[2].Cells[16, 7].PutValue(ClassSat2.SR16C7);
            ScoreWorkBook.Worksheets[2].Cells[16, 8].PutValue(ClassSat2.SR16C8);
            ScoreWorkBook.Worksheets[2].Cells[16, 9].PutValue(ClassSat2.SR16C9);
            ScoreWorkBook.Worksheets[2].Cells[16, 10].PutValue(ClassSat2.SR16C10);
            ScoreWorkBook.Worksheets[2].Cells[16, 11].PutValue(ClassSat2.SR16C11);
            ScoreWorkBook.Worksheets[2].Cells[16, 12].PutValue(ClassSat2.SR16C12);
            ScoreWorkBook.Worksheets[2].Cells[16, 13].PutValue(ClassSat2.SR16C13);
            ScoreWorkBook.Worksheets[2].Cells[16, 14].PutValue(ClassSat2.SR16C14);

            //二年級學生數總計
            ScoreWorkBook.Worksheets[2].Cells[17, 5].PutValue(ClassSat2.SR17C5);
            ScoreWorkBook.Worksheets[2].Cells[17, 6].PutValue(ClassSat2.SR17C6);
            ScoreWorkBook.Worksheets[2].Cells[17, 7].PutValue(ClassSat2.SR17C7);
            ScoreWorkBook.Worksheets[2].Cells[17, 8].PutValue(ClassSat2.SR17C8);
            ScoreWorkBook.Worksheets[2].Cells[17, 9].PutValue(ClassSat2.SR17C9);
            ScoreWorkBook.Worksheets[2].Cells[17, 10].PutValue(ClassSat2.SR17C10);
            ScoreWorkBook.Worksheets[2].Cells[17, 11].PutValue(ClassSat2.SR17C11);
            ScoreWorkBook.Worksheets[2].Cells[17, 12].PutValue(ClassSat2.SR17C12);
            ScoreWorkBook.Worksheets[2].Cells[17, 13].PutValue(ClassSat2.SR17C13);
            ScoreWorkBook.Worksheets[2].Cells[17, 14].PutValue(ClassSat2.SR17C14);


            ScoreWorkBook.Worksheets[2].Cells[18, 5].PutValue(ClassSat2.SR18C5);
            ScoreWorkBook.Worksheets[2].Cells[18, 6].PutValue(ClassSat2.SR18C6);
            ScoreWorkBook.Worksheets[2].Cells[18, 7].PutValue(ClassSat2.SR18C7);
            ScoreWorkBook.Worksheets[2].Cells[18, 8].PutValue(ClassSat2.SR18C8);
            ScoreWorkBook.Worksheets[2].Cells[18, 9].PutValue(ClassSat2.SR18C9);
            ScoreWorkBook.Worksheets[2].Cells[18, 10].PutValue(ClassSat2.SR18C10);
            ScoreWorkBook.Worksheets[2].Cells[18, 11].PutValue(ClassSat2.SR18C11);
            ScoreWorkBook.Worksheets[2].Cells[18, 12].PutValue(ClassSat2.SR18C12);
            ScoreWorkBook.Worksheets[2].Cells[18, 13].PutValue(ClassSat2.SR18C13);
            ScoreWorkBook.Worksheets[2].Cells[18, 14].PutValue(ClassSat2.SR18C14);

            //三年級學生數總計
            ScoreWorkBook.Worksheets[2].Cells[19, 5].PutValue(ClassSat2.SR19C5);
            ScoreWorkBook.Worksheets[2].Cells[19, 6].PutValue(ClassSat2.SR19C6);
            ScoreWorkBook.Worksheets[2].Cells[19, 7].PutValue(ClassSat2.SR19C7);
            ScoreWorkBook.Worksheets[2].Cells[19, 8].PutValue(ClassSat2.SR19C8);
            ScoreWorkBook.Worksheets[2].Cells[19, 9].PutValue(ClassSat2.SR19C9);
            ScoreWorkBook.Worksheets[2].Cells[19, 10].PutValue(ClassSat2.SR19C10);
            ScoreWorkBook.Worksheets[2].Cells[19, 11].PutValue(ClassSat2.SR19C11);
            ScoreWorkBook.Worksheets[2].Cells[19, 12].PutValue(ClassSat2.SR19C12);
            ScoreWorkBook.Worksheets[2].Cells[19, 13].PutValue(ClassSat2.SR19C13);
            ScoreWorkBook.Worksheets[2].Cells[19, 14].PutValue(ClassSat2.SR19C14);

            ScoreWorkBook.Worksheets[2].Cells[20, 5].PutValue(ClassSat2.SR20C5);
            ScoreWorkBook.Worksheets[2].Cells[20, 6].PutValue(ClassSat2.SR20C6);
            ScoreWorkBook.Worksheets[2].Cells[20, 7].PutValue(ClassSat2.SR20C7);
            ScoreWorkBook.Worksheets[2].Cells[20, 8].PutValue(ClassSat2.SR20C8);
            ScoreWorkBook.Worksheets[2].Cells[20, 9].PutValue(ClassSat2.SR20C9);
            ScoreWorkBook.Worksheets[2].Cells[20, 10].PutValue(ClassSat2.SR20C10);
            ScoreWorkBook.Worksheets[2].Cells[20, 11].PutValue(ClassSat2.SR20C11);
            ScoreWorkBook.Worksheets[2].Cells[20, 12].PutValue(ClassSat2.SR20C12);
            ScoreWorkBook.Worksheets[2].Cells[20, 13].PutValue(ClassSat2.SR20C13);
            ScoreWorkBook.Worksheets[2].Cells[20, 14].PutValue(ClassSat2.SR20C14);

            //四年級學生數總計
            ScoreWorkBook.Worksheets[2].Cells[21, 5].PutValue(ClassSat2.SR21C5);
            ScoreWorkBook.Worksheets[2].Cells[21, 6].PutValue(ClassSat2.SR21C6);
            ScoreWorkBook.Worksheets[2].Cells[21, 7].PutValue(ClassSat2.SR21C7);
            ScoreWorkBook.Worksheets[2].Cells[21, 8].PutValue(ClassSat2.SR21C8);
            ScoreWorkBook.Worksheets[2].Cells[21, 9].PutValue(ClassSat2.SR21C9);
            ScoreWorkBook.Worksheets[2].Cells[21, 10].PutValue(ClassSat2.SR21C10);
            ScoreWorkBook.Worksheets[2].Cells[21, 11].PutValue(ClassSat2.SR21C11);
            ScoreWorkBook.Worksheets[2].Cells[21, 12].PutValue(ClassSat2.SR21C12);
            ScoreWorkBook.Worksheets[2].Cells[21, 13].PutValue(ClassSat2.SR21C13);
            ScoreWorkBook.Worksheets[2].Cells[21, 14].PutValue(ClassSat2.SR21C14);

            ScoreWorkBook.Worksheets[2].Cells[22, 5].PutValue(ClassSat2.SR22C5);
            ScoreWorkBook.Worksheets[2].Cells[22, 6].PutValue(ClassSat2.SR22C6);
            ScoreWorkBook.Worksheets[2].Cells[22, 7].PutValue(ClassSat2.SR22C7);
            ScoreWorkBook.Worksheets[2].Cells[22, 8].PutValue(ClassSat2.SR22C8);
            ScoreWorkBook.Worksheets[2].Cells[22, 9].PutValue(ClassSat2.SR22C9);
            ScoreWorkBook.Worksheets[2].Cells[22, 10].PutValue(ClassSat2.SR22C10);
            ScoreWorkBook.Worksheets[2].Cells[22, 11].PutValue(ClassSat2.SR22C11);
            ScoreWorkBook.Worksheets[2].Cells[22, 12].PutValue(ClassSat2.SR22C12);
            ScoreWorkBook.Worksheets[2].Cells[22, 13].PutValue(ClassSat2.SR22C13);
            ScoreWorkBook.Worksheets[2].Cells[22, 14].PutValue(ClassSat2.SR22C14);

            //延修生學生數總計
            ScoreWorkBook.Worksheets[2].Cells[23, 5].PutValue(ClassSat2.SR23C5);
            ScoreWorkBook.Worksheets[2].Cells[23, 6].PutValue(ClassSat2.SR23C6);
            ScoreWorkBook.Worksheets[2].Cells[23, 7].PutValue(ClassSat2.SR23C7);
            ScoreWorkBook.Worksheets[2].Cells[23, 8].PutValue(ClassSat2.SR23C8);
            ScoreWorkBook.Worksheets[2].Cells[23, 9].PutValue(ClassSat2.SR23C9);
            ScoreWorkBook.Worksheets[2].Cells[23, 10].PutValue(ClassSat2.SR23C10);
            ScoreWorkBook.Worksheets[2].Cells[23, 11].PutValue(ClassSat2.SR23C11);
            ScoreWorkBook.Worksheets[2].Cells[23, 12].PutValue(ClassSat2.SR23C12);
            ScoreWorkBook.Worksheets[2].Cells[23, 13].PutValue(ClassSat2.SR23C13);
            ScoreWorkBook.Worksheets[2].Cells[23, 14].PutValue(ClassSat2.SR23C14);

            ScoreWorkBook.Worksheets[2].Cells[24, 5].PutValue(ClassSat2.SR24C5);
            ScoreWorkBook.Worksheets[2].Cells[24, 6].PutValue(ClassSat2.SR24C6);
            ScoreWorkBook.Worksheets[2].Cells[24, 7].PutValue(ClassSat2.SR24C7);
            ScoreWorkBook.Worksheets[2].Cells[24, 8].PutValue(ClassSat2.SR24C8);
            ScoreWorkBook.Worksheets[2].Cells[24, 9].PutValue(ClassSat2.SR24C9);
            ScoreWorkBook.Worksheets[2].Cells[24, 10].PutValue(ClassSat2.SR24C10);
            ScoreWorkBook.Worksheets[2].Cells[24, 11].PutValue(ClassSat2.SR24C11);
            ScoreWorkBook.Worksheets[2].Cells[24, 12].PutValue(ClassSat2.SR24C12);
            ScoreWorkBook.Worksheets[2].Cells[24, 13].PutValue(ClassSat2.SR24C13);
            ScoreWorkBook.Worksheets[2].Cells[24, 14].PutValue(ClassSat2.SR24C14);
            #endregion


            try
            {

                if (Program.ErrorList.Count > 0)
                {
                    int rowIdx = 1;
                    foreach (string str in Program.ErrorList)
                    {
                        ScoreWorkBook.Worksheets[3].Cells[rowIdx, 0].PutValue(str);
                        rowIdx++;
                    }
                    System.Windows.Forms.MessageBox.Show("產生過程有發生問題，請到工作表Error檢視。");
                }
                if (Program.ErrorList.Count == 0)
                    ScoreWorkBook.Worksheets.RemoveAt(3);
                ScoreWorkBook.Save(filename);
                System.Diagnostics.Process.Start(filename);
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            btnRun.Enabled = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;

            DeptGroupNameDict.Clear();
            // 取得部別設定資料
            List<DeptGroupInfo> deptGroupList =  DataAccess.GetDeptGroupInfo();
            foreach (DeptGroupInfo info in deptGroupList)
            {
                if (!DeptGroupNameDict.ContainsKey(info.DeptGroupName))
                    DeptGroupNameDict.Add(info.DeptGroupName, new List<string>());

                DeptGroupNameDict[info.DeptGroupName].Add(info.DeptName);
            }

            // 填入部別選單
            comboDeptGroup.Items.Clear();
            foreach (string deptGroupName in DeptGroupNameDict.Keys)
            {
                comboDeptGroup.Items.Add(deptGroupName);
            }

            comboDeptGroup.DropDownStyle = ComboBoxStyle.DropDownList;
            
        }        
    }
}